<#
Configura IIS como proxy inverso con Windows Authentication delante de la
app ExtractorFacturas (uvicorn en 127.0.0.1:$BackendPort).

REQUIERE ejecutarse en PowerShell "Como administrador" en YBR-VM-SCANFATC.
Seguro de volver a ejecutar entero las veces que haga falta (cada paso
comprueba el estado actual antes de actuar).

Qué hace, en orden:
  1. Instala el rol IIS + módulo Windows Authentication.
  2. Instala ARR y URL Rewrite (MSIs ya descargados y verificados en esta
     misma carpeta, firmados por Microsoft Corporation).
  3. Se asegura de que exista "Default Web Site" con un Application Pool
     válido y arrancado (si el servidor se reinstaló desde cero, esto no
     viene creado solo).
  4. Habilita el proxy de ARR a nivel de servidor.
  5. Desbloquea las secciones de configuración que IIS trae bloqueadas por
     defecto (autenticación, rewrite).
  6. Escribe el web.config con la regla de proxy inverso hacia
     127.0.0.1:$BackendPort, y DESPUÉS configura Windows Authentication
     (activa Windows Auth + sus proveedores Negotiate/NTLM, desactiva
     Anonymous) -en ese orden: escribir el web.config es un volcado del
     archivo entero, así que si fuera al revés borraría los ajustes de
     autenticación que IIS acaba de guardar ahí mismo-.
  7. iisreset.
  8. Verificación final: relee todo lo anterior y lo imprime junto, para
     confirmar de un vistazo que quedó bien sin tener que ir comprobando
     cosa por cosa a mano.

Cambia $BackendPort más abajo si tu app corre en otro puerto.
#>

$ErrorActionPreference = "Stop"
$SitePath = "IIS:\Sites\Default Web Site"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BackendPort = 8000

function Comprobar-MsiExitCode($proceso, $nombre) {
    # 0 = éxito. 3010 = éxito, pero pide reinicio (no bloqueante para seguir
    # configurando IIS en la misma sesión, msiexec ya deja los archivos listos).
    if ($proceso.ExitCode -ne 0 -and $proceso.ExitCode -ne 3010) {
        throw "La instalación de $nombre terminó con código $($proceso.ExitCode) (no es 0 ni 3010). Revisa manualmente antes de continuar."
    }
    if ($proceso.ExitCode -eq 3010) {
        Write-Host "   ($nombre pide reinicio del servidor; se puede seguir configurando ahora, pero conviene reiniciar cuando sea posible)" -ForegroundColor Yellow
    }
}

Write-Host "== 1. Instalando rol IIS + Windows Authentication ==" -ForegroundColor Cyan
$resultadoFeature = Install-WindowsFeature -Name Web-Server, Web-Windows-Auth, Web-Basic-Auth, Web-Mgmt-Console -IncludeManagementTools
if (-not $resultadoFeature.Success) {
    throw "Install-WindowsFeature no terminó con éxito (ExitCode: $($resultadoFeature.ExitCode)). No se sigue con el resto del script."
}
if ($resultadoFeature.RestartNeeded -eq "Yes") {
    Write-Host "   AVISO: Windows pide reiniciar el servidor. Termina este script y reinicia antes de probar nada." -ForegroundColor Yellow
}

Write-Host "== 2. Instalando URL Rewrite ==" -ForegroundColor Cyan
$p = Start-Process msiexec.exe -ArgumentList "/i `"$ScriptDir\rewrite_amd64_en-US.msi`" /quiet /norestart" -Wait -PassThru
Comprobar-MsiExitCode $p "URL Rewrite"

Write-Host "== 2b. Instalando Application Request Routing (ARR) ==" -ForegroundColor Cyan
$p = Start-Process msiexec.exe -ArgumentList "/i `"$ScriptDir\requestRouter_amd64.msi`" /quiet /norestart" -Wait -PassThru
Comprobar-MsiExitCode $p "Application Request Routing"

Import-Module WebAdministration

Write-Host "== 3. Comprobando 'Default Web Site' y su Application Pool ==" -ForegroundColor Cyan
if (-not (Test-Path "C:\inetpub\wwwroot")) {
    New-Item -ItemType Directory -Path "C:\inetpub\wwwroot" -Force | Out-Null
}
if (-not (Test-Path "IIS:\AppPools\DefaultAppPool")) {
    Write-Host "   Creando DefaultAppPool (no existía)" -ForegroundColor Yellow
    New-WebAppPool -Name "DefaultAppPool" | Out-Null
}
if ((Get-WebAppPoolState -Name "DefaultAppPool").Value -ne "Started") {
    Start-WebAppPool -Name "DefaultAppPool"
}
if (-not (Test-Path $SitePath)) {
    Write-Host "   Creando 'Default Web Site' (no existía)" -ForegroundColor Yellow
    New-Website -Name "Default Web Site" -Port 80 -PhysicalPath "C:\inetpub\wwwroot" -ApplicationPool "DefaultAppPool" -Force | Out-Null
}
if ((Get-Item $SitePath).applicationPool -ne "DefaultAppPool") {
    Set-ItemProperty $SitePath -Name applicationPool -Value "DefaultAppPool"
}
if ((Get-Website -Name "Default Web Site").State -ne "Started") {
    Start-Website -Name "Default Web Site"
}

Write-Host "== 4. Habilitando el proxy de ARR a nivel de servidor ==" -ForegroundColor Cyan
Set-WebConfigurationProperty -pspath 'MACHINE/WEBROOT/APPHOST' -filter "system.webServer/proxy" -name "enabled" -value "True"

Write-Host "== 5. Desbloqueando secciones de configuración (vienen bloqueadas por defecto en IIS) ==" -ForegroundColor Cyan
# anonymousAuthentication/windowsAuthentication/rewrite traen overrideModeDefault="Deny" de fábrica
# en applicationHost.config: sin este desbloqueo, Set-WebConfigurationProperty/Add-WebConfiguration
# sobre el sitio fallan con "Esta sección de configuración no puede utilizarse en esta ruta" aunque
# el módulo correspondiente esté bien instalado.
$appcmd = Join-Path $env:windir "system32\inetsrv\appcmd.exe"
$seccionesADesbloquear = @(
    "system.webServer/security/authentication/anonymousAuthentication",
    "system.webServer/security/authentication/windowsAuthentication",
    "system.webServer/rewrite/allowedServerVariables",
    "system.webServer/rewrite/rules"
)
foreach ($seccion in $seccionesADesbloquear) {
    $salida = & $appcmd unlock config "-section:$seccion" 2>&1
    Write-Host "   $seccion -> $salida"
}

Write-Host "== 6. Escribiendo web.config con la regla de proxy inverso (puerto $BackendPort) ==" -ForegroundColor Cyan
# IMPORTANTE: este Set-Content vuelca el archivo ENTERO -no fusiona con lo que ya
# hubiera-. Por eso va ANTES de tocar la autenticación: si fuera al revés, esta
# escritura borraría los valores de anonymousAuthentication/windowsAuthentication
# que IIS acaba de guardar en este mismo web.config (nos pasó de verdad: el script
# no daba ningún error, pero el sitio se quedaba con Anonymous activa igualmente).
$sitePhysicalPath = [System.Environment]::ExpandEnvironmentVariables((Get-Item $SitePath).PhysicalPath)
if (-not (Test-Path $sitePhysicalPath)) {
    throw "La ruta física del sitio '$sitePhysicalPath' no existe. Revisa manualmente antes de continuar."
}
$webConfigPath = Join-Path $sitePhysicalPath "web.config"

$webConfigXml = @"
<configuration>
  <system.webServer>
    <rewrite>
      <allowedServerVariables>
        <add name="HTTP_X_FORWARDED_USER" />
      </allowedServerVariables>
      <rules>
        <rule name="ReverseProxyToUvicorn" stopProcessing="true">
          <match url="(.*)" />
          <serverVariables>
            <set name="HTTP_X_FORWARDED_USER" value="{LOGON_USER}" />
          </serverVariables>
          <action type="Rewrite" url="http://127.0.0.1:$BackendPort/{R:1}" />
        </rule>
      </rules>
    </rewrite>
  </system.webServer>
</configuration>
"@

if (Test-Path $webConfigPath) {
    $backup = "$webConfigPath.bak_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Copy-Item $webConfigPath $backup
    Write-Host "   (web.config existente respaldado en $backup)" -ForegroundColor Yellow
}
Set-Content -Path $webConfigPath -Value $webConfigXml -Encoding UTF8

Write-Host "== 7. Configurando Windows Authentication en $SitePath ==" -ForegroundColor Cyan
# A partir de aqui usamos los cmdlets de IIS (no Set-Content): estos leen el
# web.config que acabamos de escribir y le AÑADEN los nodos de autenticación,
# sin borrar la regla de rewrite que ya está en el archivo.
Set-WebConfigurationProperty -Filter /system.webServer/security/authentication/anonymousAuthentication `
    -Name enabled -Value false -PSPath $SitePath
Set-WebConfigurationProperty -Filter /system.webServer/security/authentication/windowsAuthentication `
    -Name enabled -Value true -PSPath $SitePath

# Set-WebConfiguration -Value @("Negotiate","NTLM") NO escribe bien esta colección en
# este servidor (se probó y se quedaba vacía sin dar ningún error) -> Clear + Add uno
# a uno, que sí es fiable.
Clear-WebConfiguration -Filter "/system.webServer/security/authentication/windowsAuthentication/providers" -PSPath $SitePath
Add-WebConfiguration -Filter "/system.webServer/security/authentication/windowsAuthentication/providers" -PSPath $SitePath -Value @{value="Negotiate"}
Add-WebConfiguration -Filter "/system.webServer/security/authentication/windowsAuthentication/providers" -PSPath $SitePath -Value @{value="NTLM"}

Write-Host "== 8. iisreset ==" -ForegroundColor Cyan
iisreset

Write-Host ""
Write-Host "== 9. Verificación final (todo junto) ==" -ForegroundColor Cyan
Write-Host "Sitio:" -NoNewline; Write-Host " $((Get-Website -Name 'Default Web Site').State), pool $((Get-Item $SitePath).applicationPool)"
Write-Host "Anonymous Authentication (debe ser False):" -NoNewline
Write-Host " $((Get-WebConfigurationProperty -Filter /system.webServer/security/authentication/anonymousAuthentication -Name enabled -PSPath $SitePath).Value)"
Write-Host "Windows Authentication (debe ser True):" -NoNewline
Write-Host " $((Get-WebConfigurationProperty -Filter /system.webServer/security/authentication/windowsAuthentication -Name enabled -PSPath $SitePath).Value)"
Write-Host "Providers (debe haber Negotiate y NTLM):"
Get-WebConfiguration /system.webServer/security/authentication/windowsAuthentication/providers -PSPath $SitePath | Select-Object -ExpandProperty Collection | ForEach-Object { Write-Host "   - $($_.value)" }
Write-Host "web.config:"
Get-Content $webConfigPath | Write-Host

Write-Host ""
try {
    $resp = Invoke-WebRequest -Uri "http://localhost/whoami" -UseDefaultCredentials -UseBasicParsing -TimeoutSec 5
    Write-Host "/whoami respondió $($resp.StatusCode): $($resp.Content)" -ForegroundColor Green
} catch {
    Write-Host "/whoami dio error: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "(normal si uvicorn no está arrancado en 127.0.0.1:$BackendPort ahora mismo)" -ForegroundColor Yellow
}
