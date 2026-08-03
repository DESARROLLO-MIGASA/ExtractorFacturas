# Instala FacturaHelper en este PC: copia el ejecutable y registra el
# protocolo facturahelper:// para que la web de ExtractorFacturas pueda
# abrir Outlook localmente. Ejecutar una vez por PC de usuario (no hace
# falta ser administrador: todo se instala en el perfil del usuario actual).
#
# Uso: colocar facturahelper.exe en la misma carpeta que este script y
# ejecutar:  powershell -ExecutionPolicy Bypass -File instalar.ps1

$ErrorActionPreference = "Stop"

$origen = Join-Path $PSScriptRoot "facturahelper.exe"
if (-not (Test-Path $origen)) {
    Write-Error "No se encuentra facturahelper.exe junto a este script. Compílalo primero (ver README del helper)."
    exit 1
}

$carpetaDestino = Join-Path $env:LOCALAPPDATA "FacturaHelper"
New-Item -ItemType Directory -Force -Path $carpetaDestino | Out-Null

$destino = Join-Path $carpetaDestino "facturahelper.exe"
Copy-Item -Path $origen -Destination $destino -Force

$claveProtocolo = "HKCU:\Software\Classes\facturahelper"
New-Item -Path $claveProtocolo -Force | Out-Null
Set-ItemProperty -Path $claveProtocolo -Name "(default)" -Value "URL:FacturaHelper Protocol"
Set-ItemProperty -Path $claveProtocolo -Name "URL Protocol" -Value ""

New-Item -Path "$claveProtocolo\shell\open\command" -Force | Out-Null
Set-ItemProperty -Path "$claveProtocolo\shell\open\command" -Name "(default)" -Value "`"$destino`" `"%1`""

Write-Host "FacturaHelper instalado en $destino y protocolo facturahelper:// registrado."
