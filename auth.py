"""
Identidad y roles a partir de Windows Authentication (IIS por delante de
uvicorn, ver README). IIS valida al usuario del dominio (Kerberos/NTLM) y
manda quién es en la cabecera X-Forwarded-User (formato DOMINIO\\usuario);
esta app NUNCA autentica contraseñas, solo confía en esa cabecera.

Por eso es crítico que, fuera de IIS, nada pueda llegar directamente a
uvicorn (debe escuchar solo en 127.0.0.1) y que ARR no reenvíe una
X-Forwarded-User que venga ya puesta por el propio cliente: la cabecera la
pone IIS con lo que sacó de Windows Authentication, pisando cualquier valor
que llegara en la petición original (ver la regla de servidor en el README).

El rol se calcula por pertenencia a dos grupos de Active Directory que ya
gestiona IT (AD_GRUPO_ADMINS / AD_GRUPO_USUARIOS en .env) -de alta o baja de
gente se encarga IT añadiendo/quitando del grupo, sin tocar esta app ni su
.env-. Quien no esté en ninguno de los dos grupos se queda fuera (403),
aunque sea un usuario válido del dominio: son grupos creados específicamente
para dar acceso a esta app, no "cualquiera de la empresa".
"""

import os

import pythoncom
import win32com.client
from dotenv import load_dotenv
from fastapi import Depends, Header, HTTPException

_ROOT = os.path.dirname(os.path.abspath(__file__))
load_dotenv(dotenv_path=os.path.join(_ROOT, ".env"))

DOMINIO = os.getenv("AD_DOMINIO", "DOMINT").strip()
GRUPO_ADMINS = os.getenv("AD_GRUPO_ADMINS", "").strip()
GRUPO_USUARIOS = os.getenv("AD_GRUPO_USUARIOS", "").strip()


def _normalizar_usuario(valor_cabecera):
    """"DOMINIO\\usuario" o "usuario@dominio" -> "usuario" en minúsculas.
    Se descarta el dominio porque todos los usuarios de esta app son del
    mismo dominio; comparar solo el nombre de cuenta evita problemas si IIS
    lo entrega alguna vez como NetBIOS y otras como UPN."""
    valor = valor_cabecera.strip()
    if "\\" in valor:
        valor = valor.split("\\", 1)[1]
    elif "@" in valor:
        valor = valor.split("@", 1)[0]
    return valor.strip().lower()


def _dn_desde_nombre_nt4(nombre, dominio):
    """"DOMINIO\\nombre" -> distinguishedName completo (CN=...,OU=...,DC=...),
    vía ADSI NameTranslate. Usa la identidad de Windows del propio proceso,
    sin necesitar ninguna contraseña de cuenta de servicio."""
    nt = win32com.client.Dispatch("NameTranslate")
    nt.Init(3, "")  # 3 = ADS_NAME_INITTYPE_GC (localiza el Global Catalog solo)
    nt.Set(3, dominio + "\\" + nombre)  # 3 = ADS_NAME_TYPE_NT4 ("DOMINIO\nombre")
    return nt.Get(1)  # 1 = ADS_NAME_TYPE_1779 (distinguishedName)


def _es_miembro_de_grupo(usuario, grupo):
    """Pertenencia DIRECTA a un grupo de Active Directory, consultada por
    LDAP (no por el proveedor WinNT: para estos grupos concretos, su
    IADsGroup.Members()/IsMember() da "No implementado" -fallo conocido del
    proveedor WinNT con ciertos grupos de AD-, LDAP sí funciona bien).

    Solo mira miembros directos, no grupos anidados dentro de `grupo`: si
    algún día se anida otro grupo ahí dentro, a esa gente hay que añadirla
    igualmente como miembro directo de `grupo`, o cambiar esta función.

    No propaga excepciones (grupo no encontrado, AD no disponible, etc.):
    se trata como "no es miembro", nunca como "es admin por si acaso".
    """
    if not grupo:
        return False
    try:
        # FastAPI ejecuta esta dependencia (síncrona) en un hilo del thread
        # pool, no en el principal -y COM/ADSI exige inicializarse en cada
        # hilo que lo use, o falla con "No se ha llamado a CoInitialize"-.
        # Llamarlo de más no hace daño: en un hilo ya inicializado es un
        # no-op (S_FALSE), así que no hay que preocuparse de emparejarlo
        # con CoUninitialize en cada request.
        pythoncom.CoInitialize()
        dn_usuario = _dn_desde_nombre_nt4(usuario, DOMINIO)
        dn_grupo = _dn_desde_nombre_nt4(grupo, DOMINIO)
        grupo_obj = win32com.client.GetObject("LDAP://" + dn_grupo)
        return bool(grupo_obj.IsMember("LDAP://" + dn_usuario))
    except Exception as e:
        print(f"AVISO: no se pudo comprobar el grupo de AD '{grupo}' para '{usuario}': {e}")
        return False


def _rol_por_grupo(usuario):
    """"admin" / "usuario" / None (no pertenece a ninguno de los dos grupos
    con acceso a esta app)."""
    if _es_miembro_de_grupo(usuario, GRUPO_ADMINS):
        return "admin"
    if _es_miembro_de_grupo(usuario, GRUPO_USUARIOS):
        return "usuario"
    return None


def get_current_user(x_forwarded_user: str = Header(default=None, alias="X-Forwarded-User")):
    """Dependencia FastAPI: identidad puesta por IIS tras Windows
    Authentication, con el rol calculado por pertenencia a grupo de AD.
    401 si falta la cabecera (la petición no pasó por IIS con Windows
    Authentication activada); 403 si el usuario es válido pero no pertenece
    a ninguno de los dos grupos con acceso a esta app."""
    if not x_forwarded_user or not x_forwarded_user.strip():
        raise HTTPException(
            status_code=401,
            detail="No autenticado: falta la cabecera X-Forwarded-User. "
                   "Comprueba que la petición pasa por IIS con Windows Authentication activada.",
        )

    usuario = _normalizar_usuario(x_forwarded_user)
    if not usuario:
        raise HTTPException(status_code=401, detail="Cabecera X-Forwarded-User vacía o inválida.")

    rol = _rol_por_grupo(usuario)
    if rol is None:
        raise HTTPException(
            status_code=403,
            detail=f"Tu usuario ({usuario}) no pertenece a ningún grupo con acceso a esta app. "
                   f"Pide a IT que te añada a {GRUPO_USUARIOS or '(grupo de usuarios sin configurar)'}.",
        )

    return {
        "username": usuario,
        "usuario_dominio": x_forwarded_user.strip(),
        "role": rol,
    }


def requerir_admin(usuario: dict = Depends(get_current_user)):
    """Igual que get_current_user, pero además exige rol admin (403 si no)."""
    if usuario["role"] != "admin":
        raise HTTPException(status_code=403, detail="Esta acción requiere permisos de administrador.")
    return usuario
