"""Ayudante local para el motivo "Otro" de ExtractorFacturas: abre Outlook
en ESTE PC (no en el servidor) con el destinatario y el PDF ya adjuntos.

Se invoca como facturahelper.exe "facturahelper://abrir?servidor=...&archivo=...&asunto=...",
registrado como protocolo de URL (ver instalar.ps1) para que el navegador
del usuario pueda lanzarlo con un simple enlace.

Build: pip install -r requirements-build.txt && pip install pyinstaller
       pyinstaller --onefile --noconsole --name facturahelper facturahelper.py
"""

import ctypes
import os
import sys
import tempfile
from urllib.parse import urlparse, parse_qs, quote
from urllib.request import urlopen
import json

import win32com.client
import pywintypes


def mostrar_error(mensaje):
    ctypes.windll.user32.MessageBoxW(0, mensaje, "FacturaHelper", 0x10)  # MB_ICONERROR


def leer_json(url):
    with urlopen(url, timeout=15) as resp:
        return json.loads(resp.read().decode("utf-8"))


def descargar_bytes(url):
    with urlopen(url, timeout=30) as resp:
        return resp.read()


def abrir_outlook(email, asunto, ruta_adjunto):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = email
    mail.Subject = asunto
    mail.Attachments.Add(ruta_adjunto)
    mail.Display()


def main():
    if len(sys.argv) < 2:
        mostrar_error("FacturaHelper se invoca desde la web de ExtractorFacturas, no directamente.")
        return

    partes = urlparse(sys.argv[1])
    parametros = parse_qs(partes.query)
    servidor = (parametros.get("servidor") or [""])[0].rstrip("/")
    archivo = (parametros.get("archivo") or [""])[0]
    asunto_solicitado = (parametros.get("asunto") or [""])[0]

    if not servidor or not archivo:
        mostrar_error("Enlace de FacturaHelper incompleto (falta servidor o archivo).")
        return

    try:
        archivo_url = quote(archivo, safe="")
        datos = leer_json(
            f"{servidor}/datos-correo-outlook/{archivo_url}"
            + (f"?asunto={quote(asunto_solicitado)}" if asunto_solicitado else "")
        )
        pdf_bytes = descargar_bytes(f"{servidor}/pdf-factura/{archivo_url}")

        nombre_adjunto = os.path.basename(datos["nombre_adjunto"])
        carpeta_tmp = tempfile.mkdtemp(prefix="facturahelper_")
        ruta_adjunto = os.path.join(carpeta_tmp, nombre_adjunto)
        with open(ruta_adjunto, "wb") as f:
            f.write(pdf_bytes)

        abrir_outlook(datos["email"], datos["asunto"], ruta_adjunto)
    except pywintypes.com_error:
        mostrar_error(
            "No se pudo conectar con Outlook. Comprueba que Microsoft Outlook "
            "de escritorio (no la versión 'New Outlook') está instalado y "
            "configurado en este PC."
        )
    except Exception as e:
        mostrar_error(f"No se pudo abrir el correo: {e}")


if __name__ == "__main__":
    main()
