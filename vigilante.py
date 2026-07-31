"""
Vigilante automático de las carpetas "entrada" e "imagenes": revisa cada
INTERVALO_VIGILANCIA segundos si hay PDFs nuevos y los procesa.

Corre como proceso independiente del servidor web (main.py/uvicorn), en su
propia terminal. Antes este bucle vivía dentro del proceso de FastAPI; al
compartir intérprete (y por tanto el GIL) con las peticiones HTTP, procesar
una tanda larga de facturas dejaba la página web sin responder hasta que
terminaba. Al ser un proceso de Windows aparte, uno ya no puede bloquear al
otro.

LOCK_PROCESAMIENTO_AUTOMATICO (logic.py) sigue coordinando este proceso con
los reprocesos manuales disparados desde la web, ahora como candado de
fichero en vez de en memoria, precisamente para que seguir funcionando
entre dos procesos distintos.

Uso:
    python vigilante.py
"""

import os
import sys
import time

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from logic import FACTURAS_DIR, LOCK_PROCESAMIENTO_AUTOMATICO
from imagenes import procesar_carpeta_imagenes
from routers import auto as auto_router

CARPETA_ENTRADA = os.path.join(FACTURAS_DIR, "entrada")
CARPETA_IMAGENES = os.path.join(FACTURAS_DIR, "imagenes")

INTERVALO = int(os.getenv("INTERVALO_VIGILANCIA", "30"))


def vigilar_entrada():
    while True:
        time.sleep(INTERVALO)

        if not LOCK_PROCESAMIENTO_AUTOMATICO.acquire(blocking=False):
            continue

        # Todo el cuerpo del bucle va protegido: las carpetas viven en
        # OneDrive, así que un listado (os.listdir) puede fallar de forma
        # transitoria si en ese instante se está sincronizando. Se registra
        # y se sigue vigilando en la siguiente vuelta en vez de tumbar el
        # proceso entero.
        try:
            try:
                pdfs_entrada = []
                if os.path.exists(CARPETA_ENTRADA):
                    pdfs_entrada = [f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith(".pdf")]

                pdfs_imagenes = []
                if os.path.exists(CARPETA_IMAGENES):
                    pdfs_imagenes = [f for f in os.listdir(CARPETA_IMAGENES) if f.lower().endswith(".pdf")]

                if not pdfs_entrada and not pdfs_imagenes:
                    continue

                if pdfs_entrada:
                    print(f"[Vigilante] {len(pdfs_entrada)} PDF(s) en entrada — procesando...")
                    auto_router.procesar_carpeta()

                if pdfs_imagenes:
                    print(f"[Vigilante] {len(pdfs_imagenes)} PDF(s) en imagenes — procesando con vision...")
                    procesar_carpeta_imagenes()

            except Exception as e:
                print(f"[Vigilante] AVISO: fallo en la vuelta de vigilancia, se reintenta en {INTERVALO}s: {e}")
        finally:
            LOCK_PROCESAMIENTO_AUTOMATICO.release()


if __name__ == "__main__":
    print(f"[Vigilante] Activo — comprobando cada {INTERVALO}s")
    vigilar_entrada()
