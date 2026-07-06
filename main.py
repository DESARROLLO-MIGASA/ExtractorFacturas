import os
import asyncio
from contextlib import asynccontextmanager
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from logic import FACTURAS_DIR
from imagenes import procesar_carpeta_imagenes
from routers import auto as auto_router
from routers import manual as manual_router

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")

CARPETA_ENTRADA = os.path.join(FACTURAS_DIR, "entrada")
CARPETA_IMAGENES = os.path.join(FACTURAS_DIR, "imagenes")

# Intervalo de vigilancia en segundos (configurable en .env)
INTERVALO = int(os.getenv("INTERVALO_VIGILANCIA", "30"))

_executor = ThreadPoolExecutor(max_workers=1)
_procesando = False


async def vigilar_entrada():
    global _procesando
    while True:
        await asyncio.sleep(INTERVALO)
        if _procesando:
            continue

        pdfs_entrada = []
        if os.path.exists(CARPETA_ENTRADA):
            pdfs_entrada = [f for f in os.listdir(CARPETA_ENTRADA) if f.lower().endswith(".pdf")]

        pdfs_imagenes = []
        if os.path.exists(CARPETA_IMAGENES):
            pdfs_imagenes = [f for f in os.listdir(CARPETA_IMAGENES) if f.lower().endswith(".pdf")]

        if not pdfs_entrada and not pdfs_imagenes:
            continue

        _procesando = True
        try:
            loop = asyncio.get_event_loop()

            if pdfs_entrada:
                print(f"[Vigilante] {len(pdfs_entrada)} PDF(s) en entrada — procesando...")
                await loop.run_in_executor(_executor, auto_router.procesar_carpeta)

            if pdfs_imagenes:
                print(f"[Vigilante] {len(pdfs_imagenes)} PDF(s) en imagenes — procesando con vision...")
                await loop.run_in_executor(_executor, procesar_carpeta_imagenes)
        finally:
            _procesando = False


@asynccontextmanager
async def lifespan(_app: FastAPI):
    tarea = asyncio.create_task(vigilar_entrada())
    print(f"[Vigilante] Activo — comprobando cada {INTERVALO}s")
    yield
    tarea.cancel()


# =========================================================
# APP
# =========================================================

app = FastAPI(title="MIGASA — Extractor de facturas", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

app.include_router(auto_router.router)
app.include_router(manual_router.router)


@app.get("/")
def home():
    return FileResponse(os.path.join(TEMPLATES_DIR, "index.html"))
