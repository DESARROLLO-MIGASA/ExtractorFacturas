import os
import tempfile
from datetime import datetime
from typing import List

from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel

from logic import (
    CORREGIR_DIR,
    PRIORIDAD_PREFIX,
    EXPECTED_HEADERS,
    read_pdf_text,
    extract_invoice_with_agent,
    combinar_csvs,
    csv_to_matrix,
    guardar_historial,
    cargar_pdf_pendiente_individual,
    confirmar_y_mover_factura,
    descartar_pendiente,
    historial_completo,
    tabla_a_pipe_csv,
    export_to_excel,
)

router = APIRouter()


class ConfirmacionFactura(BaseModel):
    archivo: str
    fila_completa: list
    usuario: str


# =========================================================
# PENDIENTES (facturas en corregir_manualmente)
# =========================================================

@router.get("/pendientes-lista")
def pendientes_lista():
    if not os.path.exists(CORREGIR_DIR):
        return {"archivos": []}
    archivos = sorted(
        (f for f in os.listdir(CORREGIR_DIR) if f.lower().endswith(".pdf")),
        key=lambda f: (0 if f.startswith(PRIORIDAD_PREFIX) else 1, f.lower()),
    )
    return {"archivos": archivos}


@router.post("/subir-factura")
async def subir_factura(factura: UploadFile = File(...), urgente: bool = Form(False)):
    if not factura.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Solo se admiten archivos PDF.")

    os.makedirs(CORREGIR_DIR, exist_ok=True)

    nombre = os.path.basename(factura.filename)
    if urgente:
        nombre = PRIORIDAD_PREFIX + nombre

    destino = os.path.join(CORREGIR_DIR, nombre)
    if os.path.exists(destino):
        base, ext = os.path.splitext(nombre)
        nombre = f"{base}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
        destino = os.path.join(CORREGIR_DIR, nombre)

    contenido = await factura.read()
    with open(destino, "wb") as fp:
        fp.write(contenido)

    return {"ok": True, "archivo": nombre, "urgente": urgente}


@router.post("/extraer-pendiente")
async def extraer_pendiente(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    if not archivo:
        raise HTTPException(status_code=400, detail="Falta el nombre del archivo.")

    fila, fuente, es_no_factura_flag = cargar_pdf_pendiente_individual(archivo)

    if es_no_factura_flag:
        return {"tabla": None, "fuente": fuente, "es_no_factura": True}

    if fila is None:
        raise HTTPException(status_code=500, detail="No se pudieron extraer datos del PDF.")

    return {"tabla": [EXPECTED_HEADERS, fila], "fuente": fuente, "es_no_factura": False}


@router.post("/confirmar-factura")
async def confirmar_factura(body: ConfirmacionFactura):
    try:
        confirmar_y_mover_factura(
            archivo=body.archivo,
            fila_completa=body.fila_completa,
            usuario=body.usuario,
        )
        return {"ok": True}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {body.archivo}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/descartar-pendiente")
async def descartar_pendiente_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    if not archivo:
        raise HTTPException(status_code=400, detail="Falta el nombre del archivo.")

    if os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        descartar_pendiente(archivo)
        return {"ok": True}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/pdf-pendiente/{archivo}")
def ver_pdf_pendiente(archivo: str):
    if os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    ruta = os.path.join(CORREGIR_DIR, archivo)
    if not os.path.exists(ruta):
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")

    return FileResponse(ruta, media_type="application/pdf")


# =========================================================
# EXTRACCIÓN DE PDF SUBIDO MANUALMENTE (flujo secundario)
# =========================================================

@router.post("/extraer")
async def extraer(facturas: List[UploadFile] = File(...)):
    uploads_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "uploads")
    os.makedirs(uploads_dir, exist_ok=True)
    results = []

    for f in facturas:
        ruta_temp = os.path.join(uploads_dir, f.filename)
        content = await f.read()

        with open(ruta_temp, "wb") as fp:
            fp.write(content)

        try:
            text = read_pdf_text(ruta_temp)
            result_csv = extract_invoice_with_agent(
                file_name=f.filename,
                invoice_text=text,
            )
            results.append(result_csv)
            guardar_historial(result_csv, "usuario")
        except Exception as e:
            print(f"ERROR {f.filename}: {e}")
        finally:
            try:
                os.remove(ruta_temp)
            except Exception:
                pass

    if not results:
        raise HTTPException(status_code=500, detail="No se pudieron extraer datos.")

    combined = combinar_csvs(results)
    tabla = csv_to_matrix(combined)
    return {"tabla": tabla}


# =========================================================
# HISTORIAL COMPLETO (extracción automática + visión + correcciones manuales)
# =========================================================

@router.get("/historial-json")
def obtener_historial():
    return {"tabla": historial_completo()}


@router.get("/historial")
def descargar_historial():
    tabla = historial_completo()

    if len(tabla) < 2:
        raise HTTPException(status_code=404, detail="No hay historial todavía.")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    export_to_excel(tabla_a_pipe_csv(tabla), tmp_path)

    return FileResponse(
        tmp_path,
        filename="historial_completo.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
