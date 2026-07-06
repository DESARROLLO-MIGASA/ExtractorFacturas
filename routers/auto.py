import os
import shutil

from fastapi import APIRouter, UploadFile, File
from fastapi.responses import JSONResponse

from logic import (
    FACTURAS_DIR,
    NO_FACTURA_DIR,
    read_pdf_text,
    es_no_factura,
    extract_invoice_with_agent,
    csv_to_matrix,
    clasificar_factura,
    mover_pdf,
    procesar_carpeta,
    guardar_historial,
    guardar_factura_examinada_sql,
)

CARPETA_ENTRADA = os.path.join(FACTURAS_DIR, "entrada")

router = APIRouter()


# =========================================================
# PROCESAR CARPETA (manual desde la UI)
# =========================================================

@router.post("/procesar")
def procesar():
    resultados = procesar_carpeta()
    return JSONResponse({"resultados": resultados})


# =========================================================
# UPLOAD PDF (llamado desde Power Automate)
# =========================================================

@router.post("/upload-pdf")
async def upload_pdf(file: UploadFile = File(...)):

    os.makedirs(CARPETA_ENTRADA, exist_ok=True)
    pdf_path = os.path.join(CARPETA_ENTRADA, file.filename)

    with open(pdf_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        text = read_pdf_text(pdf_path)

        if es_no_factura(text):
            mover_pdf(pdf_path, "no_es_factura")
            return JSONResponse({"archivo": file.filename, "estado": "no_es_factura"})

        if len(text) < 80:
            mover_pdf(pdf_path, "imagen")
            return JSONResponse({"archivo": file.filename, "estado": "imagen"})

        result_csv = extract_invoice_with_agent(
            file_name=file.filename,
            invoice_text=text,
        )

        tabla = csv_to_matrix(result_csv)

        if len(tabla) < 2:
            mover_pdf(pdf_path, "error")
            return JSONResponse({"archivo": file.filename, "estado": "error", "detalle": "respuesta vacía del modelo"})

        fila = tabla[1]
        tipo = clasificar_factura(fila)
        mover_pdf(pdf_path, tipo)
        guardar_historial(result_csv, "auto")

        if tipo == "examinada":
            guardar_factura_examinada_sql(fila, "auto")

        return JSONResponse({"archivo": file.filename, "estado": tipo})

    except Exception as e:
        try:
            mover_pdf(pdf_path, "error")
        except Exception:
            pass
        return JSONResponse(
            {"archivo": file.filename, "estado": "error", "detalle": str(e)},
            status_code=500,
        )


# =========================================================
# ESTADISTICAS
# =========================================================

@router.get("/estadisticas")
def estadisticas():

    carpetas = {
        "entrada":         os.path.join(FACTURAS_DIR, "entrada"),
        "procesadas":      os.path.join(FACTURAS_DIR, "procesadas"),
        "imagenes":        os.path.join(FACTURAS_DIR, "imagenes"),
        "error":           os.path.join(FACTURAS_DIR, "error"),
        "examinadas":      os.path.join(FACTURAS_DIR, "examinadas"),
        "manual":          os.path.join(FACTURAS_DIR, "corregir_manualmente"),
        "no_es_factura":   NO_FACTURA_DIR,
        "reenviar_pedido": os.path.join(FACTURAS_DIR, "reenviar_falta_pedidocliente"),
    }

    datos = {}
    for nombre, ruta in carpetas.items():
        if not os.path.exists(ruta):
            datos[nombre] = 0
        else:
            datos[nombre] = len([f for f in os.listdir(ruta) if f.lower().endswith(".pdf")])

    return JSONResponse(datos)
