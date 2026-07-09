import os
import shutil

from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse

from logic import (
    FACTURAS_DIR,
    NO_FACTURA_DIR,
    COMPLETADAS_DIR,
    FACTURAS_REVISADAS_DIR,
    REENVIAR_PEDIDO_DIR,
    REENVIADAS_PEDIDO_DIR,
    REENVIAR_OTRO_MOTIVO_DIR,
    REENVIADAS_OTRO_MOTIVO_DIR,
    MOTIVOS_ENVIO_CORREO,
    CAMPOS_OBLIGATORIOS_FACTURA,
    read_pdf_text,
    es_no_factura,
    extract_invoice_with_agent,
    csv_to_matrix,
    clasificar_factura,
    mover_pdf,
    procesar_carpeta,
    listar_reenviar_pedido,
    solicitar_envio_correo,
    listar_solicitudes_envio_correo,
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

        if tipo == "completada":
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
        "completadas":     COMPLETADAS_DIR,
        "revisadas":       FACTURAS_REVISADAS_DIR,
        "manual":          os.path.join(FACTURAS_DIR, "corregir_manualmente"),
        "no_es_factura":   NO_FACTURA_DIR,
        "reenviar_pedido": REENVIAR_PEDIDO_DIR,
        "reenviadas_pedido": REENVIADAS_PEDIDO_DIR,
        "reenviar_otro_motivo": REENVIAR_OTRO_MOTIVO_DIR,
        "reenviadas_otro_motivo": REENVIADAS_OTRO_MOTIVO_DIR,
    }

    datos = {}
    for nombre, ruta in carpetas.items():
        if not os.path.exists(ruta):
            datos[nombre] = 0
        elif nombre in ("reenviar_otro_motivo", "reenviadas_otro_motivo"):
            # Estas carpetas tienen subcarpetas por campo para las solicitudes
            # de "falta_dato_obligatorio", así que hay que contar recursivamente.
            datos[nombre] = sum(
                len([f for f in archivos if f.lower().endswith(".pdf")])
                for _, _, archivos in os.walk(ruta)
            )
        else:
            datos[nombre] = len([f for f in os.listdir(ruta) if f.lower().endswith(".pdf")])

    return JSONResponse(datos)


# =========================================================
# REENVIAR_PEDIDO (solo lectura): facturas a las que les falta el pedido de
# cliente. La app nunca las mueve a "reenviadas" — ese archivado es cosa de
# Power Automate una vez procesa la solicitud generada en /solicitar-envio-correo.
# =========================================================

@router.get("/reenviar-pedido-lista")
def reenviar_pedido_lista():
    return {"archivos": listar_reenviar_pedido()}


# =========================================================
# ENVIAR CORREO POR OTROS MOTIVOS (revisión manual)
# =========================================================

@router.get("/motivos-envio-correo")
def motivos_envio_correo():
    return {
        "motivos": [{"clave": k, "etiqueta": v} for k, v in MOTIVOS_ENVIO_CORREO.items()],
        "campos_obligatorios": CAMPOS_OBLIGATORIOS_FACTURA,
    }


@router.post("/solicitar-envio-correo")
async def solicitar_envio_correo_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    motivo = str(body.get("motivo", "")).strip()
    motivo_otro = body.get("motivo_otro")
    campo_obligatorio = body.get("campo_obligatorio")

    if not archivo or os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        archivo_generado = solicitar_envio_correo(archivo, motivo, motivo_otro, campo_obligatorio)
        return {"ok": True, "archivo_generado": archivo_generado}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/solicitudes-envio-correo-lista")
def solicitudes_envio_correo_lista():
    return {"archivos": listar_solicitudes_envio_correo()}
