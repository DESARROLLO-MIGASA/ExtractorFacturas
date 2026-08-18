import os
import shutil

from dotenv import get_key
from fastapi import APIRouter, Depends, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse

from auth import requerir_admin, get_current_user
from logic import (
    FACTURAS_DIR,
    ERROR_DIR,
    INCIDENCIAS_DIR,
    ESPERANDO_ALTA_DIR,
    NO_FACTURA_DIR,
    COMPLETADAS_DIR,
    FACTURAS_REVISADAS_DIR,
    REENVIAR_PEDIDO_DIR,
    REENVIADAS_PEDIDO_DIR,
    REENVIAR_DOS_FACTURAS_DIR,
    REENVIAR_OTRO_MOTIVO_DIR,
    REENVIADAS_OTRO_MOTIVO_DIR,
    REENVIAR_ERROR_PESA_MUCHO_DIR,
    REENVIAR_ERROR_OTRO_DIR,
    MOTIVOS_ENVIO_CORREO,
    MOTIVOS_ERROR_EXTRACCION,
    read_pdf_text,
    es_no_factura,
    extract_invoice_with_agent,
    rechazar_si_pesa_demasiado,
    csv_to_matrix,
    clasificar_factura,
    mover_pdf,
    procesar_carpeta,
    listar_reenviar_pedido,
    solicitar_envio_correo,
    listar_solicitudes_envio_correo,
    listar_errores_completo,
    listar_no_factura_completo,
    abrir_carpeta_no_factura,
    clasificar_error,
    reprocesar_error,
    reprocesar_errores,
    LOCK_PROCESAMIENTO_AUTOMATICO,
    guardar_historial,
    guardar_factura_examinada_sql,
    detectar_y_marcar_duplicados,
    contar_reenviadas,
    contar_duplicados_pendientes,
    listar_reenviadas_detalle,
    eliminar_reenvios_de_factura_repetida,
    crear_borrador_outlook_graph,
    datos_correo_outlook,
    asunto_base_correo_otro_motivo,
    solicitar_correo_otro_motivo_pa,
)

CARPETA_ENTRADA = os.path.join(FACTURAS_DIR, "entrada")
ENV_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env")

router = APIRouter()


def procesamiento_pausado():
    """Se relee el .env en cada llamada (no os.getenv) para que activar/desactivar
    la pausa cambiando PAUSAR_API no requiera reiniciar uvicorn."""
    return (get_key(ENV_PATH, "PAUSAR_API") or "").strip().lower() == "true"


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
def upload_pdf(file: UploadFile = File(...)):

    os.makedirs(CARPETA_ENTRADA, exist_ok=True)
    pdf_path = os.path.join(CARPETA_ENTRADA, file.filename)

    with open(pdf_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    if procesamiento_pausado():
        return JSONResponse({"archivo": file.filename, "estado": "pausado"})

    try:
        if rechazar_si_pesa_demasiado(pdf_path, file.filename):
            return JSONResponse({"archivo": file.filename, "estado": "pesa_mucho"})

        text = read_pdf_text(pdf_path)

        if es_no_factura(text):
            mover_pdf(pdf_path, "no_es_factura")
            return JSONResponse({"archivo": file.filename, "estado": "no_es_factura"})

        if len(text) < 80:
            mover_pdf(pdf_path, "imagen")
            return JSONResponse({"archivo": file.filename, "estado": "imagen"})

        result_csv, multiples_facturas = extract_invoice_with_agent(
            file_name=file.filename,
            invoice_text=text,
            pdf_path=pdf_path,
        )

        if multiples_facturas:
            mover_pdf(pdf_path, "incidencia")
            return JSONResponse({"archivo": file.filename, "estado": "incidencia"})

        tabla = csv_to_matrix(result_csv)

        if len(tabla) < 2:
            mover_pdf(pdf_path, "error")
            return JSONResponse({"archivo": file.filename, "estado": "error", "detalle": "respuesta vacía del modelo"})

        fila = tabla[1]
        tipo = clasificar_factura(fila)
        mover_pdf(pdf_path, tipo, fila=fila)
        guardar_historial(result_csv, "auto")

        if tipo == "completada":
            guardar_factura_examinada_sql(fila, "auto")
            detectar_y_marcar_duplicados(fila)

        eliminar_reenvios_de_factura_repetida(fila, file.filename)

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
        "error":           ERROR_DIR,
        "incidencias":     INCIDENCIAS_DIR,
        "esperando_alta":  ESPERANDO_ALTA_DIR,
        "completadas":     COMPLETADAS_DIR,
        "revisadas":       FACTURAS_REVISADAS_DIR,
        "manual":          os.path.join(FACTURAS_DIR, "corregir_manualmente"),
        "no_es_factura":   NO_FACTURA_DIR,
        "reenviar_pedido": REENVIAR_PEDIDO_DIR,
        "reenviadas_pedido": REENVIADAS_PEDIDO_DIR,
        "reenviar_dos_facturas": REENVIAR_DOS_FACTURAS_DIR,
        "reenviar_otro_motivo": REENVIAR_OTRO_MOTIVO_DIR,
        "reenviadas_otro_motivo": REENVIADAS_OTRO_MOTIVO_DIR,
        "reenviar_error_pesa_mucho": REENVIAR_ERROR_PESA_MUCHO_DIR,
        "reenviar_error_otro": REENVIAR_ERROR_OTRO_DIR,
    }

    datos = {}
    for nombre, ruta in carpetas.items():
        if not os.path.exists(ruta):
            datos[nombre] = 0
        else:
            datos[nombre] = len([f for f in os.listdir(ruta) if f.lower().endswith(".pdf")])

    datos["reenviadas"] = contar_reenviadas()
    datos["duplicados"] = contar_duplicados_pendientes()

    return JSONResponse(datos)


# =========================================================
# REENVIADAS: detalle (factura, email, fecha, motivo) de lo ya reenviado
# =========================================================

@router.get("/reenviadas-detalle")
def reenviadas_detalle(usuario: dict = Depends(requerir_admin)):
    return {"facturas": listar_reenviadas_detalle()}


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
    }


@router.post("/solicitar-envio-correo")
def solicitar_envio_correo_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    motivo = str(body.get("motivo", "")).strip()
    motivo_otro = body.get("motivo_otro")
    usuario = str(body.get("usuario", "")).strip() or "desconocido"

    if not archivo or os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        archivo_generado, aviso = solicitar_envio_correo(archivo, motivo, usuario, motivo_otro)
        return {"ok": True, "archivo_generado": archivo_generado, "aviso": aviso}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/solicitudes-envio-correo-lista")
def solicitudes_envio_correo_lista():
    return {"archivos": listar_solicitudes_envio_correo()}


@router.post("/crear-borrador-outlook")
def crear_borrador_outlook_endpoint(body: dict):
    """Crea el borrador (destinatario + PDF adjunto) en el buzón compartido
    vía Microsoft Graph y devuelve el enlace para abrirlo en Outlook Web."""
    archivo = str(body.get("archivo", "")).strip()
    asunto = body.get("asunto")
    usuario = str(body.get("usuario", "")).strip() or "desconocido"

    if not archivo or os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        return crear_borrador_outlook_graph(archivo, usuario, asunto)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except RuntimeError as e:
        raise HTTPException(status_code=500, detail=str(e))


# =========================================================
# ENVIAR CORREO "OTRO MOTIVO" VÍA POWER AUTOMATE (modal propio de EscanerIA)
# =========================================================

@router.get("/datos-correo-otro-motivo/{archivo}")
def datos_correo_otro_motivo_endpoint(archivo: str, usuario: dict = Depends(get_current_user)):
    """Destinatario, nombre de adjunto y asunto para pintar (de solo lectura)
    el modal de correo "otro motivo" antes de enviar. Los mismos datos se
    vuelven a resolver en el servidor al confirmar el envío: esto es solo
    para que la persona vea a quién y con qué PDF se va a enviar."""
    if not archivo or os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        datos = datos_correo_outlook(archivo)
        return {**datos, "asunto": asunto_base_correo_otro_motivo(archivo, datos)}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")


@router.post("/enviar-correo-otro-motivo")
def enviar_correo_otro_motivo_endpoint(body: dict, usuario: dict = Depends(get_current_user)):
    """Envía el correo del motivo "otros" vía Power Automate. El destinatario,
    el PDF y el asunto los vuelve a resolver el backend (nunca se confía en
    nada que no sea `archivo`, `motivo` y `cuerpo` del propio body), y el
    usuario se identifica por la sesión de EscanerIA (X-Forwarded-User), no
    por lo que mande el navegador."""
    archivo = str(body.get("archivo", "")).strip()
    motivo = str(body.get("motivo", "")).strip()
    cuerpo = str(body.get("cuerpo", "")).strip()

    if not archivo or os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        solicitar_correo_otro_motivo_pa(archivo, motivo, cuerpo, usuario["username"])
        return {"ok": True}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except RuntimeError:
        raise HTTPException(
            status_code=502,
            detail="No se pudo enviar el correo. Inténtalo de nuevo o contacta con IT si el problema persiste.",
        )


# =========================================================
# ERRORES DE EXTRACCIÓN (carpeta "error", sin clasificar)
# =========================================================

@router.get("/motivos-error-extraccion")
def motivos_error_extraccion():
    return {"motivos": [{"clave": k, "etiqueta": v} for k, v in MOTIVOS_ERROR_EXTRACCION.items()]}


@router.get("/errores-completo-json")
def errores_completo_json():
    return {"tabla": listar_errores_completo()}


@router.get("/no-factura-completo-json")
def no_factura_completo_json(usuario: dict = Depends(requerir_admin)):
    return {"tabla": listar_no_factura_completo()}


@router.post("/abrir-carpeta-no-factura")
def abrir_carpeta_no_factura_endpoint():
    try:
        abrir_carpeta_no_factura()
        return {"ok": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No se pudo abrir la carpeta: {e}")


@router.post("/clasificar-error")
def clasificar_error_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    motivo = str(body.get("motivo", "")).strip()
    usuario = str(body.get("usuario", "")).strip() or "desconocido"

    if not archivo or os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        email = clasificar_error(archivo, motivo, usuario)
        return {"ok": True, "email": email}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/reprocesar-error")
def reprocesar_error_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()

    if not archivo or os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    if not LOCK_PROCESAMIENTO_AUTOMATICO.acquire(blocking=False):
        raise HTTPException(status_code=409, detail="Hay un procesamiento en curso, inténtalo de nuevo en unos segundos.")

    try:
        exito, estado = reprocesar_error(archivo)
        return {"ok": True, "exito": exito, "estado": estado}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    finally:
        LOCK_PROCESAMIENTO_AUTOMATICO.release()


@router.post("/reprocesar-errores")
def reprocesar_errores_endpoint():
    if not LOCK_PROCESAMIENTO_AUTOMATICO.acquire(blocking=False):
        raise HTTPException(status_code=409, detail="Hay un procesamiento en curso, inténtalo de nuevo en unos segundos.")

    try:
        resultados = reprocesar_errores()
        exitosos = sum(1 for r in resultados if r["exito"])
        return {
            "ok": True,
            "resultados": resultados,
            "exitosos": exitosos,
            "fallidos": len(resultados) - exitosos,
        }
    finally:
        LOCK_PROCESAMIENTO_AUTOMATICO.release()
