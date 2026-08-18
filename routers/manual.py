import os
import csv
import base64
import tempfile
from datetime import datetime
from typing import List

import fitz  # PyMuPDF
from fastapi import APIRouter, Depends, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel

import posiciones
from auth import get_current_user, requerir_admin
from logic import (
    CORREGIR_DIR,
    COLUMNAS_RESUMEN_AUDITORIA,
    EXPECTED_HEADERS,
    EXCEL_DEFINITIVO_PATH,
    read_pdf_text,
    extract_invoice_with_agent,
    combinar_csvs,
    csv_to_matrix,
    guardar_historial,
    cargar_pdf_pendiente_individual,
    confirmar_y_mover_factura,
    guardar_cambios_pendiente,
    descartar_pendiente,
    listar_pendientes_lista,
    listar_pendientes_completo,
    listar_incidencias_completo,
    listar_esperando_alta_completo,
    marcar_esperando_alta,
    reprocesar_esperando_alta,
    listar_facturas_completadas,
    marcar_factura_definitiva,
    marcar_revisada,
    actualizar_factura_completada,
    descartar_factura_completada,
    listar_duplicados_completo,
    resolver_duplicado,
    listar_vista_global_completo,
    facturas_definitivas_tabla,
    buscar_pdf_por_nombre,
    buscar_lineas_csv_por_nombre,
    guardar_lineas_csv,
    ruta_lineas_csv,
    extraer_lineas_de_region,
    tabla_a_pipe_csv,
    export_to_excel,
    buscar_empresas_por_prefijo,
    buscar_proveedores_por_prefijo,
    reservar_facturas_sql,
    liberar_reservas_sql,
    latido_reservas_sql,
    estado_reservas_sql,
    mis_reservas_sql,
    liberar_todas_las_reservas_sql,
    listar_auditoria_sql,
    resumen_auditoria_por_archivo,
)

router = APIRouter()

_ETIQUETAS_AUDITORIA_POR_FILA = {etiqueta for _clave, etiqueta in COLUMNAS_RESUMEN_AUDITORIA}


def _quitar_columnas_auditoria(tabla):
    """Quita del listado las columnas "Editado por"/"Confirmada por"/etc.
    (quién hizo cada acción): son de solo administrador, el resto de la fila
    sigue disponible para cualquier usuario."""
    headers, *filas = tabla
    indices_a_quitar = [i for i, h in enumerate(headers) if h in _ETIQUETAS_AUDITORIA_POR_FILA]
    if not indices_a_quitar:
        return tabla

    def _sin_columnas(fila):
        return [v for i, v in enumerate(fila) if i not in indices_a_quitar]

    return [_sin_columnas(headers)] + [_sin_columnas(fila) for fila in filas]


class ConfirmacionFactura(BaseModel):
    archivo: str
    fila_completa: list
    usuario: str


class ActualizacionFacturaCompletada(BaseModel):
    archivo: str
    fila_completa: list
    origen: str
    usuario: str


class PosicionesFactura(BaseModel):
    archivo: str
    fila_completa: list


class LineaFacturaPayload(BaseModel):
    Descripcion: str = "-"
    Cantidad: str = "-"
    Precio: str = "-"
    Importe: str = "-"
    Otros: str = "-"


class LineasFacturaPayload(BaseModel):
    lineas: List[LineaFacturaPayload]


class OcrRegion(BaseModel):
    archivo: str
    pagina: int
    x0: float
    y0: float
    x1: float
    y1: float


# =========================================================
# PENDIENTES (facturas en corregir_manualmente)
# =========================================================

@router.get("/pendientes-lista")
def pendientes_lista():
    return {"archivos": listar_pendientes_lista()}


@router.get("/pendientes-completo-json")
def pendientes_completo_json(usuario: dict = Depends(get_current_user)):
    tabla = listar_pendientes_completo()
    if usuario["role"] != "admin":
        tabla = _quitar_columnas_auditoria(tabla)
    return {"tabla": tabla}


# =========================================================
# INCIDENCIAS (comprador o proveedor no reconocidos en la base de datos)
# =========================================================

@router.get("/incidencias-completo-json")
def incidencias_completo_json(usuario: dict = Depends(get_current_user)):
    tabla = listar_incidencias_completo()
    if usuario["role"] != "admin":
        tabla = _quitar_columnas_auditoria(tabla)
    return {"tabla": tabla}


@router.post("/marcar-revisada")
def marcar_revisada_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    revisada = bool(body.get("revisada"))
    usuario = str(body.get("usuario", "")).strip() or "desconocido"

    if not archivo:
        raise HTTPException(status_code=400, detail="Falta el nombre del archivo.")

    try:
        confirmada = marcar_revisada(archivo, revisada, usuario)
        return {"ok": True, "accion": "definitiva" if confirmada else "revisada"}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# =========================================================
# ESPERANDO ALTA (proveedor o empresa aún no dados de alta)
# =========================================================

@router.get("/esperando-alta-completo-json")
def esperando_alta_completo_json():
    return {"tabla": listar_esperando_alta_completo()}


@router.post("/marcar-esperando-alta")
def marcar_esperando_alta_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    usuario = str(body.get("usuario", "")).strip() or "desconocido"

    if not archivo:
        raise HTTPException(status_code=400, detail="Falta el nombre del archivo.")

    try:
        marcar_esperando_alta(archivo, usuario)
        return {"ok": True}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/reprocesar-esperando-alta")
def reprocesar_esperando_alta_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()

    if not archivo:
        raise HTTPException(status_code=400, detail="Falta el nombre del archivo.")

    try:
        estado = reprocesar_esperando_alta(archivo)
        return {"ok": True, "movida": estado is not None, "estado": estado}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/extraer-pendiente")
def extraer_pendiente(body: dict):
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
def confirmar_factura(body: ConfirmacionFactura):
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


@router.post("/guardar-cambios-pendiente")
def guardar_cambios_pendiente_endpoint(body: ConfirmacionFactura):
    try:
        guardar_cambios_pendiente(fila_completa=body.fila_completa, usuario=body.usuario)
        return {"ok": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/descartar-pendiente")
def descartar_pendiente_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    usuario = str(body.get("usuario", "")).strip() or "desconocido"
    if not archivo:
        raise HTTPException(status_code=400, detail="Falta el nombre del archivo.")

    if os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        descartar_pendiente(archivo, usuario)
        return {"ok": True}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/pdf-factura/{archivo}")
def ver_pdf_factura(archivo: str):
    if os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    ruta = buscar_pdf_por_nombre(archivo)
    if ruta is None:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")

    return FileResponse(ruta, media_type="application/pdf")


@router.get("/lineas-factura/{archivo}")
def ver_lineas_factura(archivo: str):
    """Líneas de una factura (descripción/cantidad/precio/importe/otros),
    extraídas por el LLM y guardadas como CSV junto al PDF (ver
    guardar_lineas_csv en logic.py). No todas las facturas tienen: si no
    desglosaban líneas, o el CSV se perdió al mover el PDF, se devuelve 404
    en vez de una lista vacía, para que la UI pueda distinguir ambos casos."""
    if os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    ruta_csv = buscar_lineas_csv_por_nombre(archivo)
    if ruta_csv is None:
        raise HTTPException(status_code=404, detail=f"Esta factura no tiene líneas extraídas: {archivo}")

    with open(ruta_csv, "r", newline="", encoding="utf-8-sig") as f:
        lineas = list(csv.DictReader(f, delimiter=";"))

    return {"lineas": lineas}


@router.post("/lineas-factura/{archivo}")
def guardar_lineas_factura(archivo: str, payload: LineasFacturaPayload):
    """Sobrescribe el CSV de líneas de una factura con el contenido final del
    modal de edición (ver cargarLineasEnEdicion / registrarEditLinea /
    anadirLineaEdicion / eliminarLineaEdicion en el frontend): admite
    corregir valores, añadir líneas nuevas y quitar líneas existentes, sea
    cual sea el número de líneas resultante. Una lista vacía borra el CSV en
    vez de dejar un archivo vacío."""
    if os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    ruta_pdf = buscar_pdf_por_nombre(archivo)
    if ruta_pdf is None:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")

    lineas = [linea.dict() for linea in payload.lineas]

    if not lineas:
        ruta_csv = ruta_lineas_csv(ruta_pdf)
        if os.path.exists(ruta_csv):
            os.remove(ruta_csv)
    else:
        guardar_lineas_csv(ruta_pdf, lineas)

    return {"ok": True}


@router.get("/pdf-factura-paginas/{archivo}")
def pdf_factura_paginas(archivo: str):
    """Todas las páginas del PDF como imagen (PNG en base64), para el visor
    propio de la vista de revisión (un iframe con el PDF nativo del
    navegador no permite pintar el recuadro encima). Se renderiza a la
    misma DPI que usa posiciones.py para calcular las cajas, así encajan
    en píxeles sin tener que reescalar nada en el frontend."""
    if os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    ruta = buscar_pdf_por_nombre(archivo)
    if ruta is None:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")

    zoom = posiciones.DPI_CACHE / 72
    paginas = []
    try:
        doc = fitz.open(ruta)
        try:
            for page in doc:
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                paginas.append({
                    "ancho": pix.width,
                    "alto": pix.height,
                    "imagen_base64": base64.b64encode(pix.tobytes("png")).decode("utf-8"),
                })
        finally:
            doc.close()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No se pudo renderizar el PDF: {e}")

    return {"paginas": paginas, "dpi": posiciones.DPI_CACHE}


@router.post("/posiciones-factura")
def posiciones_factura(body: PosicionesFactura):
    """Cajas por campo (página + coordenadas en píxeles a posiciones.DPI_CACHE)
    para pintar el recuadro; se sirven de caché o se calculan al vuelo si es
    la primera vez que se pide esta factura. `fila_completa` son los valores
    que se están mostrando en el formulario (no necesariamente los que
    quedaron guardados), para que el recuadro siga los datos ya corregidos."""
    if os.path.basename(body.archivo) != body.archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    datos_fila = dict(zip(EXPECTED_HEADERS, body.fila_completa))

    try:
        cajas = posiciones.obtener_o_calcular_cajas(body.archivo, datos_fila)
    except Exception:
        cajas = {"paginas_render": {}, "campos": {}}

    return cajas


@router.get("/posiciones-lineas-factura/{archivo}")
def posiciones_lineas_factura(archivo: str):
    """Igual que /posiciones-factura, pero una caja por cada línea de la
    factura (ver /lineas-factura), para resaltarla en el visor al pinchar
    sobre ella. Si la factura no tiene líneas extraídas, devuelve una lista
    vacía en vez de 404 (no es un error, simplemente no hay nada que
    resaltar)."""
    if os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    ruta_csv = buscar_lineas_csv_por_nombre(archivo)
    if ruta_csv is None:
        return {"cajas": []}

    with open(ruta_csv, "r", newline="", encoding="utf-8-sig") as f:
        lineas = list(csv.DictReader(f, delimiter=";"))

    try:
        cajas = posiciones.obtener_o_calcular_cajas_lineas(archivo, lineas)
    except Exception:
        cajas = [None] * len(lineas)

    return {"cajas": cajas}


@router.post("/ocr-region")
def ocr_region(body: OcrRegion):
    """Selección manual con arrastre: recorta esa región de esa página y le
    pasa Tesseract, para prellenar el campo activo con lo que se lea ahí."""
    if os.path.basename(body.archivo) != body.archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        texto = posiciones.recortar_y_ocr_region(
            body.archivo, body.pagina, body.x0, body.y0, body.x1, body.y1
        )
    except posiciones.TesseractNoDisponibleError as e:
        raise HTTPException(status_code=503, detail=str(e))

    return {"texto": texto}


@router.post("/reextraer-lineas-region")
def reextraer_lineas_region(body: OcrRegion):
    """Vuelve a llamar a la IA (vision), pero solo con la región de la tabla
    de líneas que se ha seleccionado a mano en el visor -no con la factura
    entera-, para poder corregir de golpe muchas líneas mal extraídas (o
    ninguna) sin arrastrar campo a campo. No guarda nada por sí solo: el
    resultado sustituye a edicionLineas.lineas en el frontend, y se guarda
    junto con el resto de la factura al pulsar "Guardar"."""
    if os.path.basename(body.archivo) != body.archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    ruta_pdf = buscar_pdf_por_nombre(body.archivo)
    if ruta_pdf is None:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {body.archivo}")

    lineas = extraer_lineas_de_region(ruta_pdf, body.pagina, body.x0, body.y0, body.x1, body.y1)
    return {"lineas": lineas}


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
            result_csv, _ = extract_invoice_with_agent(
                file_name=f.filename,
                invoice_text=text,
                pdf_path=ruta_temp,
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
# REVISAR FACTURAS (facturas ya completadas, en FacturasExaminadas)
# =========================================================

@router.get("/facturas-completadas-json")
def facturas_completadas_json(usuario: dict = Depends(get_current_user)):
    """Pendientes de verificación (Definitiva = No) para todos; las ya
    marcadas como definitivas ("facturas revisadas") solo para admin."""
    tabla = listar_facturas_completadas()
    if usuario["role"] == "admin":
        return {"tabla": tabla}

    headers, *filas = tabla
    idx_definitiva = headers.index("Definitiva")
    filas_pendientes = [fila for fila in filas if fila[idx_definitiva] != "Sí"]
    return {"tabla": [headers] + filas_pendientes}


@router.post("/marcar-factura-definitiva")
def marcar_factura_definitiva_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    definitiva = bool(body.get("definitiva"))
    usuario = str(body.get("usuario", "")).strip() or "desconocido"

    if not archivo:
        raise HTTPException(status_code=400, detail="Falta el nombre del archivo.")

    try:
        marcar_factura_definitiva(archivo, definitiva, usuario)
        return {"ok": True}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró la factura: {archivo}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/actualizar-factura-completada")
def actualizar_factura_completada_endpoint(body: ActualizacionFacturaCompletada):
    try:
        actualizar_factura_completada(
            fila_completa=body.fila_completa,
            origen=body.origen,
            usuario=body.usuario,
        )
        return {"ok": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/descartar-factura-completada")
def descartar_factura_completada_endpoint(body: dict):
    archivo = str(body.get("archivo", "")).strip()
    usuario = str(body.get("usuario", "")).strip() or "desconocido"
    if not archivo or os.path.basename(archivo) != archivo:
        raise HTTPException(status_code=400, detail="Nombre de archivo no válido.")

    try:
        descartar_factura_completada(archivo, usuario)
        return {"ok": True}
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {archivo}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/duplicados-completo-json")
def duplicados_completo_json():
    return {"tabla": listar_duplicados_completo()}


@router.post("/resolver-duplicado")
def resolver_duplicado_endpoint(body: dict):
    archivo_mantener = str(body.get("archivo_mantener", "")).strip()
    archivos_eliminar = [str(a).strip() for a in body.get("archivos_eliminar", []) if str(a).strip()]
    usuario = str(body.get("usuario", "")).strip() or "desconocido"

    if not archivo_mantener:
        raise HTTPException(status_code=400, detail="Falta el archivo a mantener.")
    if not archivos_eliminar:
        raise HTTPException(status_code=400, detail="Falta al menos un archivo a eliminar.")

    try:
        resolver_duplicado(archivo_mantener, archivos_eliminar, usuario)
        return {"ok": True}
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=f"No se encontró el PDF: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


class ExportarExcelBody(BaseModel):
    tabla: List[list]
    nombre: str = "listado"


@router.post("/exportar-excel-listado")
def exportar_excel_listado(body: ExportarExcelBody, usuario: dict = Depends(requerir_admin)):
    if len(body.tabla) < 2:
        raise HTTPException(status_code=400, detail="No hay filas para exportar.")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    export_to_excel(tabla_a_pipe_csv(body.tabla), tmp_path)

    return FileResponse(
        tmp_path,
        filename=f"{body.nombre}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================================================
# VISTA GLOBAL (todas las colas juntas)
# =========================================================

@router.get("/vista-global-json")
def vista_global_json(usuario: dict = Depends(requerir_admin)):
    return {"tabla": listar_vista_global_completo()}


@router.get("/auditoria-json")
def auditoria_json(archivo: str = "", usuario: str = "", desde: str = "", hasta: str = "", limite: int = 500, admin: dict = Depends(requerir_admin)):
    filas = listar_auditoria_sql(
        archivo=archivo.strip() or None,
        usuario=usuario.strip() or None,
        desde=desde.strip() or None,
        hasta=hasta.strip() or None,
        limite=limite,
    )
    cabeceras = ["Fecha", "Usuario", "Accion", "Archivo", "Detalle"]
    return {"tabla": [cabeceras, *filas]}


@router.get("/auditoria-por-archivo-json")
def auditoria_por_archivo_json(archivo: str = "", usuario: str = "", desde: str = "", hasta: str = "", limite: int = 200, admin: dict = Depends(requerir_admin)):
    return {
        "tabla": resumen_auditoria_por_archivo(
            archivo=archivo.strip() or None,
            usuario=usuario.strip() or None,
            desde=desde.strip() or None,
            hasta=hasta.strip() or None,
            limite=limite,
        )
    }


@router.get("/descargar-facturas-definitivas")
def descargar_facturas_definitivas(usuario: dict = Depends(requerir_admin)):
    if not os.path.exists(EXCEL_DEFINITIVO_PATH):
        raise HTTPException(status_code=404, detail="Todavía no hay ninguna factura marcada como definitiva.")

    tabla = facturas_definitivas_tabla()

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    export_to_excel(tabla_a_pipe_csv(tabla), tmp_path)

    return FileResponse(
        tmp_path,
        filename="facturas_100_definitivas.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================================================
# AUTOCOMPLETADO DE CIF (Buyer/Proveedor en el modal de edición)
# =========================================================

@router.get("/empresas-buscar")
def empresas_buscar(q: str = ""):
    return {"resultados": buscar_empresas_por_prefijo(q)}


@router.get("/proveedores-buscar")
def proveedores_buscar(q: str = ""):
    return {"resultados": buscar_proveedores_por_prefijo(q)}


# =========================================================
# RESERVAS (varias personas revisando a la vez, sin pisarse)
# =========================================================

class ReservaBody(BaseModel):
    archivos: List[str]
    usuario: str


@router.post("/reservas")
def reservas_endpoint(body: ReservaBody):
    if not body.usuario.strip():
        raise HTTPException(status_code=400, detail="Falta el usuario.")

    resultado = reservar_facturas_sql(body.archivos, body.usuario.strip().lower())
    if resultado is None:
        raise HTTPException(status_code=500, detail="No se pudo conectar con la base de datos para reservar.")
    return {"reservas": resultado}


@router.post("/reservas-liberar")
def reservas_liberar_endpoint(body: ReservaBody):
    liberar_reservas_sql(body.archivos, body.usuario.strip().lower())
    return {"ok": True}


@router.post("/reservas-latido")
def reservas_latido_endpoint(body: ReservaBody):
    latido_reservas_sql(body.archivos, body.usuario.strip().lower())
    return {"ok": True}


@router.get("/reservas-estado")
def reservas_estado_endpoint(archivos: str = ""):
    lista = [a for a in archivos.split(",") if a]
    return {"reservas": estado_reservas_sql(lista)}


@router.get("/reservas-mias")
def reservas_mias_endpoint(usuario: str = ""):
    return {"archivos": mis_reservas_sql(usuario.strip().lower())}


@router.post("/reservas-liberar-todas")
def reservas_liberar_todas_endpoint(usuario: dict = Depends(requerir_admin)):
    """Vía de escape: vacía TODA la tabla de reservas (de cualquier
    usuario). Pensada para desatascar reservas que se hayan quedado
    "colgadas" en vez de esperar a que caduquen solas. Solo admins: es una
    acción que afecta al trabajo de todo el mundo revisando a la vez."""
    ok = liberar_todas_las_reservas_sql()
    if not ok:
        raise HTTPException(status_code=500, detail="No se pudieron liberar las reservas.")
    return {"ok": True}
