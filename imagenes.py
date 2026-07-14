"""
Segunda pasada para PDFs sin texto extraible (facturas escaneadas o
fotografiadas), que app_auto y app_manual dejan en la carpeta "imagenes".

Este script concentra las llamadas a la API con imagenes (vision). El
resto del pipeline (app_auto, app_manual) sigue trabajando solo con el
texto extraido por pdfplumber, salvo por un único respaldo puntual: si ni
el texto ni el regex de detectar_cifs_con_etiqueta encuentran el CIF del
comprador o del proveedor, logic.py importa detectar_cif_con_vision_desde_pdf
de este módulo para intentarlo mirando la imagen real de la factura (útil
cuando el CIF está en un logo o pegado sin etiqueta a una URL/teléfono).
Es la única excepción: se activa solo cuando de verdad falta ese dato, no
en cada factura con texto.

Uso:
    python imagenes.py            -> procesa la carpeta "imagenes" una vez
    python imagenes.py --watch    -> repite el proceso cada INTERVALO_VIGILANCIA_IMAGENES segundos
"""

import os
import sys
import csv
import json
import base64
import shutil
import time
from io import StringIO
from datetime import datetime

import fitz  # PyMuPDF

# Reutiliza toda la lógica ya existente (prompt, cliente OpenAI, CSV,
# clasificación, historial, guardado en SQL...) en vez de duplicarla.
import logic as auto_logic

FACTURAS_DIR = auto_logic.FACTURAS_DIR
CARPETA_IMAGENES = os.path.join(FACTURAS_DIR, "imagenes")
CARPETA_SIN_DATOS = os.path.join(FACTURAS_DIR, "imagenes_sin_datos")

MAX_PAGINAS = int(os.getenv("MAX_PAGINAS_IMAGENES", "3"))
DPI_IMAGENES = int(os.getenv("DPI_IMAGENES", "200"))


# =========================================================
# PDF -> IMAGENES
# =========================================================

def pdf_a_imagenes_base64(pdf_path, max_paginas=MAX_PAGINAS, dpi=DPI_IMAGENES, paginas=None):
    """
    Renderiza páginas del PDF como PNG y las devuelve codificadas en
    base64, listas para enviar a la API de vision. Por defecto son las
    primeras `max_paginas`; si se pasa `paginas` (índices 0-based), se
    renderizan esas en concreto (p.ej. para incluir también la última
    página, donde suele estar el pie con los datos del proveedor).
    """
    imagenes = []
    zoom = dpi / 72
    matriz = fitz.Matrix(zoom, zoom)

    doc = fitz.open(pdf_path)

    try:
        indices = paginas if paginas is not None else range(min(max_paginas, doc.page_count))
        for i in indices:
            if i < 0 or i >= doc.page_count:
                continue
            pix = doc[i].get_pixmap(matrix=matriz)
            imagenes.append(base64.b64encode(pix.tobytes("png")).decode("utf-8"))
    finally:
        doc.close()

    return imagenes


def _paginas_cabecera_y_pie(pdf_path, max_total=3):
    """
    Índices de página (0-based) útiles para buscar el CIF: las primeras dos
    (donde suele estar la cabecera con comprador/proveedor) y la última
    (donde muchas plantillas meten el pie con los datos registrales del
    proveedor), sin repetir páginas si el documento es corto.
    """
    try:
        doc = fitz.open(pdf_path)
        total = doc.page_count
        doc.close()
    except Exception:
        return [0, 1]

    if total <= max_total:
        return list(range(total))
    return [0, 1, total - 1]


# =========================================================
# RESPALDO POR VISION: SOLO EL CIF (comprador/proveedor)
# =========================================================
#
# A veces el CIF que el texto (o su respaldo por regex) no encuentra está
# metido en un logo, sello, o pegado sin ninguna etiqueta a una URL/teléfono
# en el pie de página; un modelo mirando la imagen real de la factura suele
# reconocerlo mejor que leyendo el texto plano ya aplanado por pdfplumber.
# Se pide SOLO el CIF (no los 18 campos) porque un prompt más simple y
# centrado tiene mejores resultados que repetir la extracción completa.

def detectar_cif_con_vision(imagenes_b64):
    """
    Devuelve (cif_buyer, cif_proveedor) mirando las imágenes ya renderizadas
    de la factura, o ("-", "-") si no encuentra alguno o falla la llamada.
    No decide qué CIF corresponde a qué campo: eso lo hace igual que
    siempre resolver_empresa_y_proveedor(), cruzando contra las tablas
    maestras.
    """
    if not imagenes_b64:
        return "-", "-"

    client = auto_logic.build_client()
    model = auto_logic.get_model()

    contenido = [
        {
            "type": "text",
            "text": (
                "Esta es la imagen de una factura de proveedor. Devuelve "
                'SOLO un objeto JSON con dos campos: "Buyer" (el CIF/NIF/VAT '
                'del comprador/cliente) y "Proveedor" (el CIF/NIF/VAT del '
                "proveedor/vendedor que emite la factura).\n"
                "Alguno de los dos puede estar metido dentro de un logo, "
                "sello, marca de agua, o en letra muy pequeña sin ninguna "
                "etiqueta explícita (por ejemplo pegado a una URL o a un "
                "teléfono en el pie de página): míralo con atención antes "
                'de rendirte. Si de verdad no aparece, devuelve "-" en ese '
                "campo. No añadas explicaciones ni ningún otro campo."
            ),
        }
    ]

    for img_b64 in imagenes_b64:
        contenido.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{img_b64}"},
        })

    schema = {
        "type": "object",
        "properties": {"Buyer": {"type": "string"}, "Proveedor": {"type": "string"}},
        "required": ["Buyer", "Proveedor"],
        "additionalProperties": False,
    }

    try:
        response = client.chat.completions.create(
            model=model,
            temperature=0,
            messages=[
                {"role": "system", "content": "Devuelves solo el JSON solicitado, sin texto adicional."},
                {"role": "user", "content": contenido},
            ],
            response_format={
                "type": "json_schema",
                "json_schema": {"name": "cif_factura", "schema": schema, "strict": True},
            },
        )
        datos = json.loads(response.choices[0].message.content or "{}")
    except Exception as e:
        print(f"AVISO: fallo en el respaldo por visión del CIF: {e}")
        return "-", "-"

    # extraer_cif_de_texto_libre (no normalizar_cif a secas) porque la
    # visión a veces devuelve más texto del pedido (p.ej. "OLEO VERDE S.L.
    # NIF B91580142" en vez de solo el CIF).
    return (
        auto_logic.extraer_cif_de_texto_libre(datos.get("Buyer", "-")),
        auto_logic.extraer_cif_de_texto_libre(datos.get("Proveedor", "-")),
    )


def detectar_cif_con_vision_desde_pdf(pdf_path):
    """Como detectar_cif_con_vision, pero renderizando el propio PDF
    (cabecera + última página) en vez de partir de imágenes ya generadas."""
    try:
        paginas = _paginas_cabecera_y_pie(pdf_path)
        imagenes_b64 = pdf_a_imagenes_base64(pdf_path, paginas=paginas)
    except Exception as e:
        print(f"AVISO: no se pudo renderizar {pdf_path} para el respaldo por visión: {e}")
        return "-", "-"

    return detectar_cif_con_vision(imagenes_b64)


# =========================================================
# LLAMADA AL MODELO (VISION)
# =========================================================

def extract_invoice_from_images(file_name, imagenes_b64, agent_prompt=None):
    agent_prompt = agent_prompt or auto_logic.DEFAULT_PROMPT

    client = auto_logic.build_client()
    model = auto_logic.get_model()

    contenido = [
        {
            "type": "text",
            "text": (
                f"{agent_prompt}\n\n"
                f"NOMBRE DEL ARCHIVO:\n\"\"\"\n{file_name}\n\"\"\"\n\n"
                "La factura no tiene texto extraible: se adjunta como "
                "imagen (una o varias páginas escaneadas). Lee los datos "
                "directamente de la imagen."
            ),
        }
    ]

    for img_b64 in imagenes_b64:
        contenido.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{img_b64}"},
        })

    response = client.chat.completions.create(
        model=model,
        temperature=0,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0,
        messages=[
            {
                "role": "system",
                "content": (
                    "Eres un extractor documental muy estricto especializado "
                    "en leer facturas escaneadas o fotografiadas. Devuelve los "
                    "datos en el objeto JSON solicitado, sin explicaciones ni "
                    "texto adicional."
                ),
            },
            {"role": "user", "content": contenido},
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "factura",
                "schema": auto_logic.FACTURA_JSON_SCHEMA,
                "strict": True,
            },
        },
    )

    raw = response.choices[0].message.content or "{}"
    headers = auto_logic.EXPECTED_HEADERS

    try:
        datos_json = json.loads(raw)
    except ValueError:
        datos_json = {}

    # Igual que en logic.py.extract_invoice_with_agent: el JSON Schema en
    # modo "strict" garantiza los 18 campos exactos, así que no hace falta
    # validar ni corregir desplazamientos de columnas.
    data = [auto_logic.normalizar_valor(datos_json.get(h, "-")) for h in headers]

    datos = dict(zip(headers, data))
    datos["Archivo"] = file_name
    datos["FEscaneo"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    datos["FFactura"] = auto_logic.normalizar_fecha(datos["FFactura"])
    datos["FOperacion"] = auto_logic.normalizar_fecha(datos["FOperacion"])

    if datos["FFactura"] != "-":
        datos["FOperacion"] = datos["FFactura"]

    # Buyer/Proveedor y Empresa/NombreProveedor: igual que en logic.py, cuál
    # de los dos CIF extraídos es el comprador y cuál el proveedor lo decide
    # resolver_empresa_y_proveedor() cruzando contra las tablas clasificadas,
    # no la posición/contexto que haya usado el modelo al etiquetarlos, para
    # que una factura escaneada se trate igual que una con texto
    # (incidencias incluidas si el CIF no se reconoce).
    cif_1 = auto_logic.normalizar_cif(datos["Buyer"])
    cif_2 = auto_logic.normalizar_cif(datos["Proveedor"])

    # Un solo intento de vision (pidiendo los 18 campos a la vez) a veces
    # pasa por alto el CIF aunque sí lo tenga delante: un segundo intento
    # con un prompt más simple, centrado solo en el CIF, tiene mejores
    # resultados. Solo se repite si de verdad falta alguno de los dos.
    if cif_1 == "-" or cif_2 == "-":
        cif_1_vision, cif_2_vision = detectar_cif_con_vision(imagenes_b64)
        if cif_1 == "-":
            cif_1 = cif_1_vision
        if cif_2 == "-":
            cif_2 = cif_2_vision

    datos["Buyer"], datos["Empresa"], datos["Proveedor"], datos["NombreProveedor"] = (
        auto_logic.resolver_empresa_y_proveedor(cif_1, cif_2)
    )

    output = StringIO()
    writer = csv.writer(output, delimiter="|", lineterminator="\n")
    writer.writerow(headers)
    writer.writerow([datos[h] for h in headers])

    return output.getvalue().strip()


# =========================================================
# MOVER PDF TRAS EL RESULTADO
# =========================================================

def mover_pdf_imagen(pdf_path, tipo):
    """
    Mueve el PDF (ya copiado a "procesadas" en la primera pasada) desde
    "imagenes" a la carpeta que corresponda según el resultado de vision.

    Si sigue sin poder leerse ("imagen"), va a "imagenes_sin_datos" en
    vez de volver a "imagenes", para no reprocesarlo en cada pasada.
    """
    carpetas = {
        "completada":      auto_logic.COMPLETADAS_DIR,
        "manual":          os.path.join(FACTURAS_DIR, "corregir_manualmente"),
        "reenviar_pedido": os.path.join(FACTURAS_DIR, "reenviar_falta_pedidocliente"),
        "incidencia":      auto_logic.INCIDENCIAS_DIR,
        "imagen":          CARPETA_SIN_DATOS,
        "error":           os.path.join(FACTURAS_DIR, "error"),
    }

    carpeta_destino = carpetas.get(tipo, os.path.join(FACTURAS_DIR, "error"))
    os.makedirs(carpeta_destino, exist_ok=True)

    destino = os.path.join(carpeta_destino, os.path.basename(pdf_path))
    shutil.move(pdf_path, destino)

    print(f"PDF MOVIDO A {tipo.upper()}:", destino)


# =========================================================
# PROCESAR CARPETA "imagenes"
# =========================================================

def procesar_carpeta_imagenes():
    resultados = []

    if not os.path.exists(CARPETA_IMAGENES):
        return resultados

    for archivo in os.listdir(CARPETA_IMAGENES):

        if not archivo.lower().endswith(".pdf"):
            continue

        pdf_path = os.path.join(CARPETA_IMAGENES, archivo)

        try:
            imagenes_b64 = pdf_a_imagenes_base64(pdf_path)

            if not imagenes_b64:
                mover_pdf_imagen(pdf_path, "error")
                resultados.append({"archivo": archivo, "estado": "error", "detalle": "PDF sin páginas"})
                continue

            result_csv = extract_invoice_from_images(archivo, imagenes_b64)
            tabla = auto_logic.csv_to_matrix(result_csv)

            if len(tabla) < 2:
                mover_pdf_imagen(pdf_path, "error")
                resultados.append({"archivo": archivo, "estado": "error", "detalle": "respuesta vacía del modelo"})
                continue

            fila = tabla[1]
            tipo = auto_logic.clasificar_factura(fila)

            mover_pdf_imagen(pdf_path, tipo)
            auto_logic.guardar_historial(result_csv, "imagenes")

            if tipo == "completada":
                auto_logic.guardar_factura_examinada_sql(fila, "imagenes")

            resultados.append({"archivo": archivo, "estado": tipo})

        except Exception as e:
            print(f"ERROR procesando imagen {archivo}: {e}")

            try:
                mover_pdf_imagen(pdf_path, "error")
            except Exception as e2:
                print(f"No se pudo mover {archivo} a error: {e2}")

            resultados.append({"archivo": archivo, "estado": f"ERROR: {e}"})

    return resultados


if __name__ == "__main__":
    if "--watch" in sys.argv:
        intervalo = int(os.getenv("INTERVALO_VIGILANCIA_IMAGENES", "300"))
        print(f"[imagenes.py] Vigilando '{CARPETA_IMAGENES}' cada {intervalo}s...")
        while True:
            for r in procesar_carpeta_imagenes():
                print(r)
            time.sleep(intervalo)
    else:
        for r in procesar_carpeta_imagenes():
            print(r)
