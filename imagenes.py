"""
Segunda pasada para PDFs sin texto extraible (facturas escaneadas o
fotografiadas), que app_auto y app_manual dejan en la carpeta "imagenes".

Este script es el UNICO punto del proyecto que llama a la API con
imagenes (vision). El resto del pipeline (app_auto, app_manual) sigue
trabajando solo con el texto extraido por pdfplumber, así que la llamada
a vision solo se produce para los PDFs que de verdad son imagenes.

Uso:
    python imagenes.py            -> procesa la carpeta "imagenes" una vez
    python imagenes.py --watch    -> repite el proceso cada INTERVALO_VIGILANCIA_IMAGENES segundos
"""

import os
import sys
import csv
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

def pdf_a_imagenes_base64(pdf_path, max_paginas=MAX_PAGINAS, dpi=DPI_IMAGENES):
    """
    Renderiza las primeras `max_paginas` páginas del PDF como PNG y las
    devuelve codificadas en base64, listas para enviar a la API de vision.
    """
    imagenes = []
    zoom = dpi / 72
    matriz = fitz.Matrix(zoom, zoom)

    doc = fitz.open(pdf_path)

    try:
        for pagina in doc[:max_paginas]:
            pix = pagina.get_pixmap(matrix=matriz)
            imagenes.append(base64.b64encode(pix.tobytes("png")).decode("utf-8"))
    finally:
        doc.close()

    return imagenes


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
                    "en leer facturas escaneadas o fotografiadas. Tu única "
                    "salida permitida es una tabla válida separada por |, "
                    "sin explicaciones, sin markdown y sin texto adicional."
                ),
            },
            {"role": "user", "content": contenido},
        ],
    )

    raw = response.choices[0].message.content or ""
    cleaned = auto_logic.clean_llm_csv_response(raw)
    rows = auto_logic.csv_to_matrix(cleaned)

    headers = auto_logic.EXPECTED_HEADERS

    if len(rows) < 2:
        data = ["-"] * len(headers)
    else:
        data = auto_logic.limpiar_fila(rows[1])

    datos = dict(zip(headers, data))
    datos["Archivo"] = file_name
    datos["FEscaneo"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    if datos["FFactura"] != "-":
        datos["FOperacion"] = datos["FFactura"]

    if datos["Buyer"] != "-":
        datos["Buyer"] = datos["Buyer"].replace(" ", "").upper()

    if datos["Proveedor"] != "-":
        datos["Proveedor"] = datos["Proveedor"].replace(" ", "").upper()

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
        "examinada":       os.path.join(FACTURAS_DIR, "examinadas"),
        "manual":          os.path.join(FACTURAS_DIR, "corregir_manualmente"),
        "reenviar_pedido": os.path.join(FACTURAS_DIR, "reenviar_falta_pedidocliente"),
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

            if tipo == "examinada":
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
