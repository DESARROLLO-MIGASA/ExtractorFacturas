"""
Localiza en qué posición de la página del PDF aparece cada dato ya
extraído (NumeroFactura, TotalFact, etc.), para poder pintar un recuadro
encima en el visor de la vista de revisión. Es un cálculo aparte de la
extracción en sí (logic.py / imagenes.py siguen extrayendo solo el valor):
este módulo se limita a buscar dónde está ese valor en la página.

Dos fuentes de palabras con posición, según la página tenga texto o no:
- Con texto: PyMuPDF (fitz) ya da la posición de cada palabra sin coste
  extra (page.get_text("words")).
- Sin texto (factura escaneada): se renderiza la página como imagen (igual
  que hace imagenes.py para la segunda pasada con visión) y se le pasa a
  Tesseract (pytesseract), que sí da posición por palabra en una imagen.

El resultado se cachea en disco por nombre de archivo (no por carpeta:
igual que buscar_pdf_por_nombre, "archivo" es la clave estable aunque el
PDF cambie de carpeta según avanza por el flujo), para no tener que
recalcularlo cada vez que se abre la factura en la vista de revisión.
"""

import os
import re
import io
import json
import unicodedata
import difflib

import fitz  # PyMuPDF

import logic as auto_logic

try:
    import pytesseract
    from PIL import Image
    _TESSERACT_DISPONIBLE = True
except Exception:
    _TESSERACT_DISPONIBLE = False

_tesseract_cmd = os.getenv("TESSERACT_CMD", "").strip()
if _TESSERACT_DISPONIBLE and _tesseract_cmd:
    pytesseract.pytesseract.tesseract_cmd = _tesseract_cmd


class TesseractNoDisponibleError(Exception):
    """El binario de Tesseract (no el paquete pytesseract, que puede estar
    instalado sin problema) no está instalado o no se encuentra en el PATH
    ni en TESSERACT_CMD. Se distingue de "no se encontró texto en la
    región" para no confundir un problema de instalación con una selección
    vacía."""


# _TESSERACT_DISPONIBLE solo confirma que el paquete Python se importa; el
# binario de Tesseract es un programa aparte (ver README) que puede faltar
# aunque el paquete esté instalado. Se comprueba una sola vez aquí, no en
# cada llamada de OCR (que ya lanza su propio proceso), para poder avisar
# con un mensaje claro en vez de que cada intento devuelva simplemente "no
# encontré texto ahí".
_TESSERACT_BINARIO_DISPONIBLE = False
if _TESSERACT_DISPONIBLE:
    try:
        pytesseract.get_tesseract_version()
        _TESSERACT_BINARIO_DISPONIBLE = True
    except Exception as e:
        print(
            f"AVISO: pytesseract está instalado pero el binario de Tesseract no responde ({e}). "
            "Los recuadros en facturas escaneadas y la selección manual de texto no funcionarán "
            "hasta instalarlo (ver README) o configurar TESSERACT_CMD en .env."
        )

FACTURAS_DIR = auto_logic.FACTURAS_DIR
CARPETA_CACHE = os.path.join(FACTURAS_DIR, "cache_posiciones")

DPI_CACHE = int(os.getenv("DPI_CACHE_POSICIONES", os.getenv("DPI_IMAGENES", "200")))

# Campos que nunca tiene sentido intentar localizar en la página (el propio
# nombre de archivo, y la fecha/hora de procesamiento, que no aparece en la
# factura).
CAMPOS_SIN_CAJA = {"Archivo", "FEscaneo"}

# Por debajo de este parecido (0-1, difflib ratio) se descarta la mejor
# ventana encontrada: mejor no pintar recuadro que pintarlo en el sitio
# equivocado.
UMBRAL_COINCIDENCIA = 0.68

# Nº máximo de palabras consecutivas que se prueban como ventana candidata
# (los valores de estos campos rara vez ocupan más de 6-7 palabras seguidas).
MAX_PALABRAS_VENTANA = 8


# =========================================================
# NORMALIZACIÓN PARA COMPARAR
# =========================================================

def _normalizar_comparacion(texto):
    """Dos valores que a simple vista son "el mismo dato" (1.234,56 € vs
    1234,56, o Nº-2024/001 vs N2024001) deben compararse ignorando mayúsculas,
    acentos y cualquier separador/puntuación/símbolo de moneda."""
    texto = str(texto or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", texto)


# =========================================================
# PALABRAS + POSICIÓN POR PÁGINA
# =========================================================

def _palabras_pagina_con_texto(page, zoom):
    palabras = [
        {"x0": x0 * zoom, "y0": y0 * zoom, "x1": x1 * zoom, "y1": y1 * zoom, "texto": palabra}
        for (x0, y0, x1, y1, palabra, _bloque, _linea, _num) in page.get_text("words")
    ]
    palabras.sort(key=lambda p: (p["y0"], p["x0"]))
    return palabras


def _palabras_pagina_ocr(png_bytes):
    if not _TESSERACT_BINARIO_DISPONIBLE:
        return []

    try:
        imagen = Image.open(io.BytesIO(png_bytes))
        datos = pytesseract.image_to_data(imagen, output_type=pytesseract.Output.DICT)
    except Exception as e:
        print(f"AVISO: fallo de OCR al calcular posiciones: {e}")
        return []

    palabras = []
    for i in range(len(datos.get("text", []))):
        texto = (datos["text"][i] or "").strip()
        if not texto:
            continue
        try:
            confianza = float(datos["conf"][i])
        except (ValueError, TypeError):
            confianza = -1
        if confianza < 30:  # descarta ruido de OCR de baja confianza
            continue
        x, y, w, h = datos["left"][i], datos["top"][i], datos["width"][i], datos["height"][i]
        palabras.append({"x0": x, "y0": y, "x1": x + w, "y1": y + h, "texto": texto})

    palabras.sort(key=lambda p: (p["y0"], p["x0"]))
    return palabras


def _mapa_paginas(pdf_path, dpi=DPI_CACHE):
    """{pagina_idx: {"ancho":, "alto":, "palabras": [...]}}, con las palabras
    y su posición en píxeles a `dpi` (el mismo con el que se renderiza la
    página para el visor, así los dos encajan sin reescalar aparte)."""
    zoom = dpi / 72
    mapa = {}

    doc = fitz.open(pdf_path)
    try:
        for i, page in enumerate(doc):
            texto_pagina = (page.get_text() or "").strip()

            if len(texto_pagina) >= 20:
                rect = page.rect
                ancho, alto = rect.width * zoom, rect.height * zoom
                palabras = _palabras_pagina_con_texto(page, zoom)
            else:
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                ancho, alto = pix.width, pix.height
                palabras = _palabras_pagina_ocr(pix.tobytes("png"))

            mapa[i] = {"ancho": ancho, "alto": alto, "palabras": palabras}
    finally:
        doc.close()

    return mapa


# =========================================================
# EMPAREJAR VALOR EXTRAÍDO -> MEJOR VENTANA DE PALABRAS
# =========================================================

def _mejor_caja_para_valor(valor, mapa_paginas):
    objetivo = _normalizar_comparacion(valor)
    if len(objetivo) < 2:
        return None

    mejor = None  # (ratio, pagina, x0, y0, x1, y1)

    for pagina, info in mapa_paginas.items():
        palabras = info["palabras"]
        n = len(palabras)

        for inicio in range(n):
            acumulado = ""
            for fin in range(inicio, min(inicio + MAX_PALABRAS_VENTANA, n)):
                acumulado += _normalizar_comparacion(palabras[fin]["texto"])
                if not acumulado:
                    continue

                ratio = difflib.SequenceMatcher(None, objetivo, acumulado).ratio()
                if mejor is None or ratio > mejor[0]:
                    ventana = palabras[inicio:fin + 1]
                    mejor = (
                        ratio, pagina,
                        min(p["x0"] for p in ventana), min(p["y0"] for p in ventana),
                        max(p["x1"] for p in ventana), max(p["y1"] for p in ventana),
                    )

                # Alargar la ventana ya no va a mejorar el parecido si el
                # acumulado dobla holgadamente la longitud del objetivo.
                if len(acumulado) > len(objetivo) * 2 + 6:
                    break

    if mejor is None or mejor[0] < UMBRAL_COINCIDENCIA:
        return None

    _ratio, pagina, x0, y0, x1, y1 = mejor
    margen = 4  # recuadro "que abarca un poco más", no un subrayado ajustado
    return {
        "pagina": pagina,
        "x0": max(x0 - margen, 0),
        "y0": max(y0 - margen, 0),
        "x1": x1 + margen,
        "y1": y1 + margen,
    }


def calcular_cajas_campos(pdf_path, datos_fila):
    """`datos_fila`: dict {campo: valor} con las claves de EXPECTED_HEADERS.
    Nunca lanza excepción hacia quien la llama: cualquier fallo se traduce
    en "sin cajas" para no bloquear la revisión manual."""
    try:
        mapa_paginas = _mapa_paginas(pdf_path)
    except Exception as e:
        print(f"AVISO: no se pudieron calcular posiciones para {pdf_path}: {e}")
        return {"paginas_render": {}, "campos": {}}

    campos = {}
    for campo, valor in (datos_fila or {}).items():
        if campo in CAMPOS_SIN_CAJA:
            continue

        valor = str(valor or "").strip()
        if valor in ("", "-"):
            continue

        try:
            caja = _mejor_caja_para_valor(valor, mapa_paginas)
        except Exception as e:
            print(f"AVISO: fallo al localizar el campo {campo} de {pdf_path}: {e}")
            caja = None

        if caja:
            campos[campo] = caja

    paginas_render = {
        str(i): {"ancho": info["ancho"], "alto": info["alto"], "dpi": DPI_CACHE}
        for i, info in mapa_paginas.items()
    }

    return {"paginas_render": paginas_render, "campos": campos}


# =========================================================
# CACHÉ EN DISCO (por nombre de archivo, no por carpeta de flujo)
# =========================================================

def _ruta_cache(archivo):
    base, _ext = os.path.splitext(archivo)
    return os.path.join(CARPETA_CACHE, base + ".json")


def guardar_cajas_en_cache(archivo, cajas):
    os.makedirs(CARPETA_CACHE, exist_ok=True)
    try:
        with open(_ruta_cache(archivo), "w", encoding="utf-8") as f:
            json.dump(cajas, f, ensure_ascii=False)
    except Exception as e:
        print(f"AVISO: no se pudo guardar la caché de posiciones de {archivo}: {e}")


def cargar_cajas_de_cache(archivo):
    ruta = _ruta_cache(archivo)
    if not os.path.exists(ruta):
        return None
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def calcular_cajas_lineas(pdf_path, lineas):
    """`lineas`: lista de dicts {Descripcion, Cantidad, Precio, Importe, Otros}
    (ver guardar_lineas_csv en logic.py). Devuelve una lista del mismo largo
    y orden que `lineas`, con la caja de cada una o None si no se localizó.

    Cada línea se localiza por su Descripción (el texto más distintivo de la
    fila); si además se encuentra el Importe en la misma página, el recuadro
    se amplía para cubrir hasta ahí, de modo que abarque toda la fila de la
    tabla (descripción a la izquierda, importe a la derecha) y no solo la
    descripción. Nunca lanza excepción: cualquier fallo se traduce en "sin
    caja" para esa línea, igual que calcular_cajas_campos."""
    try:
        mapa_paginas = _mapa_paginas(pdf_path)
    except Exception as e:
        print(f"AVISO: no se pudieron calcular posiciones de líneas para {pdf_path}: {e}")
        return [None] * len(lineas or [])

    cajas = []
    for linea in lineas or []:
        descripcion = str((linea or {}).get("Descripcion", "-") or "-").strip()

        caja = None
        if descripcion not in ("", "-"):
            try:
                caja = _mejor_caja_para_valor(descripcion, mapa_paginas)
            except Exception as e:
                print(f"AVISO: fallo al localizar la línea '{descripcion}' de {pdf_path}: {e}")

        importe = str((linea or {}).get("Importe", "-") or "-").strip()
        if caja and importe not in ("", "-"):
            try:
                caja_importe = _mejor_caja_para_valor(importe, mapa_paginas)
            except Exception:
                caja_importe = None

            if caja_importe and caja_importe["pagina"] == caja["pagina"]:
                caja = {
                    "pagina": caja["pagina"],
                    "x0": min(caja["x0"], caja_importe["x0"]),
                    "y0": min(caja["y0"], caja_importe["y0"]),
                    "x1": max(caja["x1"], caja_importe["x1"]),
                    "y1": max(caja["y1"], caja_importe["y1"]),
                }

        cajas.append(caja)

    return cajas


def calcular_y_cachear_lineas(archivo, lineas, pdf_path=None):
    """Igual que calcular_y_cachear, pero para las líneas de factura. Se
    guarda bajo la clave "lineas" del mismo fichero de caché que usan las
    cajas por campo (ver _ruta_cache), para no tener que duplicar
    paginas_render ni gestionar un segundo fichero por factura."""
    ruta = pdf_path or auto_logic.buscar_pdf_por_nombre(archivo)
    if not ruta:
        return [None] * len(lineas or [])

    cajas_lineas = calcular_cajas_lineas(ruta, lineas)

    cache = cargar_cajas_de_cache(archivo) or {"paginas_render": {}, "campos": {}}
    cache["lineas"] = cajas_lineas
    guardar_cajas_en_cache(archivo, cache)

    return cajas_lineas


def obtener_o_calcular_cajas_lineas(archivo, lineas):
    """Sirve la caché si ya tiene calculadas las cajas de líneas para esta
    factura; si no (primera vez, o caché de antes de que existiera esta
    función), las calcula."""
    cache = cargar_cajas_de_cache(archivo)
    if cache is not None and "lineas" in cache:
        return cache["lineas"]
    return calcular_y_cachear_lineas(archivo, lineas)


def calcular_y_cachear(archivo, datos_fila, pdf_path=None):
    """Calcula las cajas y las deja en caché. Si se conoce ya `pdf_path`
    (como en la extracción, donde ya se tiene la ruta a mano) se evita
    volver a localizar el PDF con buscar_pdf_por_nombre."""
    ruta = pdf_path or auto_logic.buscar_pdf_por_nombre(archivo)
    if not ruta:
        return {"paginas_render": {}, "campos": {}}

    cajas = calcular_cajas_campos(ruta, datos_fila)
    guardar_cajas_en_cache(archivo, cajas)
    return cajas


def obtener_o_calcular_cajas(archivo, datos_fila):
    """Sirve la caché si existe; si no, la calcula (cubre también facturas
    procesadas antes de que existiera este módulo, sin necesitar migración)."""
    cajas = cargar_cajas_de_cache(archivo)
    if cajas is not None:
        return cajas
    return calcular_y_cachear(archivo, datos_fila)


# =========================================================
# SELECCIÓN MANUAL: OCR DE UNA REGIÓN ARRASTRADA A MANO
# =========================================================

def recortar_region_png(pdf_path, pagina, x0, y0, x1, y1, dpi=DPI_CACHE):
    """Bytes PNG de esa región de esa página (en píxeles a `dpi`, el mismo
    sistema de coordenadas que ve el visor), o None si la página no existe.
    Comparte el recorte con recortar_y_ocr_region, pero sin pasar por
    Tesseract: para usos que necesiten la imagen en sí, no texto plano (ver
    logic.extraer_lineas_de_region)."""
    zoom = dpi / 72
    doc = fitz.open(pdf_path)
    try:
        if pagina < 0 or pagina >= doc.page_count:
            return None
        page = doc[pagina]
        # clip va en el sistema de coordenadas de la página (puntos), no en
        # los píxeles ya escalados que llegan del visor.
        recorte = fitz.Rect(x0 / zoom, y0 / zoom, x1 / zoom, y1 / zoom)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=recorte)
        return pix.tobytes("png")
    finally:
        doc.close()


def recortar_y_ocr_region(archivo, pagina, x0, y0, x1, y1, dpi=DPI_CACHE):
    """Recorta esa región (en píxeles a `dpi`, el mismo sistema de
    coordenadas que ve el visor) y le pasa Tesseract. Funciona igual para
    una factura con texto que para una escaneada: siempre vuelve a
    renderizar la página con fitz, así que no depende de si esa página
    tenía o no capa de texto."""
    if not _TESSERACT_BINARIO_DISPONIBLE:
        raise TesseractNoDisponibleError(
            "Tesseract OCR no está instalado en el servidor (o TESSERACT_CMD no apunta a su ubicación)."
        )

    ruta = auto_logic.buscar_pdf_por_nombre(archivo)
    if not ruta:
        return ""

    png_bytes = recortar_region_png(ruta, pagina, x0, y0, x1, y1, dpi)
    if png_bytes is None:
        return ""

    try:
        imagen = Image.open(io.BytesIO(png_bytes))
        # --psm 7: la región recortada es una sola línea de texto.
        return pytesseract.image_to_string(imagen, config="--psm 7").strip()
    except Exception as e:
        print(f"AVISO: fallo de OCR de región para {archivo}: {e}")
        return ""
