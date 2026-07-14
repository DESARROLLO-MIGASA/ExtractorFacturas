import os
import re
import sys
import csv
import json
import tempfile
from io import StringIO
from datetime import datetime
import httpx
import shutil

import pdfplumber
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from dotenv import load_dotenv
load_dotenv(dotenv_path=os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sql_historial import (
    guardar_factura_examinada_sql,
    listar_facturas_examinadas_sql,
    marcar_factura_definitiva_sql,
    eliminar_factura_examinada_sql,
    buscar_empresa_por_cif,
    buscar_proveedor_por_cif,
    marcar_revisada_sql,
    archivos_revisados_sql,
)

_fallback = os.path.join(os.path.dirname(os.path.abspath(__file__)), "facturas")
FACTURAS_DIR = os.getenv("FACTURAS_DIR", _fallback).strip('"').strip("'")

CORREGIR_DIR         = os.path.join(FACTURAS_DIR, "corregir_manualmente")
COMPLETADAS_DIR      = os.path.join(FACTURAS_DIR, "completadas")
FACTURAS_REVISADAS_DIR = os.path.join(FACTURAS_DIR, "facturas_revisadas")
HISTORIAL_DIR        = os.path.join(FACTURAS_DIR, "historial")
NO_FACTURA_DIR       = os.path.join(FACTURAS_DIR, "no_es_factura")
ERROR_DIR            = os.path.join(FACTURAS_DIR, "error")
INCIDENCIAS_DIR      = os.path.join(FACTURAS_DIR, "incidencias")
REENVIAR_PEDIDO_DIR  = os.path.join(FACTURAS_DIR, "reenviar_falta_pedidocliente")
REENVIADAS_PEDIDO_DIR = os.path.join(FACTURAS_DIR, "reenviadas_falta_pedidocliente")
REENVIAR_DOS_FACTURAS_DIR = os.path.join(FACTURAS_DIR, "reenviar_dos_factura_una_pagina")
REENVIAR_OTRO_MOTIVO_DIR   = os.path.join(FACTURAS_DIR, "reenviar_otro_motivo")
REENVIADAS_OTRO_MOTIVO_DIR = os.path.join(FACTURAS_DIR, "reenviadas_otro_motivo")
REENVIAR_ERROR_PESA_MUCHO_DIR = os.path.join(FACTURAS_DIR, "reenviar_error_pesa_mucho")
REENVIAR_ERROR_OTRO_DIR       = os.path.join(FACTURAS_DIR, "reenviar_error_otro")

# Prefijo con el que se marcan en disco las facturas subidas como urgentes,
# para que aparezcan primero en la lista de pendientes sin necesitar una tabla aparte.
PRIORIDAD_PREFIX = "URGENTE__"

# Motivos seleccionables para pedir por correo una aclaración de la factura.
# Ninguno se envía solo: siempre hace falta que una persona revise la factura
# y elija el motivo antes de que se copie a la carpeta que vigila Power Automate
# (salvo "otros", que además guía a redactar el correo a mano con el PDF).
MOTIVOS_ENVIO_CORREO = {
    "falta_pedido_cliente":     "Falta el número de pedido de cliente",
    "dos_facturas_una_pagina":  "Hay dos o más facturas en la misma página/PDF",
    "otros":                    "Otro",
}

# Motivos de clasificación de un error de extracción (carpeta "error").
MOTIVOS_ERROR_EXTRACCION = {
    "pesa_mucho": "El fichero pesa demasiado",
    "otro":       "Otro",
}


# =========================================================
# CONFIG / PROMPT
# =========================================================

EXPECTED_HEADERS = [
    "Archivo",
    "NumeroFactura",
    "Buyer",
    "Empresa",
    "Proveedor",
    "NombreProveedor",
    "PedidoCliente",
    "BaseImp",
    "BaseIRPF",
    "TipoIVA",
    "TipoIVA2",
    "TipoIVA3",
    "ImporIVA",
    "TotalFact",
    "Moneda",
    "FFactura",
    "FOperacion",
    "FEscaneo",
]

DEFAULT_PROMPT = """Eres un extractor estricto de datos de facturas de proveedor.

Debes devolver los datos como un objeto JSON con un valor de texto por cada
campo indicado más abajo (el formato exacto del JSON ya viene forzado por el
esquema de la petición; tú solo tienes que rellenar bien cada campo).

Reglas generales:
- Si un dato no aparece claramente, devuelve "-".
- No inventes ningún dato.
- No deduzcas datos que no estén explícitos en la factura.
- Mantén el texto limpio, sin saltos de línea.
- No añadas comentarios.
- No añadas unidades ni símbolos de moneda salvo que formen parte inseparable del dato.
- Para importes, devuelve solo el número, usando coma decimal si aparece así en el documento.
- Para fechas, devuelve el formato que aparezca en el documento. Si puedes normalizar con seguridad, usa DD/MM/AAAA.

Definición de campos:

IDFactura:
- Identificador único de la factura.
- Normalmente coincide con el ID documental o identificador principal de la factura.
- Si no aparece claramente devuelve "-".

BaseImp:
- Base imponible total de la factura.
- Devuelve únicamente el importe.

BaseIRPF:
- Base del IRPF si existe.
- Si no existe devuelve "-".

Buyer:
- CIF/NIF/VAT del comprador/cliente (nunca del proveedor/vendedor que emite la factura).
- Puede aparecer bajo etiquetas como:
  NIF
  CIF
  VAT
  VAT Number
  Tax ID
  Tax Number
  N° TVA
  Nº Contribuinte
  V/ Nº Contribuinte
  Vosso Contribuinte
  Partita IVA
  USt-IdNr
  BTW-nummer
- Devuelve únicamente el identificador fiscal.
- Nunca devuelvas el nombre de la empresa.
- NUNCA devuelvas un código de cliente/cuenta interno (lo que aparece tras
  etiquetas como "Cliente:", "Customer:", "Account:", "Nº Cliente", "Código
  Cliente"). Esos códigos son una referencia interna del emisor de la
  factura, no un CIF/NIF/VAT, aunque tengan un formato parecido (letra +
  números).
- Es muy habitual que un código de cliente ("Cliente: 003565") y el
  CIF/NIF/VAT real del comprador ("NIF/CIF: B16709305") aparezcan juntos en
  el mismo bloque de dirección. En ese caso ignora el código de cliente y
  usa el valor de la etiqueta NIF/CIF/VAT explícita, aunque esté más abajo
  o parezca menos destacado que el código de cliente.
- Si el único dato disponible junto al comprador es un código de cliente de
  este tipo y no hay ninguna etiqueta NIF/CIF/VAT explícita para él en todo
  el documento, devuelve "-".
- Si no aparece claramente devuelve "-".

Empresa:
- Siempre devuelve "-".
- Este campo se rellena posteriormente por el sistema a partir del CIF/NIF/VAT del comprador (Buyer), nunca lo extraigas del texto.

FEscaneo:
- Siempre devuelve "-".
- Este campo será completado posteriormente por el sistema.

FFactura:
- Fecha de factura.

FOperacion:
- Si no aparece claramente devuelve "-".

ImporIVA:
- Si la factura indica que no existe IVA, VAT o impuesto aplicable, devuelve 0.

Moneda:
- Devuelve EUR, USD, GBP o la moneda indicada en la factura.
- Si aparece € devuelve EUR.

NombreProveedor:
- Siempre devuelve "-".
- Este campo se rellena posteriormente por el sistema a partir del CIF/NIF/VAT del proveedor, nunca lo extraigas del texto.

NumeroFactura:
- Número de factura del proveedor.

PedidoCliente:
- Número de pedido del cliente.
- Puede aparecer como:
  Pedido Cliente
  Su Pedido
  Su referencia
  Customer Order
  Customer PO
  Purchase Order
  PO Number
  Order Number
  Nº Pedido
  Pedido
  Order Ref
  Customer Reference
  Your Order
  Your Reference
  Ref. Cliente
  Referencia Cliente
  Votre commande
  Ihre Bestellung
  Uw order
  Vostro ordine
- Si hay varios pedidos en la factura (una línea por pedido), devuelve todos separados por punto y coma (;).
- Prioriza el campo cuya etiqueta sea "Su Pedido", "Su referencia" o "Customer PO" sobre otros campos de referencia genéricos.
- NO devolver el número de factura ni el número de albarán.
- Devuelve únicamente el pedido del cliente.
- Si no existe devuelve "-".

Proveedor:
- CIF/NIF/VAT del proveedor/vendedor que emite la factura (nunca del comprador/cliente).
- Puede aparecer bajo etiquetas como:
  NIF
  CIF
  VAT
  VAT Number
  Tax ID
  Tax Number
  N° TVA
  Nº Contribuinte
  Partita IVA
  USt-IdNr
  BTW-nummer
- Suele aparecer junto al nombre y dirección del vendedor en la cabecera, o en el pie de página junto a los datos legales/registrales de la empresa emisora (registro mercantil, capital social, etc.).
- El pie de página con los datos registrales (Registro Mercantil, capital
  social, protección de datos) identifica casi siempre a la empresa que
  EMITE la factura, aunque en la cabecera aparezca destacado el nombre o el
  CIF del comprador (algunas facturas imprimen primero los datos de envío/
  facturación del cliente y solo mencionan al emisor en esa letra pequeña
  del pie). Si hay conflicto entre un CIF de la cabecera y uno del pie de
  página junto a "Registro Mercantil"/"C.I.F.-" de una empresa distinta,
  prioriza el del pie de página como Proveedor.
- Devuelve únicamente el identificador fiscal.
- Si no aparece claramente devuelve "-".

TipoIVA:
- Porcentaje de IVA aplicado a la factura (por ejemplo: 21, 10, 4).
- Devuelve únicamente el número del porcentaje, sin el símbolo %.
- Si la factura indica exención, inversión del sujeto pasivo o impuesto 0%, devuelve 0.
- No devuelvas "-" cuando pueda determinarse que el IVA es cero.
- Si no se puede determinar ningún porcentaje, devuelve "-".

TipoIVA2:
- Segundo tipo de IVA (porcentaje) si la factura desglosa un segundo tipo distinto al de TipoIVA.
- Si no existe un segundo tipo de IVA, devuelve "-". No omitas esta columna bajo ningún concepto.

TipoIVA3:
- Tercer tipo de IVA (porcentaje) si la factura desglosa un tercer tipo distinto a los anteriores.
- Si no existe un tercer tipo de IVA, devuelve "-". No omitas esta columna bajo ningún concepto.

TotalFact:
- Importe total final de la factura, impuestos incluidos.
- Es un importe en dinero, nunca un porcentaje de IVA.


IMPORTANTE:

Las facturas pueden estar en cualquier idioma.

Debes reconocer automáticamente los campos aunque aparezcan en:
- español
- inglés
- portugués
- francés
- italiano
- alemán
- neerlandés
- chino
- griego
- otros idiomas

No dependas del idioma para localizar la información.

Identifica los conceptos por su significado semántico y no por palabras exactas.
"""


# =========================================================
# CLIENTE OPENAI / AZURE
# =========================================================

def paginas_pdf(pdf_path):

    paginas = []

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            texto = page.extract_text() or ""

            paginas.append(texto)

    return paginas


def detectar_numero_factura(texto):

    client = build_client()

    response = client.chat.completions.create(
        model=get_model(),
        temperature=0,
        messages=[
            {
                "role": "system",
                "content": """
Devuelve únicamente el número de factura.

Si no existe devuelve -.

Sin explicaciones.
"""
            },
            {
                "role": "user",
                "content": texto[:15000]
            }
        ]
    )

    return response.choices[0].message.content.strip()

def agrupar_facturas(pdf_path):

    paginas = paginas_pdf(pdf_path)

    grupos = {}

    for pagina in paginas:

        numero = detectar_numero_factura(pagina)

        if numero not in grupos:
            grupos[numero] = []

        grupos[numero].append(pagina)

    return [
        "\n".join(paginas_factura)
        for paginas_factura in grupos.values()
    ]


def build_client():
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    base_url = os.getenv("OPENAI_BASE_URL", "").strip()

    if not api_key:
        raise Exception("No se ha encontrado OPENAI_API_KEY. Revisa el .env")

    http_client = httpx.Client(verify=False)

    if base_url:
        return OpenAI(
            api_key=api_key,
            base_url=base_url,
            http_client=http_client
        )

    return OpenAI(
        api_key=api_key,
        http_client=http_client
    )


def get_model():
    return os.getenv("OPENAI_MODEL", "gpt-4.1").strip()


# =========================================================
# PDF
# =========================================================

def read_pdf_text(path):

    text = ""

    with pdfplumber.open(path) as pdf:

        paginas = pdf.pages

        if len(paginas) > 10:

            seleccion = (
                paginas[:2]
                + paginas[-5:]
            )

        else:
            seleccion = paginas

        for page in seleccion:

            content = page.extract_text()

            if content:
                text += content + "\n"

    return text.strip()


def clean_pdf_text(text):
    if not text:
        return ""

    text = text.replace("\x00", " ")
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def pdf_a_markdown_factura(text):
    """
    Convierte el texto extraído de la factura en pseudo-markdown
    para ayudar al modelo a localizar bloques relevantes.
    """
    if not text:
        return ""

    text = clean_pdf_text(text)
    lineas = [l.strip() for l in text.splitlines() if l.strip()]
    resultado = []

    claves_seccion = [
        "FACTURA",
        "INVOICE",
        "PROVEEDOR",
        "SUPPLIER",
        "VENDOR",
        "CLIENTE",
        "CUSTOMER",
        "NIF",
        "CIF",
        "VAT",
        "BASE IMPONIBLE",
        "BASE",
        "IVA",
        "VAT",
        "TOTAL",
        "FECHA",
        "DATE",
        "VENCIMIENTO",
        "DUE DATE",
        "DIRECCIÓN",
        "DIRECCION",
        "ADDRESS",
        "POSTAL",
        "COUNTRY",
        "PAÍS",
        "PAIS",
        # Order / pedido keywords
        "PEDIDO",
        "ORDER",
        "PURCHASE ORDER",
        "PO NUMBER",
        "PO NO",
        "ORDER NUMBER",
        "ORDER REF",
        "CUSTOMER ORDER",
        "CUSTOMER PO",
        "CUSTOMER REFERENCE",
        "REFERENCIA",
        "REFERENCE",
        "Nº PEDIDO",
        "NO. PEDIDO",
        "NR. ORDER",
        "COMMANDE",
        "BESTELLUNG",
        "ORDINE",
    ]

    for linea in lineas:
        up = linea.upper()

        if any(k in up for k in claves_seccion):
            resultado.append(f"\n## {linea}\n")
        else:
            resultado.append(linea)

    md = "\n".join(resultado)
    md = re.sub(r"\n{3,}", "\n\n", md).strip()

    return md


# =========================================================
# CSV / UTILIDADES
# =========================================================

def csv_to_matrix(csv_text):
    lines = [line.strip() for line in str(csv_text).splitlines() if line.strip()]
    return [line.split("|") for line in lines]


def normalizar_valor(v):
    v = str(v).strip()

    if v in ["", "N/A", "NA", "No aplica", "NO APLICA", "n/a", "na", "None", "null"]:
        return "-"

    return v


# Nombres de mes (sin acentos) en los idiomas que aparecen en las facturas
# que procesamos (ES/EN/FR/IT/PT/DE), para poder normalizar fechas del tipo
# "12 de marzo de 2024" o "12 March 2024" además de las puramente numéricas.
_MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12,
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7, "aug": 8,
    "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
    "janvier": 1, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
    "juillet": 7, "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11,
    "decembre": 12,
    "gennaio": 1, "febbraio": 2, "aprile": 4, "maggio": 5, "giugno": 6,
    "luglio": 7, "settembre": 9, "ottobre": 10, "dicembre": 12,
    "janeiro": 1, "fevereiro": 2, "marco": 3, "maio": 5, "junho": 6,
    "julho": 7, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
    "januar": 1, "februar": 2, "marz": 3, "juni": 6, "juli": 7,
    "oktober": 10, "dezember": 12,
}


def _quitar_acentos(s):
    import unicodedata
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()


def normalizar_fecha(v):
    """Reescribe una fecha extraída por el LLM (en cualquier formato en que
    haya venido) como DD/MM/AAAA, igual que FEscaneo. Si no se reconoce el
    formato, se devuelve el valor tal cual para no perder el dato."""
    v = str(v).strip()

    if v in ("", "-"):
        return "-"

    # DD/MM/AAAA, DD-MM-AAAA, DD.MM.AAAA (año de 2 o 4 dígitos)
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$", v)
    if m:
        dia, mes, anio = m.groups()
        if len(anio) == 2:
            anio = ("20" if int(anio) <= 79 else "19") + anio
        try:
            return datetime(int(anio), int(mes), int(dia)).strftime("%d/%m/%Y")
        except ValueError:
            return v

    # AAAA-MM-DD, AAAA/MM/DD
    m = re.match(r"^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$", v)
    if m:
        anio, mes, dia = m.groups()
        try:
            return datetime(int(anio), int(mes), int(dia)).strftime("%d/%m/%Y")
        except ValueError:
            return v

    texto = _quitar_acentos(v.lower())

    # "12 de marzo de 2024", "12 march 2024", "12-mar-2024"
    m = re.match(r"^(\d{1,2})\s*(?:de)?\s*[/\-.\s]\s*([a-z]+)\.?\s*[/\-.,]?\s*(?:de)?\s*(\d{4})$", texto)
    if m:
        dia, mes_txt, anio = m.groups()
        mes = _MESES.get(mes_txt)
        if mes:
            try:
                return datetime(int(anio), mes, int(dia)).strftime("%d/%m/%Y")
            except ValueError:
                return v

    # "march 12, 2024", "march 12 2024"
    m = re.match(r"^([a-z]+)\.?\s+(\d{1,2}),?\s+(\d{4})$", texto)
    if m:
        mes_txt, dia, anio = m.groups()
        mes = _MESES.get(mes_txt)
        if mes:
            try:
                return datetime(int(anio), mes, int(dia)).strftime("%d/%m/%Y")
            except ValueError:
                return v

    return v


def limpiar_fila(row):
    row = (row + [""] * len(EXPECTED_HEADERS))[:len(EXPECTED_HEADERS)]
    return [normalizar_valor(x) for x in row]


_RE_CIF_NO_ALFANUM = re.compile(r"[^A-Za-z0-9]")


def normalizar_cif(valor):
    """Deja un CIF/NIF/VAT listo para comparar contra la base de datos:
    quita espacios, guiones, puntos y cualquier otro carácter que no sea
    letra o número (las facturas los escriben con formatos muy distintos,
    p.ej. "B-90.207.085"), y pone todo en mayúsculas."""
    if valor is None or valor == "-":
        return "-"
    limpio = _RE_CIF_NO_ALFANUM.sub("", str(valor)).upper()
    return limpio if limpio else "-"


# El modelo a veces no extrae un CIF que sí aparece con etiqueta clara en
# el documento: en letra pequeña de pie de página, con la etiqueta
# abreviada ("NIF ES B14092902", "R.M. de Córdoba" en vez de "Registro
# Mercantil"), o con dos CIF (comprador y proveedor) compartiendo una sola
# etiqueta en la misma línea ("C.I.F. B41510223 B91950253"). Pedirle al
# modelo que preste más atención por prompt resultó frágil (en pruebas, un
# modelo llegaba a devolver "-" en Buyer y Proveedor a la vez con solo
# tocar el texto del prompt), así que en vez de eso se buscan por regex
# TODAS las etiquetas CIF/NIF/VAT del documento, como candidatos de
# respaldo. No se asigna ningún CIF a un campo por sí sola: solo aporta
# candidatos que resolver_empresa_y_proveedor comprobará contra las tablas
# maestras (y a los que no coincidan con ninguna se les da visibilidad en
# vez de descartarlos en silencio, ver esa función).
_RE_ETIQUETA_CIF = re.compile(r"\b(?:C\.?I\.?F\.?|N\.?I\.?F\.?|VAT)\.?\s*(?:ES)?[:\-\.\s]+", re.IGNORECASE)
_RE_TOKEN_CIF = re.compile(r"[A-Z][\-\. ]?\d{7,8}[0-9A-Z]?\b")
_VENTANA_ETIQUETA_CIF = 35


def detectar_cifs_con_etiqueta(texto):
    """
    Devuelve la lista (sin duplicados, en orden de aparición) de todos los
    CIF/NIF/VAT que aparecen etiquetados en cualquier parte de `texto`, no
    solo el primero. Devuelve [] si no se encuentra ninguno.
    """
    if not texto:
        return []

    vistos = []
    for m in _RE_ETIQUETA_CIF.finditer(texto):
        ventana = texto[m.end():m.end() + _VENTANA_ETIQUETA_CIF]

        for tm in _RE_TOKEN_CIF.finditer(ventana):
            cif = normalizar_cif(tm.group())
            if cif != "-" and cif not in vistos:
                vistos.append(cif)

    return vistos


# Algunas facturas meten el CIF del proveedor suelto, sin ninguna etiqueta
# CIF/NIF, pegado justo a la web o al teléfono de contacto en el pie de
# página (p.ej. "+34 954 18 66 80 https://www.procisa.es A41071465"). No
# hay ninguna palabra ancla como "CIF" o "Registro Mercantil" cerca, así
# que se busca específicamente justo después de una URL/www.
_RE_CIF_TRAS_URL = re.compile(
    r"(?:https?://|www\.)\S+\s+([A-Z][\-\. ]?\d{7,8}[0-9A-Z]?)\b",
    re.IGNORECASE,
)


def detectar_cif_tras_url(texto):
    """Devuelve la lista de CIF encontrados justo después de una URL/web
    en `texto`, sin ninguna etiqueta CIF/NIF de por medio. Devuelve [] si
    no se encuentra ninguno."""
    if not texto:
        return []

    vistos = []
    for m in _RE_CIF_TRAS_URL.finditer(texto):
        cif = normalizar_cif(m.group(1))
        if cif != "-" and cif not in vistos:
            vistos.append(cif)

    return vistos


_RE_CIF_SUELTO = re.compile(r"[A-Z][\-\. ]?\d{7,8}[0-9A-Z]?")


def extraer_cif_de_texto_libre(texto):
    """
    Busca el primer CIF con forma válida dentro de una cadena libre. Sirve
    para sanear respuestas de vision que a veces devuelven más texto del
    pedido (p.ej. "OLEO VERDE S.L. NIF B91580142" en vez de solo el CIF),
    que normalizar_cif por sí sola dejaría todo pegado en un único token
    inválido. Devuelve el CIF normalizado, o "-" si no encuentra ninguno.
    """
    if not texto or texto == "-":
        return "-"
    m = _RE_CIF_SUELTO.search(texto)
    return normalizar_cif(m.group()) if m else "-"


def _es_eco_corrupto(candidato, referencia):
    """
    Heurística barata para descartar una alucinación típica del respaldo
    por visión: en vez de reconocer el CIF que de verdad falta, el modelo
    devuelve el MISMO CIF que ya se conoce del otro campo pero con un
    carácter de más o de menos (p.ej. referencia="B91616227", candidato=
    "B916162227"). Si `candidato` es exactamente `referencia` con un solo
    carácter insertado (o viceversa), se considera sospechoso.
    """
    if not candidato or not referencia or candidato in ("-", referencia):
        return False

    largo, corto = (candidato, referencia) if len(candidato) > len(referencia) else (referencia, candidato)
    if len(largo) - len(corto) != 1:
        return False

    return any(largo[:i] + largo[i + 1:] == corto for i in range(len(largo)))


# Se muestra en Empresa/NombreProveedor cuando el CIF sí se ha detectado en la
# factura pero no existe (o no está activo/sin bloquear) en la tabla maestra
# que le corresponde, para distinguirlo de un "-" que significaría que no se
# encontró ningún CIF. clasificar_factura también reconoce este mensaje como
# "no resuelto" a la hora de decidir si la factura es una incidencia.
MENSAJE_CIF_NO_ENCONTRADO = "El CIF no se encuentra en la base de datos"


def resolver_empresa_y_proveedor(cif_buyer, cif_proveedor, candidatos_extra=None):
    """
    Cuál de los CIF/NIF/VAT de la factura es el comprador (Buyer/Empresa) y
    cuál el proveedor (Proveedor/NombreProveedor) lo decide esta función
    consultando las tablas maestras, no la posición/contexto que haya usado
    el modelo para etiquetarlos. Así, si el modelo confunde cabecera y pie
    de página y etiqueta los CIF al revés, la factura se sigue clasificando
    bien mientras cada CIF exista en la tabla que le corresponde.

    cif_buyer y cif_proveedor son los CIF tal como los etiquetó el modelo;
    candidatos_extra es la lista de respaldo que haya encontrado
    detectar_cifs_con_etiqueta (CIF con etiqueta clara en el documento que
    el modelo pasó por alto). Por defecto se asume que son del proveedor
    -es el caso más habitual: el CIF del comprador casi siempre lo
    encuentra ya el modelo-, pero si alguno resulta ser el del comprador lo
    decide igualmente el cruce contra las tablas maestras, no esta
    suposición. Un "-"/[] significa que no hay candidato en esa posición.

    Un candidato que no coincida con ninguna tabla (p.ej. un proveedor nuevo
    que todavía no se ha dado de alta en ProveedoresClasificados) no se
    descarta en silencio: se deja visible en el hueco de Buyer/Proveedor
    que le corresponda SEGÚN SU ROL (no en el primer hueco libre) — así, si
    solo se ha encontrado un CIF y era el del proveedor, no se cuela como
    si fuera el del comprador. Empresa/NombreProveedor muestra entonces
    MENSAJE_CIF_NO_ENCONTRADO en vez de un "-" que no distingue "no se
    encontró ningún CIF" de "se encontró pero no está de alta".

    Si el mismo CIF aparece con roles distintos (p.ej. porque quedó mal
    etiquetado como Buyer en una extracción antigua, pero también aparece
    en candidatos_extra), gana el rol de candidatos_extra: es una señal más
    fiable porque viene de una etiqueta explícita en el documento, en vez
    de una etiqueta sin verificar.

    Devuelve (buyer_cif, nombre_empresa, proveedor_cif, nombre_proveedor).
    """
    candidatos_por_prioridad = [(cif_proveedor, "proveedor")]
    candidatos_por_prioridad.extend((extra, "proveedor") for extra in (candidatos_extra or []))
    candidatos_por_prioridad.append((cif_buyer, "buyer"))

    rol_de = {}
    for cif, rol in candidatos_por_prioridad:
        if cif and cif != "-" and cif not in rol_de:
            rol_de[cif] = rol

    buyer_cif, nombre_empresa = "-", None
    proveedor_cif, nombre_proveedor = "-", None

    # 1) Resolver contra las tablas maestras sin importar el rol.
    for cif in rol_de:
        if nombre_empresa is None:
            empresa = buscar_empresa_por_cif(cif)
            if empresa:
                buyer_cif, nombre_empresa = cif, empresa
                continue

        if nombre_proveedor is None:
            proveedor = buscar_proveedor_por_cif(cif)
            if proveedor:
                proveedor_cif, nombre_proveedor = cif, proveedor

    # 2) Candidatos sin resolver: se colocan en el hueco de su rol.
    for cif, rol in rol_de.items():
        if cif == buyer_cif or cif == proveedor_cif:
            continue

        if rol == "buyer" and buyer_cif == "-":
            buyer_cif = cif
        elif rol == "proveedor" and proveedor_cif == "-":
            proveedor_cif = cif

    nombre_empresa_final = nombre_empresa or (MENSAJE_CIF_NO_ENCONTRADO if buyer_cif != "-" else "-")
    nombre_proveedor_final = nombre_proveedor or (MENSAJE_CIF_NO_ENCONTRADO if proveedor_cif != "-" else "-")

    return buyer_cif, nombre_empresa_final, proveedor_cif, nombre_proveedor_final


def combinar_csvs(lista_csv):
    """
    Une varios CSV individuales en una única tabla.
    Mantiene una sola cabecera.
    """
    todas = []

    for csv_text in lista_csv:
        rows = csv_to_matrix(csv_text)

        if len(rows) < 2:
            continue

        if not todas:
            todas.append(EXPECTED_HEADERS)

        todas.append(limpiar_fila(rows[1]))

    output = StringIO()
    writer = csv.writer(output, delimiter="|", lineterminator="\n")

    for row in todas:
        writer.writerow(row)

    return output.getvalue().strip()


# =========================================================
# LLAMADA AL MODELO
# =========================================================

# Esquema JSON estricto para la respuesta del modelo: con "strict": true la
# API garantiza que el objeto devuelto tiene EXACTAMENTE estas 18 claves (ni
# de menos ni de más), así que a diferencia del antiguo formato de texto
# separado por "|" es imposible que el modelo se salte un campo a mitad de
# la fila y desplace los siguientes.
FACTURA_JSON_SCHEMA = {
    "type": "object",
    "properties": {campo: {"type": "string"} for campo in EXPECTED_HEADERS},
    "required": EXPECTED_HEADERS,
    "additionalProperties": False,
}


def extract_invoice_with_agent(file_name, invoice_text, agent_prompt=DEFAULT_PROMPT, pdf_path=None):
    client = build_client()
    model = get_model()

    raw_text = clean_pdf_text(invoice_text)
    llm_text = pdf_a_markdown_factura(raw_text)
    MAX_CHARS = 35000
    if len(raw_text) > MAX_CHARS:

        mitad = MAX_CHARS // 2

        raw_text = (
                raw_text[:mitad]
                + "\n\n"
                + raw_text[-mitad:]
            )


    final_prompt = f"""
{agent_prompt}

NOMBRE DEL ARCHIVO:
\"\"\"
{file_name}
\"\"\"

FACTURA:
\"\"\"
{llm_text}
\"\"\"
"""

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
                    "Eres un extractor documental muy estricto. Devuelve los datos "
                    "de la factura en el objeto JSON solicitado, sin explicaciones "
                    "ni texto adicional."
                )
            },
            {
                "role": "user",
                "content": final_prompt
            }
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "factura",
                "schema": FACTURA_JSON_SCHEMA,
                "strict": True,
            },
        },
    )

    raw = response.choices[0].message.content or "{}"

    try:
        datos_json = json.loads(raw)
    except ValueError:
        datos_json = {}

    # Con response_format en modo "strict", el JSON Schema garantiza que la
    # API solo devuelve los 18 campos exactos (o falla la petición) — a
    # diferencia del antiguo formato de texto separado por "|", aquí es
    # estructuralmente imposible que el modelo se salte un campo y desplace
    # los siguientes (el bug que hacía que el CIF del proveedor terminara en
    # la columna del TipoIVA).
    #
    # Se trabaja por nombre de campo (dict), no por posición: así el orden
    # de EXPECTED_HEADERS se puede cambiar sin tener que revisar índices
    # numéricos a mano en todo este bloque; la lista posicional solo se
    # construye al final, para el CSV de salida.
    datos = {h: normalizar_valor(datos_json.get(h, "-")) for h in EXPECTED_HEADERS}

    datos["Archivo"] = file_name
    datos["FEscaneo"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    datos["FFactura"] = normalizar_fecha(datos["FFactura"])
    datos["FOperacion"] = normalizar_fecha(datos["FOperacion"])

    # FOperacion = FFactura
    if datos["FFactura"] != "-":
        datos["FOperacion"] = datos["FFactura"]

    # ImporIVA a 0 si BaseImp coincide con TotalFact (factura sin IVA)
    if datos["ImporIVA"] == "-" and datos["BaseImp"] != "-" and datos["TotalFact"] != "-":
        try:
            base = float(datos["BaseImp"].replace(".", "").replace(",", "."))
            total = float(datos["TotalFact"].replace(".", "").replace(",", "."))

            if abs(base - total) < 0.01:
                datos["ImporIVA"] = "0"
        except:
            pass

    # TipoIVA a 0 si ImporIVA es 0
    if datos["TipoIVA"] == "-" and datos["ImporIVA"] == "0":
        datos["TipoIVA"] = "0"

    # Moneda
    if datos["Moneda"] == "-":
        texto_up = raw_text.upper()

        if "€" in raw_text or "EUR" in texto_up:
            datos["Moneda"] = "EUR"

        elif "$" in raw_text or "USD" in texto_up:
            datos["Moneda"] = "USD"

    # Buyer/Proveedor + Empresa/NombreProveedor: el modelo extrae los dos
    # CIF/NIF/VAT de la factura, pero cuál es el comprador y cuál el
    # proveedor lo decide resolver_empresa_y_proveedor() consultando
    # EmpresasClasificadas/ProveedoresClasificados, no la posición/contexto
    # que haya usado el modelo para etiquetarlos (ver esa función). Si un
    # CIF no aparece o está inactivo/bloqueado en ninguna tabla, se deja en
    # "-" y clasificar_factura la mandará a incidencias.
    cif_1 = normalizar_cif(datos["Buyer"])
    cif_2 = normalizar_cif(datos["Proveedor"])

    # Respaldo determinista: además de los dos CIF que haya devuelto el
    # modelo (los use bien, los intercambie o directamente no encuentre
    # alguno), se buscan por regex todos los CIF con etiqueta clara en el
    # documento, más los que aparezcan sueltos justo después de una URL/web
    # (sin etiqueta CIF/NIF de por medio), que el modelo pudo pasar por
    # alto. resolver_empresa_y_proveedor ya ignora los candidatos "-" y los
    # que no encuentre en ninguna tabla, así que pasar estos de más no hace
    # daño cuando no aplican.
    candidatos_extra = detectar_cifs_con_etiqueta(raw_text) + detectar_cif_tras_url(raw_text)

    buyer_cif, nombre_empresa, proveedor_cif, nombre_proveedor = resolver_empresa_y_proveedor(
        cif_1, cif_2, candidatos_extra
    )

    # Último respaldo: si ni el modelo ni el regex han dado con el CIF del
    # comprador o del proveedor, es posible que esté metido en un logo,
    # sello o pegado sin etiqueta a una URL/teléfono — cosas que un modelo
    # mirando la imagen real de la factura reconoce mejor que leyendo el
    # texto ya aplanado. Solo se llama a vision (más caro y lento) cuando
    # de verdad hace falta, no en cada factura.
    if pdf_path and (buyer_cif == "-" or proveedor_cif == "-"):
        try:
            import imagenes as _imagenes_vision
            cif_vision_buyer, cif_vision_proveedor = _imagenes_vision.detectar_cif_con_vision_desde_pdf(pdf_path)
        except Exception as e:
            print(f"AVISO: fallo al intentar el respaldo por visión del CIF para {file_name}: {e}")
            cif_vision_buyer, cif_vision_proveedor = "-", "-"

        # La visión a veces "alucina": en vez de encontrar el CIF que
        # falta, repite el que ya se conoce del otro campo con un carácter
        # de más/menos. Se descarta ese eco en vez de dejarlo ensuciar el
        # campo con un CIF inventado.
        if _es_eco_corrupto(cif_vision_buyer, proveedor_cif) or _es_eco_corrupto(cif_vision_buyer, cif_2):
            cif_vision_buyer = "-"
        if _es_eco_corrupto(cif_vision_proveedor, buyer_cif) or _es_eco_corrupto(cif_vision_proveedor, cif_1):
            cif_vision_proveedor = "-"

        if (buyer_cif == "-" and cif_vision_buyer != "-") or (proveedor_cif == "-" and cif_vision_proveedor != "-"):
            cif_1_con_vision = cif_1 if buyer_cif != "-" else cif_vision_buyer
            candidatos_extra_con_vision = candidatos_extra + (
                [cif_vision_proveedor] if proveedor_cif == "-" and cif_vision_proveedor != "-" else []
            )
            buyer_cif, nombre_empresa, proveedor_cif, nombre_proveedor = resolver_empresa_y_proveedor(
                cif_1_con_vision, cif_2, candidatos_extra_con_vision
            )

    datos["Buyer"], datos["Empresa"], datos["Proveedor"], datos["NombreProveedor"] = (
        buyer_cif, nombre_empresa, proveedor_cif, nombre_proveedor
    )

    output = StringIO()
    writer = csv.writer(output, delimiter="|", lineterminator="\n")
    writer.writerow(EXPECTED_HEADERS)
    writer.writerow([datos[h] for h in EXPECTED_HEADERS])

    return output.getvalue().strip()


# =========================================================
# EXCEL EXPORT
# =========================================================

def export_to_excel(csv_text, path):
    rows = csv_to_matrix(csv_text)

    if len(rows) < 2:
        raise ValueError("No hay datos válidos para exportar.")

    headers = rows[0]
    data_rows = rows[1:]

    wb = Workbook()
    ws = wb.active
    ws.title = "Facturas"

    fill_header = PatternFill("solid", fgColor="DCE6F1")
    fill_empty = PatternFill("solid", fgColor="E7E6E6")

    border = Border(
        left=Side(style="thin", color="D9E2EC"),
        right=Side(style="thin", color="D9E2EC"),
        top=Side(style="thin", color="D9E2EC"),
        bottom=Side(style="thin", color="D9E2EC"),
    )

    for col_idx, name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font = Font(bold=True, color="0F2D52")
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row_idx, row in enumerate(data_rows, start=2):
        expanded = (row + [""] * len(headers))[:len(headers)]

        for col_idx, value in enumerate(expanded, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

            if value == "-":
                cell.fill = fill_empty
                cell.font = Font(color="666666")

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)

    ws.freeze_panes = "A2"

    last_col = ws.cell(row=1, column=len(headers)).column_letter
    last_row = len(data_rows) + 1

    table = Table(
        displayName="TablaFacturas",
        ref=f"A1:{last_col}{last_row}"
    )

    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    table.tableStyleInfo = style
    ws.add_table(table)

    wb.save(path)


# =========================================================
# HISTORIAL
# =========================================================

def _guardar_excel_atomico(wb, path):
    """Guarda un Workbook primero en un archivo temporal del mismo directorio
    y luego lo renombra sobre el destino con os.replace() (operación atómica
    en el mismo volumen). Estos .xlsx viven en una carpeta sincronizada con
    OneDrive y se reescriben muy a menudo; guardar directamente sobre el
    archivo final deja una ventana en la que OneDrive (u otro proceso) puede
    leerlo a medio escribir, corrompiendo el .zip interno del xlsx."""
    directorio = os.path.dirname(path) or "."
    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir=directorio)
    os.close(fd)
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, path)
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise


def guardar_historial(csv_text, user_id):
    os.makedirs(HISTORIAL_DIR, exist_ok=True)

    rows = csv_to_matrix(csv_text)

    if not rows or len(rows) < 2:
        print("AVISO: CSV invalido, no se guarda historial")
        return

    fecha = datetime.now().strftime("%Y-%m-%d")
    hora = datetime.now().strftime("%H:%M:%S")
    path = os.path.join(HISTORIAL_DIR, f"historial_facturas_{user_id}.xlsx")

    headers = rows[0]
    data_rows = rows[1:]

    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Historial"

    fill_bloque = PatternFill("solid", fgColor="BDD7EE")
    fill_header = PatternFill("solid", fgColor="DCE6F1")
    fill_empty = PatternFill("solid", fgColor="E7E6E6")

    border = Border(
        left=Side(style="thin", color="D9E2EC"),
        right=Side(style="thin", color="D9E2EC"),
        top=Side(style="thin", color="D9E2EC"),
        bottom=Side(style="thin", color="D9E2EC"),
    )

    start_row = ws.max_row + 1 if ws.max_row > 1 or ws["A1"].value else 1

    titulo = f"Extracción facturas | Fecha: {fecha} | Hora: {hora} | Nº facturas: {len(data_rows)}"
    ws.cell(row=start_row, column=1, value=titulo)

    total_cols = len(headers)

    ws.merge_cells(
        start_row=start_row,
        start_column=1,
        end_row=start_row,
        end_column=total_cols
    )

    title_cell = ws.cell(row=start_row, column=1)
    title_cell.font = Font(bold=True, color="0F2D52")
    title_cell.fill = fill_bloque
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    header_row = start_row + 1

    for col_idx, name in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=name)
        cell.font = Font(bold=True, color="0F2D52")
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    current_row = header_row + 1

    for row in data_rows:
        expanded = (row + [""] * len(headers))[:len(headers)]

        for col_idx, value in enumerate(expanded, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

            if value == "-":
                cell.fill = fill_empty
                cell.font = Font(color="666666")

        current_row += 1

    current_row += 1

    for col_idx in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0

        for cell in ws[col_letter]:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)

    _guardar_excel_atomico(wb, path)


def guardar_factura_corregida_completa(fila_completa, usuario):
    path = os.path.join(HISTORIAL_DIR, "facturas_corregidas.xlsx")

    headers = (
        EXPECTED_HEADERS
        + [
            "UsuarioUltimaModificacion",
            "FechaUltimaModificacion"
        ]
    )

    os.makedirs(HISTORIAL_DIR, exist_ok=True)

    if not os.path.exists(path):

        wb = Workbook()
        ws = wb.active
        ws.title = "FacturasCorregidas"

        ws.append(headers)

        _guardar_excel_atomico(wb, path)

    wb = load_workbook(path)
    ws = wb.active

    archivo = str(fila_completa[0]).strip()

    fila_objetivo = None

    for fila in range(2, ws.max_row + 1):

        archivo_excel = str(
            ws.cell(fila, 1).value or ""
        ).strip()

        if archivo_excel == archivo:

            fila_objetivo = fila
            break

    if fila_objetivo is None:

        fila_objetivo = ws.max_row + 1

    fila_completa = [
        "-" if str(x).strip() == "" else str(x).strip()
        for x in fila_completa
    ]

    while len(fila_completa) < len(EXPECTED_HEADERS):
        fila_completa.append("-")

    fila_completa = fila_completa[:len(EXPECTED_HEADERS)]

    for col, valor in enumerate(fila_completa, start=1):

        ws.cell(
            row=fila_objetivo,
            column=col
        ).value = valor

    ws.cell(
        row=fila_objetivo,
        column=len(EXPECTED_HEADERS) + 1
    ).value = usuario

    ws.cell(
        row=fila_objetivo,
        column=len(EXPECTED_HEADERS) + 2
    ).value = datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )

    _guardar_excel_atomico(wb, path)


# =========================================================
# EXCEL "100% DEFINITIVO"
# Espejo, en un Excel aparte, de las facturas de "Revisar facturas" que un
# humano ha marcado explícitamente como definitivas. FacturasExaminadas (SQL)
# es la fuente de verdad del estado (columna Definitiva); este Excel es solo
# una copia de conveniencia para exportar/consultar sin acceso a la BD.
# =========================================================

EXCEL_DEFINITIVO_PATH = os.path.join(HISTORIAL_DIR, "facturas_100_definitivas.xlsx")
HEADERS_DEFINITIVO = EXPECTED_HEADERS + ["Origen", "UsuarioDefinitiva", "FechaDefinitiva"]


def _abrir_o_crear_excel_definitivo():
    if os.path.exists(EXCEL_DEFINITIVO_PATH):
        wb = load_workbook(EXCEL_DEFINITIVO_PATH)
        return wb, wb.active

    wb = Workbook()
    ws = wb.active
    ws.title = "FacturasDefinitivas"
    ws.append(HEADERS_DEFINITIVO)
    return wb, ws


def agregar_o_actualizar_factura_definitiva(fila_completa, origen, usuario):
    os.makedirs(HISTORIAL_DIR, exist_ok=True)
    wb, ws = _abrir_o_crear_excel_definitivo()

    archivo = str(fila_completa[0]).strip()

    fila_objetivo = None
    for fila_idx in range(2, ws.max_row + 1):
        if str(ws.cell(fila_idx, 1).value or "").strip() == archivo:
            fila_objetivo = fila_idx
            break

    if fila_objetivo is None:
        fila_objetivo = ws.max_row + 1

    valores = ["-" if str(x).strip() == "" else str(x).strip() for x in fila_completa]
    while len(valores) < len(EXPECTED_HEADERS):
        valores.append("-")
    valores = valores[:len(EXPECTED_HEADERS)]

    for col, valor in enumerate(valores, start=1):
        ws.cell(row=fila_objetivo, column=col).value = valor

    ws.cell(row=fila_objetivo, column=len(EXPECTED_HEADERS) + 1).value = origen
    ws.cell(row=fila_objetivo, column=len(EXPECTED_HEADERS) + 2).value = usuario
    ws.cell(row=fila_objetivo, column=len(EXPECTED_HEADERS) + 3).value = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    _guardar_excel_atomico(wb, EXCEL_DEFINITIVO_PATH)


def quitar_factura_definitiva(archivo):
    if not os.path.exists(EXCEL_DEFINITIVO_PATH):
        return

    wb = load_workbook(EXCEL_DEFINITIVO_PATH)
    ws = wb.active

    for fila_idx in range(2, ws.max_row + 1):
        if str(ws.cell(fila_idx, 1).value or "").strip() == archivo:
            ws.delete_rows(fila_idx)
            _guardar_excel_atomico(wb, EXCEL_DEFINITIVO_PATH)
            return


# =========================================================
# REVISAR FACTURAS (facturas ya completadas, en FacturasExaminadas)
# =========================================================

HEADERS_FACTURAS_COMPLETADAS = EXPECTED_HEADERS + ["Origen", "Definitiva"]


def _es_definitiva(fila_dict):
    return str(fila_dict.get("Definitiva") or "0").strip() in ("1", "True", "true")


def listar_facturas_completadas():
    """[headers, *filas] de todas las facturas de FacturasExaminadas (SQL),
    para la pestaña "Revisar facturas". No incluye las pendientes de corregir."""
    filas_dict = listar_facturas_examinadas_sql()

    filas = []
    for d in filas_dict:
        fila = [str(d.get(c, "-") or "-") for c in EXPECTED_HEADERS]
        origen = str(d.get("Origen") or "-")
        definitiva = "Sí" if _es_definitiva(d) else "No"
        filas.append(fila + [origen, definitiva])

    return [HEADERS_FACTURAS_COMPLETADAS] + filas


def _mover_pdf_definitiva(archivo, carpeta_destino):
    """Mueve el PDF de `archivo` a `carpeta_destino` (buscándolo en cualquiera
    de las carpetas del flujo), evitando colisiones de nombre."""
    src = buscar_pdf_por_nombre(archivo)
    if src is None:
        raise FileNotFoundError(archivo)

    if os.path.normpath(os.path.dirname(src)) == os.path.normpath(carpeta_destino):
        return

    os.makedirs(carpeta_destino, exist_ok=True)
    dest = os.path.join(carpeta_destino, archivo)

    if os.path.exists(dest):
        base, ext = os.path.splitext(archivo)
        dest = os.path.join(carpeta_destino, f"{base}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}")

    shutil.move(src, dest)


def marcar_factura_definitiva(archivo, definitiva, usuario):
    """Marca/desmarca en SQL una factura examinada como 100% definitiva,
    refleja el cambio en el Excel definitivo aparte y mueve el PDF entre
    "completadas" y "facturas_revisadas"."""
    ok = marcar_factura_definitiva_sql(archivo, definitiva, usuario)
    if not ok:
        raise RuntimeError("No se pudo actualizar el estado en la base de datos.")

    if definitiva:
        filas_dict = listar_facturas_examinadas_sql()
        fila_dict = next((f for f in filas_dict if f.get("Archivo") == archivo), None)
        if fila_dict is None:
            raise FileNotFoundError(archivo)
        fila_completa = [fila_dict.get(c, "-") for c in EXPECTED_HEADERS]
        agregar_o_actualizar_factura_definitiva(fila_completa, fila_dict.get("Origen", "-"), usuario)
        _mover_pdf_definitiva(archivo, FACTURAS_REVISADAS_DIR)
    else:
        quitar_factura_definitiva(archivo)
        _mover_pdf_definitiva(archivo, COMPLETADAS_DIR)


def actualizar_factura_completada(fila_completa, origen, usuario):
    """Corrige los datos de una factura ya examinada (acción "Editar" en
    "Revisar facturas"). Si ya estaba marcada como definitiva, actualiza
    también su copia en el Excel definitivo para que no quede desfasada."""
    ok = guardar_factura_examinada_sql(fila_completa, origen)
    if not ok:
        raise RuntimeError("No se pudo actualizar la factura en la base de datos.")

    archivo = str(fila_completa[0]).strip()
    filas_dict = listar_facturas_examinadas_sql()
    fila_dict = next((f for f in filas_dict if f.get("Archivo") == archivo), None)
    if fila_dict and _es_definitiva(fila_dict):
        agregar_o_actualizar_factura_definitiva(fila_completa, origen, usuario)


def descartar_factura_completada(archivo):
    """Papelera de "Revisar facturas": quita la factura de FacturasExaminadas
    y del Excel definitivo (si estaba ahí), y mueve el PDF a no_es_factura."""
    ok = eliminar_factura_examinada_sql(archivo)
    if not ok:
        raise RuntimeError("No se pudo eliminar la factura de la base de datos.")

    quitar_factura_definitiva(archivo)

    src = buscar_pdf_por_nombre(archivo)
    if src is None:
        raise FileNotFoundError(archivo)

    os.makedirs(NO_FACTURA_DIR, exist_ok=True)
    dest = os.path.join(NO_FACTURA_DIR, archivo)

    if os.path.exists(dest):
        base, ext = os.path.splitext(archivo)
        dest = os.path.join(NO_FACTURA_DIR, f"{base}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}")

    shutil.move(src, dest)


CAMPOS_EXCLUIR_IMAGEN = {"Archivo", "FEscaneo", "ImporIVA"}

CAMPOS_OBLIGATORIOS = [
    "BaseImp",
    "Buyer",
    "Empresa",
    "FFactura",
    "ImporIVA",
    "Moneda",
    "NombreProveedor",
    "NumeroFactura",
    "PedidoCliente",
    "Proveedor",
    "TipoIVA",
    "TotalFact",
]


def clasificar_factura(fila):

    """
    Devuelve:

    completada → extracción correcta, todos los campos obligatorios presentes
    imagen     → PDF sin texto legible (todos los campos "-" salvo Archivo, FEscaneo e ImporIVA)
    manual     → extracción parcial, algún campo obligatorio falta
    """

    archivo = str(fila[0]).strip()

    # ==========================================
    # YA COMPLETADA (corregida manualmente)
    # ==========================================

    path_corregidas = os.path.join(HISTORIAL_DIR, "facturas_corregidas.xlsx")

    if os.path.exists(path_corregidas):

        wb = load_workbook(path_corregidas)
        ws = wb.active

        for row in range(2, ws.max_row + 1):

            archivo_excel = str(
                ws.cell(row, 1).value or ""
            ).strip()

            if archivo_excel == archivo:

                return "completada"

    datos = dict(zip(EXPECTED_HEADERS, fila))

    # ==========================================
    # IMAGEN: todos los campos son "-" salvo
    # Archivo, FEscaneo e ImporIVA
    # ==========================================

    todos_vacios = all(
        str(datos.get(campo, "-")).strip() == "-"
        for campo in EXPECTED_HEADERS
        if campo not in CAMPOS_EXCLUIR_IMAGEN
    )

    if todos_vacios:
        return "imagen"

    # ==========================================
    # INCIDENCIA: el comprador (Empresa) o el
    # proveedor (NombreProveedor) no se pudieron
    # resolver contra su base de datos (CIF no
    # encontrado, o encontrado pero inactivo/
    # bloqueado -> MENSAJE_CIF_NO_ENCONTRADO).
    # Estos dos campos ya no los extrae la API,
    # siempre vienen de ese cruce, así que si no
    # se resolvieron es un problema de datos
    # maestros, no de extracción.
    # ==========================================

    NO_RESUELTO = ("-", MENSAJE_CIF_NO_ENCONTRADO)

    if str(datos.get("Empresa", "-")).strip() in NO_RESUELTO or str(datos.get("NombreProveedor", "-")).strip() in NO_RESUELTO:
        return "incidencia"

    # ==========================================
    # Comprobar campos obligatorios (incluido
    # PedidoCliente: si falta, pasa por corregir
    # manualmente en vez de reenviarse solo)
    # ==========================================

    for campo in CAMPOS_OBLIGATORIOS:

        valor = str(datos.get(campo, "-")).strip()

        if valor == "-":
            return "manual"

    # ==========================================
    # COMPLETADA: extracción correcta
    # ==========================================

    return "completada"


PATRONES_ALBARAN_TITULO = [
    r"ALBAR[AÁ]N\s+DE\s+ENTREGA",
    r"NOTA\s+DE\s+ENTREGA",
    r"DELIVERY\s+NOTE",
    r"VALE\s+DE\s+ENTREGA",
    r"NOTA\s+DE\s+ENV[IÍ]O",
    r"PACKING\s+LIST",
    r"BON\s+DE\s+LIVRAISON",
    r"LIEFERSCHEIN",
    # Tickets de báscula/pesaje emitidos como "TICKET - ALBARAN"
    r"TICKET\s*-\s*ALBAR[AÁ]N",
    # Documentos de transporte (CMR) en varios idiomas — no son facturas
    r"CARTA\s+DE\s+PORTE",
    r"LETTRE\s+DE\s+VOITURE",
    r"VRACHTBRIEF",
    r"FRACHTBRIEF",
    r"DOCUMENTO\s+DE\s+CONTROL",
    r"BILL\s+OF\s+LOADING",
    r"OUTTAKE\s+ORDER",
    # Partes de horas / actividad — documentos internos de RRHH, no facturas
    r"PARTE\s+DE\s+HORAS",
]

PATRONES_FACTURA = [
    r"\bFACTURA\b",
    r"\bINVOICE\b",
    r"\bFATTURA\b",
    r"\bRECHNUNG\b",
    r"\bFACTURE\b",
]

# Detecta líneas del tipo: "Albarán: 1-000033 22/06/2026 PCMI26-24422"
# donde hay contenido adicional (código de pedido) en la MISMA línea tras la fecha.
# Las facturas que sólo referencian un albarán acaban en la fecha y no tienen
# nada más en esa misma línea → no deben detectarse como albarán.
_RE_ALBARAN_LINEA = re.compile(
    r"Albar[aá]n:\s+[\w\-\.]+\s+\d{2}/\d{2}/\d{4}[ \t]+\S",
    re.IGNORECASE,
)

# Tickets de báscula/pesaje sin título explícito (p.ej. albaranes de almazara):
# se reconocen por traer a la vez peso bruto, tara, neto y matrícula del vehículo.
_PATRONES_TICKET_PESAJE = [
    r"\bBRUTO\b",
    r"\bTARA\b",
    r"\bNETO\b",
    r"MATR[IÍ]CULA",
]


def es_no_factura(texto):
    if not texto:
        return False

    texto_up = texto.upper()
    cabecera = texto_up[:1500]

    # Regla 1: el documento se titula explícitamente como albarán/delivery note/
    # documento de transporte y NO como factura
    if any(re.search(p, cabecera) for p in PATRONES_ALBARAN_TITULO):
        if not any(re.search(p, cabecera) for p in PATRONES_FACTURA):
            return True

    # Regla 2: el cuerpo contiene "Albarán: <ref> <DD/MM/AAAA>"
    # — formato típico de documentos que describen un albarán como contenido
    if _RE_ALBARAN_LINEA.search(texto):
        return True

    # Regla 3: ticket de báscula/pesaje sin título reconocible — trae peso
    # bruto/tara/neto y matrícula, y no se autodenomina factura
    if all(re.search(p, texto_up) for p in _PATRONES_TICKET_PESAJE):
        if not any(re.search(p, cabecera) for p in PATRONES_FACTURA):
            return True

    return False


def mover_pdf(pdf_path, tipo):

    import time

    carpetas = {
        "completada":      COMPLETADAS_DIR,
        "manual":          os.path.join(FACTURAS_DIR, "corregir_manualmente"),
        "imagen":          os.path.join(FACTURAS_DIR, "imagenes"),
        "no_es_factura":   NO_FACTURA_DIR,
        "reenviar_pedido": REENVIAR_PEDIDO_DIR,
        "incidencia":      INCIDENCIAS_DIR,
        "error":           ERROR_DIR,
    }

    carpeta_destino = carpetas.get(tipo, ERROR_DIR)

    def _copiar_con_reintentos(origen, destino, intentos=10, espera=1):
        # El PDF origen puede estar en una carpeta sincronizada con OneDrive:
        # si en ese momento OneDrive lo está subiendo/descargando, la copia
        # falla con "WinError 32: en uso por otro proceso". Es transitorio,
        # así que se reintenta unas cuantas veces antes de darse por vencido.
        for intento in range(intentos):
            try:
                shutil.copy2(origen, destino)
                return
            except (PermissionError, OSError):
                if intento == intentos - 1:
                    raise
                time.sleep(espera)

    # Carpeta maestra: TODOS los PDF procesados van aquí siempre
    carpeta_procesadas = os.path.join(FACTURAS_DIR, "procesadas")
    os.makedirs(carpeta_procesadas, exist_ok=True)
    dest_procesadas = os.path.join(carpeta_procesadas, os.path.basename(pdf_path))
    _copiar_con_reintentos(pdf_path, dest_procesadas)
    print("PDF COPIADO A PROCESADAS:", dest_procesadas)

    # Carpeta específica según categoría
    os.makedirs(carpeta_destino, exist_ok=True)
    destino = os.path.join(carpeta_destino, os.path.basename(pdf_path))
    _copiar_con_reintentos(pdf_path, destino)
    print("PDF COPIADO A", tipo.upper() + ":", destino)

    for _ in range(10):

        try:

            if os.path.exists(pdf_path):

                os.remove(pdf_path)

                print("PDF ELIMINADO DE ENTRADA:", pdf_path)

                break

        except Exception:

            time.sleep(1)


def _extraer_y_clasificar_pdf(pdf_path, archivo):
    """Lee, extrae y clasifica un PDF, moviéndolo a la carpeta que
    corresponda según el resultado. Común a procesar_carpeta() (carpeta
    "entrada") y a reprocesar_error() (un PDF que ya estaba en ERROR_DIR).
    Devuelve el tipo de destino, o None si la extracción no devolvió datos
    (el LLM no lanzó excepción pero tampoco hay fila que clasificar)."""

    text = read_pdf_text(pdf_path)

    if es_no_factura(text):
        mover_pdf(pdf_path, "no_es_factura")
        return "no_es_factura"

    if len(text) < 80:
        mover_pdf(pdf_path, "imagen")
        return "imagen"

    result_csv = extract_invoice_with_agent(
        file_name=archivo,
        invoice_text=text,
        pdf_path=pdf_path,
    )

    tabla = csv_to_matrix(result_csv)

    if len(tabla) < 2:
        return None

    fila = tabla[1]
    tipo = clasificar_factura(fila)
    mover_pdf(pdf_path, tipo)
    guardar_historial(result_csv, "auto")

    if tipo == "completada":
        guardar_factura_examinada_sql(fila, "auto")

    eliminar_reenvios_de_factura_repetida(fila, archivo)

    return tipo


def procesar_carpeta():

    carpeta = os.path.join(FACTURAS_DIR, "entrada")

    resultados = []

    if not os.path.exists(carpeta):
        return resultados

    for archivo in os.listdir(carpeta):

        if not archivo.lower().endswith(".pdf"):
            continue

        pdf_path = os.path.join(
            carpeta,
            archivo
        )

        try:
            tipo = _extraer_y_clasificar_pdf(pdf_path, archivo)
            if tipo is None:
                continue

            resultados.append({
                "archivo": archivo,
                "estado": tipo
            })

        except Exception as e:

            print(f"ERROR procesando {archivo}: {e}")

            try:
                mover_pdf(pdf_path, "error")
            except Exception as e2:
                print(f"No se pudo mover {archivo} a error: {e2}")

            resultados.append({
                "archivo": archivo,
                "estado": f"ERROR: {e}"
            })

    return resultados


def reprocesar_error(archivo):
    """
    Reintenta la extracción de un PDF que quedó en ERROR_DIR (p.ej. un fallo
    puntual de la API). Si esta vez tiene éxito, el PDF se mueve a la carpeta
    que corresponda, igual que en el flujo automático normal. Si vuelve a
    fallar, el PDF se queda donde estaba (en ERROR_DIR) para clasificarlo a
    mano, y se informa del fallo para que la UI pueda avisar.

    Devuelve (exito, estado): con exito=True, estado es el tipo de destino;
    con exito=False, estado es el motivo del fallo.
    """
    pdf_path = os.path.join(ERROR_DIR, archivo)
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(archivo)

    try:
        tipo = _extraer_y_clasificar_pdf(pdf_path, archivo)
        if tipo is None:
            return False, "La extracción no devolvió datos."
        return True, tipo
    except Exception as e:
        print(f"ERROR reprocesando {archivo}: {e}")
        return False, str(e)


# =========================================================
# FLUJO MANUAL (corrección de pendientes)
# =========================================================

_indice_historial_cache = {"firma": None, "indice": {}}


def _indice_historial(directorio):
    """Índice {archivo: fila} de todos los historial_facturas_*.xlsx de
    `directorio`, cacheado en memoria. Se reconstruye solo si cambia el
    conjunto de ficheros o su fecha de modificación, para no tener que
    reabrir y re-escanear los Excel en cada búsqueda."""
    if not os.path.exists(directorio):
        return {}

    nombres = sorted(
        f for f in os.listdir(directorio)
        if f.startswith("historial_facturas_") and f.endswith(".xlsx")
    )
    firma = tuple((n, os.path.getmtime(os.path.join(directorio, n))) for n in nombres)

    if firma == _indice_historial_cache["firma"]:
        return _indice_historial_cache["indice"]

    indice = {}
    for nombre in nombres:
        hist_path = os.path.join(directorio, nombre)
        try:
            wb = load_workbook(hist_path, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                if not row or row[0] is None:
                    continue
                val = str(row[0]).strip()
                if val == "Archivo" or "Extracción facturas" in val:
                    continue
                # Última aparición gana: si una factura se reprocesa (mismo
                # nombre de archivo), debe mostrarse la extracción más
                # reciente, no la primera que se guardó.
                indice[val] = [str(c) if c is not None else "-" for c in row]
            wb.close()
        except Exception as e:
            print(f"Error leyendo historial {hist_path}: {e}")

    _indice_historial_cache["firma"] = firma
    _indice_historial_cache["indice"] = indice
    return indice


def _buscar_en_directorio_historial(archivo, directorio):
    return _indice_historial(directorio).get(archivo)


def buscar_en_historial(archivo):
    return _buscar_en_directorio_historial(archivo, HISTORIAL_DIR)


def cargar_pdf_pendiente_individual(archivo):
    """Devuelve (fila, fuente, es_no_factura):
    - fuente es 'historial' o 'api'.
    - fila es None si falla la extracción o si el PDF no parece una factura
      (en ese caso es_no_factura es True)."""
    fila = buscar_en_historial(archivo)
    if fila is not None:
        print(f"[historial] {archivo} encontrado en historial, sin llamada a la API")
        return limpiar_fila(fila), "historial", False

    ruta = os.path.join(CORREGIR_DIR, archivo)

    try:
        text = read_pdf_text(ruta)
    except Exception as e:
        print(f"ERROR leyendo {archivo}: {e}")
        return None, "api", False

    if es_no_factura(text):
        print(f"[no_es_factura] {archivo} no parece una factura, se omite la extracción")
        return None, "api", True

    print(f"[api] {archivo} no encontrado en historial, extrayendo con LLM")
    try:
        result_csv = extract_invoice_with_agent(file_name=archivo, invoice_text=text, pdf_path=ruta)
        tabla = csv_to_matrix(result_csv)
        if len(tabla) > 1:
            guardar_historial(result_csv, "manual")
            return tabla[1], "api", False
    except Exception as e:
        print(f"ERROR extrayendo {archivo}: {e}")

    return None, "api", False


def listar_pendientes_lista():
    if not os.path.exists(CORREGIR_DIR):
        return []
    return sorted(
        (f for f in os.listdir(CORREGIR_DIR) if f.lower().endswith(".pdf")),
        key=lambda f: (0 if f.startswith(PRIORIDAD_PREFIX) else 1, f.lower()),
    )


HEADERS_PENDIENTES_COMPLETO = EXPECTED_HEADERS + ["Revisada"]
HEADERS_INCIDENCIAS_COMPLETO = EXPECTED_HEADERS + ["Revisada"]


def _refrescar_buyer_proveedor(fila):
    """
    Vuelve a resolver Buyer/Empresa/Proveedor/NombreProveedor de `fila`
    (formato EXPECTED_HEADERS) contra el estado ACTUAL de
    EmpresasClasificadas/ProveedoresClasificados, en vez de quedarse con el
    valor que se guardó en el historial en el momento de la extracción.

    Esto es necesario porque "pendientes" e "incidencias" leen del
    historial en vez de volver a llamar al LLM, así que sin este refresco
    seguirían mostrando "-" en Empresa/NombreProveedor aunque las tablas
    maestras se actualizasen después (alta de una empresa/proveedor nuevo,
    o corrección de un CIF), y aunque el propio resolver ya sepa reubicar
    un Buyer/Proveedor que el modelo etiquetó al revés.

    También vuelve a buscar candidatos con etiqueta CIF/NIF leyendo el PDF
    de nuevo (sin llamar al LLM): así, si se mejora
    detectar_cifs_con_etiqueta más adelante, las incidencias ya atascadas
    por ese motivo se corrigen solas la próxima vez que se listan, en vez
    de quedarse ancladas al resultado de la extracción original.
    """
    i_buyer, i_empresa = EXPECTED_HEADERS.index("Buyer"), EXPECTED_HEADERS.index("Empresa")
    i_proveedor, i_nombreprov = EXPECTED_HEADERS.index("Proveedor"), EXPECTED_HEADERS.index("NombreProveedor")
    i_archivo = EXPECTED_HEADERS.index("Archivo")

    cif_1 = normalizar_cif(fila[i_buyer])
    cif_2 = normalizar_cif(fila[i_proveedor])

    candidatos_extra = []
    ruta_pdf = buscar_pdf_por_nombre(fila[i_archivo])
    if ruta_pdf:
        try:
            texto_pdf = clean_pdf_text(read_pdf_text(ruta_pdf))
            candidatos_extra = detectar_cifs_con_etiqueta(texto_pdf) + detectar_cif_tras_url(texto_pdf)
        except Exception as e:
            print(f"AVISO: no se pudo releer {fila[i_archivo]} para refrescar sus CIF: {e}")

    fila[i_buyer], fila[i_empresa], fila[i_proveedor], fila[i_nombreprov] = (
        resolver_empresa_y_proveedor(cif_1, cif_2, candidatos_extra)
    )
    return fila


def listar_pendientes_completo():
    """
    [headers, *filas] con los datos ya extraídos de cada factura pendiente de
    corregir, para pintar "Corregir manualmente" como tabla filtrable (igual
    que "todas las extracciones"). Reutiliza cargar_pdf_pendiente_individual,
    que casi siempre lee del historial ya existente; solo llama al LLM para
    PDFs subidos a mano que todavía no se hayan extraído nunca.

    Añade la columna "Revisada" (Sí/No): un simple marcador de que alguien ya
    la miró, independiente de si ya se ha completado o no.
    """
    revisados = archivos_revisados_sql()

    filas = []
    for archivo in listar_pendientes_lista():
        fila, _fuente, es_no_factura_flag = cargar_pdf_pendiente_individual(archivo)

        if fila is None:
            fila = [archivo] + ["-"] * (len(EXPECTED_HEADERS) - 1)
        else:
            fila = (list(fila) + ["-"] * len(EXPECTED_HEADERS))[:len(EXPECTED_HEADERS)]
            fila[0] = archivo
            fila = _refrescar_buyer_proveedor(fila)

        fila.append("Sí" if archivo in revisados else "No")
        filas.append(fila)

    return [HEADERS_PENDIENTES_COMPLETO] + filas


def listar_incidencias_lista():
    if not os.path.exists(INCIDENCIAS_DIR):
        return []
    return sorted(f for f in os.listdir(INCIDENCIAS_DIR) if f.lower().endswith(".pdf"))


def listar_incidencias_completo():
    """
    [headers, *filas] con los datos ya extraídos de cada factura en
    incidencias (comprador o proveedor sin resolver contra su base de
    datos), igual que listar_pendientes_completo. A diferencia de
    "corregir_manualmente", aquí no hace falta llamar al LLM: la factura ya
    se extrajo y se guardó en el historial antes de clasificarla como
    incidencia, así que basta con leerla de ahí.

    Añade también la columna "Revisada" (Sí/No), igual que listar_pendientes_completo.
    """
    revisados = archivos_revisados_sql()

    filas = []
    for archivo in listar_incidencias_lista():
        fila = buscar_en_historial(archivo)

        if fila is None:
            fila = [archivo] + ["-"] * (len(EXPECTED_HEADERS) - 1)
        else:
            fila = limpiar_fila(list(fila))
            fila[0] = archivo
            fila = _refrescar_buyer_proveedor(fila)

        fila.append("Sí" if archivo in revisados else "No")
        filas.append(fila)

    return [HEADERS_INCIDENCIAS_COMPLETO] + filas


def marcar_revisada(archivo, revisada, usuario):
    """Marca/desmarca en SQL una factura de "Corregir manualmente" o
    "Incidencias" como revisada por una persona."""
    ok = marcar_revisada_sql(archivo, revisada, usuario)
    if not ok:
        raise RuntimeError("No se pudo actualizar el estado de revisión en la base de datos.")


def confirmar_y_mover_factura(archivo, fila_completa, usuario):
    """Corregir manualmente ya es la revisión humana de la factura, así que
    al confirmar se marca directamente como definitiva (sin pasar por
    "completadas"): se guarda en el historial/SQL y el PDF se mueve
    directamente a facturas_revisadas."""
    guardar_factura_corregida_completa(fila_completa=fila_completa, usuario=usuario)
    guardar_factura_examinada_sql(fila_completa, "manual")
    marcar_factura_definitiva(archivo, True, usuario)


def descartar_pendiente(archivo):
    """
    "Papelera" de corregir_manualmente (y de incidencias): para documentos
    que la extracción clasificó como pendientes de corrección o como
    incidencia pero que en realidad no son una factura (un albarán, un
    ticket, etc. que se coló). No se borra el PDF, se mueve a la carpeta
    "no_es_factura" por si hubiera que revisarlo más tarde.
    """
    src = buscar_pdf_por_nombre(archivo)

    if src is None:
        raise FileNotFoundError(archivo)

    os.makedirs(NO_FACTURA_DIR, exist_ok=True)
    dest = os.path.join(NO_FACTURA_DIR, archivo)

    if os.path.exists(dest):
        base, ext = os.path.splitext(archivo)
        dest = os.path.join(NO_FACTURA_DIR, f"{base}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}")

    shutil.move(src, dest)


def listar_reenviar_pedido():
    """
    Lista de solo lectura de REENVIAR_PEDIDO_DIR. La app nunca mueve nada de
    aquí a REENVIADAS_PEDIDO_DIR: ese traslado es responsabilidad exclusiva
    del flujo de Power Automate una vez gestiona (envía y archiva) la
    solicitud, para evitar que la app y Power Automate compitan moviendo el
    mismo archivo.
    """
    if not os.path.exists(REENVIAR_PEDIDO_DIR):
        return []
    return sorted(f for f in os.listdir(REENVIAR_PEDIDO_DIR) if f.lower().endswith(".pdf"))


# =========================================================
# SOLICITUDES DE ENVÍO DE CORREO POR OTROS MOTIVOS
# (fallo del reenvío automático, falta de otro dato obligatorio,
# varias facturas en un mismo PDF, u otro motivo redactado a mano).
#
# A diferencia de "reenviar_pedido", aquí el PDF puede venir de
# cualquier punto del flujo (pendiente de corregir, ya completada...),
# así que se localiza por nombre y se COPIA (no se mueve) a la carpeta
# de salida, dejando intacto el original en su carpeta de origen.
# El destinatario se reutiliza del marcador __EMAIL__...__ENDMAIL__ que
# ya trae el nombre del archivo desde la ingesta original; el motivo se
# codifica en el nombre con el mismo estilo para que el flujo de Power
# Automate que vigile esta carpeta pueda redactar el correo adecuado.
# =========================================================

_CARACTERES_INVALIDOS_ARCHIVO = re.compile(r'[\\/:*?"<>|\r\n]+')


def _sanear_motivo_para_archivo(texto):
    texto = _CARACTERES_INVALIDOS_ARCHIVO.sub(" ", texto).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto[:120]


def _carpetas_busqueda_pdf():
    return [
        CORREGIR_DIR,
        COMPLETADAS_DIR,
        FACTURAS_REVISADAS_DIR,
        NO_FACTURA_DIR,
        ERROR_DIR,
        INCIDENCIAS_DIR,
        REENVIAR_PEDIDO_DIR,
        REENVIADAS_PEDIDO_DIR,
        REENVIAR_DOS_FACTURAS_DIR,
        REENVIAR_ERROR_PESA_MUCHO_DIR,
        REENVIAR_ERROR_OTRO_DIR,
        # "procesadas" guarda una copia de auditoría de TODO lo que pasa por
        # mover_pdf(), así que cualquier archivo que haya sido clasificado
        # alguna vez tiene ahí una copia aunque ya no esté en esa cola. Debe
        # mirarse en último lugar: si se buscara antes que la carpeta real,
        # "No es factura"/reenvíos moverían esa copia de auditoría en vez
        # del PDF que de verdad está pendiente, y este reaparecería en su
        # cola original al recargar.
        os.path.join(FACTURAS_DIR, "procesadas"),
    ]


# REENVIAR_OTRO_MOTIVO_DIR y REENVIADAS_OTRO_MOTIVO_DIR ya no son planas: las
# solicitudes de "falta_dato_obligatorio" se guardan en una subcarpeta por
# campo (una por cada valor de CAMPOS_OBLIGATORIOS_FACTURA), así que hay que
# recorrerlas recursivamente en vez de mirar solo el nivel superior.
def _carpetas_busqueda_pdf_recursiva():
    return [REENVIAR_OTRO_MOTIVO_DIR, REENVIADAS_OTRO_MOTIVO_DIR]


def buscar_pdf_por_nombre(archivo):
    """Busca `archivo` en todas las carpetas del flujo y devuelve su ruta completa,
    o None si no se encuentra en ninguna."""
    for carpeta in _carpetas_busqueda_pdf():
        ruta = os.path.join(carpeta, archivo)
        if os.path.exists(ruta):
            return ruta

    for carpeta in _carpetas_busqueda_pdf_recursiva():
        if not os.path.exists(carpeta):
            continue
        for raiz, _, archivos in os.walk(carpeta):
            if archivo in archivos:
                return os.path.join(raiz, archivo)

    return None


_CARPETA_POR_MOTIVO_CORREO = {
    "falta_pedido_cliente":    REENVIAR_PEDIDO_DIR,
    "dos_facturas_una_pagina": REENVIAR_DOS_FACTURAS_DIR,
    "otros":                   REENVIAR_OTRO_MOTIVO_DIR,
}


#Motivos en los que Power Automate se queda con la factura a partir de aquí
# (la gestiona y archiva por su cuenta), así que el PDF se MUEVE fuera de su
# cola actual (corregir_manualmente, completadas...) para que no siga
# apareciendo ahí como pendiente. "otros" no tiene flujo automático detrás
# (lo redacta una persona a mano), así que ese se copia y el original se
# queda donde estaba.
_MOTIVOS_QUE_MUEVEN = {"falta_pedido_cliente", "dos_facturas_una_pagina"}


def solicitar_envio_correo(archivo, motivo, motivo_otro=None):
    """
    Traslada el PDF a la carpeta correspondiente al motivo elegido,
    codificando el motivo en el nombre de archivo para que el flujo de Power
    Automate que vigile esa carpeta pueda redactar el correo adecuado.
    Devuelve el nombre final generado.
    """
    if motivo not in MOTIVOS_ENVIO_CORREO:
        raise ValueError(f"Motivo no reconocido: {motivo}")

    carpeta_destino = _CARPETA_POR_MOTIVO_CORREO[motivo]

    if motivo == "otros":
        texto_motivo = (motivo_otro or "").strip()
        if not texto_motivo:
            raise ValueError("Debes redactar el motivo cuando seleccionas 'Otro'.")
    else:
        texto_motivo = MOTIVOS_ENVIO_CORREO[motivo]

    texto_motivo = _sanear_motivo_para_archivo(texto_motivo)

    src = buscar_pdf_por_nombre(archivo)
    if src is None:
        raise FileNotFoundError(archivo)

    base, ext = os.path.splitext(archivo)
    nuevo_nombre = f"{base}__MOTIVOENVIO__{texto_motivo}__ENDMOTIVOENVIO__{ext}"

    os.makedirs(carpeta_destino, exist_ok=True)
    dest = os.path.join(carpeta_destino, nuevo_nombre)

    if os.path.exists(dest):
        sufijo = datetime.now().strftime("%Y%m%d%H%M%S")
        nuevo_nombre = f"{base}__MOTIVOENVIO__{texto_motivo}__ENDMOTIVOENVIO__{sufijo}{ext}"
        dest = os.path.join(carpeta_destino, nuevo_nombre)

    if motivo in _MOTIVOS_QUE_MUEVEN:
        shutil.move(src, dest)
    else:
        shutil.copy2(src, dest)

    return nuevo_nombre


_RE_EMAIL_MARCADOR = re.compile(r"__EMAIL__(.*?)__ENDMAIL__")


def extraer_email_de_nombre(archivo):
    """Recupera el email de quien envió la factura a partir del marcador
    __EMAIL__...__ENDMAIL__ que trae el nombre desde la ingesta original
    (Power Automate). Devuelve None si el archivo no lo trae."""
    m = _RE_EMAIL_MARCADOR.search(archivo)
    return m.group(1) if m else None


def listar_errores():
    """Lista de solo lectura de ERROR_DIR: PDFs cuya extracción falló y que
    todavía no se han clasificado (fichero muy grande / otro motivo)."""
    if not os.path.exists(ERROR_DIR):
        return []
    return sorted(f for f in os.listdir(ERROR_DIR) if f.lower().endswith(".pdf"))


HEADERS_ERRORES = ["Archivo", "FechaError"]


def listar_errores_completo():
    """[headers, *filas] de ERROR_DIR, para mostrarla como tabla filtrable
    (igual que "Corregir manualmente"/"Incidencias") en vez de una simple
    lista. No hay datos de factura que mostrar -la extracción no llegó a
    completarse-, así que la única columna aparte del archivo es la fecha
    en la que quedó en error (mtime del PDF)."""
    filas = []
    for archivo in listar_errores():
        ruta = os.path.join(ERROR_DIR, archivo)
        try:
            fecha = datetime.fromtimestamp(os.path.getmtime(ruta)).strftime("%d/%m/%Y %H:%M:%S")
        except OSError:
            fecha = "-"
        filas.append([archivo, fecha])
    return [HEADERS_ERRORES] + filas


def clasificar_error(archivo, motivo):
    """
    Clasifica un PDF de ERROR_DIR y lo mueve a la carpeta correspondiente:
    - "pesa_mucho" -> REENVIAR_ERROR_PESA_MUCHO_DIR
    - "otro"       -> REENVIAR_ERROR_OTRO_DIR (para que quede constancia;
      el correo en sí se redacta a mano con el PDF adjunto).
    Devuelve el email (si el nombre lo trae) para poder prellenar el "Para"
    del correo cuando el motivo es "otro".
    """
    if motivo not in MOTIVOS_ERROR_EXTRACCION:
        raise ValueError(f"Motivo no reconocido: {motivo}")

    carpeta_destino = (
        REENVIAR_ERROR_PESA_MUCHO_DIR if motivo == "pesa_mucho" else REENVIAR_ERROR_OTRO_DIR
    )

    src = os.path.join(ERROR_DIR, archivo)
    if not os.path.exists(src):
        raise FileNotFoundError(archivo)

    os.makedirs(carpeta_destino, exist_ok=True)
    dest = os.path.join(carpeta_destino, archivo)

    if os.path.exists(dest):
        base, ext = os.path.splitext(archivo)
        dest = os.path.join(carpeta_destino, f"{base}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}")

    shutil.move(src, dest)
    return extraer_email_de_nombre(archivo)


def listar_solicitudes_envio_correo():
    """
    Lista de solo lectura de las solicitudes de correo que no son "falta
    pedido cliente" (esa ya tiene su propia lista en REENVIAR_PEDIDO_DIR):
    "dos facturas en una página" y "otro". Igual que con reenviar_pedido, la
    app nunca mueve nada de aquí a REENVIADAS_OTRO_MOTIVO_DIR: el archivo
    desaparece de esta lista solo cuando Power Automate lo procesa (envía el
    correo) y lo archiva por su cuenta.
    """
    archivos = []
    for carpeta in (REENVIAR_DOS_FACTURAS_DIR, REENVIAR_OTRO_MOTIVO_DIR):
        if not os.path.exists(carpeta):
            continue
        archivos.extend(f for f in os.listdir(carpeta) if f.lower().endswith(".pdf"))
    return sorted(archivos)


# =========================================================
# REENVIADAS: contador unificado y detalle (email/motivo/fecha) de las
# facturas ya reenviadas y archivadas por Power Automate en las carpetas
# "reenviadas_*" (no las colas "reenviar_*", que siguen pendientes).
# =========================================================

def _contar_pdfs_recursivo(carpeta):
    """Cuenta los .pdf de `carpeta` incluidas las subcarpetas (necesario para
    REENVIADAS_OTRO_MOTIVO_DIR, que guarda una subcarpeta por campo)."""
    if not os.path.exists(carpeta):
        return 0
    total = 0
    for _, _, archivos in os.walk(carpeta):
        total += sum(1 for f in archivos if f.lower().endswith(".pdf"))
    return total


def contar_reenviadas():
    """Total de facturas ya reenviadas y archivadas, sumando las dos carpetas
    'reenviadas_*'. No incluye las colas 'reenviar_*' (todavía pendientes de
    que Power Automate las procese), que ya tienen su propio aviso."""
    return _contar_pdfs_recursivo(REENVIADAS_PEDIDO_DIR) + _contar_pdfs_recursivo(REENVIADAS_OTRO_MOTIVO_DIR)


_RE_MOTIVO_MARCADOR = re.compile(r"__MOTIVOENVIO__(.*?)__ENDMOTIVOENVIO__")


def _archivo_original_de_reenvio(nombre):
    """Recupera el nombre de archivo tal y como se guardó en el historial
    (sin el marcador __MOTIVOENVIO__...__ENDMOTIVOENVIO__ que se añade al
    copiarlo/moverlo a una carpeta de reenvío)."""
    return _RE_MOTIVO_MARCADOR.sub("", nombre)


def _motivo_de_nombre_reenvio(nombre, motivo_por_defecto=None):
    m = _RE_MOTIVO_MARCADOR.search(nombre)
    return m.group(1) if m else motivo_por_defecto


def listar_reenviadas_detalle():
    """[{archivo, archivo_real, email, motivo, fecha}] de las facturas ya
    reenviadas y archivadas (carpetas 'reenviadas_*'), para el panel de
    detalle de la visualización automática. La fecha es la de archivado
    (mtime), que es cuando Power Automate movió el PDF aquí tras enviar el
    correo. "archivo" es el nombre limpio para mostrar; "archivo_real" es el
    nombre tal cual está en disco (con marcadores), necesario para abrir el
    PDF vía /pdf-factura/{archivo}."""
    filas = []

    def _agregar(carpeta, motivo_por_defecto):
        if not os.path.exists(carpeta):
            return
        for raiz, _, archivos in os.walk(carpeta):
            for nombre in archivos:
                if not nombre.lower().endswith(".pdf"):
                    continue
                ruta = os.path.join(raiz, nombre)
                archivo_limpio = _RE_EMAIL_MARCADOR.sub("", _archivo_original_de_reenvio(nombre))
                try:
                    fecha = datetime.fromtimestamp(os.path.getmtime(ruta)).strftime("%d/%m/%Y %H:%M:%S")
                except OSError:
                    fecha = "-"
                filas.append({
                    "archivo": archivo_limpio,
                    "archivo_real": nombre,
                    "email": extraer_email_de_nombre(nombre),
                    "motivo": _motivo_de_nombre_reenvio(nombre, motivo_por_defecto),
                    "fecha": fecha,
                })

    _agregar(REENVIADAS_PEDIDO_DIR, MOTIVOS_ENVIO_CORREO["falta_pedido_cliente"])
    _agregar(REENVIADAS_OTRO_MOTIVO_DIR, None)

    filas.sort(key=lambda f: f["fecha"], reverse=True)
    return filas


# =========================================================
# FACTURA REENVIADA QUE VUELVE A LLEGAR
# Si el proveedor reenvía por su cuenta (fuera de este flujo) la misma
# factura que ya estaba pendiente/archivada en una carpeta de reenvío, esa
# copia antigua deja de tener sentido: se borra para que no quede duplicada
# ni siga contando como pendiente de reenvío.
#
# Esto es una excepción deliberada a la regla de que la app nunca toca las
# carpetas reenviar_*/reenviadas_* (ver comentarios de listar_reenviar_pedido
# y listar_solicitudes_envio_correo): aquí no se compite con Power Automate
# por mover el archivo, se borra una copia que ya no hace falta porque la
# factura ha vuelto a llegar por su cuenta.
# =========================================================

_CARPETAS_REENVIO_CON_DATOS = [
    (REENVIAR_PEDIDO_DIR, False),
    (REENVIADAS_PEDIDO_DIR, False),
    (REENVIAR_DOS_FACTURAS_DIR, False),
    (REENVIAR_OTRO_MOTIVO_DIR, True),
    (REENVIADAS_OTRO_MOTIVO_DIR, True),
]

# En estas dos no hay NumeroFactura disponible: la extracción nunca llegó a
# completarse (por eso el PDF terminó en error), así que solo se puede
# comparar por nombre de archivo.
_CARPETAS_REENVIO_SIN_DATOS = [REENVIAR_ERROR_PESA_MUCHO_DIR, REENVIAR_ERROR_OTRO_DIR]


def eliminar_reenvios_de_factura_repetida(fila, archivo_entrante):
    """
    Si la factura recién extraída (`fila`, en el mismo orden que
    EXPECTED_HEADERS) ya tenía una copia esperando/archivada en alguna
    carpeta de reenvío, la borra de ahí. Se identifica por NumeroFactura +
    Proveedor (más fiable que el nombre de archivo, que puede cambiar si el
    proveedor reenvía con un fichero distinto).

    Devuelve la lista de rutas borradas.
    """
    borrados = []

    numero_factura = fila[1] if len(fila) > 1 else None
    proveedor = fila[4] if len(fila) > 4 else None

    if numero_factura and numero_factura != "-":
        for carpeta, recursiva in _CARPETAS_REENVIO_CON_DATOS:
            if not os.path.exists(carpeta):
                continue
            paseo = os.walk(carpeta) if recursiva else [(carpeta, [], os.listdir(carpeta))]
            for raiz, _, archivos in paseo:
                for nombre in archivos:
                    if not nombre.lower().endswith(".pdf"):
                        continue
                    fila_existente = buscar_en_historial(_archivo_original_de_reenvio(nombre))
                    if not fila_existente or len(fila_existente) <= 4:
                        continue
                    if fila_existente[1] == numero_factura and fila_existente[4] == proveedor:
                        ruta = os.path.join(raiz, nombre)
                        try:
                            os.remove(ruta)
                            borrados.append(ruta)
                        except OSError as e:
                            print(f"AVISO: no se pudo borrar {ruta} tras detectar reenvío repetido: {e}")

    for carpeta in _CARPETAS_REENVIO_SIN_DATOS:
        ruta = os.path.join(carpeta, archivo_entrante)
        if not os.path.exists(ruta):
            continue
        try:
            os.remove(ruta)
            borrados.append(ruta)
        except OSError as e:
            print(f"AVISO: no se pudo borrar {ruta} tras detectar reenvío repetido: {e}")

    return borrados


# =========================================================
# HISTORIAL COMPLETO (todo lo procesado, sin filtrar por fecha)
# =========================================================

def leer_historial_bloques(path):
    """
    Lee un Excel con el formato de guardar_historial(): bloques repetidos de
    [título de lote] [cabecera] [una o varias filas de datos]. La cabecera
    puede variar de un bloque a otro si el esquema de campos ha cambiado con
    el tiempo, así que se relee en cada bloque.
    Devuelve una lista de dicts {columna: valor}.
    """
    if not os.path.exists(path):
        return []

    wb = load_workbook(path, read_only=True)
    ws = wb.active

    filas = []
    cabecera_actual = None

    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue

        primera = str(row[0]).strip()

        if primera.startswith("Extracci"):  # "Extracción facturas | Fecha: ..."
            cabecera_actual = None
            continue

        if primera == "Archivo":
            cabecera_actual = [str(c).strip() if c is not None else "" for c in row]
            continue

        if cabecera_actual is None:
            continue

        filas.append(dict(zip(cabecera_actual, row)))

    wb.close()
    return filas


def leer_historial_plano(path):
    """Lee un Excel de tabla simple (una cabecera, filas de datos).
    Devuelve una lista de dicts {columna: valor}."""
    if not os.path.exists(path):
        return []

    wb = load_workbook(path, read_only=True)
    ws = wb.active

    filas_iter = ws.iter_rows(values_only=True)
    try:
        cabecera = [str(c).strip() if c is not None else "" for c in next(filas_iter)]
    except StopIteration:
        wb.close()
        return []

    filas = []
    for row in filas_iter:
        if not row or row[0] is None:
            continue
        filas.append(dict(zip(cabecera, row)))

    wb.close()
    return filas


def fila_desde_dict(datos):
    fila = [
        "-" if datos.get(campo) is None else str(datos.get(campo))
        for campo in EXPECTED_HEADERS
    ]
    return limpiar_fila(fila)


def facturas_definitivas_tabla():
    """[headers, *filas] leyendo el Excel "100% definitivo" ya escrito en
    disco, para poder ofrecerlo como descarga con el formato/estilo habitual."""
    filas_dict = leer_historial_plano(EXCEL_DEFINITIVO_PATH)
    filas = [[str(d.get(c, "-") or "-") for c in HEADERS_DEFINITIVO] for d in filas_dict]
    return [HEADERS_DEFINITIVO] + filas


def tabla_a_pipe_csv(tabla):
    output = StringIO()
    writer = csv.writer(output, delimiter="|", lineterminator="\n")
    for row in tabla:
        writer.writerow(row)
    return output.getvalue().strip()
