# Extractor de Facturas MIGASA

Sistema que extrae automáticamente los datos estructurados de facturas en PDF usando IA (OpenAI GPT-4.1 o compatible). Es una única aplicación FastAPI con tres vistas (Visualización automática, Corregir manualmente, Subir facturas) más un par de utilidades:

- **`main.py`** — app FastAPI única: vigila una carpeta de entrada, expone el endpoint que usa Power Automate, y sirve la interfaz web con las tres vistas.
- **`logic.py`** — extracción de PDF, llamada al LLM, clasificación, Excel, historial. Toda la lógica de negocio consolidada en un solo módulo.
- **`routers/`** — endpoints agrupados por área: `auto.py` (procesamiento automático) y `manual.py` (corrección manual, subida, historial).
- **`imagenes.py`** — segunda pasada con visión (GPT vision) para los PDFs escaneados o fotografiados que no tienen texto extraíble.
- **`sql_historial.py`** — guarda cada factura "examinada" (ya sea automática o corregida a mano) en una base de datos (SQL Server o SQLite), como complemento del histórico en Excel.
- **`migrar_historial_excel.py`** — migración puntual y segura de re-ejecutar del histórico Excel acumulado hacia SQL Server.

---

## Funcionalidades

- **Extracción automática** de datos de factura PDF mediante LLM (OpenAI GPT-4.1 o compatible)
- **Soporte multiidioma**: detecta facturas en español, inglés, portugués, francés, italiano, alemán, neerlandés, chino, griego y otros
- **Vigilancia automática de carpeta**: sondea la carpeta `entrada` cada `INTERVALO_VIGILANCIA` segundos y procesa lo que encuentra, sin intervención manual
- **Endpoint de subida** (`/upload-pdf`) pensado para integrarse con **Power Automate**
- **Detección de "no es factura"**: los documentos que no son facturas (albaranes, tickets de báscula/pesaje, documentos de transporte/CMR, partes de horas...) se identifican y se apartan del flujo, tanto en el procesamiento automático como al abrir un pendiente en "Corregir manualmente"
- **Segunda pasada con visión** (`imagenes.py`) para PDFs escaneados/fotografiados sin texto extraíble, reutilizando toda la lógica de extracción, clasificación e historial de `logic.py`
- **Clasificación automática de PDFs** en carpetas según el resultado de la extracción:
  - `procesadas` — copia maestra de todo lo que pasa por el flujo automático
  - `examinadas` — extracción correcta (automática o corregida manualmente)
  - `corregir_manualmente` — falta algún campo obligatorio
  - `imagenes` — PDF sin texto extraíble, pendiente de la segunda pasada con visión
  - `imagenes_sin_datos` — tampoco se pudo leer con visión
  - `no_es_factura` — el documento no es una factura (albarán, ticket de báscula/pesaje, documento de transporte/CMR, parte de horas, etc.)
  - `reenviar_falta_pedidocliente` — todo correcto salvo el número de pedido de cliente
  - `error` — fallo de procesamiento
- **Corrección manual**: vista web para completar los campos que faltan en `corregir_manualmente`, confirmar y mover la factura a `examinadas`
- **Subida manual**: vista para añadir una factura directamente a la cola de corrección manual, con opción de marcarla como urgente (prioridad en la lista)
- **Historial diario en Excel**, acumulado por fecha con marca de hora
- **Persistencia en base de datos** de las facturas "examinada" (tabla `FacturasExaminadas`), con upsert por nombre de archivo — no rompe el flujo si la conexión falla

---

## Campos extraídos

| Campo | Descripción |
|---|---|
| Archivo | Nombre del archivo PDF |
| BaseImp | Base imponible |
| BaseIRPF | Base IRPF (si aplica) |
| Buyer | CIF/NIF/VAT del comprador |
| Empresa | Razón social del comprador |
| FEscaneo | Fecha y hora de procesamiento |
| FFactura | Fecha de la factura |
| FOperacion | Fecha de operación |
| ImporIVA | Importe del IVA |
| Moneda | Moneda (EUR, USD, GBP…) |
| NombreProveedor | Razón social del proveedor |
| NumeroFactura | Número de factura |
| PedidoCliente | Número de pedido del cliente |
| Proveedor | CIF/NIF/VAT del proveedor |
| TipoIVA | Tipo de IVA principal |
| TipoIVA2 | Segundo tipo de IVA |
| TipoIVA3 | Tercer tipo de IVA |
| TotalFact | Importe total de la factura |

---

## Flujo de procesamiento

```
Power Automate ──POST /upload-pdf──▶ main.py
                                        │
                       ¿no es factura?──┴──▶ facturas/no_es_factura
                                        │
                       ¿sin texto (<80 car.)?──▶ facturas/imagenes ──▶ imagenes.py (visión)
                                        │                                      │
                              extracción con LLM                    clasifica igual que el flujo automático
                                        │                                      │
                                 clasificar_factura                           mueve a examinadas /
                                        │                                corregir_manualmente / etc.
                    ┌───────────────────┼────────────────────┐
                    ▼                   ▼                    ▼
              examinadas       corregir_manualmente      reenviar_falta_pedidocliente / error
                    │                   │
          guardar_historial      vista "Corregir manualmente" revisa y confirma
                    │                   │
       guardar_factura_examinada_sql ◀──┘
```

La misma app también vigila la carpeta `entrada` de forma periódica (no solo vía Power Automate), procesando cualquier PDF que aparezca ahí.

---

## Requisitos

- Python 3.10+
- API key de OpenAI (o endpoint compatible); con soporte de visión si se va a usar `imagenes.py`
- Opcional: SQL Server accesible + ODBC Driver 18, o usar el motor `sqlite` (sin instalación adicional) mientras no haya acceso a SQL Server

---

## Instalación

```bash
# Clonar el repositorio
git clone <url-del-repo>
cd Extraer-texto-facturas

# Crear entorno virtual
python -m venv .venv
.venv\Scripts\activate      # Windows
# source .venv/bin/activate  # Linux/Mac

# Instalar dependencias
pip install -r requirements.txt
```

---

## Configuración

Crear un fichero `.env` en la raíz del proyecto:

```env
# --- OpenAI ---
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4.1
# Opcional: para usar un endpoint alternativo (Azure OpenAI, proxy, etc.)
# OPENAI_BASE_URL=https://...

# --- Rutas y vigilancia ---
FACTURAS_DIR=C:\ruta\a\facturas          # carpeta base con entrada/procesadas/examinadas/...
INTERVALO_VIGILANCIA=30                  # segundos entre pasadas del vigilante

# --- Segunda pasada con visión (imagenes.py) ---
# MAX_PAGINAS_IMAGENES=3
# DPI_IMAGENES=200
# INTERVALO_VIGILANCIA_IMAGENES=300      # solo si se lanza con --watch

# --- Base de datos (guardado de facturas "examinada") ---
SQL_ENGINE=sqlserver                     # "sqlserver" (por defecto) o "sqlite"

# SQL Server (Trusted Connection, no hace falta usuario/contraseña)
SQL_SERVER=
SQL_DATABASE=
SQL_DRIVER=ODBC Driver 18 for SQL Server
# SQL_TRUST_SERVER_CERTIFICATE=yes

# SQLite (alternativa temporal mientras no hay acceso al SQL Server corporativo)
# SQLITE_PATH=historial_facturas.db
```

Si la conexión a la base de datos falla o no está configurada, se registra un aviso por consola y el procesamiento de facturas continúa con normalidad (el histórico Excel no depende de ella).

---

## Uso

Una sola app FastAPI, un solo puerto (es el mismo puerto que debe tener configurado el flujo de Power Automate para `/upload-pdf`):

```bash
uvicorn main:app --reload --port 8000
```

| URL | Descripción |
|---|---|
| http://localhost:8000 | Interfaz web con las tres vistas: Visualización automática / Corregir manualmente / Subir facturas |
| http://localhost:8000/docs | Documentación interactiva (Swagger) de todos los endpoints |
| http://localhost:8000/upload-pdf | Endpoint que debe apuntar el flujo de Power Automate |

### Segunda pasada con visión (PDFs escaneados)

```bash
python imagenes.py            # procesa la carpeta "imagenes" una vez
python imagenes.py --watch    # repite cada INTERVALO_VIGILANCIA_IMAGENES segundos
```

### Migración del histórico Excel a SQL Server

```bash
python migrar_historial_excel.py --dry-run   # solo cuenta y lista, no escribe nada
python migrar_historial_excel.py             # migra de verdad
```

### Flujo de uso (vista "Corregir manualmente")

1. Abre el navegador en `http://localhost:8000` y pulsa "Corregir manualmente"
2. Revisa las facturas pendientes en `corregir_manualmente`
3. Completa o corrige los campos que falten (si el PDF no parece una factura, se avisa en vez de mostrar un formulario vacío)
4. Confirma: la factura se guarda en el historial y en SQL, y se mueve a `examinadas`
5. En la vista "Subir facturas" puedes añadir una factura suelta a la cola, marcándola como urgente si necesita prioridad

---

## Estructura del proyecto

```
├── main.py                   # App FastAPI única: vigilante, CORS, /static, incluye los routers
├── logic.py                  # Extracción PDF, llamada al LLM, clasificación, Excel, historial
├── routers/
│   ├── auto.py                # /upload-pdf, /procesar, /estadisticas
│   └── manual.py               # /pendientes-lista, /subir-factura, /extraer-pendiente, /confirmar-factura, /extraer, /historial*
├── templates/
│   └── index.html             # Shell con las 3 vistas (pestañas)
├── imagenes.py                # Segunda pasada con visión para PDFs escaneados
├── sql_historial.py           # Guardado en SQL Server / SQLite de facturas "examinada"
├── migrar_historial_excel.py  # Migración puntual del histórico Excel a SQL Server
├── sql/
│   └── crear_tabla_facturas_examinadas.sql   # DDL opcional de la tabla FacturasExaminadas
├── static/                    # Recursos estáticos (logo, etc.)
├── uploads/                   # Temporales de subida (flujo secundario /extraer)
├── requirements.txt
└── .env                       # Variables de entorno (no subir al repositorio)
```

---

## API endpoints

| Método | Ruta | Descripción |
|---|---|---|
| `GET` | `/` | Interfaz web (3 vistas) |
| `POST` | `/upload-pdf` | Recibe un PDF (p. ej. desde Power Automate), lo extrae, clasifica y mueve |
| `POST` | `/procesar` | Lanza manualmente el procesado de la carpeta `entrada` |
| `GET` | `/estadisticas` | Nº de PDFs por carpeta de destino |
| `GET` | `/pendientes-lista` | Lista los PDFs en `corregir_manualmente` |
| `POST` | `/subir-factura` | Sube un PDF directamente a la cola de corrección manual (opción "urgente") |
| `POST` | `/extraer-pendiente` | Extrae (o recupera del historial) los datos de un PDF pendiente; avisa si no parece una factura |
| `POST` | `/confirmar-factura` | Guarda la corrección, la persiste en SQL y mueve el PDF a `examinadas` |
| `POST` | `/extraer` | Extrae datos de uno o varios PDFs sueltos (flujo secundario, sin UI asociada) |
| `GET` | `/historial` | Descarga el historial Excel del día actual |
| `GET` | `/historial-json` | Devuelve el historial del día en formato JSON |

---

## Dependencias principales

| Paquete | Uso |
|---|---|
| `fastapi` | Framework web |
| `uvicorn` | Servidor ASGI |
| `pdfplumber` | Extracción de texto de PDFs |
| `PyMuPDF` (`fitz`) | Renderizado de páginas PDF a imagen para la segunda pasada con visión |
| `openai` | Cliente OpenAI / Azure OpenAI |
| `openpyxl` | Generación y lectura de ficheros Excel |
| `pyodbc` | Conexión a SQL Server |
| `python-dotenv` | Carga de variables de entorno |
| `python-multipart` | Subida de ficheros |
