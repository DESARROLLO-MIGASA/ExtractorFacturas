# Extractor de Facturas MIGASA

Sistema que extrae automáticamente los datos estructurados de facturas en PDF usando IA (OpenAI GPT-4.1 o compatible). Está compuesto por dos aplicaciones FastAPI independientes más un par de utilidades:

- **`app_auto`** — procesamiento desatendido: vigila una carpeta de entrada y expone un endpoint pensado para integrarse con Power Automate.
- **`app_manual`** — interfaz web para revisar/corregir las facturas que la extracción automática no pudo clasificar, y para extraer facturas sueltas bajo demanda.
- **`imagenes.py`** — segunda pasada con visión (GPT vision) para los PDFs escaneados o fotografiados que no tienen texto extraíble.
- **`sql_historial.py`** — guarda cada factura "examinada" (ya sea automática o corregida a mano) en una base de datos (SQL Server o SQLite), como complemento del histórico en Excel.
- **`migrar_historial_excel.py`** — migración puntual y segura de re-ejecutar del histórico Excel acumulado hacia SQL Server.

---

## Funcionalidades

- **Extracción automática** de datos de factura PDF mediante LLM (OpenAI GPT-4.1 o compatible)
- **Soporte multiidioma**: detecta facturas en español, inglés, portugués, francés, italiano, alemán, neerlandés, chino, griego y otros
- **Vigilancia automática de carpeta** (`app_auto`): sondea la carpeta `entrada` cada `INTERVALO_VIGILANCIA` segundos y procesa lo que encuentra, sin intervención manual
- **Endpoint de subida** (`/upload-pdf`) pensado para integrarse con **Power Automate**
- **Detección de albaranes**: los documentos que no son facturas se identifican y se apartan del flujo
- **Segunda pasada con visión** (`imagenes.py`) para PDFs escaneados/fotografiados sin texto extraíble, reutilizando toda la lógica de extracción, clasificación e historial de `app_auto`
- **Clasificación automática de PDFs** en carpetas según el resultado de la extracción:
  - `procesadas` — copia maestra de todo lo que pasa por el flujo automático
  - `examinadas` — extracción correcta (automática o corregida manualmente)
  - `corregir_manualmente` — falta algún campo obligatorio
  - `imagenes` — PDF sin texto extraíble, pendiente de la segunda pasada con visión
  - `imagenes_sin_datos` — tampoco se pudo leer con visión
  - `albaran` — el documento es un albarán, no una factura
  - `reenviar_falta_pedidocliente` — todo correcto salvo el número de pedido de cliente
  - `error` — fallo de procesamiento
- **Corrección manual** (`app_manual`): interfaz web para completar los campos que faltan en `corregir_manualmente`, confirmar y mover la factura a `examinadas`
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
Power Automate ──POST /upload-pdf──▶ app_auto
                                        │
                          ¿es albarán?──┴──▶ facturas/albaran
                                        │
                       ¿sin texto (<80 car.)?──▶ facturas/imagenes ──▶ imagenes.py (visión)
                                        │                                      │
                              extracción con LLM                    clasifica igual que app_auto
                                        │                                      │
                                 clasificar_factura                           mueve a examinadas /
                                        │                                corregir_manualmente / etc.
                    ┌───────────────────┼────────────────────┐
                    ▼                   ▼                    ▼
              examinadas       corregir_manualmente      reenviar_falta_pedidocliente / error
                    │                   │
          guardar_historial      app_manual (UI) revisa y confirma
                    │                   │
       guardar_factura_examinada_sql ◀──┘
```

`app_auto` también vigila la carpeta `entrada` de forma periódica (no solo vía Power Automate), procesando cualquier PDF que aparezca ahí.

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
INTERVALO_VIGILANCIA=30                  # segundos entre pasadas del vigilante (app_auto)

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

Cada aplicación es un FastAPI independiente; hay que lanzarlas por separado (con puertos distintos si se ejecutan a la vez):

```bash
# Procesamiento automático + endpoint para Power Automate
cd app_auto
uvicorn main:app --reload --port 8000

# Interfaz de corrección manual
cd app_manual
uvicorn main:app --reload --port 8001
```

| App | URL | Descripción |
|---|---|---|
| `app_auto` | http://localhost:8000 | Panel del vigilante automático (ver `/estadisticas`) |
| `app_auto` | http://localhost:8000/docs | Documentación interactiva (Swagger) de sus endpoints |
| `app_auto` | http://localhost:8000/upload-pdf | Endpoint que debe apuntar el flujo de Power Automate |
| `app_manual` | http://localhost:8001 | Interfaz web de corrección manual |
| `app_manual` | http://localhost:8001/docs | Documentación interactiva (Swagger) de sus endpoints |

> Los puertos `8000`/`8001` son solo el ejemplo de este README (pasados con `--port`); uvicorn usa `8000` por defecto si no se indica ninguno, así que si vas a levantar las dos apps a la vez en la misma máquina, asegúrate de usar puertos distintos.

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

### Flujo de uso (app_manual)

1. Abre el navegador en `http://localhost:8001`
2. Revisa las facturas pendientes en `corregir_manualmente`
3. Completa o corrige los campos que falten
4. Confirma: la factura se guarda en el historial y en SQL, y se mueve a `examinadas`
5. También puedes subir facturas sueltas para extracción puntual y descargar el historial en Excel

---

## Estructura del proyecto

```
├── app_auto/
│   ├── main.py            # API FastAPI: vigilante, /upload-pdf, /procesar, /estadisticas
│   ├── logic.py            # Extracción PDF, llamada al LLM, clasificación, Excel, historial
│   ├── templates/
│   │   └── index.html
│   └── facturas/           # entrada/procesadas/examinadas/corregir_manualmente/imagenes/...
├── app_manual/
│   ├── main.py             # API FastAPI: pendientes, confirmación, extracción puntual, historial
│   ├── logic.py
│   └── templates/
│       └── index.html
├── imagenes.py              # Segunda pasada con visión para PDFs escaneados
├── sql_historial.py          # Guardado en SQL Server / SQLite de facturas "examinada"
├── migrar_historial_excel.py # Migración puntual del histórico Excel a SQL Server
├── sql/
│   └── crear_tabla_facturas_examinadas.sql   # DDL opcional de la tabla FacturasExaminadas
├── static/                  # Recursos estáticos compartidos
├── uploads/                 # Temporales de subida (app_manual)
├── requirements.txt
└── .env                     # Variables de entorno (no subir al repositorio)
```

---

## API endpoints

### `app_auto`

| Método | Ruta | Descripción |
|---|---|---|
| `GET` | `/` | Interfaz web |
| `POST` | `/upload-pdf` | Recibe un PDF (p. ej. desde Power Automate), lo extrae, clasifica y mueve |
| `POST` | `/procesar` | Lanza manualmente el procesado de la carpeta `entrada` |
| `GET` | `/estadisticas` | Nº de PDFs por carpeta de destino |

### `app_manual`

| Método | Ruta | Descripción |
|---|---|---|
| `GET` | `/` | Interfaz web |
| `GET` | `/pendientes-lista` | Lista los PDFs en `corregir_manualmente` |
| `POST` | `/extraer-pendiente` | Extrae (o recupera del historial) los datos de un PDF pendiente |
| `POST` | `/confirmar-factura` | Guarda la corrección, la persiste en SQL y mueve el PDF a `examinadas` |
| `POST` | `/extraer` | Extrae datos de uno o varios PDFs sueltos (flujo secundario) |
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
