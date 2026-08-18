# Extractor de Facturas MIGASA

Sistema que extrae automáticamente los datos estructurados de facturas en PDF usando IA (OpenAI GPT-4.1 o compatible). Es una aplicación FastAPI (interfaz web con ocho pestañas) más un proceso de vigilancia independiente y un puñado de utilidades:

- **`main.py`** — app FastAPI: expone el endpoint que usa Power Automate, sirve la interfaz web y monta los routers. Ya **no** vigila la carpeta de entrada dentro de su propio proceso (ver más abajo).
- **`vigilante.py`** — proceso independiente (se lanza aparte con `python vigilante.py`, en su propia terminal) que sondea las carpetas `entrada` e `imagenes` cada `INTERVALO_VIGILANCIA` segundos y dispara el procesamiento. Antes este bucle vivía dentro del proceso de `uvicorn`; al compartir intérprete (y GIL) con las peticiones HTTP, una tanda larga de facturas dejaba la página web sin responder hasta que terminaba. Al ser un proceso de Windows aparte, uno ya no bloquea al otro. Se coordina con los reprocesos manuales lanzados desde la web mediante un candado de fichero, `LOCK_PROCESAMIENTO_AUTOMATICO` (`logic.py`), que vive en `.lock_procesamiento_automatico`.
- **`logic.py`** — extracción de PDF, llamada al LLM, clasificación, detección de duplicados, reenvíos por correo, Excel, historial. Toda la lógica de negocio consolidada en un solo módulo (~127 KB).
- **`routers/`** — endpoints agrupados por área: `auto.py` (procesamiento automático, estadísticas, reenvíos por correo, errores de extracción) y `manual.py` (corrección manual, incidencias, esperando alta, revisar facturas, duplicados, reservas, autocompletado de CIF, visor de PDF).
- **`imagenes.py`** — segunda pasada con visión (GPT vision) para los PDFs escaneados o fotografiados que no tienen texto extraíble.
- **`posiciones.py`** — localiza en qué posición de la página aparece cada dato ya extraído (vía PyMuPDF si el PDF tiene texto, vía OCR con Tesseract si es una factura escaneada), para pintar el recuadro en la vista de revisión. Cachea el resultado por nombre de archivo en `cache_posiciones/`.
- **`sql_historial.py`** — capa de acceso a SQL Server o SQLite: guarda cada factura completada (tabla `FacturasExaminadas`), gestiona el check de "revisada" (`FacturasRevisadas`), las reservas multiusuario (`ReservasFacturas`) y la caché de pendientes/incidencias (`ColaRevision`).
- **`migrar_historial_excel.py`** — migración puntual (segura de re-ejecutar) del histórico Excel acumulado antes de tener acceso a SQL Server, hacia la tabla `FacturasExaminadas`.
- **`migrar_orden_columnas.py`** — migración puntual e histórica que reordenó las columnas de los ficheros Excel ya escritos para que coincidieran con el orden actual de `EXPECTED_HEADERS` (esos ficheros se leen por posición, no por nombre de columna). Solo sería necesario volver a usarla si ese orden cambiara de nuevo.
- **`migrar_cola_revision.py`** — migración puntual (segura de re-ejecutar) que puebla `ColaRevision` con el contenido ya acumulado en `corregir_manualmente/` e `incidencias/` la primera vez que se despliega esa caché; a partir de ahí la caché se mantiene sola.
- **`cargar_empresas.py`** — utilidad independiente del circuito de facturas: carga `Empresas granel.xlsx` y `Empresas envasado.xlsb` en la tabla `EmpresasClasificadas` (CIF, código de empresa, nombre, activa, clasificación Granel/Envasado); al activar un CIF, dispara también la reclasificación automática de las incidencias que lo tenían pendiente (`reclasificar_cola_por_cif`).
- **`cargar_proveedores.py`** — utilidad independiente del circuito de facturas: carga `ProveedoresGranel.xlsx` y `ProveedoresEnvasado.xlsx` en la tabla `ProveedoresClasificados` (CIF, nombre, dirección, población, clasificación, bloqueado); igual que `cargar_empresas.py`, al activar un CIF dispara `reclasificar_cola_por_cif`.
- **`helper/`** — `facturahelper`, un ejecutable local (no forma parte de la app FastAPI) que cada usuario instala una vez en su propio PC para que el botón "abrir correo" (motivo "Otro") abra Outlook de escritorio ahí, y no en el servidor. Ver [FacturaHelper](#facturahelper-correo-con-outlook-en-el-pc-del-usuario).

---

## Funcionalidades

- **Extracción automática** de datos de factura PDF mediante LLM (OpenAI GPT-4.1 o compatible)
- **Soporte multiidioma**: detecta facturas en español, inglés, portugués, francés, italiano, alemán, neerlandés, chino, griego y otros
- **Vigilancia automática de carpeta**: el proceso independiente `vigilante.py` sondea las carpetas `entrada` e `imagenes` cada `INTERVALO_VIGILANCIA` segundos y procesa lo que encuentra, sin intervención manual
- **Endpoint de subida** (`/upload-pdf`) pensado para integrarse con **Power Automate**
- **Límite de tamaño**: un PDF de más de 6 MB no se intenta procesar; se aparta directamente a la cola de reenvío "pesa mucho" (ver más abajo)
- **Detección de "no es factura"**: los documentos que no son facturas (albaranes, tickets de báscula/pesaje, documentos de transporte/CMR, partes de horas...) se identifican y se apartan del flujo, tanto en el procesamiento automático como al abrir un pendiente en "Corregir manualmente"
- **Segunda pasada con visión** (`imagenes.py`) para PDFs escaneados/fotografiados sin texto extraíble, reutilizando toda la lógica de extracción, clasificación e historial de `logic.py`. Para evitar que esa pasada "invente" un CIF de comprador/proveedor que en realidad no aparece en el documento, se exige que dos lecturas independientes por visión coincidan antes de darlo por válido; si no coinciden, el campo queda sin resolver y la factura cae en Incidencias en vez de asignarse a una empresa equivocada
- **Detección de facturas duplicadas**: al completar una factura se comprueba si ya existe otra con el mismo número de factura, proveedor, comprador, base imponible y total (con tolerancia de redondeo); si coincide, ambas quedan marcadas como duplicado en SQL y aparecen en la pestaña "Duplicados" para que una persona decida cuál conservar (`/resolver-duplicado` descarta las demás sin borrar los PDFs, los mueve a `no_es_factura`)
- **Solicitud de reenvío por correo** con tres motivos posibles, seleccionables desde casi cualquier pestaña (Corregir manualmente, Incidencias, Esperando alta, Revisar facturas, Errores):
  - *Falta el número de pedido de cliente* y *Hay dos o más facturas en la misma página/PDF* — el PDF se traslada a una cola dedicada que vigila un flujo de Power Automate, que compone y envía el correo automáticamente al remitente original
  - *Otro motivo* (texto libre) — se abre un modal propio de EscanerIA donde la persona escribe el motivo y redacta el cuerpo del correo; el destinatario, el PDF y el asunto los resuelve el backend (no el navegador) y el envío se hace vía un flujo de Power Automate (`POWER_AUTOMATE_CORREO_URL`, ver más abajo), con el PDF adjunto en Base64. Si el envío se confirma, el PDF se mueve a `reenviadas_otro_motivo` y queda constancia en `HistorialCorreoOtroMotivo`; si falla, el PDF no se mueve y el modal se queda abierto sin perder lo escrito
- **Errores de extracción**: los PDFs que fallan al procesarse caen en la carpeta `error`; desde la pestaña "Errores" se pueden reprocesar (uno o todos a la vez) o clasificar con un motivo ("el fichero pesa demasiado" / "otro"), que los traslada a su propia cola de reenvío
- **Cola "Esperando alta"**: cualquier factura de Corregir manualmente, Incidencias o Revisar facturas cuyo comprador o proveedor todavía no esté dado de alta en `EmpresasClasificadas`/`ProveedoresClasificados` se puede apartar manualmente a esta cola independiente; cuando se da de alta el CIF que faltaba, el botón "Reprocesar" la reintegra al flujo normal (completadas o corregir manualmente)
- **Caché de pendientes/incidencias** (`ColaRevision`): "Corregir manualmente" e "Incidencias" leen de esta tabla en vez de releer cada PDF y volver a resolver Buyer/Proveedor en cada carga de página, que con una cola larga tardaba varios minutos; la caché se mantiene sola en el momento en que cada PDF entra, sale o se corrige en una de las dos colas. La reclasificación automática de incidencias (por si el CIF que faltaba se acaba de dar de alta) ya no recorre toda la cola en cada recarga: se dispara una sola vez, solo para las incidencias afectadas, justo cuando `cargar_empresas.py`/`cargar_proveedores.py` activan ese CIF concreto
- **Marcar como revisada** (en Corregir manualmente o en Incidencias) da la factura por buena tal como está: la guarda en `FacturasExaminadas` y mueve el PDF directamente a `facturas_revisadas`, sin exigir que todos los campos obligatorios estén completos
- **Revisar facturas / marcar como definitiva**: las facturas completadas automáticamente se pueden revisar, editar y marcar como "definitiva"; al marcarlas se reflejan también en un Excel aparte (`facturas_100_definitivas.xlsx`, descargable) y el PDF se traslada de `completadas` a `facturas_revisadas`
- **Reservas multiusuario**: para que varias personas puedan revisar facturas a la vez sin pisarse el trabajo, cada factura abierta para edición queda "reservada" a nombre de quien la está viendo (nombre libre, sin login real); la reserva se refresca mientras la pestaña sigue abierta y expira sola si se cierra o se queda inactiva. Un botón "Liberar todas las reservas" en la cabecera actúa como válvula de escape si alguna se queda colgada
- **Autocompletado de CIF**: al editar el comprador o el proveedor de una factura, un buscador sugiere coincidencias contra `EmpresasClasificadas`/`ProveedoresClasificados` por prefijo
- **Clasificación automática de PDFs** en carpetas según el resultado de la extracción:
  - `procesadas` — copia maestra de todo lo que pasa por el flujo automático
  - `completadas` — extracción correcta (automática o corregida manualmente), pendiente de marcar como definitiva
  - `facturas_revisadas` — factura ya revisada/marcada como definitiva o confirmada como "Revisada"; destino final
  - `corregir_manualmente` — falta algún campo obligatorio
  - `imagenes` — PDF sin texto extraíble, pendiente de la segunda pasada con visión
  - `imagenes_sin_datos` — tampoco se pudo leer con visión
  - `incidencias` — el comprador o el proveedor no se resuelven contra `EmpresasClasificadas`/`ProveedoresClasificados` (CIF no encontrado, checksum inválido, inactivo/bloqueado)
  - `esperando_alta` — apartada manualmente porque el comprador o el proveedor todavía no está dado de alta en la base de datos; espera a que alguien pulse "Reprocesar" tras darlo de alta
  - `no_es_factura` — el documento no es una factura (albarán, ticket de báscula/pesaje, documento de transporte/CMR, parte de horas, etc.), o un duplicado descartado
  - `error` — fallo de procesamiento, pendiente de reprocesar o clasificar con un motivo
  - `reenviar_falta_pedidocliente` / `reenviadas_falta_pedidocliente` — cola / archivo de solicitudes por falta de pedido de cliente
  - `reenviar_dos_factura_una_pagina` — cola de solicitudes por "dos facturas en un mismo PDF"
  - `reenviar_otro_motivo` (sin uso actualmente) / `reenviadas_otro_motivo` — archivo de solicitudes con motivo libre (vía Outlook, abierto en el PC del usuario con `facturahelper`); no hay cola intermedia, el PDF va directo al archivo en cuanto se abre Outlook
  - `reenviar_error_pesa_mucho` / `reenviar_error_otro` — colas de errores de extracción clasificados con un motivo
- **Alertas grandes** en "Visualización automática" cuando hay facturas pendientes de revisión manual, incidencias, errores u otras colas con elementos esperando
- **Filtro de fecha por pestaña**: las tablas de Corregir manualmente, Incidencias, Esperando alta, Errores, Revisar facturas, Duplicados y No es factura se pueden acotar por fecha de procesamiento, cada una sobre su propia cola (no existe una tabla única con todo el histórico)
- **Corrección manual**: vista web para completar los campos que faltan en `corregir_manualmente`, confirmar y mover la factura a `completadas`
- **Revisión en pantalla dividida**: al editar una factura (pendiente, incidencia, esperando alta o completada) se abre un visor propio del PDF junto al formulario, con un recuadro pastel que salta automáticamente sobre la posición de cada dato; con ↓/Enter se pasa al siguiente campo y con ↑ al anterior — la navegación no se para en el último campo: si la factura tiene líneas, sigue por ellas y solo vuelve al principio al pasar la última. Si el recuadro no acierta o el dato no se detectó, se puede arrastrar un recuadro a mano sobre el PDF (funciona igual en facturas con texto que en escaneadas, vía OCR) para rellenar el campo con lo que haya ahí
- **Líneas de factura editables y localizables en el PDF**: ver la sección [Líneas de factura](#líneas-de-factura) más abajo — añadir/quitar líneas a mano, corregirlas celda a celda con recuadro en el visor, o reextraer con IA la región de la tabla que se seleccione (sin tener que reextraer la factura entera)
- **Papelera de corrección manual**: si un PDF de la cola de `corregir_manualmente` resulta no ser una factura, se puede apartar a `no_es_factura` con un clic, sin borrarlo
- **Ver el PDF sin salir de la app**: cada pendiente de corrección se puede abrir directamente en el navegador, sin ir a buscarlo a la carpeta
- **Exportar cualquier listado a Excel**: las tablas de las distintas pestañas se pueden descargar tal cual se están viendo (filtradas), además de la descarga específica de "facturas definitivas"
- **Persistencia en base de datos** de las facturas completadas (tabla `FacturasExaminadas`), con upsert por nombre de archivo — no rompe el flujo si la conexión falla

> Los endpoints `POST /extraer` y `POST /extraer-pendiente` (subida/extracción manual de PDFs sueltos) siguen existiendo en `routers/manual.py`, pero ninguno de los dos tiene ya una vista asociada en la interfaz actual (la antigua pestaña "Subir facturas" se retiró, y la carga de cada pendiente en "Corregir manualmente" reutiliza los datos ya cargados en la tabla en vez de volver a llamar a este endpoint); son código heredado, no alcanzables desde la UI hoy.

---

## Campos extraídos

Orden real de `EXPECTED_HEADERS` en `logic.py` (es también el orden de las columnas en los Excel de historial):

| # | Campo | Descripción |
|---|---|---|
| 1 | Archivo | Nombre del archivo PDF |
| 2 | NumeroFactura | Número de factura |
| 3 | Buyer | CIF/NIF/VAT del comprador |
| 4 | Empresa | Razón social del comprador |
| 5 | Proveedor | CIF/NIF/VAT del proveedor |
| 6 | NombreProveedor | Razón social del proveedor |
| 7 | PedidoCliente | Número de pedido del cliente |
| 8 | BaseImp | Base imponible |
| 9 | BaseIRPF | Base IRPF (si aplica) |
| 10 | TipoIVA | Tipo de IVA principal |
| 11 | TipoIVA2 | Segundo tipo de IVA |
| 12 | TipoIVA3 | Tercer tipo de IVA |
| 13 | ImporIVA | Importe del IVA |
| 14 | TotalFact | Importe total de la factura |
| 15 | Moneda | Moneda (EUR, USD, GBP…) |
| 16 | FFactura | Fecha de la factura |
| 17 | FOperacion | Fecha de operación |
| 18 | FEscaneo | Fecha y hora de procesamiento |
| 19 | NumerosAlbaran | Número(s) de albarán del proveedor (varios valores posibles, separados por `;`) |

`PedidoCliente` y `NumerosAlbaran` admiten varios valores en una misma factura, concatenados con `;` (p.ej. `PEDIDO-001;PEDIDO-002`); en SQL Server ambas columnas son `NVARCHAR(MAX)` (se ha visto en producción una factura real con ~85-96 valores, muy por encima de los 1000 caracteres que tenían antes).

### Líneas de factura

Además de los 19 campos de cabecera de arriba, el LLM también extrae las **líneas/conceptos individuales** de la factura (descripción, cantidad, precio, importe, y "otros" para cualquier dato adicional como número de bultos/palés). No son una columna más de `EXPECTED_HEADERS`: se guardan como un **CSV aparte junto al PDF**, con el mismo nombre y el sufijo `_lineas.csv` (delimitador `;`, con BOM UTF-8 para que Excel en español lo abra bien), y viajan con el PDF cuando este se mueve entre las carpetas del flujo (`ruta_lineas_csv`/`guardar_lineas_csv` en `logic.py`).

Se pueden consultar y corregir desde la vista de edición de cualquier factura:
- Se muestran al final de los campos, editables celda a celda igual que el resto de campos (mismo estilo, mismo botón ✎).
- Al pinchar una línea se resalta en el visor del PDF, igual que ocurre con los campos normales (`posiciones.calcular_cajas_lineas`, localiza cada línea por su Descripción+Importe).
- Se pueden añadir líneas nuevas en blanco o eliminar una existente (por si la extracción se dejó alguna o detectó de más).
- **Reextraer líneas por región**: si la extracción automática se equivocó en bloque (o no encontró ninguna), se puede seleccionar a mano la zona de la tabla en el PDF y pedirle a la IA que la vuelva a leer -mandando solo esa imagen recortada, no la factura entera-. Cada región seleccionada **añade** líneas a las que ya haya (no las sustituye), para poder repetirlo página a página en facturas de varias hojas; un botón "Vaciar líneas" permite empezar de cero.

---

## Flujo de procesamiento

```
Power Automate ──POST /upload-pdf──▶ main.py (routers/auto.py)
                                        │
                    ¿pesa más de 6MB?───┴──▶ facturas/reenviar_error_pesa_mucho
                                        │
                       ¿no es factura?──┴──▶ facturas/no_es_factura
                                        │
                       ¿sin texto (<80 car.)?──▶ facturas/imagenes ──▶ imagenes.py (visión)
                                        │                                      │
                              extracción con LLM                    clasifica igual que el flujo automático
                                        │                                      │
                                 clasificar_factura                           mueve a completadas /
                                        │                                corregir_manualmente / incidencias / etc.
        ┌───────────────┬──────────────┼───────────────┬───────────────────┐
        ▼                ▼             ▼                ▼                    ▼
  completadas   corregir_manualmente  incidencias   esperando_alta        error
        │                │                                                   │
  detección de    vista "Corregir             (marcar-esperando-alta       reprocesar o
  duplicados      manualmente" o                manual, cualquier         clasificar-error
        │         "Incidencias"                 pestaña)                  (pesa_mucho / otro)
        ▼         revisa/confirma/                    │
  marcar-factura-  marca-revisada                reprocesar-esperando-alta
  definitiva              │                            │
        │        guardar_factura_examinada_sql ◀───────┘
        ▼
facturas_revisadas
```

La app FastAPI (`main.py`, servida con `uvicorn`) **no** vigila carpetas por sí misma. Ese trabajo lo hace `vigilante.py`, un proceso aparte que hay que arrancar por separado (ver [Uso](#uso)) y que sondea `entrada`/`imagenes` cada `INTERVALO_VIGILANCIA` segundos, además de lo que llegue directamente vía `/upload-pdf`.

---

## Requisitos

- Python 3.10+
- API key de OpenAI (o endpoint compatible); con soporte de visión si se va a usar `imagenes.py`
- Opcional: SQL Server accesible + ODBC Driver 18, o usar el motor `sqlite` (sin instalación adicional) mientras no haya acceso a SQL Server
- **Tesseract OCR** (el programa, no solo la librería `pytesseract` de `requirements.txt`): necesario para que el recuadro automático de la vista de revisión funcione en facturas escaneadas/sin texto, y para la selección manual con arrastre (`posiciones.py`). Sin él, esas dos cosas simplemente no hacen nada (no rompen el resto de la app). En Windows: instalar el ejecutable de Tesseract y, si no queda en el PATH, indicar la ruta en `.env` con `TESSERACT_CMD`.
- **FacturaHelper instalado en el PC de cada usuario** (no en el servidor): necesario únicamente para el botón "abrir correo" de las solicitudes de reenvío con motivo "Otro". El servidor ya no llama a Outlook él mismo (ver [FacturaHelper](#facturahelper-correo-con-outlook-en-el-pc-del-usuario)); sin el helper instalado y sin Microsoft Outlook de escritorio (no "New Outlook") configurado en ese PC, esa acción concreta falla con un aviso, pero el resto de la app sigue funcionando igual.

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

## FacturaHelper (correo con Outlook en el PC del usuario)

El botón "abrir correo" del motivo "Otro" necesita que Outlook se abra en el PC de **quien está usando la app**, no en el servidor donde corre `uvicorn`. Por eso esa acción ya no la hace el backend con `win32com`: el navegador navega a un enlace `facturahelper://...` que dispara un ayudante local, instalado una vez por PC, el cual pide los datos del correo y el PDF al servidor (`GET /datos-correo-outlook/{archivo}`, `GET /pdf-factura/{archivo}`) y abre Outlook ahí mismo con `win32com`.

Instalación (una vez por PC, no requiere administrador):

```powershell
# Copiar helper/facturahelper.exe y helper/instalar.ps1 a una carpeta del PC del usuario, luego:
powershell -ExecutionPolicy Bypass -File instalar.ps1
```

Esto copia el ejecutable a `%LOCALAPPDATA%\FacturaHelper\` y registra el protocolo `facturahelper://` en `HKCU:\Software\Classes` (solo para el usuario actual). Requiere Microsoft Outlook de escritorio (no "New Outlook") instalado y configurado en ese PC; si no lo está, el helper avisa con un mensaje en vez de fallar en silencio.

Para reconstruir `facturahelper.exe` desde `helper/facturahelper.py` tras un cambio:

```bash
pip install -r helper/requirements-build.txt   # pywin32
pip install pyinstaller
pyinstaller --onefile --noconsole --name facturahelper helper/facturahelper.py
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
FACTURAS_DIR=C:\ruta\a\facturas          # carpeta base con entrada/procesadas/completadas/...
INTERVALO_VIGILANCIA=30                  # segundos entre pasadas de vigilante.py

# --- Segunda pasada con visión (imagenes.py) ---
# MAX_PAGINAS_IMAGENES=3
# DPI_IMAGENES=200
# INTERVALO_VIGILANCIA_IMAGENES=300      # solo si se lanza con --watch

# --- Recuadro del visor de revisión (posiciones.py) ---
# TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe   # solo si no está en el PATH
# DPI_CACHE_POSICIONES=200                                     # por defecto, igual que DPI_IMAGENES

# --- Base de datos (guardado de facturas completadas, revisadas y reservas) ---
SQL_ENGINE=sqlserver                     # "sqlserver" (por defecto) o "sqlite"

# SQL Server (Trusted Connection, no hace falta usuario/contraseña)
SQL_SERVER=
SQL_DATABASE=
SQL_DRIVER=ODBC Driver 18 for SQL Server
# SQL_TRUST_SERVER_CERTIFICATE=yes

# SQLite (alternativa temporal mientras no hay acceso al SQL Server corporativo)
# SQLITE_PATH=historial_facturas.db

# --- Microsoft Graph (reenvío "Otro": crea el borrador de correo directamente
# en el buzón compartido, con el PDF ya adjunto, vía API en vez de abrir
# Outlook por COM -que solo funcionaría en el propio servidor, no en el PC
# del usuario-). Requiere que IT registre una app en Azure AD y conceda,
# con consentimiento de administrador, el permiso de Graph "Mail.ReadWrite"
# de tipo Application (Mail.Send no basta: hace falta poder crear/editar el
# borrador antes de enviarlo). Conviene además que IT restrinja el alcance
# de la app a un único buzón con una Application Access Policy en Exchange
# Online (New-ApplicationAccessPolicy), en vez de dejarla con acceso a todos
# los buzones del tenant. Mientras estas tres variables no estén rellenas,
# el endpoint /crear-borrador-outlook responde con un aviso de "Microsoft
# Graph no está configurado todavía" en vez de fallar en silencio. El botón
# "Otro" del modal de envío de correo ya no lo usa (ver POWER_AUTOMATE_CORREO_URL
# más abajo); se deja disponible por compatibilidad, sin romperlo.
GRAPH_TENANT_ID=
GRAPH_CLIENT_ID=
GRAPH_CLIENT_SECRET=
GRAPH_BUZON_ENVIO=escanerIA@migasa.com

# --- Envío de correo "Otro motivo" vía Power Automate: al elegir "Otro" en
# el modal de envío de correo, EscanerIA abre su propio modal (motivo +
# cuerpo) y el backend hace un POST a esta URL con el PDF en Base64 y el
# resto de datos ya resueltos (destinatario, asunto, usuario...). La URL del
# flujo solo vive aquí -nunca se expone al navegador ni se escribe en los
# logs-. Mientras esté vacía, esa acción responde con un aviso de "no
# configurado todavía" en vez de fallar en silencio.
POWER_AUTOMATE_CORREO_URL=

# --- Login con Windows Authentication (IIS por delante de uvicorn, ver
# "Despliegue en producción" más abajo) — el rol se calcula por pertenencia
# a estos dos grupos de Active Directory (los gestiona IT: alta/baja de
# gente es cosa suya, no de este .env). Quien no esté en ninguno de los dos
# se queda sin acceso (403).
AD_DOMINIO=DOMINT
AD_GRUPO_ADMINS=G-HQ-TEC-INF-SG-Admins_Escaner_Facturas
AD_GRUPO_USUARIOS=G-SAT-TEC-INF-SG-Users_Escaner_Facturas
```

Si la conexión a la base de datos falla o no está configurada, se registra un aviso por consola y el procesamiento de facturas continúa con normalidad (el histórico Excel no depende de ella). Ojo: si el login SQL configurado no tiene permiso de `CREATE TABLE`/`ALTER TABLE`, la app intenta crear y migrar las tablas (incluidas `ReservasFacturas` y `ColaRevision`, y columnas nuevas como `NumerosAlbaran`) al vuelo y ese fallo es silencioso (solo queda un aviso por consola) — sin `ReservasFacturas` la función de reservas multiusuario deja de funcionar, sin `ColaRevision` las pestañas "Corregir manualmente" e "Incidencias" dejan de poder listar nada, y si falta una columna nueva el guardado de **todas** las facturas empieza a fallar con un error `Invalid column name`, hasta que alguien con permisos de DDL ejecute a mano los scripts de `sql/` o el `ALTER TABLE` que corresponda.

---

## Uso

Dos procesos independientes (ambos deben estar corriendo para tener el flujo completo):

```bash
# Terminal 1: servidor web (interfaz + endpoint de Power Automate)
uvicorn main:app --reload --port 8000

# Terminal 2: vigilancia de las carpetas entrada / imagenes
python vigilante.py
```

| URL | Descripción |
|---|---|
| http://localhost:8000 | Interfaz web con las ocho pestañas |
| http://localhost:8000/docs | Documentación interactiva (Swagger) de todos los endpoints |
| http://localhost:8000/upload-pdf | Endpoint que debe apuntar el flujo de Power Automate |

### Despliegue en producción (IIS + proxy .NET/YARP delante de uvicorn)

En el VM de producción, uvicorn **no** se expone directamente a la red: escucha
solo en `127.0.0.1:8800` y delante hay un sitio IIS llamado **`EscanerIA`**
(`C:\sites\EscanerIAProxy\publish`) que ejecuta una app ASP.NET Core aparte,
**`EscanerIA.AuthProxy`**, vía `AspNetCoreModuleV2`. Esa app es la que atiende
a los usuarios, se apoya en Windows Authentication de IIS para autenticar
contra el dominio MIGASA y usa `Yarp.ReverseProxy` para reenviar cada petición
a FastAPI, añadiendo la cabecera `X-Forwarded-User` con el usuario de Windows
ya autenticado (y quitando primero cualquier valor de esa cabecera que viniera
del cliente, para que no se pueda suplantar). Es la app que lee `auth.py`.

> El código fuente de `EscanerIA.AuthProxy` **no vive en este repositorio**:
> está en `C:\sites\EscanerIAProxy\EscanerIA.AuthProxy\` en el propio VM
> (`Program.cs` + `appsettings.json`, este último con la URL de destino
> `http://127.0.0.1:8800/`). El script `deploy/iis_setup/configurar_iis_windows_auth.ps1`
> de esta carpeta describe un montaje distinto y **más antiguo** (IIS +
> Application Request Routing + URL Rewrite haciendo el proxy inverso
> directamente, sin esta app .NET intermedia); quedó desactualizado cuando se
> migró a `EscanerIA.AuthProxy` y no refleja lo que hay desplegado hoy. Antes
> de volver a ejecutarlo en el VM, confirmar primero si sigue siendo el
> método vigente o si hay que actualizarlo para que arranque este proxy.

`GET /proxy-health` (servido por el propio `EscanerIA.AuthProxy`, sin pasar a
FastAPI) devuelve si la petición llegó autenticada y con qué usuario; útil
para comprobar la parte de IIS/Windows Auth sin depender de que uvicorn esté
levantado. `GET /whoami` (ya dentro de FastAPI) comprueba el mismo dato una
vez atravesado el proxy.

`get_current_user`/`requerir_admin` (`auth.py`) leen la cabecera
`X-Forwarded-User` y devuelven `{"username", "usuario_dominio", "role"}`;
`role` se calcula consultando por LDAP (vía ADSI/`pywin32`, con la identidad
de Windows del propio proceso, sin contraseña de ninguna cuenta de servicio)
si el usuario es miembro directo de `AD_GRUPO_ADMINS` (`.env`) -> `"admin"`,
de `AD_GRUPO_USUARIOS` -> `"usuario"`, o ninguno de los dos -> 403.

Nota de threading: `_es_miembro_de_grupo` llama a `pythoncom.CoInitialize()`
antes de usar ADSI porque FastAPI ejecuta esta dependencia (síncrona) en un
hilo del thread pool, no en el principal -sin esa llamada falla con "No se ha
llamado a CoInitialize" en cada request-.

### Arranque y parada en producción

En el VM, uvicorn no lo arranca nadie a mano: lo hace la **tarea programada de
Windows `EscanerIA-FastAPI`**, configurada directamente en el servidor (no
forma parte de este repositorio, no hay ningún script que la cree ni la
reproduzca). Sus características:

- **Disparador**: al iniciar el sistema (`MSFT_TaskBootTrigger`) — se lanza
  sola en cada arranque del VM, sin necesidad de que nadie inicie sesión.
- **Cuenta**: `SYSTEM`, con privilegio "Highest".
- **Acción**: `...\.venv\Scripts\python.exe -m uvicorn main:app --host 127.0.0.1 --port 8800`
- **Reintentos**: hasta 10 veces, cada 1 minuto, dentro de una ventana de 72h,
  si el proceso muere.

Comandos útiles (PowerShell, como Administrador, en el propio VM):

```powershell
# Ver estado, última ejecución y resultado
Get-ScheduledTask -TaskName "EscanerIA-FastAPI" | Get-ScheduledTaskInfo

# Parar (mata el proceso de uvicorn en marcha)
Stop-ScheduledTask -TaskName "EscanerIA-FastAPI"

# Arrancar (o relanzar tras un cambio de código)
Start-ScheduledTask -TaskName "EscanerIA-FastAPI"
```

El sitio IIS `EscanerIA` (el proxy `EscanerIA.AuthProxy`) es independiente de
esta tarea y normalmente no hace falta tocarlo; si hiciera falta, se gestiona
como cualquier sitio de IIS (`Start-Website`/`Stop-Website -Name "EscanerIA"`,
o desde el Administrador de IIS).

`vigilante.py` **no** tiene ninguna tarea programada equivalente: hoy en el VM
depende de que alguien lo deje corriendo a mano (`python vigilante.py`, en su
propia ventana/sesión) y se para en cuanto esa sesión se cierra. Sin él
corriendo, `entrada`/`imagenes` no se vigilan solas — solo se procesa lo que
llegue directamente vía `/upload-pdf` o lo que se dispare a mano desde
"Procesar carpeta entrada" en la interfaz. Antes de dar por sentado que el
vigilante está activo en el VM, conviene comprobarlo (por ejemplo, si
`vigilante_out.log`/`vigilante_err.log` en la raíz del proyecto se están
actualizando).

`deploy/iis_setup/` contiene el script `configurar_iis_windows_auth.ps1`
(instala IIS + Windows Authentication, ARR y URL Rewrite, y configuraba la
regla de proxy inverso hacia `127.0.0.1:8000` directamente) junto con los dos
instaladores MSI oficiales de Microsoft que necesita, ya descargados. Ver el
aviso más arriba: describe el montaje anterior a `EscanerIA.AuthProxy`, no el
actual.

### Pestañas de la interfaz

1. **Visualización automática** — panel con el nº de PDFs pendientes en cada carpeta/cola, botón "Procesar carpeta entrada", botón "Ver facturas reenviadas" y alertas
2. **Corregir manualmente** — cola de facturas con algún campo obligatorio sin resolver
3. **Incidencias** — facturas cuyo comprador o proveedor no se resuelve contra las tablas maestras
4. **Esperando alta** — facturas apartadas manualmente a la espera de que se dé de alta el CIF
5. **Errores** — PDFs cuya extracción falló; reprocesar uno o todos, o clasificar con un motivo
6. **Revisar facturas** — facturas completadas automáticamente; editar, marcar como definitiva, descargar el Excel de definitivas
7. **Duplicados** — pares de facturas que coinciden en número, importes y CIFs; decidir cuál conservar
8. **No es factura** — documentos descartados (no eran factura, o duplicados resueltos)

### Segunda pasada con visión (PDFs escaneados)

```bash
python imagenes.py            # procesa la carpeta "imagenes" una vez
python imagenes.py --watch    # repite cada INTERVALO_VIGILANCIA_IMAGENES segundos
```

En producción, `vigilante.py` ya llama a `procesar_carpeta_imagenes()` en cada vuelta, así que normalmente no hace falta lanzar `imagenes.py --watch` por separado.

### Migración del histórico Excel a SQL Server

```bash
python migrar_historial_excel.py --dry-run   # solo cuenta y lista, no escribe nada
python migrar_historial_excel.py             # migra de verdad
```

### Empresas clasificadas (Granel / Envasado)

Utilidad puntual, independiente del circuito de facturas, para volcar a SQL el listado de empresas compradoras exportado desde Business Central (`Empresas granel.xlsx`, `Empresas envasado.xlsb`) a la tabla `EmpresasClasificadas` (CIF, código de empresa, nombre, si está activa, clasificación). Es seguro repetir la carga: hace upsert por CIF + Clasificación.

```bash
python cargar_empresas.py --dry-run   # solo cuenta y lista, no escribe nada
python cargar_empresas.py             # carga de verdad
```

### Proveedores clasificados (Granel / Envasado)

Utilidad puntual, independiente del circuito de facturas, para volcar a SQL el listado de proveedores de Business Central a la tabla `ProveedoresClasificados` (CIF, nombre, dirección, población, clasificación y si está bloqueado). Ya no hace falta descargar y colocar a mano ningún Excel: ambas clasificaciones se leen en caliente por OData de Business Central (NTLM, credenciales en `BC_ODATA_USER`/`BC_ODATA_PASSWORD` del `.env`).
- **Envasado** — un único servicio (`ProveedoresBloq`, instancia "oleico"; URL opcional en `BC_ODATA_URL`). No trae dirección, solo población.
- **Granel** — dos servicios de la instancia "olivar" (empresa Migasa Aceites, S.L.U.): `EscanerIAListaProveedores` (`BC_ODATA_GRANEL_LISTA_URL`, lista base con dirección/población y un bloqueo "genérico") y `EscanerIAConsultaBloProv` (`BC_ODATA_GRANEL_BLOQUEO_URL`, bloqueo específico de esa empresa por proveedor, que manda sobre el genérico cuando existe).

Si un mismo CIF aparece en ambas clasificaciones, se guarda una fila por cada una. Es seguro repetir la carga: hace upsert por CIF + Clasificación.

```bash
python cargar_proveedores.py --dry-run   # solo cuenta y lista, no escribe nada
python cargar_proveedores.py             # carga de verdad
```

### Flujo de uso (vista "Corregir manualmente")

1. Abre el navegador en `http://localhost:8000` y pulsa "Corregir manualmente"
2. Revisa las facturas pendientes; puedes abrir el PDF de cualquiera con "📄 Abrir PDF" antes de decidir
3. Completa o corrige los campos que falten (si el PDF no parece una factura, se avisa en vez de mostrar un formulario vacío); el recuadro sobre el PDF salta al campo activo, y se puede arrastrar uno a mano si no acierta
4. Si el documento no es realmente una factura, pulsa "🗑️ No es factura" para apartarlo a `no_es_factura` sin borrarlo
5. Si falta el pedido de cliente, hay dos facturas en un mismo PDF, o hace falta pedir cualquier otra corrección al proveedor por correo, usa "Solicitar reenvío" y elige el motivo
6. Confirma: la factura se guarda en el historial y en SQL, y se mueve a `completadas`

---

## Estructura del proyecto

```
├── main.py                   # App FastAPI: CORS, /static, incluye los routers (ya no vigila carpetas)
├── vigilante.py               # Proceso aparte: sondea entrada/imagenes cada INTERVALO_VIGILANCIA segundos
├── logic.py                  # Extracción PDF, LLM, clasificación, duplicados, reenvíos, Excel, historial
├── routers/
│   ├── auto.py                 # /upload-pdf, /procesar, /estadisticas, reenvíos por correo, errores de extracción
│   └── manual.py                # pendientes, incidencias, esperando alta, revisar facturas, duplicados, reservas, CIF, PDF
├── templates/
│   └── index.html             # Shell con las 8 pestañas
├── imagenes.py                # Segunda pasada con visión para PDFs escaneados
├── posiciones.py               # Recuadro por campo en el visor de revisión (PyMuPDF / Tesseract)
├── sql_historial.py           # Acceso a SQL Server / SQLite: FacturasExaminadas, FacturasRevisadas, ReservasFacturas, ColaRevision
├── migrar_historial_excel.py  # Migración puntual del histórico Excel a SQL Server
├── migrar_orden_columnas.py   # Migración puntual e histórica del orden de columnas en los Excel ya escritos
├── migrar_cola_revision.py    # Migración puntual: puebla ColaRevision con lo ya acumulado en corregir_manualmente/incidencias
├── cargar_empresas.py         # Carga puntual de Empresas granel.xlsx / envasado.xlsb a EmpresasClasificadas
├── cargar_proveedores.py      # Carga Granel + Envasado (OData BC) a ProveedoresClasificados
├── sql/
│   ├── crear_tabla_empresas_clasificadas.sql     # DDL opcional de la tabla EmpresasClasificadas
│   ├── crear_tabla_facturas_examinadas.sql       # DDL opcional de la tabla FacturasExaminadas (SQL Server)
│   ├── crear_tabla_facturas_examinadas_sqlite.sql # DDL equivalente para el motor SQLite
│   ├── crear_tabla_proveedores_clasificados.sql  # DDL opcional de la tabla ProveedoresClasificados
│   ├── crear_tabla_reservas_facturas.sql         # DDL opcional de la tabla ReservasFacturas (reservas multiusuario)
│   └── crear_tabla_colarevision.sql              # DDL opcional de la tabla ColaRevision (caché de pendientes/incidencias)
├── helper/                    # FacturaHelper: se instala en el PC del usuario, no forma parte de la app FastAPI
│   ├── facturahelper.py         # Fuente del ayudante local (abre Outlook con win32com en el PC del usuario)
│   ├── facturahelper.exe        # Ejecutable ya compilado (PyInstaller)
│   ├── instalar.ps1             # Instala el .exe en %LOCALAPPDATA% y registra el protocolo facturahelper://
│   └── requirements-build.txt   # Dependencias solo para compilar el helper (pywin32)
├── static/                    # Recursos estáticos (logo, etc.)
├── uploads/                   # Temporales de subida (flujo secundario /extraer, sin vista asociada hoy)
├── requirements.txt
└── .env                       # Variables de entorno (no subir al repositorio)
```

---

## API endpoints

### Automático / Power Automate (`routers/auto.py`)

| Método | Ruta | Descripción |
|---|---|---|
| `POST` | `/upload-pdf` | Recibe un PDF desde Power Automate, lo extrae, clasifica y mueve |
| `POST` | `/procesar` | Lanza manualmente el procesado de la carpeta `entrada` |
| `GET` | `/estadisticas` | Nº de PDFs por carpeta de destino, más reenviadas y duplicados pendientes |
| `GET` | `/reenviadas-detalle` | Detalle (factura, email, fecha, motivo) de lo ya reenviado |
| `GET` | `/reenviar-pedido-lista` | Lista los PDFs en `reenviar_falta_pedidocliente` |
| `GET` | `/motivos-envio-correo` | Motivos disponibles para solicitar un reenvío |
| `POST` | `/solicitar-envio-correo` | Traslada/aparta un PDF a la cola del motivo elegido |
| `GET` | `/solicitudes-envio-correo-lista` | Lista las solicitudes que no son "falta pedido cliente" |
| `GET` | `/datos-correo-outlook/{archivo}` | Destinatario y nombre de adjunto para que `facturahelper` (en el PC del usuario) componga el correo del motivo "otros"; ya no abre Outlook desde el servidor |
| `GET` | `/datos-correo-otro-motivo/{archivo}` | Destinatario, nombre de PDF y asunto (solo lectura) para pintar el modal de "Otro motivo" antes de enviar |
| `POST` | `/enviar-correo-otro-motivo` | Envía el correo del motivo "otros" vía Power Automate (destinatario/PDF/asunto resueltos en el backend) y, si tiene éxito, archiva el PDF |
| `GET` | `/motivos-error-extraccion` | Motivos disponibles para clasificar un PDF en `error` |
| `GET` | `/errores-completo-json` | Lista los PDFs en `error` |
| `GET` | `/no-factura-completo-json` | Lista los PDFs en `no_es_factura` |
| `POST` | `/abrir-carpeta-no-factura` | Abre la carpeta `no_es_factura` en el explorador del servidor |
| `POST` | `/clasificar-error` | Clasifica un PDF de `error` con un motivo y lo mueve a su cola de reenvío |
| `POST` | `/reprocesar-error` | Reintenta la extracción de un PDF concreto de `error` |
| `POST` | `/reprocesar-errores` | Reintenta la extracción de todos los PDFs de `error` |

### Corrección manual, incidencias y esperando alta (`routers/manual.py`)

| Método | Ruta | Descripción |
|---|---|---|
| `GET` | `/pendientes-lista` | Lista los PDFs en `corregir_manualmente` |
| `GET` | `/pendientes-completo-json` | Tabla completa de pendientes; lee de la caché `ColaRevision`, no vuelve a releer los PDFs en cada carga |
| `GET` | `/incidencias-completo-json` | Lista las facturas en `incidencias`; igual que pendientes, lee de la caché `ColaRevision` |
| `POST` | `/marcar-revisada` | Marca "Revisada" (Corregir manualmente o Incidencias): completa y mueve a `facturas_revisadas` sin exigir campos obligatorios |
| `GET` | `/esperando-alta-completo-json` | Lista las facturas en `esperando_alta` |
| `POST` | `/marcar-esperando-alta` | Aparta una factura (desde cualquier pestaña) a `esperando_alta` |
| `POST` | `/reprocesar-esperando-alta` | Reintenta resolver comprador/proveedor y mueve la factura si ya procede |
| `POST` | `/extraer-pendiente` | Extrae (o recupera del historial) los datos de un PDF pendiente; avisa si no parece una factura (endpoint heredado, sin llamada desde la UI actual — la tabla de "Corregir manualmente" ya carga cada pendiente vía `/pendientes-completo-json`) |
| `POST` | `/confirmar-factura` | Guarda la corrección, la persiste en SQL y mueve el PDF a `completadas` |
| `POST` | `/guardar-cambios-pendiente` | Guarda cambios en un pendiente sin confirmarlo/moverlo todavía |
| `POST` | `/descartar-pendiente` | Aparta un PDF de `corregir_manualmente` a `no_es_factura` sin borrarlo |

### Revisar facturas y duplicados

| Método | Ruta | Descripción |
|---|---|---|
| `GET` | `/facturas-completadas-json` | Lista las facturas en `FacturasExaminadas` |
| `POST` | `/marcar-factura-definitiva` | Marca/desmarca una factura como definitiva; mueve el PDF entre `completadas` y `facturas_revisadas` |
| `POST` | `/actualizar-factura-completada` | Edita los datos de una factura ya completada |
| `POST` | `/descartar-factura-completada` | Retira una factura de `FacturasExaminadas` y mueve el PDF a `no_es_factura` |
| `GET` | `/duplicados-completo-json` | Lista los pares de facturas marcadas como duplicado |
| `POST` | `/resolver-duplicado` | Descarta las copias elegidas y limpia el flag de duplicado en la que se conserva |
| `GET` | `/descargar-facturas-definitivas` | Descarga `facturas_100_definitivas.xlsx` |
| `POST` | `/exportar-excel-listado` | Exporta a Excel cualquier tabla mostrada en la interfaz |

### Reservas (multiusuario)

| Método | Ruta | Descripción |
|---|---|---|
| `POST` | `/reservas` | Reserva una lista de archivos para un usuario |
| `POST` | `/reservas-liberar` | Libera las reservas de un usuario sobre esos archivos |
| `POST` | `/reservas-latido` | Refresca (heartbeat) las reservas activas de un usuario |
| `GET` | `/reservas-estado` | Devuelve quién tiene reservado cada archivo de una lista |
| `GET` | `/reservas-mias` | Devuelve las reservas activas de un usuario |
| `POST` | `/reservas-liberar-todas` | Libera TODAS las reservas de cualquier usuario (válvula de escape) |

### PDF, posiciones y autocompletado de CIF

| Método | Ruta | Descripción |
|---|---|---|
| `GET` | `/pdf-factura/{archivo}` | Sirve el PDF de un pendiente/completada para verlo en el navegador |
| `GET` | `/pdf-factura-paginas/{archivo}` | Todas las páginas del PDF como imagen (para el visor propio de la vista de revisión) |
| `POST` | `/posiciones-factura` | Recuadro por campo (página + coordenadas) para pintarlo en el visor; se cachea por archivo |
| `POST` | `/ocr-region` | Selección manual con arrastre: OCR de la región indicada, para rellenar el campo activo |
| `GET` | `/lineas-factura/{archivo}` | Líneas de una factura (CSV junto al PDF); 404 si no tiene ninguna extraída |
| `POST` | `/lineas-factura/{archivo}` | Sobrescribe el CSV de líneas con el contenido final del modal de edición (admite añadir/quitar líneas) |
| `GET` | `/posiciones-lineas-factura/{archivo}` | Recuadro por línea (página + coordenadas) para resaltarla en el visor al pincharla; se cachea junto con `/posiciones-factura` |
| `POST` | `/reextraer-lineas-region` | Vuelve a llamar a la IA (vision) solo con la región de la tabla de líneas seleccionada a mano, en vez de la factura entera |
| `GET` | `/empresas-buscar` | Autocompletado de CIF de comprador contra `EmpresasClasificadas` |
| `GET` | `/proveedores-buscar` | Autocompletado de CIF de proveedor contra `ProveedoresClasificados` |

### Flujo secundario heredado

| Método | Ruta | Descripción |
|---|---|---|
| `POST` | `/extraer` | Extrae datos de uno o varios PDFs sueltos (endpoint heredado, sin vista asociada en la UI actual) |

---

## Dependencias principales

| Paquete | Uso |
|---|---|
| `fastapi` | Framework web |
| `uvicorn` | Servidor ASGI |
| `pdfplumber` | Extracción de texto de PDFs |
| `PyMuPDF` (`fitz`) | Renderizado de páginas PDF a imagen (segunda pasada con visión, visor de revisión, cálculo de posiciones) |
| `pytesseract` | OCR con Tesseract: posición de los datos en facturas escaneadas y selección manual con arrastre en `posiciones.py` (requiere tener instalado el programa Tesseract, no solo esta librería) |
| `openai` | Cliente OpenAI / Azure OpenAI |
| `openpyxl` | Generación y lectura de ficheros Excel (`.xlsx`) |
| `pyxlsb` | Lectura de `Empresas envasado.xlsb` (formato binario de Excel) en `cargar_empresas.py` |
| `pyodbc` | Conexión a SQL Server |
| `python-dotenv` | Carga de variables de entorno |
| `python-multipart` | Subida de ficheros |

Nota: `Pillow` (`PIL.Image`, usada en `posiciones.py`) no aparece como dependencia directa en `requirements.txt`; llega de forma transitiva a través de `pytesseract`/`PyMuPDF`, pero conviene tenerlo presente si se reconstruye el entorno desde cero.

`pywin32` ya **no** es dependencia del servidor (la automatización de Outlook se movió a `facturahelper`, ver [FacturaHelper](#facturahelper-correo-con-outlook-en-el-pc-del-usuario)); sigue apareciendo en `helper/requirements-build.txt`, solo necesario para compilar el `.exe` del helper.
