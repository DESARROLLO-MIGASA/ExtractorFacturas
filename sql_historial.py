"""
Guarda en base de datos las facturas clasificadas como "examinada"
(extracción correcta), tanto del flujo automático como del manual.

Es un complemento del historial Excel, no un sustituto: si la conexión
a la base de datos falla o no está configurada, se registra un aviso y
el flujo de procesamiento de facturas continúa con normalidad.

Soporta dos motores, seleccionables con SQL_ENGINE en .env:
- "sqlserver" (por defecto): SQL Server vía pyodbc.
- "sqlite": archivo local, sin servidor ni permisos de red. Pensado como
  solución temporal mientras no hay acceso al SQL Server corporativo;
  misma tabla y columnas, así que migrar luego es un simple volcado.
"""

import os

from dotenv import load_dotenv

_ROOT = os.path.dirname(os.path.abspath(__file__))
load_dotenv(dotenv_path=os.path.join(_ROOT, ".env"))

MOTOR = os.getenv("SQL_ENGINE", "sqlserver").strip().lower()

SQL_SERVER = os.getenv("SQL_SERVER", "").strip()
SQL_DATABASE = os.getenv("SQL_DATABASE", "").strip()
SQL_DRIVER = os.getenv("SQL_DRIVER", "ODBC Driver 18 for SQL Server").strip()
# El Driver 18 exige cifrado y valida el certificado del servidor por defecto;
# los SQL Server internos suelen tener certificado autofirmado, así que se
# confía en él salvo que se indique lo contrario en .env.
SQL_TRUST_SERVER_CERTIFICATE = os.getenv("SQL_TRUST_SERVER_CERTIFICATE", "yes").strip()

_sqlite_path_cfg = os.getenv("SQLITE_PATH", "historial_facturas.db").strip()
SQLITE_PATH = _sqlite_path_cfg if os.path.isabs(_sqlite_path_cfg) else os.path.join(_ROOT, _sqlite_path_cfg)

FECHA_ACTUAL_SQL = "datetime('now')" if MOTOR == "sqlite" else "GETDATE()"

TABLA = "FacturasExaminadas"

COLUMNAS = [
    "Archivo",
    "BaseImp",
    "BaseIRPF",
    "Buyer",
    "Empresa",
    "FEscaneo",
    "FFactura",
    "FOperacion",
    "ImporIVA",
    "Moneda",
    "NombreProveedor",
    "NumeroFactura",
    "PedidoCliente",
    "Proveedor",
    "TipoIVA",
    "TipoIVA2",
    "TipoIVA3",
    "TotalFact",
]


def _conectar():
    if MOTOR == "sqlite":
        import sqlite3

        return sqlite3.connect(SQLITE_PATH, timeout=5)

    import pyodbc

    if not SQL_SERVER or not SQL_DATABASE:
        raise RuntimeError(
            "SQL_SERVER / SQL_DATABASE no están configurados en .env; "
            "no se puede guardar en SQL Server."
        )

    conn_str = (
        f"DRIVER={{{SQL_DRIVER}}};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={SQL_DATABASE};"
        "Trusted_Connection=yes;"
        f"TrustServerCertificate={SQL_TRUST_SERVER_CERTIFICATE};"
    )

    return pyodbc.connect(conn_str, timeout=5)


def _crear_tabla_si_no_existe(cursor):
    if MOTOR == "sqlite":
        columnas_sql = ",\n".join(f"[{c}] TEXT" for c in COLUMNAS)
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {TABLA} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                {columnas_sql},
                Origen TEXT,
                FechaInsercion TEXT NOT NULL DEFAULT ({FECHA_ACTUAL_SQL}),
                UNIQUE([Archivo])
            )
        """)
        return

    columnas_sql = ",\n".join(f"[{c}] NVARCHAR(255) NULL" for c in COLUMNAS)

    cursor.execute(f"""
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '{TABLA}')
        CREATE TABLE {TABLA} (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            {columnas_sql},
            Origen NVARCHAR(20) NULL,
            FechaInsercion DATETIME NOT NULL DEFAULT GETDATE(),
            CONSTRAINT UQ_{TABLA}_Archivo UNIQUE (Archivo)
        )
    """)


def guardar_factura_examinada_sql(fila, origen):
    """
    Inserta o actualiza (según 'Archivo') una factura "examinada" en SQL Server.

    fila:   lista de valores en el mismo orden que EXPECTED_HEADERS
            (Archivo, BaseImp, ..., TotalFact).
    origen: "auto" o "manual".

    No propaga excepciones: un fallo de SQL Server no debe romper el
    procesamiento de facturas ni el guardado del historial Excel.

    Devuelve True si se guardó correctamente, False si falló.
    """
    try:
        datos = dict(zip(COLUMNAS, (list(fila) + ["-"] * len(COLUMNAS))[:len(COLUMNAS)]))

        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_si_no_existe(cursor)
            conn.commit()

            columnas_sin_archivo = [c for c in COLUMNAS if c != "Archivo"]
            set_clause = ", ".join(f"[{c}] = ?" for c in columnas_sin_archivo)
            valores_update = [datos[c] for c in columnas_sin_archivo]

            cursor.execute(
                f"UPDATE {TABLA} SET {set_clause}, Origen = ?, FechaInsercion = {FECHA_ACTUAL_SQL} "
                f"WHERE Archivo = ?",
                (*valores_update, origen, datos["Archivo"]),
            )

            if cursor.rowcount == 0:
                columnas_insert = ", ".join(f"[{c}]" for c in COLUMNAS)
                placeholders = ", ".join("?" for _ in COLUMNAS)
                cursor.execute(
                    f"INSERT INTO {TABLA} ({columnas_insert}, Origen) "
                    f"VALUES ({placeholders}, ?)",
                    (*[datos[c] for c in COLUMNAS], origen),
                )

            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudo guardar la factura en {MOTOR} ({origen}): {e}")
        return False
