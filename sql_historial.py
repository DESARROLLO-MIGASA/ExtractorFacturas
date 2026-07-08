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
        columnas_existentes = {row[1] for row in cursor.execute(f"PRAGMA table_info({TABLA})").fetchall()}
        if "Definitiva" not in columnas_existentes:
            cursor.execute(f"ALTER TABLE {TABLA} ADD COLUMN Definitiva INTEGER DEFAULT 0")
        if "UsuarioDefinitiva" not in columnas_existentes:
            cursor.execute(f"ALTER TABLE {TABLA} ADD COLUMN UsuarioDefinitiva TEXT")
        if "FechaDefinitiva" not in columnas_existentes:
            cursor.execute(f"ALTER TABLE {TABLA} ADD COLUMN FechaDefinitiva TEXT")
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
    cursor.execute(f"""
        IF COL_LENGTH('{TABLA}', 'Definitiva') IS NULL
        ALTER TABLE {TABLA} ADD Definitiva BIT NOT NULL DEFAULT 0
    """)
    cursor.execute(f"""
        IF COL_LENGTH('{TABLA}', 'UsuarioDefinitiva') IS NULL
        ALTER TABLE {TABLA} ADD UsuarioDefinitiva NVARCHAR(100) NULL
    """)
    cursor.execute(f"""
        IF COL_LENGTH('{TABLA}', 'FechaDefinitiva') IS NULL
        ALTER TABLE {TABLA} ADD FechaDefinitiva DATETIME NULL
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


def listar_facturas_examinadas_sql():
    """
    Devuelve todas las filas de FacturasExaminadas como lista de dicts
    (una factura "completada", ya sea auto o confirmada manualmente).
    Incluye Origen, Definitiva, UsuarioDefinitiva y FechaDefinitiva.
    Devuelve [] si la consulta falla (no propaga excepciones).
    """
    columnas_todas = COLUMNAS + ["Origen", "Definitiva", "UsuarioDefinitiva", "FechaDefinitiva"]
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_si_no_existe(cursor)
            conn.commit()

            columnas_sql = ", ".join(f"[{c}]" for c in columnas_todas)
            cursor.execute(f"SELECT {columnas_sql} FROM {TABLA}")
            filas = cursor.fetchall()

        return [dict(zip(columnas_todas, fila)) for fila in filas]

    except Exception as e:
        print(f"AVISO: no se pudieron listar las facturas examinadas ({MOTOR}): {e}")
        return []


def marcar_factura_definitiva_sql(archivo, definitiva, usuario):
    """Marca (o desmarca) una factura de FacturasExaminadas como 100% definitiva."""
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_si_no_existe(cursor)
            conn.commit()

            cursor.execute(
                f"UPDATE {TABLA} SET Definitiva = ?, UsuarioDefinitiva = ?, "
                f"FechaDefinitiva = {FECHA_ACTUAL_SQL} WHERE Archivo = ?",
                (1 if definitiva else 0, usuario, archivo),
            )
            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudo marcar la factura como definitiva ({MOTOR}): {e}")
        return False


def eliminar_factura_examinada_sql(archivo):
    """Elimina una factura de FacturasExaminadas (p.ej. al descartarla como 'no es factura')."""
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_si_no_existe(cursor)
            conn.commit()

            cursor.execute(f"DELETE FROM {TABLA} WHERE Archivo = ?", (archivo,))
            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudo eliminar la factura de {MOTOR}: {e}")
        return False
