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
import time

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
    "NumerosAlbaran",
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


_ESQUEMAS_ASEGURADOS = set()


def _una_vez_por_proceso(clave, func):
    """Envuelve una función "asegurar esquema" (CREATE TABLE IF NOT EXISTS +
    ALTER TABLE de migración con su propia comprobación previa) para que solo
    se ejecute la primera vez que se llama con esta `clave` en el proceso en
    marcha: el esquema no cambia mientras el proceso sigue vivo, así que
    repetir esas comprobaciones en cada guardado/lectura -como se hacía antes,
    en cada una de las decenas de llamadas de este módulo- solo añadía una
    ida y vuelta más al servidor sin ganar nada. Si el proceso se reinicia
    (p.ej. al desplegar un cambio), se vuelve a comprobar una vez, igual que
    la primera vez que arranca."""
    def envoltorio(cursor):
        if clave in _ESQUEMAS_ASEGURADOS:
            return
        func(cursor)
        _ESQUEMAS_ASEGURADOS.add(clave)
    return envoltorio


def _alter_columna_si_procede(cursor, descripcion, sentencia_sql):
    """Ejecuta un ALTER TABLE de migración "best effort": si el login de SQL
    no tiene permiso de DDL (o cualquier otro fallo puntual), se registra un
    aviso y se continúa sin la columna nueva, en vez de tirar abajo el
    guardado/lectura de facturas que depende de esta misma tabla."""
    try:
        cursor.execute(sentencia_sql)
    except Exception as e:
        print(f"AVISO: no se pudo migrar el esquema de {TABLA} ({descripcion}, {MOTOR}): {e}")


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
        if "NumerosAlbaran" not in columnas_existentes:
            _alter_columna_si_procede(cursor, "NumerosAlbaran", f"ALTER TABLE {TABLA} ADD COLUMN NumerosAlbaran TEXT")
        if "Definitiva" not in columnas_existentes:
            _alter_columna_si_procede(cursor, "Definitiva", f"ALTER TABLE {TABLA} ADD COLUMN Definitiva INTEGER DEFAULT 0")
        if "UsuarioDefinitiva" not in columnas_existentes:
            _alter_columna_si_procede(cursor, "UsuarioDefinitiva", f"ALTER TABLE {TABLA} ADD COLUMN UsuarioDefinitiva TEXT")
        if "FechaDefinitiva" not in columnas_existentes:
            _alter_columna_si_procede(cursor, "FechaDefinitiva", f"ALTER TABLE {TABLA} ADD COLUMN FechaDefinitiva TEXT")
        if "Duplicado" not in columnas_existentes:
            _alter_columna_si_procede(cursor, "Duplicado", f"ALTER TABLE {TABLA} ADD COLUMN Duplicado INTEGER DEFAULT 0")
        return

    # PedidoCliente y NumerosAlbaran pueden traer un número de valores
    # concatenados inusualmente alto (se ha visto una factura real con ~85
    # pedidos/~96 albaranes, muy por encima de 1000 caracteres), así que sin
    # límite de tamaño en vez de un ancho fijo que tarde o temprano se vuelva
    # a quedar corto.
    ANCHOS = {"PedidoCliente": "MAX", "NumerosAlbaran": "MAX"}
    columnas_sql = ",\n".join(f"[{c}] NVARCHAR({ANCHOS.get(c, 255)}) NULL" for c in COLUMNAS)

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
    # COL_LENGTH devuelve -1 para una columna ya NVARCHAR(MAX), así que estos
    # ALTER solo se repiten mientras la columna siga en un ancho fijo (255 por
    # defecto, o 1000 en las bases ya migradas antes de pasar a MAX).
    _alter_columna_si_procede(cursor, "PedidoCliente", f"""
        IF COL_LENGTH('{TABLA}', 'PedidoCliente') IS NOT NULL AND COL_LENGTH('{TABLA}', 'PedidoCliente') <> -1
        ALTER TABLE {TABLA} ALTER COLUMN [PedidoCliente] NVARCHAR(MAX) NULL
    """)
    _alter_columna_si_procede(cursor, "NumerosAlbaran", f"""
        IF COL_LENGTH('{TABLA}', 'NumerosAlbaran') IS NULL
        ALTER TABLE {TABLA} ADD NumerosAlbaran NVARCHAR(MAX) NULL
        ELSE IF COL_LENGTH('{TABLA}', 'NumerosAlbaran') <> -1
        ALTER TABLE {TABLA} ALTER COLUMN [NumerosAlbaran] NVARCHAR(MAX) NULL
    """)
    _alter_columna_si_procede(cursor, "Definitiva", f"""
        IF COL_LENGTH('{TABLA}', 'Definitiva') IS NULL
        ALTER TABLE {TABLA} ADD Definitiva BIT NOT NULL DEFAULT 0
    """)
    _alter_columna_si_procede(cursor, "UsuarioDefinitiva", f"""
        IF COL_LENGTH('{TABLA}', 'UsuarioDefinitiva') IS NULL
        ALTER TABLE {TABLA} ADD UsuarioDefinitiva NVARCHAR(100) NULL
    """)
    _alter_columna_si_procede(cursor, "FechaDefinitiva", f"""
        IF COL_LENGTH('{TABLA}', 'FechaDefinitiva') IS NULL
        ALTER TABLE {TABLA} ADD FechaDefinitiva DATETIME NULL
    """)
    _alter_columna_si_procede(cursor, "Duplicado", f"""
        IF COL_LENGTH('{TABLA}', 'Duplicado') IS NULL
        ALTER TABLE {TABLA} ADD Duplicado BIT NOT NULL DEFAULT 0
    """)


_crear_tabla_si_no_existe = _una_vez_por_proceso(TABLA, _crear_tabla_si_no_existe)


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
    Incluye Origen, Definitiva, UsuarioDefinitiva, FechaDefinitiva y Duplicado.
    Devuelve [] si la consulta falla (no propaga excepciones).

    Si la columna Duplicado todavía no existe físicamente en la tabla (p.ej.
    porque el login de SQL no tiene permiso de ALTER TABLE y la migración
    automática no pudo crearla), se repite la consulta sin ella para no
    romper el resto de pestañas; Duplicado queda a "0" en ese caso.
    """
    columnas_base = COLUMNAS + ["Origen", "Definitiva", "UsuarioDefinitiva", "FechaDefinitiva"]
    columnas_todas = columnas_base + ["Duplicado"]
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_si_no_existe(cursor)
            conn.commit()

            try:
                columnas_sql = ", ".join(f"[{c}]" for c in columnas_todas)
                cursor.execute(f"SELECT {columnas_sql} FROM {TABLA}")
                filas = cursor.fetchall()
                return [dict(zip(columnas_todas, fila)) for fila in filas]
            except Exception as e:
                print(f"AVISO: la columna Duplicado no está disponible en {TABLA} ({MOTOR}): {e}")
                columnas_sql = ", ".join(f"[{c}]" for c in columnas_base)
                cursor.execute(f"SELECT {columnas_sql} FROM {TABLA}")
                filas = cursor.fetchall()
                return [dict(zip(columnas_base, fila), Duplicado="0") for fila in filas]

    except Exception as e:
        print(f"AVISO: no se pudieron listar las facturas examinadas ({MOTOR}): {e}")
        return []


def obtener_factura_examinada_sql(archivo):
    """
    Como listar_facturas_examinadas_sql, pero para una sola factura (por
    Archivo, columna UNIQUE): devuelve su dict, o None si no existe.

    Pensada para los sitios que necesitan una fila conocida por su nombre
    -no recorrer la tabla entera para luego quedarse con una sola- como
    marcar_factura_definitiva o actualizar_factura_completada en logic.py:
    esas llamadas a listar_facturas_examinadas_sql() traían de golpe las
    ~1000 filas de la tabla (más de 1s contra el SQL Server remoto) solo
    para buscar la fila de UN archivo, en un punto del flujo -marcar
    definitiva, guardar una corrección- que se repite muchas veces seguidas
    mientras se revisan facturas.
    """
    columnas_base = COLUMNAS + ["Origen", "Definitiva", "UsuarioDefinitiva", "FechaDefinitiva"]
    columnas_todas = columnas_base + ["Duplicado"]
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_si_no_existe(cursor)
            conn.commit()

            try:
                columnas_sql = ", ".join(f"[{c}]" for c in columnas_todas)
                cursor.execute(f"SELECT {columnas_sql} FROM {TABLA} WHERE Archivo = ?", (archivo,))
                fila = cursor.fetchone()
                return dict(zip(columnas_todas, fila)) if fila else None
            except Exception as e:
                print(f"AVISO: la columna Duplicado no está disponible en {TABLA} ({MOTOR}): {e}")
                columnas_sql = ", ".join(f"[{c}]" for c in columnas_base)
                cursor.execute(f"SELECT {columnas_sql} FROM {TABLA} WHERE Archivo = ?", (archivo,))
                fila = cursor.fetchone()
                return dict(zip(columnas_base, fila), Duplicado="0") if fila else None

    except Exception as e:
        print(f"AVISO: no se pudo consultar la factura {archivo} ({MOTOR}): {e}")
        return None


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


def marcar_duplicado_sql(archivo, valor):
    """Marca (o desmarca) una factura de FacturasExaminadas como duplicado
    de otra ya existente (mismo NumeroFactura+Proveedor+Buyer)."""
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_si_no_existe(cursor)
            conn.commit()

            cursor.execute(
                f"UPDATE {TABLA} SET Duplicado = ? WHERE Archivo = ?",
                (1 if valor else 0, archivo),
            )
            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudo marcar la factura como duplicado ({MOTOR}): {e}")
        return False


TABLA_REVISADAS = "FacturasRevisadas"


def _crear_tabla_revisadas_si_no_existe(cursor):
    if MOTOR == "sqlite":
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {TABLA_REVISADAS} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Archivo TEXT NOT NULL UNIQUE,
                Revisada INTEGER NOT NULL DEFAULT 0,
                UsuarioRevisada TEXT,
                FechaRevisada TEXT
            )
        """)
        return

    cursor.execute(f"""
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '{TABLA_REVISADAS}')
        CREATE TABLE {TABLA_REVISADAS} (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            Archivo NVARCHAR(255) NOT NULL,
            Revisada BIT NOT NULL DEFAULT 0,
            UsuarioRevisada NVARCHAR(100) NULL,
            FechaRevisada DATETIME NULL,
            CONSTRAINT UQ_{TABLA_REVISADAS}_Archivo UNIQUE (Archivo)
        )
    """)


_crear_tabla_revisadas_si_no_existe = _una_vez_por_proceso(TABLA_REVISADAS, _crear_tabla_revisadas_si_no_existe)


def marcar_revisada_sql(archivo, revisada, usuario):
    """
    Marca (o desmarca) un archivo de "Corregir manualmente" o "Incidencias"
    como revisado por una persona. Es un estado propio de estas dos colas
    (una simple constancia visual de que alguien ya la miró), independiente
    de Definitiva/FacturasExaminadas -esas facturas todavía no están
    completas-, así que vive en su propia tabla.
    """
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_revisadas_si_no_existe(cursor)
            conn.commit()

            cursor.execute(
                f"UPDATE {TABLA_REVISADAS} SET Revisada = ?, UsuarioRevisada = ?, "
                f"FechaRevisada = {FECHA_ACTUAL_SQL} WHERE Archivo = ?",
                (1 if revisada else 0, usuario, archivo),
            )

            if cursor.rowcount == 0:
                cursor.execute(
                    f"INSERT INTO {TABLA_REVISADAS} (Archivo, Revisada, UsuarioRevisada, FechaRevisada) "
                    f"VALUES (?, ?, ?, {FECHA_ACTUAL_SQL})",
                    (archivo, 1 if revisada else 0, usuario),
                )

            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudo marcar {archivo} como revisada ({MOTOR}): {e}")
        return False


def archivos_revisados_sql():
    """Devuelve el conjunto de nombres de archivo marcados actualmente como
    revisados (Revisada = 1). Devuelve un conjunto vacío si la consulta
    falla, para que el llamador simplemente los trate como "No revisada"."""
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_revisadas_si_no_existe(cursor)
            conn.commit()

            cursor.execute(f"SELECT Archivo FROM {TABLA_REVISADAS} WHERE Revisada = 1")
            filas = cursor.fetchall()

        return {fila[0] for fila in filas}

    except Exception as e:
        print(f"AVISO: no se pudieron listar los archivos revisados ({MOTOR}): {e}")
        return set()


# =========================================================
# RESERVAS (varias personas revisando a la vez, sin pisarse)
# =========================================================
#
# Varias personas (2-6) revisan facturas a la vez desde distintos
# navegadores; para que no dos elijan la misma, cada una puede "reservar"
# un lote de archivos a su nombre. Es un estado puramente transitorio (no
# hay login real: el "usuario" es el que cada persona escribe una vez y el
# navegador recuerda, ver obtenerUsuarioActual() en el frontend), así que
# vive en su propia tabla, igual que FacturasRevisadas.
#
# La caducidad "de verdad" (cerrar la pestaña, o 30 minutos sin
# interactuar) la dispara el propio navegador llamando a
# liberar_reservas_sql explícitamente (ver el latido en el frontend);
# minutos_expiracion aquí es solo la red de seguridad del servidor para
# desconexiones bruscas en las que ese aviso explícito no llegó a mandarse
# (se cierra el portátil, se corta la red...).

TABLA_RESERVAS = "ReservasFacturas"


def _crear_tabla_reservas_si_no_existe(cursor):
    if MOTOR == "sqlite":
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {TABLA_RESERVAS} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Archivo TEXT NOT NULL UNIQUE,
                Usuario TEXT NOT NULL,
                FechaReserva TEXT NOT NULL,
                UltimaActividad TEXT NOT NULL
            )
        """)
        return

    cursor.execute(f"""
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '{TABLA_RESERVAS}')
        CREATE TABLE {TABLA_RESERVAS} (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            Archivo NVARCHAR(255) NOT NULL,
            Usuario NVARCHAR(100) NOT NULL,
            FechaReserva DATETIME NOT NULL,
            UltimaActividad DATETIME NOT NULL,
            CONSTRAINT UQ_{TABLA_RESERVAS}_Archivo UNIQUE (Archivo)
        )
    """)


_crear_tabla_reservas_si_no_existe = _una_vez_por_proceso(TABLA_RESERVAS, _crear_tabla_reservas_si_no_existe)


def _cutoff_sql(minutos):
    """Expresión SQL (sin parámetro, `minutos` es siempre un valor interno
    de confianza, no algo que escriba un usuario) para "hace N minutos",
    coherente con FECHA_ACTUAL_SQL de cada motor."""
    minutos = int(minutos)
    if MOTOR == "sqlite":
        return f"datetime('now', '-{minutos} minutes')"
    return f"DATEADD(MINUTE, -{minutos}, GETDATE())"


def reservar_facturas_sql(archivos, usuario, minutos_expiracion=3):
    """
    Intenta reservar cada archivo de `archivos` para `usuario`. Se queda con
    uno ya reservado si es del propio usuario (refresca su actividad), o si
    su UltimaActividad lleva más de `minutos_expiracion` sin refrescarse
    (reserva abandonada); si es de otro usuario y sigue vigente, no se toca.

    Devuelve {archivo: usuario_que_la_tiene_ahora} para TODOS los archivos
    pedidos (no solo los conseguidos), para que quien llama compare contra lo
    que pidió y sepa distinguir "conseguida" de "ya la tiene fulano".
    """
    archivos = list(dict.fromkeys(a for a in (archivos or []) if a))
    if not archivos or not usuario:
        return {}

    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_reservas_si_no_existe(cursor)
            conn.commit()

            corte = _cutoff_sql(minutos_expiracion)
            resultado = {}

            for archivo in archivos:
                cursor.execute(f"SELECT Usuario FROM {TABLA_RESERVAS} WHERE Archivo = ?", (archivo,))
                fila = cursor.fetchone()

                if fila is None:
                    cursor.execute(
                        f"INSERT INTO {TABLA_RESERVAS} (Archivo, Usuario, FechaReserva, UltimaActividad) "
                        f"VALUES (?, ?, {FECHA_ACTUAL_SQL}, {FECHA_ACTUAL_SQL})",
                        (archivo, usuario),
                    )
                    resultado[archivo] = usuario
                    continue

                # Comparación sin distinguir mayúsculas/espacios: aunque
                # quien llama ya normalice `usuario`, esto evita que una fila
                # ya guardada con otras mayúsculas/espacios de sobra (de
                # antes de esta normalización, o de un cliente que no la
                # aplicara) se lea como "de otra persona" siendo la misma.
                usuario_actual = (fila[0] or "").strip().lower()

                if usuario_actual == usuario:
                    cursor.execute(
                        f"UPDATE {TABLA_RESERVAS} SET UltimaActividad = {FECHA_ACTUAL_SQL} WHERE Archivo = ?",
                        (archivo,),
                    )
                    resultado[archivo] = usuario
                    continue

                cursor.execute(
                    f"SELECT 1 FROM {TABLA_RESERVAS} WHERE Archivo = ? AND UltimaActividad < {corte}",
                    (archivo,),
                )
                caducada = cursor.fetchone() is not None

                if caducada:
                    cursor.execute(
                        f"UPDATE {TABLA_RESERVAS} SET Usuario = ?, FechaReserva = {FECHA_ACTUAL_SQL}, "
                        f"UltimaActividad = {FECHA_ACTUAL_SQL} WHERE Archivo = ?",
                        (usuario, archivo),
                    )
                    resultado[archivo] = usuario
                else:
                    resultado[archivo] = usuario_actual

            conn.commit()

        return resultado

    except Exception as e:
        print(f"AVISO: no se pudieron reservar facturas ({MOTOR}): {e}")
        # None (no {}) a propósito: {} también es la respuesta legítima para
        # una lista de archivos vacía, y quien llama necesita distinguir "no
        # había nada que reservar" de "ha fallado la base de datos" -si no,
        # un fallo de conexión se malinterpreta como "ya las tiene otra
        # persona", que es justo el mensaje confuso que esto reemplaza-.
        return None


def liberar_reservas_sql(archivos, usuario):
    """Libera (borra) las reservas de `archivos` que sean de `usuario`.
    Nunca toca una reserva de otra persona, aunque se pida por su archivo."""
    archivos = [a for a in (archivos or []) if a]
    if not archivos or not usuario:
        return True

    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_reservas_si_no_existe(cursor)
            conn.commit()

            marcadores = ",".join("?" for _ in archivos)
            cursor.execute(
                f"DELETE FROM {TABLA_RESERVAS} WHERE LOWER(Usuario) = ? AND Archivo IN ({marcadores})",
                (usuario, *archivos),
            )
            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudieron liberar reservas ({MOTOR}): {e}")
        return False


def latido_reservas_sql(archivos, usuario):
    """Refresca UltimaActividad de las reservas de `archivos` que sean de
    `usuario`, para que no caduquen mientras se sigue trabajando en ellas."""
    archivos = [a for a in (archivos or []) if a]
    if not archivos or not usuario:
        return True

    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_reservas_si_no_existe(cursor)
            conn.commit()

            marcadores = ",".join("?" for _ in archivos)
            cursor.execute(
                f"UPDATE {TABLA_RESERVAS} SET UltimaActividad = {FECHA_ACTUAL_SQL} "
                f"WHERE LOWER(Usuario) = ? AND Archivo IN ({marcadores})",
                (usuario, *archivos),
            )
            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudo refrescar el latido de reservas ({MOTOR}): {e}")
        return False


def estado_reservas_sql(archivos, minutos_expiracion=3):
    """{archivo: usuario} de las reservas vigentes entre `archivos` (de paso
    borra las ya caducadas: limpieza perezosa, sin tarea programada aparte)."""
    archivos = [a for a in (archivos or []) if a]
    if not archivos:
        return {}

    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_reservas_si_no_existe(cursor)
            conn.commit()

            cursor.execute(f"DELETE FROM {TABLA_RESERVAS} WHERE UltimaActividad < {_cutoff_sql(minutos_expiracion)}")
            conn.commit()

            marcadores = ",".join("?" for _ in archivos)
            cursor.execute(
                f"SELECT Archivo, Usuario FROM {TABLA_RESERVAS} WHERE Archivo IN ({marcadores})",
                archivos,
            )
            filas = cursor.fetchall()

        return {archivo: usuario for archivo, usuario in filas}

    except Exception as e:
        print(f"AVISO: no se pudo consultar el estado de reservas ({MOTOR}): {e}")
        return {}


def liberar_todas_las_reservas_sql():
    """Vía de escape: vacía TODA la tabla de reservas, sin importar de quién
    sean. Pensada para desbloquear a mano si alguna reserva se queda
    "colgada" (p.ej. por una discrepancia de mayúsculas/espacios entre el
    usuario que la reservó y el que la vuelve a comprobar, o cualquier otro
    caso raro no cubierto por la caducidad automática)."""
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_reservas_si_no_existe(cursor)
            conn.commit()
            cursor.execute(f"DELETE FROM {TABLA_RESERVAS}")
            conn.commit()
        return True
    except Exception as e:
        print(f"AVISO: no se pudieron liberar todas las reservas ({MOTOR}): {e}")
        return False


def mis_reservas_sql(usuario, minutos_expiracion=3):
    """Archivos reservados ahora mismo por `usuario` (para recomponer su
    lote de trabajo si recarga la página)."""
    if not usuario:
        return []

    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_reservas_si_no_existe(cursor)
            conn.commit()

            cursor.execute(f"DELETE FROM {TABLA_RESERVAS} WHERE UltimaActividad < {_cutoff_sql(minutos_expiracion)}")
            conn.commit()

            cursor.execute(f"SELECT Archivo FROM {TABLA_RESERVAS} WHERE LOWER(Usuario) = ?", (usuario,))
            filas = cursor.fetchall()

        return [fila[0] for fila in filas]

    except Exception as e:
        print(f"AVISO: no se pudieron listar tus reservas ({MOTOR}): {e}")
        return []


def _candidatos_cif(cif):
    """
    Variantes de `cif` a probar, de más a menos específica. Las facturas
    intracomunitarias suelen traer el CIF con el prefijo de país ISO (p.ej.
    "ESB71406318"), mientras que en Business Central los CIF españoles se
    guardan sin él ("B71406318"), así que si la versión completa no
    encuentra nada se reintenta quitando ese prefijo de 2 letras.
    """
    candidatos = [cif]
    if len(cif) > 2 and cif[:2].isalpha():
        candidatos.append(cif[2:])
    return candidatos


def _cif_o_con_prefijo_pais(candidato):
    """
    Patrón LIKE (ANSI estándar, funciona igual en SQL Server y SQLite) que
    además de coincidir con `candidato` tal cual, coincide con el mismo
    valor precedido de un prefijo de país ISO de 2 letras ("__" = exactamente
    2 caracteres cualesquiera). Algunas facturas extranjeras (sobre todo
    portuguesas) imprimen el CIF/Contribuinte SIN el prefijo de país, aunque
    en Business Central esa empresa esté dada de alta CON él (p.ej. factura:
    "Contribuinte: 510449123", tabla maestra: "PT510449123"). _candidatos_cif
    ya cubre el caso contrario (quitar un prefijo que SÍ trae el valor
    extraído); esto cubre el que falta: prefijo que solo tiene la fila de la
    tabla maestra.
    """
    return "__" + candidato


def _mejor_estado(filas, campo_malo_es_true):
    """
    Reduce varias filas (nombre, flag) para un mismo CIF -puede haber más de
    una si la tabla tiene distintas Clasificacion para el mismo CIF, ver
    UNIQUE(CIF, Clasificacion)- a la "mejor": la que no está bloqueada/
    inactiva, si la hay entre ellas, en vez de quedarse con la primera que
    devuelva la consulta sin más (que podía ser una entrada antigua
    bloqueada aunque exista otra fila activa para el mismo CIF).
    `campo_malo_es_true` indica si el flag es "malo" cuando vale 1
    (Bloqueado) o cuando vale 0 (Activa). Mismo criterio que _mejor_por_cif,
    ya usado en el autocompletado de CIF.
    """
    mejor = None
    for nombre, flag in filas:
        malo = bool(flag) if campo_malo_es_true else not bool(flag)
        if mejor is None:
            mejor = (nombre, malo)
        if not malo:
            return nombre, malo
    return mejor


def estado_empresa_por_cif(cif):
    """
    Busca `cif` en EmpresasClasificadas y devuelve (nombre, estado), donde
    estado es uno de:
    - "activa": existe y Activa = 1 -> se puede usar como comprador.
    - "inactiva": existe pero Activa = 0 -> dada de baja, no se debe usar.
    - "no_encontrada": no hay ninguna fila con ese CIF, ni si quiera cuenta
      de fallo de consulta.

    A diferencia de la antigua buscar_empresa_por_cif (que colapsaba
    "inactiva" y "no_encontrada" en el mismo None), esto permite mostrarle al
    usuario un mensaje distinto en cada caso.
    """
    cif = (cif or "").strip().upper()
    if not cif:
        return None, "no_encontrada"
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            for candidato in _candidatos_cif(cif):
                cursor.execute(
                    "SELECT NombreEmpresa, Activa FROM EmpresasClasificadas "
                    "WHERE UPPER(CIF) = ? OR UPPER(CIF) LIKE ?",
                    (candidato, _cif_o_con_prefijo_pais(candidato)),
                )
                filas = cursor.fetchall()
                if filas:
                    nombre, inactiva = _mejor_estado(filas, campo_malo_es_true=False)
                    return nombre, ("inactiva" if inactiva else "activa")
        return None, "no_encontrada"
    except Exception as e:
        print(f"AVISO: no se pudo consultar EmpresasClasificadas ({MOTOR}): {e}")
        return None, "no_encontrada"


def estado_proveedor_por_cif(cif):
    """Análogo a estado_empresa_por_cif, pero contra ProveedoresClasificados;
    el estado "bloqueada" corresponde a Bloqueado = 1 en vez de Activa = 0."""
    cif = (cif or "").strip().upper()
    if not cif:
        return None, "no_encontrada"
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            for candidato in _candidatos_cif(cif):
                cursor.execute(
                    "SELECT NombreEmpresa, Bloqueado FROM ProveedoresClasificados "
                    "WHERE UPPER(CIF) = ? OR UPPER(CIF) LIKE ?",
                    (candidato, _cif_o_con_prefijo_pais(candidato)),
                )
                filas = cursor.fetchall()
                if filas:
                    nombre, bloqueado = _mejor_estado(filas, campo_malo_es_true=True)
                    return nombre, ("bloqueada" if bloqueado else "activa")
        return None, "no_encontrada"
    except Exception as e:
        print(f"AVISO: no se pudo consultar ProveedoresClasificados ({MOTOR}): {e}")
        return None, "no_encontrada"


def buscar_empresa_por_cif(cif):
    """Compatibilidad: solo el nombre, si existe y está activa; None en
    cualquier otro caso (no distingue "no encontrada" de "inactiva"). Para
    esa distinción usar estado_empresa_por_cif."""
    nombre, estado = estado_empresa_por_cif(cif)
    return nombre if estado == "activa" else None


def buscar_proveedor_por_cif(cif):
    """Compatibilidad: solo el nombre, si existe y no está bloqueado; None en
    cualquier otro caso. Para esa distinción usar estado_proveedor_por_cif."""
    nombre, estado = estado_proveedor_por_cif(cif)
    return nombre if estado == "activa" else None


def _mejor_por_cif(filas, campo_malo_es_true):
    """
    Reduce `filas` (tuplas (cif, nombre, flag)) a una por CIF, quedándose
    con la variante "buena" si existe alguna entre las distintas
    Clasificacion que puede tener un mismo CIF (ver la UNIQUE(CIF,
    Clasificacion) de estas tablas). `campo_malo_es_true` indica si el flag
    es "malo" cuando vale 1 (Bloqueado) o cuando vale 0 (Activa).
    """
    mejores = {}
    for cif, nombre, flag in filas:
        flag = bool(flag)
        malo = flag if campo_malo_es_true else not flag
        if cif not in mejores or (not malo and mejores[cif][1]):
            mejores[cif] = (nombre, malo)
    return [(cif, nombre, malo) for cif, (nombre, malo) in mejores.items()]


def _patrones_prefijo_cif(texto):
    """
    Patrones LIKE para buscar `texto` como prefijo de CIF en el
    autocompletado de Empresas/Proveedores, cubriendo el mismo desajuste de
    prefijo de país ISO que _candidatos_cif/_cif_o_con_prefijo_pais ya
    resuelven al clasificar automáticamente (ver esas dos funciones): si se
    escribe el CIF con el prefijo ("ESA19001304") pero en la tabla maestra
    está sin él ("A19001304"), o al revés, una búsqueda literal
    `CIF LIKE texto+'%'` no encontraba nada aunque fuera el mismo proveedor
    -y el autocompletado decía "sin coincidencias" de uno que sí está dado
    de alta y ya se usa correctamente en facturas procesadas automáticamente-.
    """
    sin_prefijo = texto[2:] if len(texto) > 2 and texto[:2].isalpha() else texto
    return [texto + "%", sin_prefijo + "%", "__" + texto + "%"]


def buscar_empresas_por_prefijo(texto, limite=20):
    """
    Autocompletado: hasta `limite` empresas de EmpresasClasificadas cuyo CIF
    empieza por `texto` o cuyo nombre lo contiene (para poder buscar tanto
    escribiendo el CIF como escribiendo parte del nombre, p.ej. "ybarra"
    encuentra todas las empresas de Ybarra dadas de alta). Se exige un texto
    de al menos 2 caracteres para no barrer la tabla entera con una consulta
    casi vacía. No se envuelve CIF en UPPER() para poder aprovechar el
    índice único de (CIF, Clasificacion); la collation habitual de SQL
    Server ya compara sin distinguir mayúsculas/minúsculas.

    Devuelve una lista de dicts {"cif", "nombre", "activa"}, sin duplicados
    por CIF (una empresa puede tener varias filas de Clasificacion).
    """
    texto = (texto or "").strip().upper()
    if len(texto) < 2:
        return []
    limite = int(limite)
    p1, p2, p3 = _patrones_prefijo_cif(texto)
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            if MOTOR == "sqlite":
                cursor.execute(
                    "SELECT CIF, NombreEmpresa, Activa FROM EmpresasClasificadas "
                    "WHERE CIF LIKE ? OR CIF LIKE ? OR CIF LIKE ? OR NombreEmpresa LIKE ? ORDER BY CIF LIMIT ?",
                    (p1, p2, p3, "%" + texto + "%", limite * 3),
                )
            else:
                cursor.execute(
                    f"SELECT TOP ({limite * 3}) CIF, NombreEmpresa, Activa FROM EmpresasClasificadas "
                    "WHERE CIF LIKE ? OR CIF LIKE ? OR CIF LIKE ? OR NombreEmpresa LIKE ? ORDER BY CIF",
                    (p1, p2, p3, "%" + texto + "%"),
                )
            filas = cursor.fetchall()
        resultado = _mejor_por_cif(filas, campo_malo_es_true=False)
        return [{"cif": c, "nombre": n, "activa": not malo} for c, n, malo in resultado[:limite]]
    except Exception as e:
        print(f"AVISO: no se pudo buscar EmpresasClasificadas por prefijo ({MOTOR}): {e}")
        return []


def buscar_proveedores_por_prefijo(texto, limite=20):
    """Análogo a buscar_empresas_por_prefijo (busca por CIF o por nombre),
    pero contra ProveedoresClasificados. Devuelve dicts {"cif", "nombre",
    "direccion", "poblacion", "bloqueado"} (direccion/poblacion pueden venir
    vacías: solo se cargan desde el export de Envasado, ver
    cargar_proveedores.py).

    Si las columnas Direccion/Poblacion todavía no existen físicamente en la
    tabla (p.ej. porque no se ha vuelto a ejecutar cargar_proveedores.py tras
    añadirlas, o el login de SQL no tiene permiso de ALTER TABLE), se repite
    la consulta sin ellas para no dejar la búsqueda entera sin resultados;
    direccion/poblacion quedan vacías en ese caso."""
    texto = (texto or "").strip().upper()
    if len(texto) < 2:
        return []
    limite = int(limite)
    p1, p2, p3 = _patrones_prefijo_cif(texto)
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            try:
                if MOTOR == "sqlite":
                    cursor.execute(
                        "SELECT CIF, NombreEmpresa, Direccion, Poblacion, Bloqueado FROM ProveedoresClasificados "
                        "WHERE CIF LIKE ? OR CIF LIKE ? OR CIF LIKE ? OR NombreEmpresa LIKE ? ORDER BY CIF LIMIT ?",
                        (p1, p2, p3, "%" + texto + "%", limite * 3),
                    )
                else:
                    cursor.execute(
                        f"SELECT TOP ({limite * 3}) CIF, NombreEmpresa, Direccion, Poblacion, Bloqueado "
                        "FROM ProveedoresClasificados "
                        "WHERE CIF LIKE ? OR CIF LIKE ? OR CIF LIKE ? OR NombreEmpresa LIKE ? ORDER BY CIF",
                        (p1, p2, p3, "%" + texto + "%"),
                    )
                filas = cursor.fetchall()
            except Exception as e:
                print(f"AVISO: Direccion/Poblacion no disponibles en ProveedoresClasificados ({MOTOR}): {e}")
                if MOTOR == "sqlite":
                    cursor.execute(
                        "SELECT CIF, NombreEmpresa, Bloqueado FROM ProveedoresClasificados "
                        "WHERE CIF LIKE ? OR CIF LIKE ? OR CIF LIKE ? OR NombreEmpresa LIKE ? ORDER BY CIF LIMIT ?",
                        (p1, p2, p3, "%" + texto + "%", limite * 3),
                    )
                else:
                    cursor.execute(
                        f"SELECT TOP ({limite * 3}) CIF, NombreEmpresa, Bloqueado FROM ProveedoresClasificados "
                        "WHERE CIF LIKE ? OR CIF LIKE ? OR CIF LIKE ? OR NombreEmpresa LIKE ? ORDER BY CIF",
                        (p1, p2, p3, "%" + texto + "%"),
                    )
                filas = [(cif, nombre, "", "", bloqueado) for cif, nombre, bloqueado in cursor.fetchall()]
        # Un mismo CIF puede traer varias filas (Granel/Envasado, ver la
        # UNIQUE(CIF, Clasificacion) de esta tabla); nos quedamos con la
        # variante no bloqueada si existe alguna.
        mejores = {}
        for cif, nombre, direccion, poblacion, bloqueado in filas:
            bloqueado = bool(bloqueado)
            actual = mejores.get(cif)
            if actual is None or (not bloqueado and actual[3]):
                mejores[cif] = (nombre, direccion or "", poblacion or "", bloqueado)
        return [
            {"cif": c, "nombre": n, "direccion": d, "poblacion": p, "bloqueado": b}
            for c, (n, d, p, b) in list(mejores.items())[:limite]
        ]
    except Exception as e:
        print(f"AVISO: no se pudo buscar ProveedoresClasificados por prefijo ({MOTOR}): {e}")
        return []


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


# =========================================================
# COLA DE REVISIÓN (Corregir manualmente / Incidencias)
# =========================================================
#
# Cachea los datos ya extraídos de cada PDF que está en corregir_manualmente/
# o en incidencias/, para que listar_pendientes_completo/
# listar_incidencias_completo (logic.py) no tengan que releer cada PDF ni
# recalcular Buyer/Proveedor en cada carga de página -eso era lo que hacía
# lenta cada recarga con una cola larga, sobre todo con las carpetas dentro
# de OneDrive-. Se escribe en el momento en que el PDF entra, sale o se
# corrige en una de estas dos carpetas (ver _mover_pdf_a_carpeta, mover_pdf y
# guardar_cambios_pendiente en logic.py); leer ya no recalcula nada.

TABLA_COLA = "ColaRevision"


def _crear_tabla_cola_si_no_existe(cursor):
    if MOTOR == "sqlite":
        columnas_sql = ",\n".join(f"[{c}] TEXT" for c in COLUMNAS)
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {TABLA_COLA} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                {columnas_sql},
                Cola TEXT NOT NULL,
                FechaActualizacion TEXT NOT NULL DEFAULT ({FECHA_ACTUAL_SQL}),
                UNIQUE([Archivo])
            )
        """)
        columnas_existentes = {row[1] for row in cursor.execute(f"PRAGMA table_info({TABLA_COLA})").fetchall()}
        if "NumerosAlbaran" not in columnas_existentes:
            _alter_columna_si_procede(cursor, "NumerosAlbaran", f"ALTER TABLE {TABLA_COLA} ADD COLUMN NumerosAlbaran TEXT")
        return

    ANCHOS = {"PedidoCliente": "MAX", "NumerosAlbaran": "MAX"}
    columnas_sql = ",\n".join(f"[{c}] NVARCHAR({ANCHOS.get(c, 255)}) NULL" for c in COLUMNAS)
    cursor.execute(f"""
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '{TABLA_COLA}')
        CREATE TABLE {TABLA_COLA} (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            {columnas_sql},
            Cola NVARCHAR(20) NOT NULL,
            FechaActualizacion DATETIME NOT NULL DEFAULT GETDATE(),
            CONSTRAINT UQ_{TABLA_COLA}_Archivo UNIQUE (Archivo)
        )
    """)
    # La tabla puede ya existir de antes de añadir/ensanchar estas columnas
    # (CREATE TABLE IF NOT EXISTS no toca una tabla ya creada), así que hace
    # falta el ALTER aparte -- ver el mismo caso para PedidoCliente en
    # _crear_tabla_si_no_existe, más arriba.
    try:
        cursor.execute(f"""
            IF COL_LENGTH('{TABLA_COLA}', 'PedidoCliente') IS NOT NULL AND COL_LENGTH('{TABLA_COLA}', 'PedidoCliente') <> -1
            ALTER TABLE {TABLA_COLA} ALTER COLUMN [PedidoCliente] NVARCHAR(MAX) NULL
        """)
    except Exception as e:
        print(f"AVISO: no se pudo migrar el esquema de {TABLA_COLA} (PedidoCliente, {MOTOR}): {e}")
    try:
        cursor.execute(f"""
            IF COL_LENGTH('{TABLA_COLA}', 'NumerosAlbaran') IS NULL
            ALTER TABLE {TABLA_COLA} ADD NumerosAlbaran NVARCHAR(MAX) NULL
            ELSE IF COL_LENGTH('{TABLA_COLA}', 'NumerosAlbaran') <> -1
            ALTER TABLE {TABLA_COLA} ALTER COLUMN [NumerosAlbaran] NVARCHAR(MAX) NULL
        """)
    except Exception as e:
        print(f"AVISO: no se pudo migrar el esquema de {TABLA_COLA} (NumerosAlbaran, {MOTOR}): {e}")


_crear_tabla_cola_si_no_existe = _una_vez_por_proceso(TABLA_COLA, _crear_tabla_cola_si_no_existe)


def guardar_en_cola_revision_sql(fila, cola, intentos=3, espera=2):
    """
    Inserta o actualiza (según 'Archivo') la fila cacheada de un PDF en
    corregir_manualmente/ (cola="pendiente") o incidencias/
    (cola="incidencia"). `fila` en el mismo orden que COLUMNAS/EXPECTED_HEADERS.

    El PDF ya se movió de carpeta cuando se llama a esto (ver mover_pdf /
    _mover_pdf_a_carpeta en logic.py), y listar_pendientes_completo /
    listar_incidencias_completo ya no leen la carpeta como respaldo: si esta
    escritura falla, el PDF queda huérfano (visible en disco, invisible en
    su pestaña) hasta que alguien lo repare a mano. Por eso se reintenta
    ante un corte transitorio de conexión antes de rendirse, y si aun así
    falla se registra como ERROR (no como aviso) para que no pase
    desapercibido.
    """
    datos = dict(zip(COLUMNAS, (list(fila) + ["-"] * len(COLUMNAS))[:len(COLUMNAS)]))

    for intento in range(1, intentos + 1):
        try:
            with _conectar() as conn:
                cursor = conn.cursor()
                _crear_tabla_cola_si_no_existe(cursor)
                conn.commit()

                columnas_sin_archivo = [c for c in COLUMNAS if c != "Archivo"]
                set_clause = ", ".join(f"[{c}] = ?" for c in columnas_sin_archivo)
                valores_update = [datos[c] for c in columnas_sin_archivo]

                cursor.execute(
                    f"UPDATE {TABLA_COLA} SET {set_clause}, Cola = ?, FechaActualizacion = {FECHA_ACTUAL_SQL} "
                    f"WHERE Archivo = ?",
                    (*valores_update, cola, datos["Archivo"]),
                )

                if cursor.rowcount == 0:
                    columnas_insert = ", ".join(f"[{c}]" for c in COLUMNAS)
                    placeholders = ", ".join("?" for _ in COLUMNAS)
                    cursor.execute(
                        f"INSERT INTO {TABLA_COLA} ({columnas_insert}, Cola) "
                        f"VALUES ({placeholders}, ?)",
                        (*[datos[c] for c in COLUMNAS], cola),
                    )

                conn.commit()

            return True

        except Exception as e:
            if intento < intentos:
                print(f"AVISO: fallo guardando {datos.get('Archivo')} en {TABLA_COLA} "
                      f"({cola}, {MOTOR}), intento {intento}/{intentos}: {e}")
                time.sleep(espera)
            else:
                print(f"ERROR: no se pudo guardar {datos.get('Archivo')} en {TABLA_COLA} "
                      f"({cola}, {MOTOR}) tras {intentos} intentos: {e}. "
                      "El PDF queda en su carpeta pero sin fila en ColaRevision "
                      "(quedará huérfano hasta que se repare)."
                      )

    return False


def actualizar_campos_cola_revision_sql(fila):
    """Actualiza los campos extraídos de una fila ya cacheada en ColaRevision
    (por 'Archivo'), sin tocar su 'Cola' actual. La usa guardar_cambios_pendiente
    (logic.py) al corregir a mano una pendiente/incidencia sin moverla de
    carpeta -no sabe (ni le hace falta saber) si esa fila es "pendiente" o
    "incidencia"-. Si el archivo todavía no está en la caché, no hace nada:
    entrará en ella la próxima vez que se clasifique o se recorra el backfill."""
    try:
        datos = dict(zip(COLUMNAS, (list(fila) + ["-"] * len(COLUMNAS))[:len(COLUMNAS)]))

        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_cola_si_no_existe(cursor)
            conn.commit()

            columnas_sin_archivo = [c for c in COLUMNAS if c != "Archivo"]
            set_clause = ", ".join(f"[{c}] = ?" for c in columnas_sin_archivo)
            valores_update = [datos[c] for c in columnas_sin_archivo]

            cursor.execute(
                f"UPDATE {TABLA_COLA} SET {set_clause}, FechaActualizacion = {FECHA_ACTUAL_SQL} "
                f"WHERE Archivo = ?",
                (*valores_update, datos["Archivo"]),
            )
            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudo actualizar la fila en {TABLA_COLA} ({MOTOR}): {e}")
        return False


def eliminar_de_cola_revision_sql(archivo):
    """Quita `archivo` de la caché de pendientes/incidencias: el PDF se movió
    fuera de ambas colas (completada, esperando alta, no es factura,
    revisada, etc.)."""
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_cola_si_no_existe(cursor)
            conn.commit()

            cursor.execute(f"DELETE FROM {TABLA_COLA} WHERE Archivo = ?", (archivo,))
            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudo eliminar {archivo} de {TABLA_COLA} ({MOTOR}): {e}")
        return False


def listar_cola_revision_sql(cola):
    """Todas las filas cacheadas de una cola ("pendiente" o "incidencia"),
    como listas de valores en el mismo orden que COLUMNAS. Devuelve [] si la
    consulta falla, para que el llamador se quede con una lista vacía en vez
    de romper la pestaña."""
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_cola_si_no_existe(cursor)
            conn.commit()

            columnas_sql = ", ".join(f"[{c}]" for c in COLUMNAS)
            cursor.execute(
                f"SELECT {columnas_sql} FROM {TABLA_COLA} WHERE Cola = ? ORDER BY Archivo",
                (cola,),
            )
            filas = cursor.fetchall()

        return [list(fila) for fila in filas]

    except Exception as e:
        print(f"AVISO: no se pudo listar {TABLA_COLA} ({cola}, {MOTOR}): {e}")
        return []


def buscar_incidencias_por_cif_sql(cif):
    """Filas de ColaRevision con Cola='incidencia' cuyo Buyer o Proveedor sea
    `cif`. La usa reclasificar_cola_por_cif (logic.py) para revisar solo las
    incidencias afectadas cuando se da de alta ese CIF concreto -en
    cargar_empresas.py/cargar_proveedores.py-, en vez de recorrer toda la
    cola en cada carga de página como se hacía antes."""
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_cola_si_no_existe(cursor)
            conn.commit()

            columnas_sql = ", ".join(f"[{c}]" for c in COLUMNAS)
            cursor.execute(
                f"SELECT {columnas_sql} FROM {TABLA_COLA} "
                f"WHERE Cola = 'incidencia' AND (Buyer = ? OR Proveedor = ?)",
                (cif, cif),
            )
            filas = cursor.fetchall()

        return [list(fila) for fila in filas]

    except Exception as e:
        print(f"AVISO: no se pudo buscar incidencias por CIF en {TABLA_COLA} ({MOTOR}): {e}")
        return []


# =========================================================
# AUDITORÍA (quién hizo qué acción, cuándo, sobre qué factura)
# =========================================================

TABLA_AUDITORIA = "AuditLog"


def _crear_tabla_auditoria_si_no_existe(cursor):
    if MOTOR == "sqlite":
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {TABLA_AUDITORIA} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Fecha TEXT NOT NULL DEFAULT ({FECHA_ACTUAL_SQL}),
                Usuario TEXT NOT NULL,
                Accion TEXT NOT NULL,
                Archivo TEXT,
                Detalle TEXT
            )
        """)
        return

    cursor.execute(f"""
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '{TABLA_AUDITORIA}')
        CREATE TABLE {TABLA_AUDITORIA} (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            Fecha DATETIME NOT NULL DEFAULT GETDATE(),
            Usuario NVARCHAR(100) NOT NULL,
            Accion NVARCHAR(50) NOT NULL,
            Archivo NVARCHAR(255) NULL,
            Detalle NVARCHAR(500) NULL
        )
    """)


_crear_tabla_auditoria_si_no_existe = _una_vez_por_proceso(TABLA_AUDITORIA, _crear_tabla_auditoria_si_no_existe)


def registrar_auditoria_sql(usuario, accion, archivo=None, detalle=None):
    """Deja constancia de una acción de negocio (quién, qué, cuándo, sobre
    qué factura). Es "best effort" como el resto del módulo: si falla, se
    avisa por consola y se sigue -nunca debe impedir que la acción real
    (guardar, mover el PDF, etc.) se complete-."""
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_auditoria_si_no_existe(cursor)
            conn.commit()

            cursor.execute(
                f"INSERT INTO {TABLA_AUDITORIA} (Fecha, Usuario, Accion, Archivo, Detalle) "
                f"VALUES ({FECHA_ACTUAL_SQL}, ?, ?, ?, ?)",
                (usuario, accion, archivo, detalle),
            )
            conn.commit()

        return True

    except Exception as e:
        print(f"AVISO: no se pudo registrar la auditoría ({accion}, {archivo}, {MOTOR}): {e}")
        return False


def listar_auditoria_sql(archivo=None, usuario=None, desde=None, hasta=None, limite=500):
    """Historial de acciones para la pestaña "Auditoría", más recientes
    primero. `archivo`/`usuario` filtran por coincidencia parcial; `desde`/
    `hasta` son fechas "YYYY-MM-DD" (se compara `hasta` hasta el final de
    ese día). Cualquier filtro no informado se ignora. Devuelve [] si la
    consulta falla."""
    condiciones = []
    parametros = []

    if archivo:
        condiciones.append("Archivo LIKE ?")
        parametros.append(f"%{archivo}%")
    if usuario:
        condiciones.append("Usuario LIKE ?")
        parametros.append(f"%{usuario}%")
    if desde:
        condiciones.append("Fecha >= ?")
        parametros.append(str(desde))
    if hasta:
        condiciones.append("Fecha <= ?")
        parametros.append(f"{hasta} 23:59:59")

    where_sql = f"WHERE {' AND '.join(condiciones)}" if condiciones else ""

    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_auditoria_si_no_existe(cursor)
            conn.commit()

            if MOTOR == "sqlite":
                cursor.execute(
                    f"SELECT Fecha, Usuario, Accion, Archivo, Detalle FROM {TABLA_AUDITORIA} "
                    f"{where_sql} ORDER BY Id DESC LIMIT ?",
                    (*parametros, int(limite)),
                )
            else:
                cursor.execute(
                    f"SELECT TOP (?) Fecha, Usuario, Accion, Archivo, Detalle FROM {TABLA_AUDITORIA} "
                    f"{where_sql} ORDER BY Id DESC",
                    (int(limite), *parametros),
                )
            filas = cursor.fetchall()

        return [list(fila) for fila in filas]

    except Exception as e:
        print(f"AVISO: no se pudo listar {TABLA_AUDITORIA} ({MOTOR}): {e}")
        return []


# Grupos de acciones que resume cada columna "¿quién...?" en Vista global y
# en la vista "por factura" de Auditoría. El segundo elemento de la tupla es
# la acción que debe ser la MÁS RECIENTE del grupo para que se muestre el
# usuario -si la más reciente es la contraria (p.ej. se desmarcó después de
# marcarse), la columna queda vacía, en vez de mostrar a quien la desmarcó
# como si la hubiera marcado-. `None` significa "siempre se muestra la más
# reciente" (no hay acción contraria en el grupo).
_GRUPOS_RESUMEN_AUDITORIA = {
    "EditadoPor":       (["editar_pendiente", "editar_completada"], None),
    "ConfirmadaPor":    (["confirmar_factura"], None),
    "RevisadaPor":      (["marcar_revisada", "desmarcar_revisada"], "marcar_revisada"),
    "DefinitivaPor":    (["marcar_definitiva", "desmarcar_definitiva"], "marcar_definitiva"),
    "EsperandoAltaPor": (["marcar_esperando_alta"], None),
    "NoEsFacturaPor":   (["descartar_pendiente", "descartar_completada"], None),
}


# Margen de seguridad bajo el límite de parámetros por consulta de SQL
# Server (~2100) y de SQLite en versiones antiguas (999): Vista global puede
# tener miles de facturas activas a la vez (con una sola consulta sin
# trocear, ya se ha visto fallar en producción con "Campo COUNT erróneo o
# error de sintaxis" al superar el límite), así que los archivos se
# consultan en lotes.
_MAX_PARAMETROS_POR_LOTE = 900


def resumen_auditoria_por_archivo_sql(archivos):
    """Para cada archivo de `archivos`, el usuario responsable de la última
    vez que ocurrió cada grupo de _GRUPOS_RESUMEN_AUDITORIA. Consulta
    AuditLog en lotes (no una consulta por archivo, ni una única consulta
    con todos a la vez -ver _MAX_PARAMETROS_POR_LOTE-), pensada para listas
    como la de Vista global. Devuelve {} si `archivos` está vacío o todas
    las consultas fallan."""
    archivos = [a for a in dict.fromkeys(archivos) if a]
    if not archivos:
        return {}

    acciones = sorted({a for accs, _ in _GRUPOS_RESUMEN_AUDITORIA.values() for a in accs})
    tamano_lote = max(1, _MAX_PARAMETROS_POR_LOTE - len(acciones))

    ultima_del_grupo = {}
    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_auditoria_si_no_existe(cursor)
            conn.commit()

            placeholders_accion = ", ".join("?" for _ in acciones)
            for inicio in range(0, len(archivos), tamano_lote):
                lote = archivos[inicio:inicio + tamano_lote]
                placeholders_archivo = ", ".join("?" for _ in lote)
                cursor.execute(
                    f"SELECT Archivo, Accion, Usuario FROM {TABLA_AUDITORIA} "
                    f"WHERE Archivo IN ({placeholders_archivo}) AND Accion IN ({placeholders_accion}) "
                    # Id (no Fecha) como criterio de "más reciente": Fecha es
                    # un DATETIME sin milisegundos útiles aquí y varias
                    # acciones pueden caer en el mismo segundo, lo que con
                    # solo Fecha DESC deja el orden entre ellas sin definir.
                    # Id es autoincremental y sí refleja el orden real de
                    # inserción.
                    f"ORDER BY Id DESC",
                    (*lote, *acciones),
                )
                # Cada lote solo trae archivos de ese lote (no se solapan
                # entre lotes), así que concatenar los resultados de todos
                # los lotes -cada uno ya ordenado por Id DESC- no cambia cuál
                # es la fila más reciente por (archivo, columna).
                for archivo, accion, usuario in cursor.fetchall():
                    for columna, (accs, _requerida) in _GRUPOS_RESUMEN_AUDITORIA.items():
                        clave = (archivo, columna)
                        if accion in accs and clave not in ultima_del_grupo:
                            ultima_del_grupo[clave] = (accion, usuario)

    except Exception as e:
        print(f"AVISO: no se pudo calcular el resumen de auditoría por archivo ({MOTOR}): {e}")
        if not ultima_del_grupo:
            return {}

    resumen = {}
    for archivo in archivos:
        fila_resumen = {}
        for columna, (_accs, requerida) in _GRUPOS_RESUMEN_AUDITORIA.items():
            dato = ultima_del_grupo.get((archivo, columna))
            if dato is None:
                fila_resumen[columna] = None
            else:
                accion, usuario = dato
                fila_resumen[columna] = usuario if (requerida is None or accion == requerida) else None
        resumen[archivo] = fila_resumen

    return resumen


def listar_archivos_auditoria_sql(archivo=None, usuario=None, desde=None, hasta=None, limite=200):
    """Archivos distintos con al menos una acción registrada en AuditLog,
    con los mismos filtros que listar_auditoria_sql, ordenados por su
    acción más reciente. Alimenta resumen_auditoria_por_archivo_sql para
    construir la vista "por factura" de la pestaña Auditoría. Devuelve []
    si la consulta falla."""
    condiciones = ["Archivo IS NOT NULL"]
    parametros = []

    if archivo:
        condiciones.append("Archivo LIKE ?")
        parametros.append(f"%{archivo}%")
    if usuario:
        condiciones.append("Usuario LIKE ?")
        parametros.append(f"%{usuario}%")
    if desde:
        condiciones.append("Fecha >= ?")
        parametros.append(str(desde))
    if hasta:
        condiciones.append("Fecha <= ?")
        parametros.append(f"{hasta} 23:59:59")

    where_sql = f"WHERE {' AND '.join(condiciones)}"

    try:
        with _conectar() as conn:
            cursor = conn.cursor()
            _crear_tabla_auditoria_si_no_existe(cursor)
            conn.commit()

            if MOTOR == "sqlite":
                cursor.execute(
                    f"SELECT Archivo FROM {TABLA_AUDITORIA} {where_sql} "
                    f"GROUP BY Archivo ORDER BY MAX(Id) DESC LIMIT ?",
                    (*parametros, int(limite)),
                )
            else:
                cursor.execute(
                    f"SELECT TOP (?) Archivo FROM {TABLA_AUDITORIA} {where_sql} "
                    f"GROUP BY Archivo ORDER BY MAX(Id) DESC",
                    (int(limite), *parametros),
                )
            filas = cursor.fetchall()

        return [fila[0] for fila in filas]

    except Exception as e:
        print(f"AVISO: no se pudo listar archivos de {TABLA_AUDITORIA} ({MOTOR}): {e}")
        return []
