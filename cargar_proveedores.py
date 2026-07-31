"""
Carga puntual: vuelca a la tabla SQL ProveedoresClasificados los proveedores
de ProveedoresGranel.xlsx (Granel) y ProveedoresEnvasado.xlsx (Envasado),
con su CIF, nombre de empresa y si están bloqueados o no.

Un mismo CIF puede aparecer en los dos ficheros (proveedor que sirve a
granel y envasado); en ese caso se guardan dos filas, una por
clasificación. Es seguro re-ejecutar: cada fila se guarda por upsert sobre
(CIF, Clasificacion), así que repetir la carga no crea duplicados.

Usa el mismo motor configurado en .env para el resto de la app
(SQL_ENGINE=sqlserver o sqlite, ver sql_historial.py).

Uso:
    python cargar_proveedores.py --dry-run   (solo cuenta y lista, no escribe nada)
    python cargar_proveedores.py             (carga de verdad)
"""

import os
import argparse

import openpyxl

from sql_historial import _conectar, MOTOR, FECHA_ACTUAL_SQL

_ROOT = os.path.dirname(os.path.abspath(__file__))

FICHEROS = [
    (os.path.join(_ROOT, "ProveedoresGranel.xlsx"), "Granel"),
    (os.path.join(_ROOT, "ProveedoresEnvasado.xlsx"), "Envasado"),
]

TABLA = "ProveedoresClasificados"


def _alter_columna_si_procede(cursor, descripcion, sentencia_sql):
    """Migración de esquema "best effort": si el login de SQL no tiene
    permiso de DDL (o cualquier otro fallo puntual), se avisa y se sigue
    sin la columna nueva en vez de romper la carga completa."""
    try:
        cursor.execute(sentencia_sql)
    except Exception as e:
        print(f"AVISO: no se pudo migrar el esquema de {TABLA} ({descripcion}, {MOTOR}): {e}")


def leer_proveedores(path, clasificacion):
    """
    Lee un fichero exportado de Business Central (vista de la tabla
    Proveedor). La fila 0 es el título de la vista, la fila 1 trae la
    fecha de exportación y la fila 2 la cabecera real; los datos empiezan
    en la fila 3. El orden de columnas difiere entre Granel y Envasado,
    así que se busca por nombre de columna. La columna "Bloqueado" no es
    booleana: viene en blanco si el proveedor no está bloqueado, o con un
    motivo de bloqueo (p.ej. "Todos", "Pago") en caso contrario.

    "Dirección"/"Dirección 2" y "Población" solo existen en el export de
    Envasado; en Granel quedan vacías porque Business Central no las trae
    en esa vista.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    filas = ws.iter_rows(values_only=True)

    next(filas)  # título
    next(filas)  # fecha de exportación
    headers = next(filas)
    idx = {h: i for i, h in enumerate(headers)}

    def _col(fila, nombre):
        i = idx.get(nombre)
        return str(fila[i] or "").strip() if i is not None else ""

    proveedores = []
    for fila in filas:
        cif = str(fila[idx["CIF/NIF"]] or "").strip()
        if not cif:
            continue
        nombre = str(fila[idx["Nombre"]] or "").strip()
        bloqueado = bool(str(fila[idx["Bloqueado"]] or "").strip())
        direccion = ", ".join(p for p in (_col(fila, "Dirección"), _col(fila, "Dirección 2")) if p)
        poblacion = _col(fila, "Población")
        proveedores.append((clasificacion, cif, nombre, direccion, poblacion, bloqueado))

    return proveedores


def _crear_tabla_si_no_existe(cursor):
    if MOTOR == "sqlite":
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {TABLA} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Clasificacion TEXT NOT NULL,
                CIF TEXT NOT NULL,
                NombreEmpresa TEXT NOT NULL,
                Bloqueado INTEGER NOT NULL DEFAULT 0,
                FechaCarga TEXT NOT NULL DEFAULT ({FECHA_ACTUAL_SQL}),
                UNIQUE(CIF, Clasificacion)
            )
        """)
        columnas_existentes = {row[1] for row in cursor.execute(f"PRAGMA table_info({TABLA})").fetchall()}
        if "Direccion" not in columnas_existentes:
            cursor.execute(f"ALTER TABLE {TABLA} ADD COLUMN Direccion TEXT")
        if "Poblacion" not in columnas_existentes:
            cursor.execute(f"ALTER TABLE {TABLA} ADD COLUMN Poblacion TEXT")
        return

    cursor.execute(f"""
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '{TABLA}')
        CREATE TABLE {TABLA} (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            Clasificacion NVARCHAR(20) NOT NULL,
            CIF NVARCHAR(30) NOT NULL,
            NombreEmpresa NVARCHAR(200) NOT NULL,
            Direccion NVARCHAR(200) NULL,
            Poblacion NVARCHAR(100) NULL,
            Bloqueado BIT NOT NULL DEFAULT 0,
            FechaCarga DATETIME NOT NULL DEFAULT GETDATE(),
            CONSTRAINT UQ_{TABLA}_CIF_Clasificacion UNIQUE (CIF, Clasificacion)
        )
    """)
    _alter_columna_si_procede(cursor, "Direccion", f"""
        IF COL_LENGTH('{TABLA}', 'Direccion') IS NULL
        ALTER TABLE {TABLA} ADD Direccion NVARCHAR(200) NULL
    """)
    _alter_columna_si_procede(cursor, "Poblacion", f"""
        IF COL_LENGTH('{TABLA}', 'Poblacion') IS NULL
        ALTER TABLE {TABLA} ADD Poblacion NVARCHAR(100) NULL
    """)


def guardar_proveedor(cursor, clasificacion, cif, nombre, direccion, poblacion, bloqueado):
    bloqueado_val = 1 if bloqueado else 0

    cursor.execute(
        f"UPDATE {TABLA} SET NombreEmpresa = ?, Direccion = ?, Poblacion = ?, Bloqueado = ?, "
        f"FechaCarga = {FECHA_ACTUAL_SQL} WHERE CIF = ? AND Clasificacion = ?",
        (nombre, direccion, poblacion, bloqueado_val, cif, clasificacion),
    )
    if cursor.rowcount == 0:
        cursor.execute(
            f"INSERT INTO {TABLA} (Clasificacion, CIF, NombreEmpresa, Direccion, Poblacion, Bloqueado) "
            f"VALUES (?, ?, ?, ?, ?, ?)",
            (clasificacion, cif, nombre, direccion, poblacion, bloqueado_val),
        )


def main():
    parser = argparse.ArgumentParser(
        description="Carga ProveedoresBC.xls y ProveedoresBCEnvasado.xls en la tabla ProveedoresClasificados."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Solo cuenta y lista, no escribe nada en la base de datos.",
    )
    args = parser.parse_args()

    todos = []
    for path, clasificacion in FICHEROS:
        print(f"Leyendo {clasificacion}: {path}")
        proveedores = leer_proveedores(path, clasificacion)
        print(f"  {len(proveedores)} proveedores")
        todos.extend(proveedores)

    bloqueados = sum(1 for _, _, _, _, _, b in todos if b)
    print(f"\nTotal filas a guardar: {len(todos)} ({bloqueados} bloqueados)")

    if args.dry_run:
        for clasificacion, cif, nombre, direccion, poblacion, bloqueado in todos[:15]:
            marca = " [BLOQUEADO]" if bloqueado else ""
            lugar = f" - {direccion}, {poblacion}" if (direccion or poblacion) else ""
            print(f"  [{clasificacion}] {cif} - {nombre}{lugar}{marca}")
        if len(todos) > 15:
            print(f"  ... y {len(todos) - 15} más")
        print(f"\nDry-run: no se ha escrito nada en {MOTOR}.")
        return

    with _conectar() as conn:
        cursor = conn.cursor()
        _crear_tabla_si_no_existe(cursor)
        conn.commit()

        for i, (clasificacion, cif, nombre, direccion, poblacion, bloqueado) in enumerate(todos, start=1):
            guardar_proveedor(cursor, clasificacion, cif, nombre, direccion, poblacion, bloqueado)
            if i % 100 == 0:
                print(f"  {i}/{len(todos)} procesados...")

        conn.commit()

    print(f"\nCarga terminada: {len(todos)} proveedores guardados en {TABLA} ({MOTOR}).")

    # Con el CIF ya confirmado en la base de datos (commit hecho arriba), se
    # revisan las incidencias que estuvieran atascadas esperando justo a
    # alguno de estos CIF: así se autorresuelven en cuanto se dan de alta o
    # se desbloquean, sin esperar a que alguien recargue "Incidencias" (ver
    # logic.reclasificar_cola_por_cif). Solo tiene sentido para los CIF que
    # quedan sin bloquear; uno bloqueado no resuelve ninguna incidencia.
    cifs_utilizables = sorted({cif for _, cif, _, _, _, bloqueado in todos if not bloqueado})
    if cifs_utilizables:
        import logic

        print(f"\nRevisando incidencias afectadas por {len(cifs_utilizables)} CIF(s) desbloqueados...")
        total_movidas = sum(logic.reclasificar_cola_por_cif(cif) for cif in cifs_utilizables)
        print(f"Incidencias reclasificadas: {total_movidas}")


if __name__ == "__main__":
    main()
