"""
Carga puntual: vuelca a la tabla SQL ProveedoresClasificados los proveedores
de ProveedoresBC.xls (Granel) y ProveedoresBCEnvasado.xls (Envasado), con
su CIF, nombre de empresa y si están bloqueados o no.

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
    (os.path.join(_ROOT, "ProveedoresBC.xls"), "Granel"),
    (os.path.join(_ROOT, "ProveedoresBCEnvasado.xls"), "Envasado"),
]

TABLA = "ProveedoresClasificados"


def leer_proveedores(path, clasificacion):
    """
    Lee la hoja 'Suppliers' de un fichero exportado de Business Central
    (aunque la extensión sea .xls, el contenido es xlsx; se abre en modo
    binario para que openpyxl no lo rechace por la extensión).
    """
    with open(path, "rb") as f:
        wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb["Suppliers"]

    filas = list(ws.iter_rows(values_only=True))
    headers = filas[0]
    idx = {h: i for i, h in enumerate(headers)}

    proveedores = []
    for fila in filas[1:]:
        cif = str(fila[idx["VATRegistrationNumber"]] or "").strip()
        nombre = str(fila[idx["Name1"]] or "").strip()
        bloqueado = str(fila[idx["Blocked"]] or "").strip().upper() == "X"
        if not cif:
            continue
        proveedores.append((clasificacion, cif, nombre, bloqueado))

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
        return

    cursor.execute(f"""
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = '{TABLA}')
        CREATE TABLE {TABLA} (
            Id INT IDENTITY(1,1) PRIMARY KEY,
            Clasificacion NVARCHAR(20) NOT NULL,
            CIF NVARCHAR(30) NOT NULL,
            NombreEmpresa NVARCHAR(200) NOT NULL,
            Bloqueado BIT NOT NULL DEFAULT 0,
            FechaCarga DATETIME NOT NULL DEFAULT GETDATE(),
            CONSTRAINT UQ_{TABLA}_CIF_Clasificacion UNIQUE (CIF, Clasificacion)
        )
    """)


def guardar_proveedor(cursor, clasificacion, cif, nombre, bloqueado):
    bloqueado_val = 1 if bloqueado else 0

    cursor.execute(
        f"UPDATE {TABLA} SET NombreEmpresa = ?, Bloqueado = ?, FechaCarga = {FECHA_ACTUAL_SQL} "
        f"WHERE CIF = ? AND Clasificacion = ?",
        (nombre, bloqueado_val, cif, clasificacion),
    )
    if cursor.rowcount == 0:
        cursor.execute(
            f"INSERT INTO {TABLA} (Clasificacion, CIF, NombreEmpresa, Bloqueado) "
            f"VALUES (?, ?, ?, ?)",
            (clasificacion, cif, nombre, bloqueado_val),
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

    bloqueados = sum(1 for _, _, _, b in todos if b)
    print(f"\nTotal filas a guardar: {len(todos)} ({bloqueados} bloqueados)")

    if args.dry_run:
        for clasificacion, cif, nombre, bloqueado in todos[:15]:
            marca = " [BLOQUEADO]" if bloqueado else ""
            print(f"  [{clasificacion}] {cif} - {nombre}{marca}")
        if len(todos) > 15:
            print(f"  ... y {len(todos) - 15} más")
        print(f"\nDry-run: no se ha escrito nada en {MOTOR}.")
        return

    with _conectar() as conn:
        cursor = conn.cursor()
        _crear_tabla_si_no_existe(cursor)
        conn.commit()

        for i, (clasificacion, cif, nombre, bloqueado) in enumerate(todos, start=1):
            guardar_proveedor(cursor, clasificacion, cif, nombre, bloqueado)
            if i % 100 == 0:
                print(f"  {i}/{len(todos)} procesados...")

        conn.commit()

    print(f"\nCarga terminada: {len(todos)} proveedores guardados en {TABLA} ({MOTOR}).")


if __name__ == "__main__":
    main()
