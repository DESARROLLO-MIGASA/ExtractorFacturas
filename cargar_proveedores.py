"""
Carga puntual: vuelca a la tabla SQL ProveedoresClasificados los proveedores
de Granel y Envasado leídos en caliente vía OData de Business Central (ya no
hace falta descargar y colocar a mano ningún .xlsx), con su CIF, nombre de
empresa, dirección/población (cuando la fuente las trae) y si están
bloqueados o no.

Un mismo CIF puede aparecer en Granel y Envasado (proveedor que sirve a
ambos); en ese caso se guardan dos filas, una por clasificación. Es seguro
re-ejecutar: cada fila se guarda por upsert sobre (CIF, Clasificacion), así
que repetir la carga no crea duplicados.

Usa el mismo motor configurado en .env para el resto de la app
(SQL_ENGINE=sqlserver o sqlite, ver sql_historial.py).

Credenciales OData (mismo usuario de servicio para ambas instancias de BC):
BC_ODATA_USER / BC_ODATA_PASSWORD en .env.

- Envasado: BC_ODATA_URL (servicio ProveedoresBloq, instancia "oleico"),
  trae dirección/población/código postal.
- Granel: BC_ODATA_GRANEL_LISTA_URL (EscanerIAListaProveedores: lista base
  con dirección/población/código postal y un bloqueo "genérico") +
  BC_ODATA_GRANEL_BLOQUEO_URL (EscanerIAConsultaBloProv: bloqueo específico
  de la empresa "Migasa Aceites, S.L.U.", vía Vendor_No). Cuando un
  proveedor tiene fila en este segundo servicio, su bloqueo manda sobre el
  genérico de la lista base.

Uso:
    python cargar_proveedores.py --dry-run   (solo cuenta y lista, no escribe nada)
    python cargar_proveedores.py             (carga de verdad)
"""

import os
import argparse

import requests
from requests_ntlm import HttpNtlmAuth

from sql_historial import _conectar, MOTOR, FECHA_ACTUAL_SQL

_ROOT = os.path.dirname(os.path.abspath(__file__))

BC_ODATA_USER = os.getenv("BC_ODATA_USER", "")
BC_ODATA_PASSWORD = os.getenv("BC_ODATA_PASSWORD", "")

BC_ODATA_URL = os.getenv("BC_ODATA_URL", (
    "http://oleico.domint.local:7048/NAVBC/ODataV4/"
    "Company(%27GRUPO%20YBARRA%20ALIMENTACION%20SL%27)/ProveedoresBloq"
))

BC_ODATA_GRANEL_LISTA_URL = os.getenv("BC_ODATA_GRANEL_LISTA_URL", (
    "http://olivar.domint.local:7048/MIGASA_PROD/ODataV4/"
    "Company(%27Migasa%20Aceites%2C%20S.L.U.%27)/EscanerIAListaProveedores"
))
BC_ODATA_GRANEL_BLOQUEO_URL = os.getenv("BC_ODATA_GRANEL_BLOQUEO_URL", (
    "http://olivar.domint.local:7048/MIGASA_PROD/ODataV4/"
    "Company(%27Migasa%20Aceites%2C%20S.L.U.%27)/EscanerIAConsultaBloProv"
))

TABLA = "ProveedoresClasificados"


def _alter_columna_si_procede(cursor, descripcion, sentencia_sql):
    """Migración de esquema "best effort": si el login de SQL no tiene
    permiso de DDL (o cualquier otro fallo puntual), se avisa y se sigue
    sin la columna nueva en vez de romper la carga completa."""
    try:
        cursor.execute(sentencia_sql)
    except Exception as e:
        print(f"AVISO: no se pudo migrar el esquema de {TABLA} ({descripcion}, {MOTOR}): {e}")


def _get_odata_json(url):
    resp = requests.get(
        url,
        auth=HttpNtlmAuth(BC_ODATA_USER, BC_ODATA_PASSWORD),
        headers={"Accept": "application/json"},
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json().get("value", [])


def leer_proveedores_granel_odata():
    """
    Lee los proveedores de Granel de dos servicios OData de Business
    Central (instancia "olivar", empresa Migasa Aceites, S.L.U.):

    - EscanerIAListaProveedores: lista base (No, Name, Address, City,
      Post_Code y un bloqueo "genérico" no ligado a una empresa concreta).
    - EscanerIAConsultaBloProv: bloqueo específico de esta empresa por
      proveedor (Vendor_No -> Blocked). Cuando un proveedor aparece aquí,
      su bloqueo manda sobre el genérico de la lista base.
    """
    bloqueo_especifico = {
        str(fila.get("Vendor_No") or "").strip(): str(fila.get("Blocked") or "").strip()
        for fila in _get_odata_json(BC_ODATA_GRANEL_BLOQUEO_URL)
    }

    proveedores = []
    for fila in _get_odata_json(BC_ODATA_GRANEL_LISTA_URL):
        cif = str(fila.get("No") or "").strip()
        if not cif:
            continue
        nombre = str(fila.get("Name") or "").strip()
        direccion = str(fila.get("Address") or "").strip()
        poblacion = str(fila.get("City") or "").strip()
        codigo_postal = str(fila.get("Post_Code") or "").strip()
        bloqueo = bloqueo_especifico.get(cif, str(fila.get("Blocked") or "").strip())
        proveedores.append(("Granel", cif, nombre, direccion, poblacion, codigo_postal, bool(bloqueo)))

    return proveedores


def leer_proveedores_envasado_odata():
    """
    Lee los proveedores de Envasado directamente del servicio OData de
    Business Central (ProveedoresBloq), en vez de un .xlsx exportado a
    mano.

    "Blocked" viene en blanco si el proveedor no está bloqueado, o con un
    motivo de bloqueo (p.ej. "All") en caso contrario.
    """
    proveedores = []
    for fila in _get_odata_json(BC_ODATA_URL):
        cif = str(fila.get("VAT_Registration_No") or "").strip()
        if not cif:
            continue
        nombre = str(fila.get("Name") or "").strip()
        direccion = str(fila.get("Address") or "").strip()
        poblacion = str(fila.get("City") or "").strip()
        codigo_postal = str(fila.get("Post_Code") or "").strip()
        bloqueado = bool(str(fila.get("Blocked") or "").strip())
        proveedores.append(("Envasado", cif, nombre, direccion, poblacion, codigo_postal, bloqueado))

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
        if "CodigoPostal" not in columnas_existentes:
            cursor.execute(f"ALTER TABLE {TABLA} ADD COLUMN CodigoPostal TEXT")
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
            CodigoPostal NVARCHAR(10) NULL,
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
    _alter_columna_si_procede(cursor, "CodigoPostal", f"""
        IF COL_LENGTH('{TABLA}', 'CodigoPostal') IS NULL
        ALTER TABLE {TABLA} ADD CodigoPostal NVARCHAR(10) NULL
    """)


def guardar_proveedor(cursor, clasificacion, cif, nombre, direccion, poblacion, codigo_postal, bloqueado):
    # La fuente (OData BC, entorno de pruebas DEPURADOR1) a veces trae en
    # Post_Code texto libre en vez de un código postal real (p.ej. un
    # fragmento de población); se trunca a lo que cabe en la columna en vez
    # de reventar la carga completa por un registro sucio.
    codigo_postal = (codigo_postal or "")[:10]
    bloqueado_val = 1 if bloqueado else 0

    cursor.execute(
        f"UPDATE {TABLA} SET NombreEmpresa = ?, Direccion = ?, Poblacion = ?, CodigoPostal = ?, Bloqueado = ?, "
        f"FechaCarga = {FECHA_ACTUAL_SQL} WHERE CIF = ? AND Clasificacion = ?",
        (nombre, direccion, poblacion, codigo_postal, bloqueado_val, cif, clasificacion),
    )
    if cursor.rowcount == 0:
        cursor.execute(
            f"INSERT INTO {TABLA} (Clasificacion, CIF, NombreEmpresa, Direccion, Poblacion, CodigoPostal, Bloqueado) "
            f"VALUES (?, ?, ?, ?, ?, ?, ?)",
            (clasificacion, cif, nombre, direccion, poblacion, codigo_postal, bloqueado_val),
        )


def main():
    parser = argparse.ArgumentParser(
        description="Carga los proveedores de Granel y Envasado (OData BC) en la tabla ProveedoresClasificados."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Solo cuenta y lista, no escribe nada en la base de datos.",
    )
    args = parser.parse_args()

    todos = []

    print(f"Leyendo Granel (OData BC): {BC_ODATA_GRANEL_LISTA_URL}")
    proveedores_granel = leer_proveedores_granel_odata()
    print(f"  {len(proveedores_granel)} proveedores")
    todos.extend(proveedores_granel)

    print(f"Leyendo Envasado (OData BC): {BC_ODATA_URL}")
    proveedores_envasado = leer_proveedores_envasado_odata()
    print(f"  {len(proveedores_envasado)} proveedores")
    todos.extend(proveedores_envasado)

    bloqueados = sum(1 for _, _, _, _, _, _, b in todos if b)
    print(f"\nTotal filas a guardar: {len(todos)} ({bloqueados} bloqueados)")

    if args.dry_run:
        for clasificacion, cif, nombre, direccion, poblacion, codigo_postal, bloqueado in todos[:15]:
            marca = " [BLOQUEADO]" if bloqueado else ""
            lugar = f" - {direccion}, {poblacion} {codigo_postal}".rstrip() if (direccion or poblacion or codigo_postal) else ""
            print(f"  [{clasificacion}] {cif} - {nombre}{lugar}{marca}")
        if len(todos) > 15:
            print(f"  ... y {len(todos) - 15} más")
        print(f"\nDry-run: no se ha escrito nada en {MOTOR}.")
        return

    with _conectar() as conn:
        cursor = conn.cursor()
        _crear_tabla_si_no_existe(cursor)
        conn.commit()

        for i, (clasificacion, cif, nombre, direccion, poblacion, codigo_postal, bloqueado) in enumerate(todos, start=1):
            guardar_proveedor(cursor, clasificacion, cif, nombre, direccion, poblacion, codigo_postal, bloqueado)
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
    cifs_utilizables = sorted({cif for _, cif, _, _, _, _, bloqueado in todos if not bloqueado})
    if cifs_utilizables:
        import logic

        print(f"\nRevisando incidencias afectadas por {len(cifs_utilizables)} CIF(s) desbloqueados...")
        total_movidas = sum(logic.reclasificar_cola_por_cif(cif) for cif in cifs_utilizables)
        print(f"Incidencias reclasificadas: {total_movidas}")


if __name__ == "__main__":
    main()
