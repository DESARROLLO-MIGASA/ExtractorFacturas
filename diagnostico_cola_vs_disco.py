"""
Diagnóstico de solo lectura: compara, para "corregir_manualmente" e
"incidencias", lo que hay físicamente en la carpeta (lo que cuenta
/estadisticas) contra lo que dice la caché ColaRevision en SQL Server (lo
que cuentan las pestañas de arriba y sus listados).

No escribe nada ni en disco ni en la base de datos: solo os.listdir() y
SELECT. Se ejecuta desde esta misma carpeta para que logic.py/sql_historial.py
carguen el .env del proyecto tal cual lo hace la app.
"""

import os

from logic import CORREGIR_DIR, INCIDENCIAS_DIR
from sql_historial import listar_cola_revision_sql, COLUMNAS

IDX_ARCHIVO = COLUMNAS.index("Archivo")


def _archivos_en_disco(carpeta):
    if not os.path.exists(carpeta):
        return set()
    return {f for f in os.listdir(carpeta) if f.lower().endswith(".pdf")}


def _archivos_en_cola(cola):
    filas = listar_cola_revision_sql(cola)
    return {fila[IDX_ARCHIVO] for fila in filas}, len(filas)


def comparar(nombre, carpeta, cola):
    disco = _archivos_en_disco(carpeta)
    en_bd, total_filas_bd = _archivos_en_cola(cola)

    solo_disco = sorted(disco - en_bd)
    solo_bd = sorted(en_bd - disco)

    print(f"\n=== {nombre} ===")
    print(f"Carpeta: {carpeta}")
    print(f"  PDFs en disco:        {len(disco)}")
    print(f"  Filas en ColaRevision ({cola!r}): {total_filas_bd}"
          + (f"  [!] {total_filas_bd} != {len(en_bd)} filas únicas por Archivo" if total_filas_bd != len(en_bd) else ""))
    print(f"  En disco pero SIN fila en ColaRevision: {len(solo_disco)}")
    print(f"  En ColaRevision pero SIN pdf en la carpeta: {len(solo_bd)}")

    if solo_disco:
        print(f"\n  -- Archivos en '{nombre}' sin fila en ColaRevision (primeros 30) --")
        for f in solo_disco[:30]:
            print(f"    {f}")
        if len(solo_disco) > 30:
            print(f"    ... y {len(solo_disco) - 30} más")

    if solo_bd:
        print(f"\n  -- Filas en ColaRevision sin PDF en '{nombre}' (primeros 30) --")
        for f in solo_bd[:30]:
            print(f"    {f}")
        if len(solo_bd) > 30:
            print(f"    ... y {len(solo_bd) - 30} más")

    return solo_disco, solo_bd


if __name__ == "__main__":
    comparar("Corregir manualmente", CORREGIR_DIR, "pendiente")
    comparar("Incidencias", INCIDENCIAS_DIR, "incidencia")
