"""
Migración puntual: puebla la tabla SQL ColaRevision con el contenido actual
de corregir_manualmente/ e incidencias/, calculado con la misma lógica que
antes corría en cada carga de "Corregir manualmente"/"Incidencias"
(listar_pendientes_completo/listar_incidencias_completo en logic.py).

Necesaria una sola vez al desplegar la caché ColaRevision: la tabla nace
vacía y ambas colas pueden tener contenido ya acumulado. A partir de aquí la
caché se mantiene sola en el momento en que cada PDF entra, sale o se
corrige en una de las dos colas (ver _mover_pdf_a_carpeta, mover_pdf y
guardar_cambios_pendiente en logic.py); no hace falta volver a ejecutar este
script salvo que se quiera reconstruir la tabla desde cero.

De paso, al recalcular cada incidencia se reintenta su reclasificación
automática (por si alguna ya se hubiera podido resolver sola desde que cayó
en la cola); a partir de esta migración, esa reclasificación ya no se
repite en cada carga de página, sino solo cuando se da de alta el CIF que
le faltaba (ver reclasificar_cola_por_cif, cargar_empresas.py,
cargar_proveedores.py).

Es seguro re-ejecutar: cada fila se guarda por upsert sobre "Archivo" (ver
guardar_en_cola_revision_sql en sql_historial.py).

Uso:
    python migrar_cola_revision.py --dry-run   (solo cuenta y lista, no escribe nada)
    python migrar_cola_revision.py             (puebla la tabla de verdad)
"""

import argparse

from logic import (
    EXPECTED_HEADERS,
    listar_pendientes_lista,
    listar_incidencias_lista,
    cargar_pdf_pendiente_individual,
    buscar_en_historial,
    limpiar_fila,
    _refrescar_buyer_proveedor,
    _reclasificar_incidencia_si_procede,
)
from sql_historial import guardar_en_cola_revision_sql, MOTOR


def _fila_pendiente(archivo):
    fila, _fuente, _es_no_factura_flag = cargar_pdf_pendiente_individual(archivo)
    if fila is None:
        return [archivo] + ["-"] * (len(EXPECTED_HEADERS) - 1)
    fila = (list(fila) + ["-"] * len(EXPECTED_HEADERS))[:len(EXPECTED_HEADERS)]
    fila[0] = archivo
    return _refrescar_buyer_proveedor(fila)


def _fila_incidencia(archivo):
    """Devuelve (fila, reclasificada). `reclasificada` es la nueva
    clasificación ("completada"/"manual") si, al recalcular, la incidencia
    ya se ha resuelto sola: en ese caso _reclasificar_incidencia_si_procede
    ya la ha movido de carpeta y ha dejado la caché al día, así que no hay
    nada más que guardar aquí para ese archivo."""
    fila = buscar_en_historial(archivo)
    if fila is None:
        return [archivo] + ["-"] * (len(EXPECTED_HEADERS) - 1), None

    fila = limpiar_fila(list(fila))
    fila[0] = archivo
    fila = _refrescar_buyer_proveedor(fila)
    reclasificada = _reclasificar_incidencia_si_procede(archivo, fila)
    return fila, reclasificada


def main():
    parser = argparse.ArgumentParser(
        description="Puebla ColaRevision con el contenido actual de corregir_manualmente/ e incidencias/."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Solo cuenta y lista, no escribe nada en la base de datos.",
    )
    args = parser.parse_args()

    pendientes = listar_pendientes_lista()
    incidencias = listar_incidencias_lista()
    print(f"Pendientes en corregir_manualmente/: {len(pendientes)}")
    print(f"Incidencias en incidencias/: {len(incidencias)}")

    if args.dry_run:
        for archivo in pendientes[:15]:
            print(f"  [pendiente] {archivo}")
        if len(pendientes) > 15:
            print(f"  ... y {len(pendientes) - 15} más")
        for archivo in incidencias[:15]:
            print(f"  [incidencia] {archivo}")
        if len(incidencias) > 15:
            print(f"  ... y {len(incidencias) - 15} más")
        print(f"\nDry-run: no se ha escrito nada en {MOTOR} (tampoco se reclasifica ninguna incidencia).")
        return

    guardadas_pendientes = 0
    for i, archivo in enumerate(pendientes, start=1):
        fila = _fila_pendiente(archivo)
        if guardar_en_cola_revision_sql(fila, "pendiente"):
            guardadas_pendientes += 1
        if i % 50 == 0:
            print(f"  pendientes: {i}/{len(pendientes)} procesados...")

    guardadas_incidencias = 0
    reclasificadas = 0
    for i, archivo in enumerate(incidencias, start=1):
        fila, reclasificada = _fila_incidencia(archivo)
        if reclasificada is not None:
            reclasificadas += 1
        elif guardar_en_cola_revision_sql(fila, "incidencia"):
            guardadas_incidencias += 1
        if i % 50 == 0:
            print(f"  incidencias: {i}/{len(incidencias)} procesados...")

    print(
        f"\nMigración terminada: {guardadas_pendientes} pendientes y "
        f"{guardadas_incidencias} incidencias guardadas en ColaRevision ({MOTOR})."
    )
    if reclasificadas:
        print(f"{reclasificadas} incidencias se han reclasificado solas de paso (ya no estaban atascadas).")


if __name__ == "__main__":
    main()
