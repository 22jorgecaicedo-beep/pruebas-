"""
Valida UNICAMENTE las carpetas de los procesos que aparecen en la lista
de faltantes (procesos_faltantes_en_disco.csv, la genera
validar_renombrar_carpetas.py) -- de los que YA tengan carpeta en el
disco (por ejemplo porque buscar_faltantes_en_drive.py ya los bajo).

Por defecto (REVISAR_CONTAMINACION_Y_DEMANDADO = False), lo UNICO que
hace es ordenar cronologicamente los documentos que YA estan adentro
de cada carpeta de la lista, numerandolos "1. ", "2. ", etc (el mas
viejo primero; ver ordenar_y_enumerar_carpeta) -- no mueve NADA entre
carpetas, no fusiona nada, no toca ninguna otra carpeta del disco.

Si ademas quieres que revise que cada carpeta solo tenga documentos de
SU PROPIO proceso (sacando archivos mezclados de otro proceso con
radicado parecido, o de un demandado distinto al del Excel -- ver
revisar_contaminacion_en_disco/revisar_demandado_en_disco en
buscar_faltantes_en_drive.py), pon
REVISAR_CONTAMINACION_Y_DEMANDADO = True mas abajo. Esas dos revisiones
SI pueden mover archivos (nunca los borran, los mueven a
Duplicados_para_revisar) -- por eso quedan apagadas por defecto.

IMPORTANTE: corre primero validar_renombrar_carpetas.py (para que
procesos_faltantes_en_disco.csv este al dia) antes de correr este
script. Si un proceso de la lista TODAVIA no tiene carpeta en el disco,
simplemente se cuenta como "sin carpeta" y se omite -- este script NO
descarga nada (para eso esta buscar_faltantes_en_drive.py).

A proposito, este script NO revisa todo el disco -- solo las carpetas
de los procesos que estan en la lista de faltantes. Si quieres una
revision de contaminacion/demandado/orden de TODAS las carpetas del
disco (no solo estas), esa ya la hace buscar_faltantes_en_drive.py
al empezar cada corrida (consolidar_duplicados_en_disco,
revisar_contaminacion_en_disco, revisar_demandado_en_disco,
ordenar_todas_las_carpetas_en_disco). Este script existe para cuando
solo quieres ordenar rapido las de la lista de faltantes, sin esperar
a que se revise el disco completo.

Respeta MODO_PRUEBA (por defecto True): en modo prueba solo simula y
te dice que haria, sin mover ni renombrar nada todavia.
"""

import logging
import os
from pathlib import Path

import buscar_faltantes_en_drive as bfd
import procesos_juridicos as organizador  # noqa: F401 -- lo necesita bfd al importarlo
import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "validar_procesos_faltantes.log")

# True (por defecto): no mueve ni renombra nada de verdad, solo revisa y
# muestra que haria. False: aplica los cambios de verdad.
MODO_PRUEBA = True

# False (por defecto): este script SOLO ordena cronologicamente los
# documentos que ya estan adentro de cada carpeta de la lista -- no
# mueve nada entre carpetas, no toca ninguna otra carpeta del disco.
# True: ademas revisa contaminacion entre procesos y demandado
# equivocado (puede MOVER archivos sospechosos a Duplicados_para_revisar,
# nunca los borra -- ver el modulo docstring arriba).
REVISAR_CONTAMINACION_Y_DEMANDADO = False

# ===========================================================================


def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def _mapa_carpetas_por_radicado():
    """{radicado de 23 digitos: Path} de cada carpeta de proceso que ya haya en CARPETA_PROCESOS."""
    carpeta_procesos = Path(bfd.CARPETA_PROCESOS)
    mapa = {}
    if not carpeta_procesos.exists():
        return mapa
    try:
        hijos = list(carpeta_procesos.iterdir())
    except OSError:
        return mapa
    for h in hijos:
        if not h.is_dir() or h.name == cruce_excel.NOMBRE_CARPETA_DUPLICADOS:
            continue
        radicado = cruce_excel.radicado_de_nombre_carpeta(h.name)
        if radicado:
            mapa[radicado] = h
    return mapa


def procesar():
    # revisar_contaminacion_en_disco/revisar_demandado_en_disco (en
    # buscar_faltantes_en_drive.py) leen el MODO_PRUEBA de ESE modulo,
    # no el de este script -- se sincroniza para que respeten el mismo
    # flag configurado aqui arriba (solo importa si
    # REVISAR_CONTAMINACION_Y_DEMANDADO esta en True).
    bfd.MODO_PRUEBA = MODO_PRUEBA

    faltantes = bfd.leer_faltantes()
    logging.info("Procesos de la lista de faltantes a validar: %d", len(faltantes))

    mapa_carpetas = _mapa_carpetas_por_radicado()

    carpetas_objetivo = []
    sin_carpeta_todavia = 0
    for fila in faltantes:
        radicado = fila["radicado"]
        if not radicado:
            continue
        carpeta = mapa_carpetas.get(radicado)
        if carpeta is None:
            sin_carpeta_todavia += 1
            continue
        carpetas_objetivo.append(carpeta)

    logging.info(
        "%d de %d proceso(s) de la lista ya tienen carpeta en el disco (los otros %d todavia no -- este "
        "script no descarga nada, corre buscar_faltantes_en_drive.py para eso).",
        len(carpetas_objetivo), len(faltantes), sin_carpeta_todavia,
    )

    if not carpetas_objetivo:
        logging.info("No hay ninguna carpeta de la lista de faltantes para validar todavia.")
        return

    if REVISAR_CONTAMINACION_Y_DEMANDADO:
        bfd.revisar_contaminacion_en_disco(carpetas_objetivo)
        bfd.revisar_demandado_en_disco(carpetas_objetivo)
    else:
        logging.info(
            "(REVISAR_CONTAMINACION_Y_DEMANDADO esta en False -- no se revisa si hay archivos mezclados de "
            "otro proceso/demandado, solo se ordenan cronologicamente los documentos que ya estan en cada "
            "carpeta. Ninguna carpeta se toca aparte de las de la lista de faltantes.)"
        )

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se ordeno/renumero ni se movio nada todavia. Revisa el log y, si se "
            "ve bien, cambia MODO_PRUEBA = False al inicio de este script y vuelve a correrlo."
        )
        return

    total_ordenados = 0
    for carpeta in carpetas_objetivo:
        try:
            total_ordenados += bfd.ordenar_y_enumerar_carpeta(carpeta)
        except OSError as error:
            logging.warning("   (no se pudo ordenar '%s': %s)", carpeta.name, error)

    if total_ordenados:
        logging.info(
            "[Orden] %d documento(s), entre las carpetas de la lista de faltantes, se ordenaron "
            "cronologicamente y se enumeraron (1., 2., ...).",
            total_ordenados,
        )

    logging.info("Listo -- carpetas de la lista de faltantes validadas.")


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
