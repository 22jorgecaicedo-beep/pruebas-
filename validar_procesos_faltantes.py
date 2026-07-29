"""
Valida UNICAMENTE las carpetas de los procesos que aparecen en la lista
de faltantes (procesos_faltantes_en_disco.csv, la genera
validar_renombrar_carpetas.py) -- de los que YA tengan carpeta en el
disco (por ejemplo porque buscar_faltantes_en_drive.py ya los bajo).

Ademas de procesos_faltantes_en_disco.csv, tambien cruza contra
carpetas_vacias.csv (el otro reporte que genera
validar_renombrar_carpetas.py): si una carpeta de la lista de
faltantes ya existe en el disco pero esta COMPLETAMENTE VACIA (sin
ningun archivo adentro), se avisa aparte con claridad -- incluyendo si
ese reporte ya habia encontrado un .zip pendiente en Descargas para
ella -- en vez de tratarla en silencio como si tuviera contenido para
ordenar/revisar. Una carpeta vacia no se toca (no hay nada que ordenar
ni que borrar en ella).

Por defecto (BORRAR_ARCHIVOS_DE_OTRO_PROCESO = False), lo UNICO que
hace con las carpetas que SI tienen contenido es ordenarlas
cronologicamente, numerando sus documentos "1. ", "2. ", etc (el mas
viejo primero; ver ordenar_y_enumerar_carpeta) -- no mueve NADA entre
carpetas, no fusiona nada, no toca ninguna otra carpeta del disco.

Si ademas quieres que BORRE los archivos que parezcan de OTRO proceso
(porque mencionan un radicado corto distinto al de su propia carpeta,
o porque tienen un "CONTRA <algo>" que no corresponde al demandado real
del proceso segun el Excel), pon BORRAR_ARCHIVOS_DE_OTRO_PROCESO = True
mas abajo.

IMPORTANTE, a diferencia de TODOS los demas scripts de este proyecto
(que nunca borran nada, solo mueven a Duplicados_para_revisar): con
BORRAR_ARCHIVOS_DE_OTRO_PROCESO = True, los archivos que no coincidan
se BORRAN DE VERDAD, de forma PERMANENTE -- no quedan en
Duplicados_para_revisar, no se pueden recuperar. Esto fue pedido asi a
proposito (para no acumular carpetas de revision manual), pero es
irreversible: revisa con calma el reporte en MODO_PRUEBA antes de
correrlo con MODO_PRUEBA = False. Lo que SI se mantiene igual que en el
resto del proyecto es que la CARPETA en si nunca se mueve, renombra, ni
se fusiona con otra -- solo se borran archivos puntuales adentro.

IMPORTANTE: corre primero validar_renombrar_carpetas.py (para que
procesos_faltantes_en_disco.csv este al dia) antes de correr este
script. Si un proceso de la lista TODAVIA no tiene carpeta en el disco,
simplemente se cuenta como "sin carpeta" y se omite -- este script NO
descarga nada (para eso esta buscar_faltantes_en_drive.py).

A proposito, este script NO revisa todo el disco -- solo las carpetas
de los procesos que estan en la lista de faltantes. Si quieres una
revision de contaminacion/demandado/orden de TODAS las carpetas del
disco (no solo estas, y sin borrar nada -- ver revisar_contaminacion_en_disco/
revisar_demandado_en_disco en buscar_faltantes_en_drive.py, que mueven
a Duplicados_para_revisar en vez de borrar), esa ya la hace
buscar_faltantes_en_drive.py al empezar cada corrida. Este script existe
para cuando solo quieres revisar/ordenar rapido las de la lista de
faltantes, sin esperar a que se revise el disco completo.

Respeta MODO_PRUEBA (por defecto True): en modo prueba solo simula y
te dice que haria, sin borrar, mover ni renombrar nada todavia.
"""

import csv
import logging
import os
from pathlib import Path

import buscar_faltantes_en_drive as bfd
import procesos_juridicos as organizador  # noqa: F401 -- lo necesita bfd al importarlo
import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "validar_procesos_faltantes.log")

# True (por defecto): no borra, mueve ni renombra nada de verdad, solo
# revisa y muestra que haria. False: aplica los cambios de verdad.
MODO_PRUEBA = True

# False (por defecto): este script SOLO ordena cronologicamente los
# documentos que ya estan adentro de cada carpeta de la lista -- no
# toca ninguna otra carpeta del disco ni borra nada.
# True: ademas BORRA (de forma PERMANENTE, sin pasar por
# Duplicados_para_revisar) los archivos que parezcan de OTRO proceso --
# ver el modulo docstring arriba, es irreversible.
BORRAR_ARCHIVOS_DE_OTRO_PROCESO = False

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


def _leer_carpetas_vacias():
    """
    Lee carpetas_vacias.csv (lo genera validar_renombrar_carpetas.py --
    columnas "Carpeta;Radicado;Zip pendiente encontrado;Donde se
    encontro"). Devuelve {radicado: {"zip_encontrado":.., "zip_ubicacion":..}}.
    Si el archivo no existe (por ejemplo porque no hay ninguna carpeta
    vacia todavia), devuelve un diccionario vacio sin error.
    """
    ruta = cruce_excel.ARCHIVO_REPORTE_VACIAS
    if not ruta or not os.path.exists(ruta):
        return {}
    datos = {}
    try:
        with open(ruta, encoding="utf-8-sig") as f:
            for fila in csv.DictReader(f, delimiter=";"):
                radicado = (fila.get("Radicado") or "").strip()
                if radicado:
                    datos[radicado] = {
                        "zip_encontrado": (fila.get("Zip pendiente encontrado") or "").strip(),
                        "zip_ubicacion": (fila.get("Donde se encontro") or "").strip(),
                    }
    except OSError:
        return {}
    return datos


def _motivo_archivo_de_otro_proceso(archivo, cortos_propios, demandado_esperado):
    """
    Devuelve un texto explicando por que 'archivo' parece ser de OTRO
    proceso (para el mensaje de log), o None si no hay evidencia de eso.
      - Menciona un radicado corto DISTINTO al propio de la carpeta, y
        NUNCA el propio (si cita ambos, se asume que es un documento
        legitimo que solo referencia un caso relacionado).
      - Tiene un "CONTRA <algo>" que no corresponde al demandado
        esperado (ver _demandado_coincide_en_texto).
    """
    if cortos_propios:
        mencionados = bfd._radicados_cortos_mencionados(archivo.name)
        if mencionados and not (mencionados & cortos_propios):
            return f"menciona el radicado corto {', '.join(sorted(mencionados))}, pero no el propio de esta carpeta"

    if demandado_esperado and bfd._demandado_coincide_en_texto(archivo.name, demandado_esperado) is False:
        return f"parece ser CONTRA otro demandado, distinto a '{demandado_esperado}'"

    return None


def borrar_archivos_de_otro_proceso(carpetas):
    """
    Para cada carpeta en 'carpetas' (procesos de la lista de faltantes
    que ya tienen carpeta en el disco), revisa archivo por archivo y
    BORRA DEFINITIVAMENTE (no los mueve a Duplicados_para_revisar) los
    que parezcan de OTRO proceso -- ver _motivo_archivo_de_otro_proceso.
    La carpeta en si NUNCA se mueve, renombra, ni se fusiona con otra --
    solo se borran archivos puntuales adentro. Respeta MODO_PRUEBA.
    """
    demandados_por_radicado = {}
    if cruce_excel.RUTA_EXCEL and os.path.exists(cruce_excel.RUTA_EXCEL):
        try:
            demandados_por_radicado = cruce_excel.leer_demandados_por_radicado()
        except Exception:
            logging.exception("[Borrar] No se pudo leer el Excel para cruzar el demandado de cada proceso.")

    borrados = 0
    for carpeta in carpetas:
        radicado_carpeta = cruce_excel.radicado_de_nombre_carpeta(carpeta.name)
        if not radicado_carpeta:
            continue
        cortos_propios = set(bfd.radicados_cortos(radicado_carpeta))
        demandado_esperado = demandados_por_radicado.get(radicado_carpeta, "")

        try:
            archivos = [a for a in carpeta.rglob("*") if a.is_file()]
        except OSError:
            continue

        for archivo in archivos:
            motivo = _motivo_archivo_de_otro_proceso(archivo, cortos_propios, demandado_esperado)
            if not motivo:
                continue

            if MODO_PRUEBA:
                logging.info(
                    "[SIMULACION -- Borrar] '%s' (dentro de '%s'): %s -- se borraria PERMANENTEMENTE.",
                    archivo.name, carpeta.name, motivo,
                )
                continue

            try:
                archivo.unlink()
            except OSError as error:
                logging.warning("   (no se pudo borrar '%s': %s)", archivo, error)
                continue
            borrados += 1
            logging.warning(
                "[Borrado] '%s' (dentro de '%s'): %s -- se borro PERMANENTEMENTE.",
                archivo.name, carpeta.name, motivo,
            )

    if borrados:
        logging.info(
            "[Borrar] %d archivo(s) que parecian de otro proceso se borraron PERMANENTEMENTE (no quedaron "
            "en Duplicados_para_revisar).",
            borrados,
        )


def procesar():
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

    # --- Cruce con carpetas_vacias.csv: una carpeta que ya existe pero
    # esta COMPLETAMENTE VACIA no tiene nada que ordenar ni que
    # revisar/borrar -- se avisa aparte, en vez de tratarla en silencio
    # como si tuviera contenido. --------------------------------------------
    carpetas_con_contenido = []
    carpetas_vacias_objetivo = []
    for carpeta in carpetas_objetivo:
        if cruce_excel.contar_archivos(carpeta) == 0:
            carpetas_vacias_objetivo.append(carpeta)
        else:
            carpetas_con_contenido.append(carpeta)

    if carpetas_vacias_objetivo:
        info_vacias = _leer_carpetas_vacias()
        logging.warning(
            "%d de esas carpetas estan COMPLETAMENTE VACIAS (sin ningun archivo adentro) -- no hay nada que "
            "ordenar ni revisar en ellas todavia:",
            len(carpetas_vacias_objetivo),
        )
        for carpeta in carpetas_vacias_objetivo:
            radicado = cruce_excel.radicado_de_nombre_carpeta(carpeta.name)
            datos = info_vacias.get(radicado)
            if datos and datos.get("zip_encontrado"):
                logging.warning(
                    "   - '%s': VACIA -- carpetas_vacias.csv encontro un zip pendiente ('%s', en %s); corre "
                    "validar_renombrar_carpetas.py de nuevo (o extraelo a mano) para llenarla.",
                    carpeta.name, datos["zip_encontrado"], datos["zip_ubicacion"],
                )
            else:
                logging.warning(
                    "   - '%s': VACIA -- no hay ningun zip pendiente detectado en Descargas.",
                    carpeta.name,
                )

    carpetas_objetivo = carpetas_con_contenido
    if not carpetas_objetivo:
        logging.info("Ninguna de las carpetas de la lista tiene contenido para ordenar/revisar todavia.")
        return

    if BORRAR_ARCHIVOS_DE_OTRO_PROCESO:
        borrar_archivos_de_otro_proceso(carpetas_objetivo)
    else:
        logging.info(
            "(BORRAR_ARCHIVOS_DE_OTRO_PROCESO esta en False -- no se borra nada, solo se ordenan "
            "cronologicamente los documentos que ya estan en cada carpeta. Ninguna carpeta se toca aparte de "
            "las de la lista de faltantes.)"
        )

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se borro, ordeno ni renumero nada todavia. Revisa el log y, si se "
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
