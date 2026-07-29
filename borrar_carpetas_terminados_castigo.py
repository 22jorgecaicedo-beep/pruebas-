"""
Borra las carpetas "placeholder" que creo crear_carpetas_terminados_castigo.py
para los procesos que NO estan en un estado activo (es decir, en
cualquier estado distinto de ESTADOS_A_CONTAR: ACTIVO, ACTIVOS CON
TITULOS, SUSPENDIDO, REORGANIZACION) -- funciona como un "deshacer" de
ese script.

Por cada numero de proceso del Excel que NO este en ESTADOS_A_CONTAR,
calcula el nombre EXACTO de carpeta que crear_carpetas_terminados_castigo.py
habria creado para el (misma regla: "<numero>. <ESTADO PROCESAL exacto>"
para TERMINADO*, o solo "<numero>." para REMITIDA*/NO INICIO*) y, SOLO
SI existe una carpeta en el disco con exactamente ese nombre Y esta
COMPLETAMENTE VACIA (sin ningun archivo adentro, ni siquiera en
subcarpetas), la borra.

Nunca toca:
  - Procesos en ESTADOS_A_CONTAR (ACTIVO, ACTIVOS CON TITULOS,
    SUSPENDIDO, REORGANIZACION) -- esos jamas se tocan, sin importar
    como se llame su carpeta.
  - Numeros de proceso duplicados en el Excel -- ambiguo, se deja para
    revisar a mano.
  - Estados sin clasificar (los que crear_carpetas_terminados_castigo.py
    tampoco crea) -- no hay nada que borrar para ellos.
  - Cualquier carpeta que NO este completamente vacia -- si tiene
    aunque sea un archivo adentro (por ejemplo porque le agregaste
    documentos a mano despues de crearla), se deja intacta y se
    reporta para que decidas que hacer.
  - Cualquier carpeta cuyo nombre no coincida EXACTAMENTE con el que
    crear_carpetas_terminados_castigo.py habria creado -- una carpeta
    organizada de la forma normal "numero. radicado" nunca se borra
    aqui.

ADVERTENCIA: borrar una carpeta es IRREVERSIBLE. Respeta MODO_PRUEBA
(por defecto True): en modo prueba solo simula y te dice que carpetas
borraria.

Usa la misma configuracion (RUTA_EXCEL, CARPETA_PROCESOS, etc) de
validar_renombrar_carpetas.py -- no hay que configurarla dos veces.
"""

import logging
import os
import shutil
from pathlib import Path

import validar_renombrar_carpetas as cruce_excel
import procesos_juridicos as organizador
import crear_carpetas_terminados_castigo as creador

# ============================= CONFIGURACION =============================

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "borrar_carpetas_terminados_castigo.log")

# True (por defecto): no borra ninguna carpeta de verdad, solo revisa y
# muestra que borraria. False: borra las carpetas de verdad (irreversible).
MODO_PRUEBA = True

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


def procesar():
    if not cruce_excel.RUTA_EXCEL or not os.path.exists(cruce_excel.RUTA_EXCEL):
        logging.error("[Excel] No se encontro el archivo configurado en RUTA_EXCEL: %r", cruce_excel.RUTA_EXCEL)
        return

    filas = creador.leer_procesos_completos()
    numero_a_filas = {}
    for fila, numero, radicado, estado in filas:
        numero_a_filas.setdefault(numero, []).append((fila, radicado, estado))

    carpeta_raiz = Path(cruce_excel.CARPETA_PROCESOS)
    if not carpeta_raiz.exists():
        logging.error("[Disco] No existe la carpeta configurada en CARPETA_PROCESOS: %s", carpeta_raiz)
        return

    try:
        carpetas_por_nombre = {
            d.name: d for d in carpeta_raiz.iterdir()
            if d.is_dir() and d.name not in cruce_excel.CARPETAS_A_IGNORAR
        }
    except OSError as error:
        logging.error("[Disco] No se pudo leer %s: %s", carpeta_raiz, error)
        return

    estados_ya_manejados = {e.upper() for e in cruce_excel.ESTADOS_A_CONTAR}

    borradas = 0
    no_vacias = []
    ambiguos = []

    for numero, entradas in sorted(numero_a_filas.items()):
        if len(entradas) > 1:
            ambiguos.append((numero, entradas))
            continue

        _fila, _radicado, estado = entradas[0]
        if not estado or estado.upper() in estados_ya_manejados:
            continue  # activo, suspendido, etc (o sin estado) -- jamas se toca aqui

        nombre_esperado = creador._nombre_carpeta_para(numero, estado)
        if nombre_esperado is None:
            continue  # estado sin clasificar -- este script tampoco lo creo, no hay nada que borrar

        carpeta = carpetas_por_nombre.get(nombre_esperado)
        if carpeta is None:
            continue  # no existe (nunca se creo, o ya se borro)

        if cruce_excel.contar_archivos(carpeta) != 0:
            no_vacias.append(carpeta.name)
            continue

        if MODO_PRUEBA:
            logging.info("[SIMULACION] Se borraria '%s' (proceso %s, estado '%s').", carpeta.name, numero, estado)
            continue

        try:
            shutil.rmtree(organizador._ruta_larga_segura(str(carpeta)))
            borradas += 1
            logging.info("[Borrada] '%s' (proceso %s, estado '%s').", carpeta.name, numero, estado)
        except OSError as error:
            logging.warning("   (no se pudo borrar '%s': %s)", carpeta.name, error)

    logging.info(
        "Resumen: %d carpeta(s) %s.",
        borradas, "simuladas para borrar (MODO_PRUEBA activo)" if MODO_PRUEBA else "borradas PERMANENTEMENTE",
    )

    if no_vacias:
        logging.warning(
            "%d carpeta(s) NO se tocaron porque ya tienen contenido adentro (alguien les agrego documentos) "
            "-- revisalas a mano si de verdad quieres vaciarlas:",
            len(no_vacias),
        )
        for nombre in no_vacias:
            logging.warning("   - %s", nombre)

    if ambiguos:
        logging.warning(
            "%d numero(s) de proceso aparecen mas de una vez en el Excel -- no se toco nada para ellos, "
            "revisalos a mano:",
            len(ambiguos),
        )
        for numero, entradas in ambiguos:
            detalle = ", ".join(
                f"fila {fila} (radicado {radicado!r}, estado {estado!r})" for fila, radicado, estado in entradas
            )
            logging.warning("   - Proceso %s aparece en: %s", numero, detalle)

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se borro ninguna carpeta todavia. Revisa el log y, si se ve bien, "
            "cambia MODO_PRUEBA = False al inicio de este script y vuelve a correrlo."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
