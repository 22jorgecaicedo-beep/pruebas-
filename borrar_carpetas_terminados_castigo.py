"""
Borra las carpetas vacias de procesos que NO estan en un estado activo
(es decir, en cualquier estado distinto de ESTADOS_A_CONTAR: ACTIVO,
ACTIVOS CON TITULOS, SUSPENDIDO, REORGANIZACION) -- funciona como un
"deshacer" de crear_carpetas_terminados_castigo.py, pero no se limita
solo a las carpetas que ese script creo: borra CUALQUIER carpeta del
numero de proceso (sin importar como se llame exactamente -- "numero.
ESTADO", solo "numero.", o incluso "numero. radicado") siempre que este
COMPLETAMENTE VACIA (sin ningun archivo adentro, ni siquiera en
subcarpetas).

Esto incluye tanto los estados clasificados (TERMINADO*, REMITIDA*, NO
INICIO*) como los que no encajan en ninguna regla conocida (ej.
"DESISTIMIENTO DE PRETENSIONES", "DEVUELTA INCURRIO EN GASTOS") -- si
el proceso no es activo y su carpeta esta vacia, se borra.

Nunca toca:
  - Procesos en ESTADOS_A_CONTAR (ACTIVO, ACTIVOS CON TITULOS,
    SUSPENDIDO, REORGANIZACION) -- esos jamas se tocan, sin importar
    como se llame su carpeta.
  - Filas sin ESTADO PROCESAL diligenciado en el Excel -- no se sabe
    todavia si son activas o no, se dejan intactas.
  - Numeros de proceso duplicados en el Excel -- ambiguo, se deja para
    revisar a mano.
  - Cualquier carpeta que NO este completamente vacia -- si tiene
    aunque sea un archivo adentro (por ejemplo porque le agregaste
    documentos a mano despues de crearla), se deja intacta y se
    reporta para que decidas que hacer.

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


def _carpetas_para_numero(numero, carpetas):
    """Todas las carpetas cuyo nombre empieza por '<numero>. ' o es exactamente '<numero>.'."""
    prefijo = f"{numero}. "
    exacto = f"{numero}."
    return [c for c in carpetas if c.name.startswith(prefijo) or c.name == exacto]


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
        carpetas = [
            d for d in carpeta_raiz.iterdir()
            if d.is_dir() and d.name not in cruce_excel.CARPETAS_A_IGNORAR
        ]
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
            continue  # activo, suspendido, etc (o sin estado diligenciado) -- jamas se toca aqui

        for carpeta in _carpetas_para_numero(numero, carpetas):
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
