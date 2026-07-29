"""
Crea carpetas "placeholder" en el disco para los procesos del informe
de Excel que estan en un ESTADO PROCESAL que el resto del proyecto NO
maneja (validar_renombrar_carpetas.py/buscar_faltantes_en_drive.py solo
cruzan/organizan los estados de ESTADOS_A_CONTAR: ACTIVO, ACTIVOS CON
TITULOS, SUSPENDIDO, REORGANIZACION) -- procesos TERMINADOS (por pago,
por auto, por prepago, con contrato, etc), REMITIDOS (a castigo, a
prepago, etc), y los que nunca llegaron a iniciarse (NO INICIO).

Estos procesos normalmente no tienen (o no importa) un radicado de 23
digitos real para organizar documentos -- son mas un registro
administrativo que un expediente judicial activo -- asi que la carpeta
NO se nombra "numero. radicado" como el resto del proyecto, sino:

  - Si el ESTADO PROCESAL empieza con "TERMINADO" (por pago, por auto,
    por prepago, con contrato, etc): "<numero>. <ESTADO PROCESAL
    EXACTO del Excel>" -- ej. "123. TERMINADO POR AUTO".
  - Si el ESTADO PROCESAL empieza con "REMITIDA" (a castigo, a
    prepago, etc) o con "NO INICIO": SOLO el numero -- ej. "145.".

Los procesos con un ESTADO PROCESAL que no encaje en ninguna de las
dos reglas de arriba (ej. "DESISTIMIENTO DE PRETENSIONES", "DEVUELTA
INCURRIO EN GASTOS") se dejan FUERA a proposito -- no se crea carpeta
para ellos, y quedan listados en el log para que decidas que hacer.

Si YA existe una carpeta en el disco para ese numero (cualquier nombre
que empiece por "<numero>. " o sea exactamente "<numero>." -- por
ejemplo porque el proceso ya se organizo de la forma normal con su
radicado real), NO se crea una nueva. Este script NUNCA borra, renombra
ni mueve nada -- solo CREA carpetas vacias nuevas donde todavia no
exista ninguna para ese numero.

Usa la misma configuracion (RUTA_EXCEL, HOJA_EXCEL, CARPETA_PROCESOS,
etc) de validar_renombrar_carpetas.py -- no hay que configurarla dos
veces.

Respeta MODO_PRUEBA (por defecto True): en modo prueba solo simula y
te dice que carpetas crearia.
"""

import logging
import os
from pathlib import Path

import openpyxl

import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "crear_carpetas_terminados_castigo.log")

# True (por defecto): no crea ninguna carpeta de verdad, solo revisa y
# muestra que crearia. False: crea las carpetas de verdad.
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


def _nombre_carpeta_para(numero: int, estado: str):
    """Nombre de carpeta segun el ESTADO PROCESAL, o None si no encaja en ninguna regla conocida."""
    estado_norm = estado.strip().upper()
    if estado_norm.startswith("TERMINADO"):
        return f"{numero}. {estado.strip()}"
    if estado_norm.startswith("REMITIDA") or estado_norm.startswith("NO INICIO"):
        return f"{numero}."
    return None


def _existe_carpeta_para_numero(numero: int, carpetas_existentes) -> bool:
    """True si YA hay una carpeta en el disco que empieza por '<numero>. ' o es exactamente '<numero>.'."""
    prefijo = f"{numero}. "
    exacto = f"{numero}."
    return any(nombre.startswith(prefijo) or nombre == exacto for nombre in carpetas_existentes)


def leer_procesos_por_estado():
    """
    Lee el Excel completo y devuelve [(fila, numero, estado_procesal), ...]
    para CADA fila con un numero de proceso valido (columna No.) y un
    ESTADO PROCESAL no vacio -- sin filtrar por radicado (a diferencia
    de leer_filas_excel, que exige un radicado valido).
    """
    wb = openpyxl.load_workbook(cruce_excel.RUTA_EXCEL, data_only=True)
    if cruce_excel.HOJA_EXCEL not in wb.sheetnames:
        raise ValueError(
            f"La hoja '{cruce_excel.HOJA_EXCEL}' no existe. Hojas disponibles: {wb.sheetnames}"
        )
    ws = wb[cruce_excel.HOJA_EXCEL]
    encabezados = [ws.cell(row=cruce_excel.FILA_ENCABEZADO, column=c).value for c in range(1, ws.max_column + 1)]
    col_no = cruce_excel.encontrar_columna(encabezados, cruce_excel.COLUMNA_NO)
    col_estado = cruce_excel.encontrar_columna(encabezados, cruce_excel.COLUMNA_ESTADO)

    filas = []
    for fila in range(cruce_excel.FILA_ENCABEZADO + 1, ws.max_row + 1):
        numero = ws.cell(row=fila, column=col_no).value
        estado = ws.cell(row=fila, column=col_estado).value
        if numero is None or not isinstance(numero, (int, float)):
            continue  # fila vacia o de notas/leyenda al final de la hoja
        if estado is None or not str(estado).strip():
            continue  # sin estado procesal diligenciado todavia
        filas.append((fila, int(numero), str(estado).strip()))
    return filas


def procesar():
    if not cruce_excel.RUTA_EXCEL or not os.path.exists(cruce_excel.RUTA_EXCEL):
        logging.error("[Excel] No se encontro el archivo configurado en RUTA_EXCEL: %r", cruce_excel.RUTA_EXCEL)
        return

    filas = leer_procesos_por_estado()
    logging.info("Excel: %d fila(s) con numero de proceso y estado procesal diligenciados.", len(filas))

    carpeta_raiz = Path(cruce_excel.CARPETA_PROCESOS)
    carpeta_raiz.mkdir(parents=True, exist_ok=True)
    try:
        carpetas_existentes = {
            d.name for d in carpeta_raiz.iterdir()
            if d.is_dir() and d.name not in cruce_excel.CARPETAS_A_IGNORAR
        }
    except OSError as error:
        logging.error("[Disco] No se pudo leer %s: %s", carpeta_raiz, error)
        return

    # Los estados que ya maneja el resto del proyecto (ACTIVO, etc) se
    # omiten aqui sin avisar -- ya los organiza validar_renombrar_carpetas.py/
    # buscar_faltantes_en_drive.py de la forma normal ("numero. radicado").
    estados_ya_manejados = {e.upper() for e in cruce_excel.ESTADOS_A_CONTAR}

    creadas = 0
    ya_existian = 0
    sin_clasificar = []

    for fila, numero, estado in filas:
        if estado.upper() in estados_ya_manejados:
            continue

        nombre_carpeta = _nombre_carpeta_para(numero, estado)
        if nombre_carpeta is None:
            sin_clasificar.append((fila, numero, estado))
            continue

        if _existe_carpeta_para_numero(numero, carpetas_existentes):
            ya_existian += 1
            continue

        if MODO_PRUEBA:
            logging.info("[SIMULACION] Se crearia '%s' (fila %s, estado '%s').", nombre_carpeta, fila, estado)
            continue

        destino = carpeta_raiz / nombre_carpeta
        try:
            destino.mkdir()
        except FileExistsError:
            ya_existian += 1
            continue
        except OSError as error:
            logging.warning("   (no se pudo crear '%s': %s)", nombre_carpeta, error)
            continue
        carpetas_existentes.add(nombre_carpeta)
        creadas += 1
        logging.info("[Creada] '%s' (fila %s, estado '%s').", nombre_carpeta, fila, estado)

    logging.info(
        "Resumen: %d carpeta(s) %s, %d ya existian, %d fila(s) con un estado procesal sin clasificar.",
        creadas, "simuladas (MODO_PRUEBA activo)" if MODO_PRUEBA else "creadas", ya_existian, len(sin_clasificar),
    )

    if sin_clasificar:
        logging.warning(
            "%d fila(s) tienen un ESTADO PROCESAL que no empieza con 'TERMINADO', 'REMITIDA' ni 'NO INICIO' "
            "(y no es uno de los estados que ya maneja el resto del proyecto) -- NO se creo carpeta para "
            "ellas, revisalas a mano:",
            len(sin_clasificar),
        )
        for fila, numero, estado in sin_clasificar:
            logging.warning("   - Fila %s, proceso %s: %r", fila, numero, estado)

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se creo ninguna carpeta todavia. Revisa el log y, si se ve bien, "
            "cambia MODO_PRUEBA = False al inicio de este script y vuelve a correrlo."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
