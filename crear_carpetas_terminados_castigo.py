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

Ademas de crear, este script hace una AUDITORIA COMPLETA de todas las
filas del Excel (no solo las que va a crear) para explicar por que el
total de carpetas en el disco puede no coincidir con el total de filas
del Excel. Cada numero de proceso sin carpeta cae en una de estas
categorias, y todas se reportan en el log:

  - Duplicado en el Excel: el mismo numero aparece en mas de una fila
    (con radicado y/o estado distintos) -- no se crea nada, hay que
    corregir el Excel a mano.
  - Activo/Suspendido/Reorganizacion (ESTADOS_A_CONTAR) CON radicado:
    ya esta rastreado en el listado de procesos faltantes
    (procesos_faltantes_en_disco.csv) -- no se crea una carpeta vacia
    aqui a proposito, porque ese proceso lo organiza
    buscar_faltantes_en_drive.py buscando su contenido real.
  - Activo/Suspendido/Reorganizacion SIN radicado diligenciado en el
    Excel todavia -- no se puede buscar ni crear hasta que se llene.
  - Sin ESTADO PROCESAL diligenciado en el Excel todavia.
  - Estado que no encaja en ninguna regla conocida (ver arriba).

Usa la misma configuracion (RUTA_EXCEL, HOJA_EXCEL, CARPETA_PROCESOS,
etc) de validar_renombrar_carpetas.py -- no hay que configurarla dos
veces.

Respeta MODO_PRUEBA (por defecto True): en modo prueba solo simula y
te dice que carpetas crearia.
"""

import logging
import os
import re
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


def leer_procesos_completos():
    """
    Lee el Excel completo y devuelve [(fila, numero, radicado, estado), ...]
    para CADA fila con un numero de proceso valido (columna No.) -- sin
    filtrar por radicado ni por estado (a diferencia de leer_filas_excel,
    que exige un radicado valido de 23 digitos). radicado es None si la
    celda esta vacia; estado es "" si la celda esta vacia.
    """
    wb = openpyxl.load_workbook(cruce_excel.RUTA_EXCEL, data_only=True)
    if cruce_excel.HOJA_EXCEL not in wb.sheetnames:
        raise ValueError(
            f"La hoja '{cruce_excel.HOJA_EXCEL}' no existe. Hojas disponibles: {wb.sheetnames}"
        )
    ws = wb[cruce_excel.HOJA_EXCEL]
    encabezados = [ws.cell(row=cruce_excel.FILA_ENCABEZADO, column=c).value for c in range(1, ws.max_column + 1)]
    col_no = cruce_excel.encontrar_columna(encabezados, cruce_excel.COLUMNA_NO)
    col_radicado = cruce_excel.encontrar_columna(encabezados, cruce_excel.COLUMNA_RADICADO)
    col_estado = cruce_excel.encontrar_columna(encabezados, cruce_excel.COLUMNA_ESTADO)

    filas = []
    for fila in range(cruce_excel.FILA_ENCABEZADO + 1, ws.max_row + 1):
        numero = ws.cell(row=fila, column=col_no).value
        if numero is None or not isinstance(numero, (int, float)):
            continue  # fila vacia o de notas/leyenda al final de la hoja

        radicado_crudo = ws.cell(row=fila, column=col_radicado).value
        radicado = re.sub(r"[\s\-]", "", str(radicado_crudo).strip()) if radicado_crudo is not None else ""
        radicado = radicado if radicado and radicado != "0" else None

        estado_crudo = ws.cell(row=fila, column=col_estado).value
        estado = str(estado_crudo).strip() if estado_crudo is not None else ""

        filas.append((fila, int(numero), radicado, estado))
    return filas


def procesar():
    if not cruce_excel.RUTA_EXCEL or not os.path.exists(cruce_excel.RUTA_EXCEL):
        logging.error("[Excel] No se encontro el archivo configurado en RUTA_EXCEL: %r", cruce_excel.RUTA_EXCEL)
        return

    filas = leer_procesos_completos()
    numero_a_filas = {}
    for fila, numero, radicado, estado in filas:
        numero_a_filas.setdefault(numero, []).append((fila, radicado, estado))
    logging.info(
        "Excel: %d fila(s), correspondientes a %d numero(s) de proceso distintos.",
        len(filas), len(numero_a_filas),
    )

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

    estados_ya_manejados = {e.upper() for e in cruce_excel.ESTADOS_A_CONTAR}

    creadas = 0
    ya_existian = 0
    sin_clasificar = []
    duplicados_numero = []
    pendientes_activos = []
    activos_sin_radicado = []
    sin_estado = []

    for numero, entradas in sorted(numero_a_filas.items()):
        if len(entradas) > 1:
            duplicados_numero.append((numero, entradas))
            continue

        fila, radicado, estado = entradas[0]

        if _existe_carpeta_para_numero(numero, carpetas_existentes):
            ya_existian += 1
            continue

        if not estado:
            sin_estado.append((fila, numero))
            continue

        # Los estados que ya maneja el resto del proyecto (ACTIVO, etc) no
        # se crean aqui -- ya los organiza validar_renombrar_carpetas.py/
        # buscar_faltantes_en_drive.py de la forma normal ("numero. radicado"),
        # pero se reportan para explicar por que todavia no tienen carpeta.
        if estado.upper() in estados_ya_manejados:
            if radicado:
                pendientes_activos.append((fila, numero, radicado, estado))
            else:
                activos_sin_radicado.append((fila, numero, estado))
            continue

        nombre_carpeta = _nombre_carpeta_para(numero, estado)
        if nombre_carpeta is None:
            sin_clasificar.append((fila, numero, estado))
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
        "Resumen: %d carpeta(s) %s, %d ya existian.",
        creadas, "simuladas (MODO_PRUEBA activo)" if MODO_PRUEBA else "creadas", ya_existian,
    )

    if pendientes_activos:
        logging.warning(
            "%d proceso(s) en estado %s NO tienen carpeta todavia, pero ya estan rastreados en el listado de "
            "procesos faltantes (procesos_faltantes_en_disco.csv) -- a proposito NO se crea una carpeta vacia "
            "aqui, corre buscar_faltantes_en_drive.py para que los busque y arme la carpeta con su contenido real:",
            len(pendientes_activos), "/".join(cruce_excel.ESTADOS_A_CONTAR),
        )
        for fila, numero, radicado, estado in pendientes_activos:
            logging.warning("   - Fila %s, proceso %s (%s): radicado %s", fila, numero, estado, radicado)

    if activos_sin_radicado:
        logging.warning(
            "%d proceso(s) en estado %s no tienen carpeta NI radicado diligenciado en el Excel todavia -- "
            "hay que llenar el radicado en el Excel antes de poder buscarlos o crearlos:",
            len(activos_sin_radicado), "/".join(cruce_excel.ESTADOS_A_CONTAR),
        )
        for fila, numero, estado in activos_sin_radicado:
            logging.warning("   - Fila %s, proceso %s (%s)", fila, numero, estado)

    if sin_estado:
        logging.warning(
            "%d proceso(s) no tienen carpeta y tampoco tienen ESTADO PROCESAL diligenciado en el Excel todavia:",
            len(sin_estado),
        )
        for fila, numero in sin_estado:
            logging.warning("   - Fila %s, proceso %s", fila, numero)

    if sin_clasificar:
        logging.warning(
            "%d fila(s) tienen un ESTADO PROCESAL que no empieza con 'TERMINADO', 'REMITIDA' ni 'NO INICIO' "
            "(y no es uno de los estados que ya maneja el resto del proyecto) -- NO se creo carpeta para "
            "ellas, revisalas a mano:",
            len(sin_clasificar),
        )
        for fila, numero, estado in sin_clasificar:
            logging.warning("   - Fila %s, proceso %s: %r", fila, numero, estado)

    if duplicados_numero:
        logging.warning(
            "%d numero(s) de proceso aparecen MAS DE UNA VEZ en el Excel (con radicado y/o estado distintos) "
            "-- no se creo ni se reviso ninguna carpeta para ellos porque no se puede saber cual fila es la "
            "correcta, revisalos a mano:",
            len(duplicados_numero),
        )
        for numero, entradas in duplicados_numero:
            detalle = ", ".join(
                f"fila {fila} (radicado {radicado!r}, estado {estado!r})" for fila, radicado, estado in entradas
            )
            logging.warning("   - Proceso %s aparece en: %s", numero, detalle)

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
