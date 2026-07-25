"""
Valida que las carpetas de radicado (23 digitos) en el disco duro coincidan
con el numero de proceso anotado en el informe de Excel, y las renombra a
"<numero>. <radicado>".

Para cada fila del Excel con un numero de proceso (columna "No.") y un
radicado valido (columna "RADICADO"), busca en CARPETA_PROCESOS una carpeta
cuyo nombre contenga ese radicado y la renombra. El cruce se hace por el
valor exacto del radicado (no por posicion/orden en la hoja), asi que no
importa si faltan carpetas o si el Excel tiene huecos.

Al terminar, reporta (en pantalla y en un log):
  - Carpetas renombradas (o que se renombrarian, en modo prueba).
  - Carpetas que ya tenian el nombre correcto (se dejan igual).
  - Procesos del Excel sin carpeta correspondiente en el disco.
  - Carpetas en el disco cuyo radicado no aparece en el Excel.
  - Filas del Excel con radicado invalido o con numero/radicado repetido
    (no se tocan, para no arriesgar un cruce incorrecto).

Por defecto corre en MODO_PRUEBA (no renombra nada, solo muestra que haria).
Revisa el reporte y, cuando confies en que el cruce esta bien, cambia
MODO_PRUEBA a False para aplicar los renombrados de verdad.
"""

import logging
import os
import re
from pathlib import Path

import openpyxl

# ============================= CONFIGURACION =============================

# Ruta al informe de Excel (.xlsx o .xlsm). Esta en tu Escritorio; si el
# nombre del archivo no es exactamente este, ajustalo.
RUTA_EXCEL = os.path.join(
    os.path.expanduser("~"), "Desktop", "3._CONTROL_PROCESOS_EJECUTIVOS_ESSA_22072026.xlsm"
)

# Nombre de la hoja/pestaña donde estan los procesos a cruzar.
HOJA_EXCEL = "ACTIVOS"

# Fila donde estan los encabezados de columna (la fila que dice "No.",
# "RADICADO", etc). Los datos empiezan en la fila siguiente.
FILA_ENCABEZADO = 5

# Nombres exactos de las columnas a usar (tal como aparecen en el encabezado).
COLUMNA_NO = "No."
COLUMNA_RADICADO = "RADICADO"

# Carpeta del disco duro donde estan las carpetas de cada proceso (las que
# hoy tienen solo el radicado de 23 digitos como nombre): tu disco duro
# externo. Si las carpetas estan dentro de otra carpeta ahi (no directo en
# la raiz de D:), agrega esa carpeta aqui, ej: r"D:/Procesos".
CARPETA_PROCESOS = r"D:/"

# True: no renombra nada, solo muestra/registra que haria (recomendado la
# primera vez). False: aplica los renombrados de verdad.
MODO_PRUEBA = True

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "validar_renombrar_carpetas.log")

PATRON_RADICADO = re.compile(r"\d{23}")

# ===========================================================================


def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8", mode="w"),
            logging.StreamHandler(),
        ],
    )


def encontrar_columna(encabezados, nombre_buscado):
    for idx, valor in enumerate(encabezados, start=1):
        if valor and str(valor).strip().lower() == nombre_buscado.strip().lower():
            return idx
    raise ValueError(
        f"No se encontro la columna '{nombre_buscado}' en la fila {FILA_ENCABEZADO} de '{HOJA_EXCEL}'."
    )


def leer_filas_excel():
    """Lee todas las filas con un numero de proceso y un radicado con formato valido (23 digitos), sin descartar aun repetidos."""
    wb = openpyxl.load_workbook(RUTA_EXCEL, data_only=True)
    if HOJA_EXCEL not in wb.sheetnames:
        raise ValueError(f"La hoja '{HOJA_EXCEL}' no existe. Hojas disponibles: {wb.sheetnames}")
    ws = wb[HOJA_EXCEL]

    encabezados = [ws.cell(row=FILA_ENCABEZADO, column=c).value for c in range(1, ws.max_column + 1)]
    col_no = encontrar_columna(encabezados, COLUMNA_NO)
    col_rad = encontrar_columna(encabezados, COLUMNA_RADICADO)

    filas = []
    for fila in range(FILA_ENCABEZADO + 1, ws.max_row + 1):
        numero = ws.cell(row=fila, column=col_no).value
        radicado_crudo = ws.cell(row=fila, column=col_rad).value

        if numero is None or not isinstance(numero, (int, float)):
            continue  # fila vacia o de notas/leyenda al final de la hoja, se ignora

        numero = int(numero)

        if radicado_crudo is None:
            continue  # proceso sin radicado asignado todavia, nada que cruzar

        radicado = re.sub(r"[\s\-]", "", str(radicado_crudo).strip())
        if radicado in ("", "0"):
            continue  # radicado aun no diligenciado

        if not PATRON_RADICADO.fullmatch(radicado):
            logging.warning(
                "[Excel] Fila %s (proceso %s): radicado con formato invalido, se omite: %r",
                fila, numero, radicado_crudo,
            )
            continue

        filas.append((fila, numero, radicado))

    return filas


def quitar_repetidos(filas):
    """Excluye del cruce cualquier numero de proceso o radicado que aparezca en mas de una fila, reportando el conflicto."""
    filas_por_numero = {}
    filas_por_radicado = {}
    for fila, numero, radicado in filas:
        filas_por_numero.setdefault(numero, []).append(fila)
        filas_por_radicado.setdefault(radicado, []).append(fila)

    numeros_repetidos = {n for n, fs in filas_por_numero.items() if len(fs) > 1}
    radicados_repetidos = {r for r, fs in filas_por_radicado.items() if len(fs) > 1}

    for numero in numeros_repetidos:
        logging.warning(
            "[Excel] Numero de proceso %s aparece en varias filas (%s); esas filas se omiten del cruce.",
            numero, filas_por_numero[numero],
        )
    for radicado in radicados_repetidos:
        logging.warning(
            "[Excel] Radicado %s aparece en varias filas (%s); esas filas se omiten del cruce.",
            radicado, filas_por_radicado[radicado],
        )

    return [
        (fila, numero, radicado)
        for fila, numero, radicado in filas
        if numero not in numeros_repetidos and radicado not in radicados_repetidos
    ]


def nombre_ya_correcto(nombre_carpeta: str, numero: int, radicado: str) -> bool:
    return nombre_carpeta.strip() == f"{numero}. {radicado}"


def radicado_de_nombre_carpeta(nombre_carpeta: str):
    m = PATRON_RADICADO.search(nombre_carpeta)
    return m.group(0) if m else None


def procesar():
    procesos = quitar_repetidos(leer_filas_excel())
    logging.info("Excel: %d proceso(s) con radicado valido y sin repetir, listos para cruzar.", len(procesos))

    por_radicado = {radicado: (numero, fila) for fila, numero, radicado in procesos}

    carpeta_raiz = Path(CARPETA_PROCESOS)
    carpetas = [d for d in carpeta_raiz.iterdir() if d.is_dir()]

    radicados_encontrados_en_disco = set()
    renombradas = 0
    ya_correctas = 0
    sin_proceso_en_excel = []
    conflictos = 0

    for carpeta in carpetas:
        radicado = radicado_de_nombre_carpeta(carpeta.name)
        if not radicado:
            continue  # no parece una carpeta de proceso (no tiene un radicado de 23 digitos en el nombre)

        match = por_radicado.get(radicado)
        if not match:
            sin_proceso_en_excel.append(carpeta.name)
            continue

        radicados_encontrados_en_disco.add(radicado)
        numero, fila = match

        if nombre_ya_correcto(carpeta.name, numero, radicado):
            ya_correctas += 1
            continue

        nuevo_nombre = f"{numero}. {radicado}"
        destino = carpeta.parent / nuevo_nombre

        if destino.exists():
            conflictos += 1
            logging.warning(
                "[Conflicto] '%s' deberia renombrarse a '%s' pero ya existe una carpeta con ese nombre. Se omite, revisa manualmente.",
                carpeta.name, nuevo_nombre,
            )
            continue

        if MODO_PRUEBA:
            logging.info("[SIMULACION] '%s'  ->  '%s'", carpeta.name, nuevo_nombre)
        else:
            carpeta.rename(destino)
            logging.info("[Renombrada] '%s'  ->  '%s'", carpeta.name, nuevo_nombre)
        renombradas += 1

    sin_carpeta_en_disco = 0
    for fila, numero, radicado in procesos:
        if radicado not in radicados_encontrados_en_disco:
            sin_carpeta_en_disco += 1
            logging.warning(
                "[Sin carpeta] Proceso %s (fila %s del Excel, radicado %s) no tiene carpeta correspondiente en %s.",
                numero, fila, radicado, CARPETA_PROCESOS,
            )

    if sin_proceso_en_excel:
        logging.warning("[Sin proceso] %d carpeta(s) con radicado que no aparece en el Excel:", len(sin_proceso_en_excel))
        for nombre in sin_proceso_en_excel:
            logging.warning("   - %s", nombre)

    logging.info("-" * 60)
    logging.info(
        "Resumen: %d %s, %d ya tenian el nombre correcto, %d sin carpeta en disco, "
        "%d carpetas sin proceso en el Excel, %d conflictos de nombre.",
        renombradas,
        "carpetas simuladas (MODO_PRUEBA activo)" if MODO_PRUEBA else "carpetas renombradas",
        ya_correctas, sin_carpeta_en_disco, len(sin_proceso_en_excel), conflictos,
    )
    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se renombro nada todavia. Revisa el reporte de arriba "
            "y, si se ve bien, cambia MODO_PRUEBA = False al inicio del script y vuelve a correrlo."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
