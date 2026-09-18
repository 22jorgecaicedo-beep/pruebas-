"""
Filtra un informe de procesos concursales y deja solo los procesos que NO
estan en estado "ACUERDO CONFIRMADO".

Lee el .xlsx del informe, detecta sola la fila de encabezados y la columna
de estado, y genera un Excel nuevo con:

  - Hoja "Procesos": todas las filas cuyo estado NO es acuerdo confirmado,
    con las mismas columnas del informe original mas la fila de origen.
  - Hoja "Resumen": cuantos procesos hay por cada estado, marcando cuales
    quedaron dentro y cuales se excluyeron.

Uso (Windows):

    python filtrar_procesos.py "C:\\Users\\Owner\\Downloads\\informe procesos concursales Agosto 2026.xlsx"

Si no se pasa ninguna ruta, busca en la carpeta de Descargas el informe
mas reciente cuyo nombre empiece por "informe procesos concursales".
"""

import argparse
import os
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================= CONFIGURACION =============================

# Estado que se quiere DEJAR POR FUERA del reporte.
ESTADO_A_EXCLUIR = "ACUERDO CONFIRMADO"

# Nombres posibles de la columna de estado dentro del informe.
NOMBRES_COLUMNA_ESTADO = [
    "ESTADO DEL PROCESO",
    "ESTADO PROCESO",
    "ESTADO ACTUAL",
    "ESTADO DEL TRAMITE",
    "ESTADO",
    "ETAPA",
]

# Cuantas filas se revisan al inicio de la hoja buscando el encabezado.
MAX_FILAS_ENCABEZADO = 30

# ===========================================================================


def normalizar(valor) -> str:
    """Texto en mayusculas, sin tildes y con los espacios colapsados."""
    if valor is None:
        return ""
    texto = str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto).strip().upper()
    return texto


def fila_vacia(valores) -> bool:
    return all(v is None or str(v).strip() == "" for v in valores)


def buscar_encabezado(hoja):
    """
    Devuelve (indice_fila_encabezado, indice_columna_estado, encabezados).

    Recorre las primeras filas de la hoja buscando una que contenga alguna
    de las columnas de estado conocidas. Si no encuentra ninguna, devuelve
    (None, None, None) para que la hoja se omita.
    """
    for indice_fila, fila in enumerate(
        hoja.iter_rows(min_row=1, max_row=min(MAX_FILAS_ENCABEZADO, hoja.max_row), values_only=True),
        start=1,
    ):
        if fila_vacia(fila):
            continue
        normalizados = [normalizar(v) for v in fila]

        # Primero, coincidencia exacta con los nombres conocidos.
        for nombre in NOMBRES_COLUMNA_ESTADO:
            if nombre in normalizados:
                return indice_fila, normalizados.index(nombre), list(fila)

        # Si no, alguna columna que empiece por "ESTADO".
        for indice_col, encabezado in enumerate(normalizados):
            if encabezado.startswith("ESTADO"):
                return indice_fila, indice_col, list(fila)

    return None, None, None


def es_acuerdo_confirmado(estado: str, exacto: bool) -> bool:
    estado_normalizado = normalizar(estado)
    objetivo = normalizar(ESTADO_A_EXCLUIR)
    if exacto:
        return estado_normalizado == objetivo
    return objetivo in estado_normalizado


def limpiar_celda(valor):
    """Deja el valor listo para escribirlo en el Excel de salida."""
    if isinstance(valor, (datetime, date, int, float, bool)) or valor is None:
        return valor
    return str(valor).strip()


def leer_informe(ruta_informe: Path, exacto: bool):
    """
    Devuelve (encabezados, filas, conteo_por_estado).

    filas: lista de dicts con hoja, fila_original y valores.
    conteo_por_estado: dict estado_normalizado -> [etiqueta, total, excluido]
    """
    libro = load_workbook(ruta_informe, data_only=True, read_only=True)

    encabezados_finales = None
    filas_conservadas = []
    conteo_por_estado = {}
    hojas_omitidas = []

    for hoja in libro.worksheets:
        fila_encabezado, columna_estado, encabezados = buscar_encabezado(hoja)
        if fila_encabezado is None:
            hojas_omitidas.append(hoja.title)
            continue

        encabezados = [limpiar_celda(e) for e in encabezados]
        # Recorta columnas sobrantes vacias al final del encabezado.
        while encabezados and (encabezados[-1] is None or str(encabezados[-1]).strip() == ""):
            encabezados.pop()
        if encabezados_finales is None:
            encabezados_finales = encabezados

        ancho = max(len(encabezados_finales), len(encabezados))

        for indice_fila, fila in enumerate(
            hoja.iter_rows(min_row=fila_encabezado + 1, values_only=True), start=fila_encabezado + 1
        ):
            if fila_vacia(fila):
                continue

            estado = fila[columna_estado] if columna_estado < len(fila) else None
            estado_normalizado = normalizar(estado) or "(SIN ESTADO)"
            excluido = es_acuerdo_confirmado(estado, exacto)

            registro = conteo_por_estado.setdefault(
                estado_normalizado,
                {"etiqueta": str(estado).strip() if estado not in (None, "") else "(sin estado)",
                 "total": 0,
                 "excluido": excluido},
            )
            registro["total"] += 1

            if excluido:
                continue

            valores = [limpiar_celda(v) for v in fila[:ancho]]
            valores += [None] * (ancho - len(valores))
            filas_conservadas.append(
                {"hoja": hoja.title, "fila_original": indice_fila, "valores": valores}
            )

    libro.close()

    if encabezados_finales is None:
        columnas = ", ".join(NOMBRES_COLUMNA_ESTADO)
        raise SystemExit(
            f"No se encontro ninguna columna de estado en '{ruta_informe.name}'.\n"
            f"Se buscaron columnas llamadas: {columnas}.\n"
            f"Revisa el informe o agrega el nombre real a NOMBRES_COLUMNA_ESTADO."
        )

    return encabezados_finales, filas_conservadas, conteo_por_estado, hojas_omitidas


def ajustar_anchos(hoja, maximo=60):
    anchos = {}
    for fila in hoja.iter_rows(values_only=True):
        for indice, valor in enumerate(fila, start=1):
            if valor is None:
                continue
            largo = len(str(valor))
            if largo > anchos.get(indice, 0):
                anchos[indice] = largo
    for indice, largo in anchos.items():
        hoja.column_dimensions[get_column_letter(indice)].width = min(max(largo + 2, 10), maximo)


def escribir_resultado(ruta_salida, encabezados, filas, conteo_por_estado, exacto):
    libro = Workbook()

    relleno_encabezado = PatternFill("solid", fgColor="1F3864")
    fuente_encabezado = Font(color="FFFFFF", bold=True)

    hoja = libro.active
    hoja.title = "Procesos"
    encabezados_salida = ["Hoja origen", "Fila origen"] + [
        e if e not in (None, "") else f"Columna {i + 1}" for i, e in enumerate(encabezados)
    ]
    hoja.append(encabezados_salida)
    for celda in hoja[1]:
        celda.fill = relleno_encabezado
        celda.font = fuente_encabezado
        celda.alignment = Alignment(vertical="center", wrap_text=True)

    for registro in filas:
        hoja.append([registro["hoja"], registro["fila_original"]] + registro["valores"])

    hoja.freeze_panes = "A2"
    if hoja.max_row >= 1:
        hoja.auto_filter.ref = f"A1:{get_column_letter(len(encabezados_salida))}{max(hoja.max_row, 1)}"
    ajustar_anchos(hoja)

    resumen = libro.create_sheet("Resumen")
    modo = "coincidencia exacta" if exacto else "el estado contiene el texto"
    resumen.append(["Informe de procesos que NO estan en estado", ESTADO_A_EXCLUIR])
    resumen.append(["Criterio de exclusion", modo])
    resumen.append(["Generado", datetime.now().strftime("%Y-%m-%d %H:%M")])
    resumen.append([])
    resumen.append(["Estado", "Cantidad", "Incluido en el reporte"])
    for celda in resumen[5]:
        celda.fill = relleno_encabezado
        celda.font = fuente_encabezado

    total_general = 0
    total_incluidos = 0
    for clave in sorted(conteo_por_estado, key=lambda k: -conteo_por_estado[k]["total"]):
        registro = conteo_por_estado[clave]
        total_general += registro["total"]
        if not registro["excluido"]:
            total_incluidos += registro["total"]
        resumen.append([registro["etiqueta"], registro["total"], "NO" if registro["excluido"] else "SI"])

    resumen.append([])
    resumen.append(["Total de procesos en el informe", total_general])
    resumen.append(["Procesos en el reporte (sin acuerdo confirmado)", total_incluidos])
    resumen.append(["Procesos excluidos (acuerdo confirmado)", total_general - total_incluidos])
    for fila in resumen.iter_rows(min_row=resumen.max_row - 2, max_row=resumen.max_row, max_col=1):
        fila[0].font = Font(bold=True)
    ajustar_anchos(resumen)

    libro.save(ruta_salida)
    return total_general, total_incluidos


def carpeta_descargas() -> Path:
    return Path(os.path.expanduser("~")) / "Downloads"


def buscar_informe_por_defecto() -> Path:
    carpeta = carpeta_descargas()
    candidatos = sorted(
        (p for p in carpeta.glob("*.xlsx")
         if p.name.lower().startswith("informe procesos concursales") and not p.name.startswith("~$")),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidatos:
        raise SystemExit(
            f"No se encontro ningun 'informe procesos concursales*.xlsx' en {carpeta}.\n"
            f"Pasa la ruta del archivo como argumento:\n"
            f'    python filtrar_procesos.py "C:\\ruta\\al\\informe.xlsx"'
        )
    return candidatos[0]


def main():
    parser = argparse.ArgumentParser(
        description="Extrae los procesos que NO estan en estado 'ACUERDO CONFIRMADO' y los deja en un Excel nuevo."
    )
    parser.add_argument("informe", nargs="?", help="Ruta del .xlsx del informe de procesos concursales.")
    parser.add_argument("-o", "--salida", help="Ruta del Excel de salida.")
    parser.add_argument(
        "--exacto",
        action="store_true",
        help="Excluir solo los estados que digan exactamente 'ACUERDO CONFIRMADO' "
             "(por defecto tambien se excluyen variantes como 'Acuerdo confirmado en ejecucion').",
    )
    args = parser.parse_args()

    ruta_informe = Path(args.informe).expanduser() if args.informe else buscar_informe_por_defecto()
    if not ruta_informe.is_file():
        raise SystemExit(f"No existe el archivo: {ruta_informe}")
    if ruta_informe.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise SystemExit(
            f"'{ruta_informe.name}' no es un archivo .xlsx.\n"
            f"Abrelo en Excel y guardalo como 'Libro de Excel (*.xlsx)' antes de correr este programa."
        )

    print(f"Leyendo: {ruta_informe}")
    encabezados, filas, conteo, hojas_omitidas = leer_informe(ruta_informe, args.exacto)

    for nombre_hoja in hojas_omitidas:
        print(f"  (aviso) la hoja '{nombre_hoja}' no tiene columna de estado, se omitio.")

    if args.salida:
        ruta_salida = Path(args.salida).expanduser()
    else:
        ruta_salida = ruta_informe.with_name(f"{ruta_informe.stem} - SIN ACUERDO CONFIRMADO.xlsx")

    total, incluidos = escribir_resultado(ruta_salida, encabezados, filas, conteo, args.exacto)

    print()
    print("Estados encontrados en el informe:")
    for clave in sorted(conteo, key=lambda k: -conteo[k]["total"]):
        registro = conteo[clave]
        marca = "excluido" if registro["excluido"] else "incluido"
        print(f"  - {registro['etiqueta']}: {registro['total']} ({marca})")

    print()
    print(f"Procesos en el informe : {total}")
    print(f"Sin acuerdo confirmado : {incluidos}")
    print(f"Excel generado         : {ruta_salida}")


if __name__ == "__main__":
    sys.exit(main())
