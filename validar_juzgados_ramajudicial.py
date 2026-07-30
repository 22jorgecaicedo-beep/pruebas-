"""
Consulta cada radicado del informe de Excel en el portal publico "Consulta
de Procesos Nacional Unificada" (CPNU) de la Rama Judicial, y compara el
despacho que reporta el portal contra el juzgado anotado en la columna
JUZGADO del Excel.

Como funciona el cruce:

  El radicado de 23 digitos termina en un "consecutivo" (los ultimos 2
  digitos) que cambia cuando el proceso pasa a otra instancia/despacho
  (por ejemplo termina en 00 en primera instancia, 01 si ya se elevo a
  otro despacho, etc). El Excel puede traer el radicado en cualquiera de
  esos consecutivos, no siempre en 00.

  Para cada fila del Excel:
    1. Se consulta el radicado tal como esta en el Excel.
    2. Se prueba el siguiente consecutivo (consecutivo + 1), luego el
       siguiente, y asi sucesivamente, mientras el portal siga
       encontrando resultado (hasta MAX_CONSECUTIVOS_A_REVISAR).
    3. El despacho del ULTIMO consecutivo que si dio resultado es el que
       se considera "despacho actual" del proceso, y se compara contra la
       columna JUZGADO del Excel.

El resultado completo (coincida o no) queda en un CSV
(REPORTE_CSV) para que lo revises con calma, y al final del log se
imprime solo la lista de procesos que NO coinciden.

IMPORTANTE - no hay captcha en esta consulta (se verifico manualmente),
asi que el script consulta de corrido sin necesitar que nadie intervenga.
De todas formas:

  - Este script no se pudo probar en vivo contra el portal (bloqueado
    desde el entorno donde se escribio). Es muy probable que la primera
    corrida necesite ajustar algun selector; corre primero con pocos
    procesos (ver SOLO_ESTOS_NUMEROS) y revisa el log.
  - Se agrega una pausa entre consultas (PAUSA_ENTRE_CONSULTAS_SEGUNDOS)
    para no saturar un portal publico del Estado con cientos de consultas
    seguidas. Con ~950 procesos, corran esto en un par de tandas en vez de
    todo de una sola vez si pueden.
  - El script recuerda que radicados ya proceso (ARCHIVO_PROGRESO), asi
    que si lo interrumpes (Ctrl+C) y lo vuelves a correr, sigue donde iba
    en vez de repetir consultas ya hechas.
"""

import csv
import logging
import os
import re
import time
import unicodedata
from pathlib import Path

import openpyxl

# ============================= CONFIGURACION =============================

RUTA_EXCEL = r"C:\Users\User\Documents\RELACION_595_PROCESOS_ESSA_1S.xlsx"
HOJA_EXCEL = "RELACION PROCESOS"
FILA_ENCABEZADO = 5
COLUMNA_NO = "No."
COLUMNA_RADICADO = "RADICADO"
COLUMNA_JUZGADO = "JUZGADO"

# Si quieres probar solo con algunos procesos primero (recomendado la
# primera vez), pon aqui su numero de proceso, ej: [1, 133]. Deja []
# para procesar todos los del Excel.
SOLO_ESTOS_NUMEROS = [1, 133]

# Cuantos consecutivos adicionales probar (consecutivo+1, +2, ...) antes
# de asumir que ya no hay una instancia mas nueva.
MAX_CONSECUTIVOS_A_REVISAR = 5

# Pausa entre cada consulta al portal, para no saturarlo.
PAUSA_ENTRE_CONSULTAS_SEGUNDOS = 2.5

# Mostrar el navegador mientras corre. Dejalo en True la primera vez.
NAVEGADOR_VISIBLE = True

URL_CONSULTA = "https://consultaprocesos.ramajudicial.gov.co/Procesos/NumeroRadicacion"
TIMEOUT_CONSULTA_MS = 25000

CARPETA_SALIDA = os.path.dirname(__file__)
ARCHIVO_LOG = os.path.join(CARPETA_SALIDA, "validar_juzgados_ramajudicial.log")
REPORTE_CSV = os.path.join(CARPETA_SALIDA, "validar_juzgados_reporte.csv")
ARCHIVO_PROGRESO = os.path.join(CARPETA_SALIDA, "validar_juzgados_progreso.txt")

PATRON_RADICADO = re.compile(r"\d{23}")

# ===========================================================================


def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


# ============================ Normalizacion de texto ========================

def _sin_acentos(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c)
    )


def normalizar(texto: str) -> str:
    texto = _sin_acentos(texto).upper()
    texto = re.sub(r"\(.*?\)", "", texto)  # quita "(SANTANDER)" y similares
    texto = re.sub(r"[^A-Z0-9 ]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


UNIDADES = ["", "PRIMERO", "SEGUNDO", "TERCERO", "CUARTO", "QUINTO", "SEXTO", "SEPTIMO", "OCTAVO", "NOVENO"]
DECENAS = ["", "DECIMO", "VIGESIMO", "TRIGESIMO", "CUADRAGESIMO", "QUINCUAGESIMO"]


def _ordinal_texto(n: int) -> str:
    """Genera la forma compuesta usual en nombres de juzgados (ej 21 -> 'VIGESIMO PRIMERO')."""
    if n <= 9:
        return UNIDADES[n]
    decena, unidad = divmod(n, 10)
    if unidad == 0:
        return DECENAS[decena]
    return f"{DECENAS[decena]} {UNIDADES[unidad]}"


# Texto ordinal normalizado -> numero, para n de 1 a 50 (mas que suficiente
# para la cantidad de juzgados de una misma especialidad en una ciudad).
ORDINAL_A_NUMERO = {_ordinal_texto(n): n for n in range(1, 51)}


def extraer_numero_juzgado(texto_normalizado: str):
    """
    Busca al inicio del texto un numero de juzgado, ya sea en forma de
    palabra ordinal ("PRIMERO CIVIL...") o numerica ("JUZGADO 001 CIVIL...").
    Devuelve (numero, resto_del_texto) o (None, texto_normalizado) si no lo
    reconoce.
    """
    texto = texto_normalizado
    if texto.startswith("JUZGADO "):
        m = re.match(r"JUZGADO\s+0*(\d+)\s+(.*)", texto)
        if m:
            return int(m.group(1)), m.group(2)

    # Prueba ordinales compuestos (2 palabras) antes que simples.
    palabras = texto.split(" ")
    for cantidad_palabras in (2, 1):
        if len(palabras) >= cantidad_palabras:
            candidato = " ".join(palabras[:cantidad_palabras])
            if candidato in ORDINAL_A_NUMERO:
                resto = " ".join(palabras[cantidad_palabras:])
                return ORDINAL_A_NUMERO[candidato], resto

    return None, texto


def juzgados_coinciden(juzgado_excel: str, despacho_portal: str) -> bool:
    """
    Compara dos nombres de despacho de forma tolerante a diferencias de
    formato: "PRIMERO CIVIL MUNICIPAL..." (Excel) vs "JUZGADO 001 CIVIL
    MUNICIPAL... (SANTANDER)" (portal), acentos, mayusculas, espacios
    dobles, etc. No es infalible: revisa igual el CSV completo si tienes
    dudas en casos limite.
    """
    a = normalizar(juzgado_excel)
    b = normalizar(despacho_portal)
    if not a or not b:
        return False

    num_a, resto_a = extraer_numero_juzgado(a)
    num_b, resto_b = extraer_numero_juzgado(b)

    if num_a is not None and num_b is not None:
        if num_a != num_b:
            return False
        return resto_a.strip() == resto_b.strip() or resto_a.strip() in resto_b.strip() or resto_b.strip() in resto_a.strip()

    # Si no se pudo reconocer el numero en alguno de los dos, compara el
    # texto completo de forma laxa (uno contenido en el otro).
    return a == b or a in b or b in a


# ================================ Lectura Excel ==============================


def encontrar_columna(encabezados, nombre_buscado):
    for idx, valor in enumerate(encabezados, start=1):
        if valor and str(valor).strip().lower() == nombre_buscado.strip().lower():
            return idx
    raise ValueError(
        f"No se encontro la columna '{nombre_buscado}' en la fila {FILA_ENCABEZADO} de '{HOJA_EXCEL}'."
    )


def leer_procesos_excel():
    wb = openpyxl.load_workbook(RUTA_EXCEL, data_only=True)
    if HOJA_EXCEL not in wb.sheetnames:
        raise ValueError(f"La hoja '{HOJA_EXCEL}' no existe. Hojas disponibles: {wb.sheetnames}")
    ws = wb[HOJA_EXCEL]

    encabezados = [ws.cell(row=FILA_ENCABEZADO, column=c).value for c in range(1, ws.max_column + 1)]
    col_no = encontrar_columna(encabezados, COLUMNA_NO)
    col_rad = encontrar_columna(encabezados, COLUMNA_RADICADO)
    col_juz = encontrar_columna(encabezados, COLUMNA_JUZGADO)

    procesos = []
    for fila in range(FILA_ENCABEZADO + 1, ws.max_row + 1):
        numero = ws.cell(row=fila, column=col_no).value
        radicado_crudo = ws.cell(row=fila, column=col_rad).value
        juzgado_excel = ws.cell(row=fila, column=col_juz).value

        if numero is None or not isinstance(numero, (int, float)):
            continue
        numero = int(numero)

        if SOLO_ESTOS_NUMEROS and numero not in SOLO_ESTOS_NUMEROS:
            continue

        if radicado_crudo is None:
            continue
        radicado = re.sub(r"[\s\-]", "", str(radicado_crudo).strip())
        if radicado in ("", "0") or not PATRON_RADICADO.fullmatch(radicado):
            continue

        procesos.append((fila, numero, radicado, (juzgado_excel or "").strip()))

    return procesos


# ============================ Consulta al portal ============================


def _cantidad_resultados(pagina) -> int:
    localizador = pagina.locator("text=/Resultados encontrados/i")
    if localizador.count() == 0:
        return 0
    texto = localizador.first.inner_text()
    m = re.search(r"(\d+)", texto)
    return int(m.group(1)) if m else 0


def consultar_radicado(pagina, radicado: str):
    """Devuelve el texto de 'Despacho y Departamento' para ese radicado, o None si el portal no encontro nada."""
    pagina.goto(URL_CONSULTA, wait_until="networkidle", timeout=TIMEOUT_CONSULTA_MS)

    opcion_todos = pagina.get_by_text("Todos los Procesos", exact=False)
    if opcion_todos.count() > 0:
        opcion_todos.first.click()

    campo = pagina.get_by_role("textbox").first
    campo.fill(radicado)

    pagina.get_by_role("button", name=re.compile("consultar", re.IGNORECASE)).click()
    pagina.wait_for_selector("text=/Resultados encontrados/i", timeout=TIMEOUT_CONSULTA_MS)

    if _cantidad_resultados(pagina) == 0:
        return None

    fila = pagina.locator("table tbody tr", has_text=radicado).first
    celdas = fila.locator("td")
    if celdas.count() < 4:
        logging.warning("[Portal] Estructura de tabla inesperada para %s; revisa el selector de la celda Despacho.", radicado)
        return None
    return celdas.nth(3).inner_text().strip()


def despacho_actual(pagina, radicado_excel: str):
    """
    Sigue la cadena de consecutivos (consecutivo, +1, +2, ...) mientras el
    portal encuentre resultado. Devuelve (radicado_usado, despacho) del
    ultimo consecutivo encontrado, o (None, None) si ni el radicado del
    Excel dio resultado.
    """
    base = radicado_excel[:-2]
    consecutivo_actual = int(radicado_excel[-2:])

    despacho = consultar_radicado(pagina, radicado_excel)
    if despacho is None:
        return None, None

    radicado_usado = radicado_excel
    for salto in range(1, MAX_CONSECUTIVOS_A_REVISAR + 1):
        siguiente_consecutivo = consecutivo_actual + salto
        if siguiente_consecutivo > 99:
            break
        radicado_siguiente = f"{base}{siguiente_consecutivo:02d}"
        time.sleep(PAUSA_ENTRE_CONSULTAS_SEGUNDOS)
        despacho_siguiente = consultar_radicado(pagina, radicado_siguiente)
        if despacho_siguiente is None:
            break
        radicado_usado, despacho = radicado_siguiente, despacho_siguiente

    return radicado_usado, despacho


# ==================================== Progreso ===============================


def cargar_progreso() -> dict:
    if not os.path.exists(ARCHIVO_PROGRESO):
        return {}
    resultado = {}
    with open(ARCHIVO_PROGRESO, encoding="utf-8") as f:
        for linea in f:
            partes = linea.rstrip("\n").split("\t")
            if len(partes) == 5:
                numero, radicado_usado, despacho, coincide, juzgado_excel = partes
                resultado[int(numero)] = (radicado_usado, despacho, coincide, juzgado_excel)
    return resultado


def guardar_progreso(numero, radicado_usado, despacho, coincide, juzgado_excel):
    with open(ARCHIVO_PROGRESO, "a", encoding="utf-8") as f:
        f.write(f"{numero}\t{radicado_usado}\t{despacho}\t{coincide}\t{juzgado_excel}\n")


# ====================================== MAIN ==================================


def procesar():
    from playwright.sync_api import sync_playwright

    procesos = leer_procesos_excel()
    logging.info("Excel: %d proceso(s) para consultar en el portal.", len(procesos))

    progreso = cargar_progreso()
    pendientes = [p for p in procesos if p[1] not in progreso]
    logging.info("%d ya estaban procesados de una corrida anterior; quedan %d por consultar.", len(procesos) - len(pendientes), len(pendientes))

    filas_reporte = []
    for numero, (radicado_usado, despacho, coincide, juzgado_excel) in progreso.items():
        filas_reporte.append([numero, radicado_usado, juzgado_excel, despacho, coincide])

    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=not NAVEGADOR_VISIBLE)
        pagina = navegador.new_page()
        try:
            for fila, numero, radicado_excel, juzgado_excel in pendientes:
                try:
                    radicado_usado, despacho = despacho_actual(pagina, radicado_excel)
                except Exception:
                    logging.exception("[Portal] Error consultando el proceso %s (radicado %s); se omite por ahora.", numero, radicado_excel)
                    continue

                if despacho is None:
                    coincide = "NO_ENCONTRADO"
                    logging.warning("[Portal] Proceso %s: el portal no encontro ningun resultado para %s.", numero, radicado_excel)
                else:
                    coincide = "SI" if juzgados_coinciden(juzgado_excel, despacho) else "NO"
                    if coincide == "NO":
                        logging.warning(
                            "[Diferencia] Proceso %s (radicado consultado %s): portal dice '%s', Excel dice '%s'.",
                            numero, radicado_usado, despacho, juzgado_excel,
                        )

                guardar_progreso(numero, radicado_usado or "", despacho or "", coincide, juzgado_excel)
                filas_reporte.append([numero, radicado_usado, juzgado_excel, despacho, coincide])

                time.sleep(PAUSA_ENTRE_CONSULTAS_SEGUNDOS)
        finally:
            navegador.close()

    filas_reporte.sort(key=lambda f: f[0])
    with open(REPORTE_CSV, "w", newline="", encoding="utf-8-sig") as f:
        escritor = csv.writer(f, delimiter=";")
        escritor.writerow(["No. proceso", "Radicado consultado (ultimo consecutivo)", "Juzgado en Excel", "Despacho en portal", "Coincide"])
        escritor.writerows(filas_reporte)

    no_coinciden = [f for f in filas_reporte if f[4] == "NO"]
    no_encontrados = [f for f in filas_reporte if f[4] == "NO_ENCONTRADO"]

    logging.info("-" * 60)
    logging.info("Reporte completo guardado en: %s", REPORTE_CSV)
    logging.info("Resumen: %d procesados, %d coinciden, %d NO coinciden, %d sin resultado en el portal.",
                  len(filas_reporte), len(filas_reporte) - len(no_coinciden) - len(no_encontrados), len(no_coinciden), len(no_encontrados))

    if no_coinciden:
        logging.warning("Procesos donde el juzgado del Excel NO coincide con el portal:")
        for numero, radicado_usado, juzgado_excel, despacho, _ in no_coinciden:
            logging.warning("  - Proceso %s (radicado %s): Excel='%s'  |  Portal='%s'", numero, radicado_usado, juzgado_excel, despacho)

    if no_encontrados:
        logging.warning("Procesos que el portal no pudo encontrar (revisar el radicado a mano):")
        for numero, radicado_usado, juzgado_excel, _despacho, _ in no_encontrados:
            logging.warning("  - Proceso %s (radicado %s)", numero, radicado_usado)


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
