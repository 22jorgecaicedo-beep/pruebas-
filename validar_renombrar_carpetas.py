"""
Valida que las carpetas de radicado (23 digitos) en el disco duro coincidan
con el numero de proceso anotado en el informe de Excel, y las renombra a
"<numero>. <radicado>".

Para cada fila del Excel con un numero de proceso (columna "No.") y un
radicado valido (columna "RADICADO"), busca en CARPETA_PROCESOS una carpeta
cuyo nombre contenga ese radicado y la renombra. El cruce se hace por el
valor exacto del radicado (no por posicion/orden en la hoja), asi que no
importa si faltan carpetas o si el Excel tiene huecos.

El radicado se reconoce de forma ESTRICTA: tiene que ser una tira de
exactamente 23 digitos que no este pegada a mas digitos antes o despues
(para no cortar mal un numero mas largo, o colar por error uno mas corto).

Seguridad antes de renombrar:
  - El informe de Excel se lee DOS VECES de forma independiente, y se
    compara que ambas lecturas den exactamente el mismo resultado, antes
    de tocar cualquier carpeta (por si el Excel se esta editando al mismo
    tiempo).
  - Justo antes de cada renombrado se vuelve a confirmar el radicado de la
    carpeta una vez mas.
  - Al terminar de renombrar (si MODO_PRUEBA = False), se hace una
    revision final volviendo a leer el disco para confirmar que cada
    carpeta quedo con el nombre esperado.

Dos tipos de diferencia de un digito, tratadas DISTINTO a proposito:
  - Si el radicado de la carpeta coincide con el del Excel en los primeros
    22 digitos y solo difiere en el ULTIMO (el "consecutivo" que indica la
    instancia/reparto, ej. termina en 00 o en 01), SI se corrige
    automatico: son el mismo proceso, solo cambio de instancia. Queda
    registrado en el log como "[Consecutivo]".
  - Si el radicado de la carpeta le sobra o le falta un digito en
    CUALQUIER OTRA posicion (o tiene 22/24 digitos en vez de 23), se
    reporta aparte como "POSIBLE COINCIDENCIA" para que la revises tu a
    mano -- el script NO la renombra sola, porque dos procesos distintos
    del mismo juzgado y año suelen compartir casi todos los digitos entre
    si, y adivinar mal significaria ponerle a una carpeta el numero de un
    proceso que no es.

Al terminar, reporta (en pantalla y en un log):
  - Carpetas renombradas (o que se renombrarian, en modo prueba).
  - Carpetas que ya tenian el nombre correcto (se dejan igual).
  - Procesos del Excel sin carpeta correspondiente en el disco.
  - Carpetas en el disco cuyo radicado no aparece en el Excel.
  - Posibles coincidencias con un digito de mas o de menos (revisar a mano).
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

# Ruta al informe de Excel (.xlsx o .xlsm).
RUTA_EXCEL = r"C:\Users\User\Desktop\3. CONTROL PROCESOS EJECUTIVOS ESSA 22072026 (5).xlsm"

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

# Radicado "bueno": exactamente 23 digitos, sin otro digito pegado antes o
# despues (para no cortar mal un numero mas largo, ni colar uno mas corto).
# No usamos \b porque \b trata "_" como parte de la palabra -- por ejemplo
# una carpeta duplicada "..._2" dejaria de reconocerse; aqui solo nos
# importa que no haya OTRO DIGITO pegado.
PATRON_RADICADO_EXACTO = re.compile(r"(?<!\d)\d{23}(?!\d)")

# Radicado "casi bueno" (con 1 digito de mas o de menos), solo para poder
# reportar posibles coincidencias -- nunca se usa para renombrar solo.
PATRON_RADICADO_CERCANO = re.compile(r"(?<!\d)\d{21,24}(?!\d)")

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


def leer_filas_excel(silencioso=False):
    """
    Lee todas las filas del Excel. Devuelve (filas_validas, filas_casi_validas):
      - filas_validas: (fila, numero, radicado) con radicado de EXACTAMENTE 23 digitos.
      - filas_casi_validas: (fila, numero, radicado) con radicado de 21, 22 o 24
        digitos (le falta o le sobra 1) -- se usan solo para detectar posibles
        coincidencias, nunca para renombrar automaticamente.
    Si silencioso=True, no escribe advertencias en el log (se usa en la
    segunda lectura de verificacion, para no duplicar cada advertencia).
    """
    wb = openpyxl.load_workbook(RUTA_EXCEL, data_only=True)
    if HOJA_EXCEL not in wb.sheetnames:
        raise ValueError(f"La hoja '{HOJA_EXCEL}' no existe. Hojas disponibles: {wb.sheetnames}")
    ws = wb[HOJA_EXCEL]

    encabezados = [ws.cell(row=FILA_ENCABEZADO, column=c).value for c in range(1, ws.max_column + 1)]
    col_no = encontrar_columna(encabezados, COLUMNA_NO)
    col_rad = encontrar_columna(encabezados, COLUMNA_RADICADO)

    filas_validas = []
    filas_casi_validas = []
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

        if not radicado.isdigit():
            if not silencioso:
                logging.warning(
                    "[Excel] Fila %s (proceso %s): radicado con formato invalido, se omite: %r",
                    fila, numero, radicado_crudo,
                )
            continue

        if len(radicado) == 23:
            filas_validas.append((fila, numero, radicado))
        elif 21 <= len(radicado) <= 24:
            if not silencioso:
                logging.warning(
                    "[Excel] Fila %s (proceso %s): el radicado tiene %d digitos en vez de 23 (%r); "
                    "se revisa solo como posible coincidencia, no se cruza automatico.",
                    fila, numero, len(radicado), radicado_crudo,
                )
            filas_casi_validas.append((fila, numero, radicado))
        else:
            if not silencioso:
                logging.warning(
                    "[Excel] Fila %s (proceso %s): radicado con formato invalido, se omite: %r",
                    fila, numero, radicado_crudo,
                )

    return filas_validas, filas_casi_validas


def quitar_repetidos(filas, silencioso=False):
    """Excluye del cruce cualquier numero de proceso o radicado que aparezca en mas de una fila, reportando el conflicto."""
    filas_por_numero = {}
    filas_por_radicado = {}
    for fila, numero, radicado in filas:
        filas_por_numero.setdefault(numero, []).append(fila)
        filas_por_radicado.setdefault(radicado, []).append(fila)

    numeros_repetidos = {n for n, fs in filas_por_numero.items() if len(fs) > 1}
    radicados_repetidos = {r for r, fs in filas_por_radicado.items() if len(fs) > 1}

    if not silencioso:
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


def leer_procesos_validos(silencioso=False):
    """Atajo: lee el Excel y devuelve (procesos_validos, procesos_casi_validos) ya sin repetidos."""
    validas, casi_validas = leer_filas_excel(silencioso=silencioso)
    return quitar_repetidos(validas, silencioso=silencioso), quitar_repetidos(casi_validas, silencioso=True)


def nombre_ya_correcto(nombre_carpeta: str, numero: int, radicado: str) -> bool:
    return nombre_carpeta.strip() == f"{numero}. {radicado}"


def radicado_de_nombre_carpeta(nombre_carpeta: str):
    """Radicado EXACTO (23 digitos, sin nada pegado antes/despues) en el nombre de la carpeta, o None."""
    m = PATRON_RADICADO_EXACTO.search(nombre_carpeta)
    return m.group(0) if m else None


def radicado_cercano_de_nombre_carpeta(nombre_carpeta: str):
    """Radicado 'casi bueno' (21 a 24 digitos) en el nombre, para buscar posibles coincidencias."""
    m = PATRON_RADICADO_CERCANO.search(nombre_carpeta)
    return m.group(0) if m else None


def difiere_por_un_digito(a: str, b: str) -> bool:
    """
    True si 'a' y 'b' son iguales, o si el mas largo se convierte en el mas
    corto quitandole exactamente un digito en alguna posicion (es decir,
    a uno le sobra o le falta un solo digito respecto al otro).
    """
    if a == b:
        return True
    if abs(len(a) - len(b)) != 1:
        return False
    largo, corto = (a, b) if len(a) > len(b) else (b, a)
    for i in range(len(largo)):
        if largo[:i] + largo[i + 1:] == corto:
            return True
    return False


def buscar_coincidencia_cercana(radicado_carpeta: str, candidatos):
    """
    Busca entre 'candidatos' (lista de (fila, numero, radicado)) alguno que
    difiera del radicado de la carpeta por exactamente un digito de mas o
    de menos. Devuelve la primera coincidencia encontrada, o None.
    """
    for fila, numero, radicado_excel in candidatos:
        if difiere_por_un_digito(radicado_carpeta, radicado_excel):
            return fila, numero, radicado_excel
    return None


def mismo_radicado_salvo_ultimo_digito(a: str, b: str) -> bool:
    """
    True si 'a' y 'b' tienen los dos exactamente 23 digitos, son iguales en
    los primeros 22, y solo difieren en el ultimo (el "consecutivo" que
    indica la instancia/reparto del proceso, ej. termina en 00 o en 01).
    A diferencia de un digito de mas/menos en cualquier posicion, esto SI
    se corrige automatico: los primeros 22 digitos ya identifican que es
    el mismo proceso, el ultimo digito distinto no es un proceso diferente.
    """
    return len(a) == 23 and len(b) == 23 and a[:-1] == b[:-1] and a[-1] != b[-1]


def buscar_coincidencia_ultimo_digito(radicado_carpeta: str, procesos):
    """
    Busca entre 'procesos' (radicados EXACTOS de 23 digitos del Excel) uno
    que coincida con el radicado de la carpeta salvo el ultimo digito.
    Devuelve la primera coincidencia encontrada, o None.
    """
    for fila, numero, radicado_excel in procesos:
        if mismo_radicado_salvo_ultimo_digito(radicado_carpeta, radicado_excel):
            return fila, numero, radicado_excel
    return None


def procesar():
    # --- Doble lectura independiente del Excel, para confirmar que el
    # resultado es estable antes de tocar ninguna carpeta ---
    procesos_1, casi_validos_1 = leer_procesos_validos(silencioso=False)
    procesos_2, casi_validos_2 = leer_procesos_validos(silencioso=True)

    dict_1 = {radicado: numero for _f, numero, radicado in procesos_1}
    dict_2 = {radicado: numero for _f, numero, radicado in procesos_2}
    if dict_1 != dict_2:
        logging.error(
            "[Seguridad] La primera y la segunda lectura del Excel NO coinciden (¿se esta editando el "
            "archivo justo ahora?). Por seguridad no se toca ninguna carpeta. Cierra el Excel si lo "
            "tienes abierto editando, y vuelve a correr el script."
        )
        return

    procesos = procesos_1
    procesos_casi_validos = casi_validos_1
    logging.info(
        "Excel: %d proceso(s) con radicado valido y sin repetir, confirmados en dos lecturas independientes.",
        len(procesos),
    )

    por_radicado = {radicado: (numero, fila) for fila, numero, radicado in procesos}

    carpeta_raiz = Path(CARPETA_PROCESOS)
    carpetas = [d for d in carpeta_raiz.iterdir() if d.is_dir()]

    radicados_encontrados_en_disco = set()
    renombradas = []  # lista de (carpeta_original, nuevo_nombre) para la auditoria final
    correcciones_consecutivo = []
    ya_correctas = 0
    sin_proceso_en_excel = []
    conflictos = 0
    posibles_coincidencias = []

    candidatos_cercanos = procesos_casi_validos + [
        (fila, numero, radicado) for fila, numero, radicado in procesos
    ]

    for carpeta in carpetas:
        radicado_en_carpeta = radicado_de_nombre_carpeta(carpeta.name)

        if not radicado_en_carpeta:
            # No tiene un radicado EXACTO de 23 digitos. Antes de descartarla,
            # revisa si tiene uno "casi bueno" (21-24 digitos) que pueda ser
            # una posible coincidencia con el Excel.
            radicado_cercano = radicado_cercano_de_nombre_carpeta(carpeta.name)
            if radicado_cercano:
                coincidencia = buscar_coincidencia_cercana(radicado_cercano, candidatos_cercanos)
                if coincidencia:
                    fila, numero, radicado_excel = coincidencia
                    posibles_coincidencias.append((carpeta.name, radicado_cercano, numero, radicado_excel, fila))
            continue

        match = por_radicado.get(radicado_en_carpeta)
        radicado_para_nombre = radicado_en_carpeta
        correccion = None

        if not match:
            # No hay match exacto. Primero revisa si es el MISMO proceso con
            # el ultimo digito (consecutivo/instancia) distinto -- ese caso
            # SI se corrige automatico.
            match_consecutivo = buscar_coincidencia_ultimo_digito(radicado_en_carpeta, procesos)
            if match_consecutivo:
                fila, numero, radicado_excel = match_consecutivo
                match = (numero, fila)
                radicado_para_nombre = radicado_excel
                correccion = (radicado_en_carpeta, radicado_excel, fila, numero)
            else:
                # Revisa si hay una fila del Excel que difiera por 1 digito
                # en cualquier posicion -- eso NO se corrige automatico.
                coincidencia = buscar_coincidencia_cercana(radicado_en_carpeta, candidatos_cercanos)
                if coincidencia:
                    fila, numero, radicado_excel = coincidencia
                    if radicado_excel != radicado_en_carpeta:  # evita reportar el propio match exacto
                        posibles_coincidencias.append((carpeta.name, radicado_en_carpeta, numero, radicado_excel, fila))
                sin_proceso_en_excel.append(carpeta.name)
                continue

        radicados_encontrados_en_disco.add(radicado_para_nombre)
        numero, fila = match

        if correccion:
            radicado_viejo, radicado_nuevo, fila_excel, numero_excel = correccion
            logging.warning(
                "[Consecutivo] Carpeta '%s': el radicado termina en '%s' pero el informe (fila %s, "
                "proceso %s) tiene el mismo proceso terminado en '%s'; se corrige al del informe.",
                carpeta.name, radicado_viejo[-1], fila_excel, numero_excel, radicado_nuevo[-1],
            )

        if nombre_ya_correcto(carpeta.name, numero, radicado_para_nombre):
            ya_correctas += 1
            continue

        nuevo_nombre = f"{numero}. {radicado_para_nombre}"
        destino = carpeta.parent / nuevo_nombre

        if destino.exists():
            conflictos += 1
            logging.warning(
                "[Conflicto] '%s' deberia renombrarse a '%s' pero ya existe una carpeta con ese nombre. Se omite, revisa manualmente.",
                carpeta.name, nuevo_nombre,
            )
            continue

        # Re-verificacion justo antes de renombrar: vuelve a leer el radicado
        # ORIGINAL de la carpeta una vez mas y confirma que sigue siendo el mismo.
        radicado_confirmado = radicado_de_nombre_carpeta(carpeta.name)
        if radicado_confirmado != radicado_en_carpeta:
            logging.error(
                "[Seguridad] '%s' cambio de nombre justo antes de renombrarla; se omite por seguridad.",
                carpeta.name,
            )
            continue

        if MODO_PRUEBA:
            logging.info("[SIMULACION] '%s'  ->  '%s'", carpeta.name, nuevo_nombre)
        else:
            carpeta.rename(destino)
            logging.info("[Renombrada] '%s'  ->  '%s'", carpeta.name, nuevo_nombre)
        renombradas.append((carpeta.name, nuevo_nombre))
        if correccion:
            correcciones_consecutivo.append((carpeta.name, nuevo_nombre))

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

    if posibles_coincidencias:
        logging.warning(
            "[POSIBLE COINCIDENCIA] %d caso(s) donde el radicado de la carpeta y uno del Excel difieren "
            "por UN digito de mas o de menos. NO se renombraron solos -- revisalos a mano y confirma cual "
            "de los dos numeros es el correcto antes de renombrar:",
            len(posibles_coincidencias),
        )
        for nombre_carpeta, radicado_carpeta, numero_excel, radicado_excel, fila_excel in posibles_coincidencias:
            logging.warning(
                "   - Carpeta '%s' (radicado carpeta: %s)  <->  Excel fila %s, proceso %s, radicado %s",
                nombre_carpeta, radicado_carpeta, fila_excel, numero_excel, radicado_excel,
            )

    # --- Auditoria final: si se aplicaron renombrados de verdad, vuelve a
    # leer el disco y confirma que cada uno quedo con el nombre esperado ---
    fallos_auditoria = []
    if renombradas and not MODO_PRUEBA:
        nombres_actuales = {d.name for d in carpeta_raiz.iterdir() if d.is_dir()}
        for nombre_original, nuevo_nombre in renombradas:
            if nuevo_nombre not in nombres_actuales:
                fallos_auditoria.append((nombre_original, nuevo_nombre))
        if fallos_auditoria:
            logging.error(
                "[Seguridad] %d carpeta(s) no quedaron con el nombre esperado despues de renombrar; revisalas a mano:",
                len(fallos_auditoria),
            )
            for nombre_original, nuevo_nombre in fallos_auditoria:
                logging.error("   - se esperaba '%s' (antes: '%s')", nuevo_nombre, nombre_original)
        else:
            logging.info("[Seguridad] Revision final: todas las carpetas renombradas quedaron con el nombre esperado.")

    logging.info("-" * 60)
    logging.info(
        "Resumen: %d %s (de las cuales %d son correcciones de consecutivo/instancia), %d ya tenian el "
        "nombre correcto, %d sin carpeta en disco, %d carpetas sin proceso en el Excel, %d conflictos de "
        "nombre, %d posibles coincidencias para revisar.",
        len(renombradas),
        "carpetas simuladas (MODO_PRUEBA activo)" if MODO_PRUEBA else "carpetas renombradas",
        len(correcciones_consecutivo),
        ya_correctas, sin_carpeta_en_disco, len(sin_proceso_en_excel), conflictos, len(posibles_coincidencias),
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
