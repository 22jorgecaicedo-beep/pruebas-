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
    carpeta (y cada carpeta duplicada movida) quedo donde se esperaba.

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

Carpetas DUPLICADAS (mismo radicado exacto en mas de una carpeta, tipico
de descargas repetidas del mismo proceso): el script NUNCA borra nada.
Se queda con la carpeta que tenga MAS ARCHIVOS adentro (la mas completa),
la renombra "numero. radicado" usando el numero MAS RECIENTE del Excel
cargado en RUTA_EXCEL, y mueve las demas copias (sin tocar su contenido)
a una carpeta "Duplicados_para_revisar" dentro de CARPETA_PROCESOS, para
que las revises y borres a mano si de verdad sobran.

Carpetas ANIDADAS (una carpeta de proceso metida DENTRO de otra carpeta
de proceso, ej. "1014. radicado" dentro de "941. radicado"): tambien se
resuelven solas, sin borrar ni fusionar contenido -- solo se mueve la
carpeta completa:
  - Si el radicado de la anidada es el MISMO que el de la carpeta que la
    contiene, se mueve a "Duplicados_para_revisar" (es una copia vieja
    del mismo caso).
  - Si el radicado es DISTINTO (contenido de otro caso que quedo mal
    ubicado), se saca al nivel principal del disco para que se evalue
    normal contra el Excel en la proxima corrida.

Verificacion mas profunda (solo detecta y reporta, no modifica nada):
  - Carpetas VACIAS (sin ningun archivo adentro): se revisa si hay un
    .zip en CARPETA_DESCARGAS cuyo nombre tenga ese mismo radicado, por
    si quedo pendiente de extraer. Se reporta, no se extrae solo.
  - Carpetas SIN NINGUN radicado reconocible en el nombre (ni siquiera
    21-24 digitos): antes se ignoraban en silencio, ahora se reportan
    aparte para que las revises a mano.
  - CONTENIDO que no corresponde al nombre (VALIDAR_CONTENIDO_CONTRA_NOMBRE):
    abre los documentos dentro de cada carpeta (nombres de archivo, y si
    hace falta el texto de hasta MAX_ARCHIVOS_CONTENIDO_A_REVISAR PDF/DOCX)
    y busca que radicados aparecen. Si el radicado del NOMBRE de la
    carpeta nunca aparece adentro, pero otro radicado si aparece
    claramente, se reporta como sospechoso de contenido mal ubicado. No
    se marca si el propio radicado SI aparece (aunque tambien aparezcan
    otros, por referencias cruzadas a casos relacionados).

Al terminar, reporta (en pantalla y en un log):
  - Carpetas renombradas (o que se renombrarian, en modo prueba).
  - Carpetas duplicadas resueltas (cual se conservo, cuales se movieron).
  - Carpetas que ya tenian el nombre correcto (se dejan igual).
  - Procesos del Excel sin carpeta correspondiente en el disco.
  - Carpetas en el disco cuyo radicado no aparece en el Excel.
  - Carpetas sin ningun radicado reconocible en el nombre.
  - Carpetas vacias, y si se encontro un zip pendiente en Descargas.
  - Posibles coincidencias con un digito de mas o de menos (revisar a mano).
  - Filas del Excel con radicado invalido o con numero/radicado repetido
    (no se tocan, para no arriesgar un cruce incorrecto).

Por defecto corre en MODO_PRUEBA (no renombra ni mueve nada, solo muestra
que haria). Revisa el reporte y, cuando confies en que el cruce esta bien,
cambia MODO_PRUEBA a False para aplicar los cambios de verdad.
"""

import csv
import logging
import os
import re
from collections import Counter
from pathlib import Path

import openpyxl

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

try:
    import docx
except ImportError:
    docx = None

# ============================= CONFIGURACION =============================

# Ruta al informe de Excel (.xlsx o .xlsm).
RUTA_EXCEL = os.path.join(os.path.expanduser("~"), "Desktop", "PARA REVISION.xlsm")

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

# Carpeta donde caen tus descargas (para revisar si una carpeta vacia tiene
# un .zip pendiente de extraer ahi). Se detecta sola como "Downloads" del
# usuario de Windows actual; cambiala si tu carpeta de Descargas esta en
# otro lado. Si la ruta no existe, esta revision simplemente se omite.
CARPETA_DESCARGAS = os.path.join(os.path.expanduser("~"), "Downloads")

# True: al final, abre los documentos DENTRO de cada carpeta y revisa si
# el radicado que aparece en su contenido corresponde con el radicado del
# NOMBRE de la carpeta (detecta casos donde el contenido quedo mal
# ubicado). Esto tarda mas en correr porque tiene que leer PDFs/DOCX de
# todas las carpetas; ponlo en False si prefieres una corrida rapida.
VALIDAR_CONTENIDO_CONTRA_NOMBRE = True

# Cuantos archivos (como maximo) se leen POR CARPETA cuando ningun nombre
# de archivo trae el radicado (para no tener que abrir los 50 PDF de una
# carpeta con muchos documentos, con unos pocos alcanza para verificar).
MAX_ARCHIVOS_CONTENIDO_A_REVISAR = 5

# True: no renombra ni mueve nada, solo muestra/registra que haria
# (recomendado la primera vez). False: aplica los cambios de verdad.
MODO_PRUEBA = True

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "validar_renombrar_carpetas.log")
ARCHIVO_REPORTE_VACIAS = os.path.join(os.path.dirname(__file__), "carpetas_vacias.csv")
ARCHIVO_REPORTE_CONTENIDO = os.path.join(os.path.dirname(__file__), "contenido_no_corresponde.csv")

# Carpeta donde se mueven (nunca se borran) las copias duplicadas sobrantes.
NOMBRE_CARPETA_DUPLICADOS = "Duplicados_para_revisar"

# Radicado "bueno": exactamente 23 digitos, sin otro digito pegado antes o
# despues (para no cortar mal un numero mas largo, ni colar uno mas corto).
# No usamos \b porque \b trata "_" como parte de la palabra -- por ejemplo
# una carpeta duplicada "..._2" dejaria de reconocerse; aqui solo nos
# importa que no haya OTRO DIGITO pegado.
PATRON_RADICADO_EXACTO = re.compile(r"(?<!\d)\d{23}(?!\d)")

# Radicado "casi bueno" (con 1 digito de mas o de menos), solo para poder
# reportar posibles coincidencias -- nunca se usa para renombrar solo.
PATRON_RADICADO_CERCANO = re.compile(r"(?<!\d)\d{21,24}(?!\d)")

# Radicado escrito con separadores (ej. "68005-40-03-001-2023-00700"), tal
# como a veces aparece DENTRO del texto de un documento (no en nombres de
# carpeta). Solo se usa para la validacion de contenido.
PATRON_RADICADO_CON_SEPARADORES = re.compile(
    r"(?<!\d)\d{5}[\s\-]?\d{2}[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{4}[\s\-]?\d{5}[\s\-]?\d{2}(?!\d)"
)

EXTENSIONES_CONTENIDO = {".pdf", ".docx"}

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


def contar_archivos(carpeta: Path) -> int:
    """Cuenta cuantos archivos (no carpetas) hay dentro de una carpeta, recursivamente."""
    try:
        return sum(1 for p in carpeta.rglob("*") if p.is_file())
    except OSError:
        return 0


def _texto_de_pdf(ruta: Path) -> str:
    if PdfReader is None:
        return ""
    try:
        lector = PdfReader(str(ruta))
        return "\n".join((pagina.extract_text() or "") for pagina in lector.pages)
    except Exception:
        return ""


def _texto_de_docx(ruta: Path) -> str:
    if docx is None:
        return ""
    try:
        documento = docx.Document(str(ruta))
        return "\n".join(p.text for p in documento.paragraphs)
    except Exception:
        return ""


def _radicados_en_texto(texto: str):
    """Todos los radicados de 23 digitos encontrados en un texto (no solo el primero)."""
    encontrados = list(PATRON_RADICADO_EXACTO.findall(texto))
    encontrados += [re.sub(r"[\s\-]", "", m) for m in PATRON_RADICADO_CON_SEPARADORES.findall(texto)]
    return encontrados


def radicados_encontrados_en_carpeta(carpeta: Path, max_archivos_contenido: int) -> Counter:
    """
    Recorre los archivos de una carpeta y devuelve un Counter con todos los
    radicados de 23 digitos encontrados: primero en los NOMBRES de
    archivo (rapido, sin abrir nada); si ninguno trae uno, lee el
    contenido de hasta 'max_archivos_contenido' PDF/DOCX como muestra
    (para no tener que abrir todos los documentos de carpetas grandes).
    """
    contador = Counter()
    candidatos_contenido = []

    try:
        rutas = list(carpeta.rglob("*"))
    except OSError:
        return contador

    for ruta in rutas:
        if not ruta.is_file():
            continue
        for radicado in _radicados_en_texto(ruta.name):
            contador[radicado] += 1
        if ruta.suffix.lower() in EXTENSIONES_CONTENIDO:
            candidatos_contenido.append(ruta)

    if contador:
        return contador

    for ruta in candidatos_contenido[:max_archivos_contenido]:
        texto = _texto_de_pdf(ruta) if ruta.suffix.lower() == ".pdf" else _texto_de_docx(ruta)
        for radicado in _radicados_en_texto(texto):
            contador[radicado] += 1

    return contador


def ruta_libre(carpeta_padre: Path, nombre: str) -> Path:
    """Como el nombre sugiere: la primera ruta dentro de carpeta_padre/nombre[_N] que no exista todavia."""
    destino = carpeta_padre / nombre
    contador = 2
    while destino.exists():
        destino = carpeta_padre / f"{nombre}_{contador}"
        contador += 1
    return destino


def subcarpetas_con_radicado(carpeta_padre: Path):
    """Subcarpetas DIRECTAS (un solo nivel) de carpeta_padre que tengan un radicado EXACTO de 23 digitos en su nombre."""
    encontradas = []
    try:
        for hijo in carpeta_padre.iterdir():
            if hijo.is_dir():
                radicado_hijo = radicado_de_nombre_carpeta(hijo.name)
                if radicado_hijo:
                    encontradas.append((hijo, radicado_hijo))
    except OSError:
        pass
    return encontradas


def _buscar_zip_en_carpeta(carpeta: Path, radicado: str):
    try:
        for archivo in carpeta.iterdir():
            if archivo.is_file() and archivo.suffix.lower() == ".zip":
                if radicado_de_nombre_carpeta(archivo.stem) == radicado or radicado_cercano_de_nombre_carpeta(archivo.stem) == radicado:
                    return archivo.name
    except OSError:
        pass
    return None


def buscar_zip_con_radicado(carpeta_descargas: Path, radicado: str):
    """
    Busca un .zip cuyo nombre contenga ese radicado exacto, primero en
    carpeta_descargas directamente y despues en su subcarpeta
    "Procesados" (ahi es donde procesos_juridicos.py movia el zip incluso
    cuando la extraccion fallaba por completo, antes de que eso se
    corrigiera). Devuelve (nombre_archivo, "Descargas" o "Descargas/Procesados"), o None.
    """
    encontrado = _buscar_zip_en_carpeta(carpeta_descargas, radicado)
    if encontrado:
        return encontrado, "Descargas"

    carpeta_procesados = carpeta_descargas / "Procesados"
    if carpeta_procesados.exists():
        encontrado = _buscar_zip_en_carpeta(carpeta_procesados, radicado)
        if encontrado:
            return encontrado, "Descargas/Procesados"

    return None


def intentar_renombrar_carpeta(carpeta: Path, numero: int, radicado_final: str, radicado_original: str, reporte: dict) -> bool:
    """
    Intenta renombrar 'carpeta' a 'numero. radicado_final', con las
    verificaciones de seguridad de siempre (evita pisar una carpeta que ya
    exista, y vuelve a confirmar el radicado justo antes de tocar nada).
    Actualiza los contadores/listas del 'reporte' segun corresponda.
    Devuelve True si ya estaba bien, se renombro, o se simulo; False si
    hubo un conflicto o fallo la re-verificacion (no se toco la carpeta).
    """
    if nombre_ya_correcto(carpeta.name, numero, radicado_final):
        reporte["ya_correctas"] += 1
        return True

    nuevo_nombre = f"{numero}. {radicado_final}"
    destino = carpeta.parent / nuevo_nombre

    if destino.exists() and destino.resolve() != carpeta.resolve():
        reporte["conflictos"] += 1
        logging.warning(
            "[Conflicto] '%s' deberia renombrarse a '%s' pero ya existe una carpeta con ese nombre. Se omite, revisa manualmente.",
            carpeta.name, nuevo_nombre,
        )
        return False

    # Re-verificacion justo antes de renombrar: vuelve a leer el radicado
    # ORIGINAL de la carpeta una vez mas y confirma que sigue siendo el mismo.
    radicado_confirmado = radicado_de_nombre_carpeta(carpeta.name)
    if radicado_confirmado != radicado_original:
        logging.error(
            "[Seguridad] '%s' cambio de nombre justo antes de renombrarla; se omite por seguridad.",
            carpeta.name,
        )
        return False

    if MODO_PRUEBA:
        logging.info("[SIMULACION] '%s'  ->  '%s'", carpeta.name, nuevo_nombre)
    else:
        carpeta.rename(destino)
        logging.info("[Renombrada] '%s'  ->  '%s'", carpeta.name, nuevo_nombre)
    reporte["renombradas"].append((carpeta.name, nuevo_nombre))
    return True


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
    carpeta_duplicados = carpeta_raiz / NOMBRE_CARPETA_DUPLICADOS
    carpetas = [d for d in carpeta_raiz.iterdir() if d.is_dir() and d.name != NOMBRE_CARPETA_DUPLICADOS]

    # Agrupa las carpetas por su radicado EXACTO (ignora prefijo "numero."
    # y cualquier sufijo tipo "_2"). Las que no tengan un radicado exacto
    # de 23 digitos se procesan aparte (solo para buscar posibles
    # coincidencias por largo distinto).
    grupos_por_radicado = {}
    carpetas_sin_radicado_exacto = []
    for carpeta in carpetas:
        radicado = radicado_de_nombre_carpeta(carpeta.name)
        if radicado:
            grupos_por_radicado.setdefault(radicado, []).append(carpeta)
        else:
            carpetas_sin_radicado_exacto.append(carpeta)

    reporte = {
        "renombradas": [],  # (nombre_original, nuevo_nombre)
        "ya_correctas": 0,
        "conflictos": 0,
    }
    radicados_encontrados_en_disco = set()
    sin_proceso_en_excel = []
    posibles_coincidencias = []
    duplicados_resueltos = []  # (radicado, nombre_conservado_nuevo, [(nombre_movido, destino_dup)])
    duplicados_sin_resolver = []  # (radicado, [nombres]) -- no se pudo determinar el numero
    movidos_a_auditar = []  # (nombre_original, ruta_destino) para la revision final

    candidatos_cercanos = procesos_casi_validos + [
        (fila, numero, radicado) for fila, numero, radicado in procesos
    ]

    for radicado_en_carpeta, lista_carpetas in grupos_por_radicado.items():
        match = por_radicado.get(radicado_en_carpeta)
        radicado_final = radicado_en_carpeta
        correccion = None

        if not match:
            match_consecutivo = buscar_coincidencia_ultimo_digito(radicado_en_carpeta, procesos)
            if match_consecutivo:
                fila, numero, radicado_excel = match_consecutivo
                match = (numero, fila)
                radicado_final = radicado_excel
                correccion = (radicado_en_carpeta, radicado_excel, fila, numero)

        # --- Un solo radicado, una sola carpeta: caso normal ---
        if len(lista_carpetas) == 1:
            carpeta = lista_carpetas[0]

            if not match:
                coincidencia = buscar_coincidencia_cercana(radicado_en_carpeta, candidatos_cercanos)
                if coincidencia:
                    fila, numero, radicado_excel = coincidencia
                    if radicado_excel != radicado_en_carpeta:
                        posibles_coincidencias.append((carpeta.name, radicado_en_carpeta, numero, radicado_excel, fila))
                sin_proceso_en_excel.append(carpeta.name)
                continue

            numero, fila = match
            if correccion:
                radicado_viejo, radicado_nuevo, fila_excel, numero_excel = correccion
                logging.warning(
                    "[Consecutivo] Carpeta '%s': el radicado termina en '%s' pero el informe (fila %s, "
                    "proceso %s) tiene el mismo proceso terminado en '%s'; se corrige al del informe.",
                    carpeta.name, radicado_viejo[-1], fila_excel, numero_excel, radicado_nuevo[-1],
                )

            radicados_encontrados_en_disco.add(radicado_final)
            intentar_renombrar_carpeta(carpeta, numero, radicado_final, radicado_en_carpeta, reporte)
            continue

        # --- Mas de una carpeta con el MISMO radicado exacto: duplicadas ---
        if not match:
            duplicados_sin_resolver.append((radicado_en_carpeta, [c.name for c in lista_carpetas]))
            for carpeta in lista_carpetas:
                sin_proceso_en_excel.append(carpeta.name)
            continue

        numero, fila = match
        radicados_encontrados_en_disco.add(radicado_final)

        if correccion:
            radicado_viejo, radicado_nuevo, fila_excel, numero_excel = correccion
            logging.warning(
                "[Consecutivo] Grupo duplicado con radicado terminado en '%s': el informe (fila %s, "
                "proceso %s) tiene el mismo proceso terminado en '%s'; se corrige al del informe.",
                radicado_viejo[-1], fila_excel, numero_excel, radicado_nuevo[-1],
            )

        # Se conserva la carpeta con MAS ARCHIVOS adentro (la mas completa).
        conteos = [(carpeta, contar_archivos(carpeta)) for carpeta in lista_carpetas]
        conteos.sort(key=lambda par: par[1], reverse=True)
        carpeta_conservar, archivos_conservar = conteos[0]
        otras = conteos[1:]

        nuevo_nombre = f"{numero}. {radicado_final}"
        movidas = []

        for carpeta_extra, archivos_extra in otras:
            destino_dup = ruta_libre(carpeta_duplicados, carpeta_extra.name)
            if MODO_PRUEBA:
                logging.info(
                    "[SIMULACION-Duplicado] '%s' (%d archivo(s)) se moveria a '%s/%s' -- se conserva '%s' (%d archivo(s)) como '%s'.",
                    carpeta_extra.name, archivos_extra, NOMBRE_CARPETA_DUPLICADOS, destino_dup.name,
                    carpeta_conservar.name, archivos_conservar, nuevo_nombre,
                )
            else:
                carpeta_duplicados.mkdir(parents=True, exist_ok=True)
                carpeta_extra.rename(destino_dup)
                logging.info(
                    "[Duplicado] '%s' (%d archivo(s)) se movio a '%s/%s' -- se conserva '%s' (%d archivo(s)) como '%s'.",
                    carpeta_extra.name, archivos_extra, NOMBRE_CARPETA_DUPLICADOS, destino_dup.name,
                    carpeta_conservar.name, archivos_conservar, nuevo_nombre,
                )
                movidos_a_auditar.append((carpeta_extra.name, destino_dup))
            movidas.append((carpeta_extra.name, destino_dup.name))

        intentar_renombrar_carpeta(carpeta_conservar, numero, radicado_final, radicado_en_carpeta, reporte)
        duplicados_resueltos.append((radicado_final, nuevo_nombre, movidas))

    # --- Carpetas sin radicado EXACTO de 23 digitos: solo se revisan por
    # si tienen un radicado "casi bueno" (largo distinto) coincidente ---
    for carpeta in carpetas_sin_radicado_exacto:
        radicado_cercano = radicado_cercano_de_nombre_carpeta(carpeta.name)
        if radicado_cercano:
            coincidencia = buscar_coincidencia_cercana(radicado_cercano, candidatos_cercanos)
            if coincidencia:
                fila, numero, radicado_excel = coincidencia
                posibles_coincidencias.append((carpeta.name, radicado_cercano, numero, radicado_excel, fila))

    # --- Carpetas de proceso ANIDADAS dentro de otra carpeta de proceso
    # (ej. "1014. radicado" metida dentro de "941. radicado"). Se revisan
    # DESPUES de lo anterior, releyendo el disco de verdad, para que
    # reflejen los nombres ya corregidos en este mismo corrida. Nunca se
    # borra ni se fusiona contenido -- solo se mueve la carpeta completa.
    anidadas_mismo_caso = []  # (carpeta_padre, nombre_anidada, nombre_destino)
    anidadas_otro_caso = []   # (carpeta_padre, nombre_anidada, radicado_anidado, nombre_destino)

    carpetas_nivel_superior_ahora = [
        d for d in carpeta_raiz.iterdir() if d.is_dir() and d.name != NOMBRE_CARPETA_DUPLICADOS
    ]
    for carpeta_padre in carpetas_nivel_superior_ahora:
        radicado_padre = radicado_de_nombre_carpeta(carpeta_padre.name)
        if not radicado_padre:
            continue

        for hijo, radicado_hijo in subcarpetas_con_radicado(carpeta_padre):
            if radicado_hijo == radicado_padre:
                destino = ruta_libre(carpeta_duplicados, hijo.name)
                if MODO_PRUEBA:
                    logging.info(
                        "[SIMULACION-Anidada] '%s' esta metida dentro de '%s' y es una copia del MISMO caso "
                        "(mismo radicado); se moveria a '%s/%s'.",
                        hijo.name, carpeta_padre.name, NOMBRE_CARPETA_DUPLICADOS, destino.name,
                    )
                else:
                    carpeta_duplicados.mkdir(parents=True, exist_ok=True)
                    hijo.rename(destino)
                    logging.info(
                        "[Anidada] '%s' estaba metida dentro de '%s' (mismo radicado); se movio a '%s/%s'.",
                        hijo.name, carpeta_padre.name, NOMBRE_CARPETA_DUPLICADOS, destino.name,
                    )
                    movidos_a_auditar.append((hijo.name, destino))
                anidadas_mismo_caso.append((carpeta_padre.name, hijo.name, destino.name))
            else:
                destino = ruta_libre(carpeta_raiz, hijo.name)
                if MODO_PRUEBA:
                    logging.info(
                        "[SIMULACION-Anidada] '%s' esta metida dentro de '%s' pero tiene un radicado DISTINTO "
                        "(%s); se sacaria al nivel principal del disco como '%s' para evaluarla en la proxima corrida.",
                        hijo.name, carpeta_padre.name, radicado_hijo, destino.name,
                    )
                else:
                    hijo.rename(destino)
                    logging.info(
                        "[Anidada] '%s' estaba metida dentro de '%s' con un radicado DISTINTO (%s); se saco al "
                        "nivel principal del disco como '%s' para evaluarla en la proxima corrida.",
                        hijo.name, carpeta_padre.name, radicado_hijo, destino.name,
                    )
                    movidos_a_auditar.append((hijo.name, destino))
                anidadas_otro_caso.append((carpeta_padre.name, hijo.name, radicado_hijo, destino.name))

    # --- Verificacion mas profunda: carpetas vacias (y si hay un zip sin
    # procesar en Descargas que parezca ser el mismo caso) y carpetas sin
    # ningun radicado reconocible en el nombre (antes se ignoraban en
    # silencio) ---
    carpeta_descargas = None
    if CARPETA_DESCARGAS:
        candidata = Path(CARPETA_DESCARGAS)
        if candidata.exists():
            carpeta_descargas = candidata
        else:
            logging.warning(
                "[Descargas] CARPETA_DESCARGAS configurada (%s) no existe; se omite la busqueda de zips pendientes.",
                CARPETA_DESCARGAS,
            )

    carpetas_vacias = []       # (nombre, radicado_o_None, zip_encontrado_o_None)
    carpetas_sin_radicado = []  # nombres sin NINGUN radicado reconocible (ni exacto ni cercano)
    contenido_no_corresponde = []  # (nombre, radicado_esperado, radicado_dominante_en_contenido, veces)

    carpetas_finales = [d for d in carpeta_raiz.iterdir() if d.is_dir() and d.name != NOMBRE_CARPETA_DUPLICADOS]
    for carpeta in carpetas_finales:
        radicado_exacto = radicado_de_nombre_carpeta(carpeta.name)
        radicado_actual = radicado_exacto or radicado_cercano_de_nombre_carpeta(carpeta.name)
        if not radicado_actual:
            carpetas_sin_radicado.append(carpeta.name)

        numero_archivos = contar_archivos(carpeta)
        if numero_archivos == 0:
            zip_encontrado, zip_ubicacion = None, None
            if radicado_actual and carpeta_descargas:
                resultado = buscar_zip_con_radicado(carpeta_descargas, radicado_actual)
                if resultado:
                    zip_encontrado, zip_ubicacion = resultado
            carpetas_vacias.append((carpeta.name, radicado_actual, zip_encontrado, zip_ubicacion))
        elif VALIDAR_CONTENIDO_CONTRA_NOMBRE and radicado_exacto:
            radicados_hallados = radicados_encontrados_en_carpeta(carpeta, MAX_ARCHIVOS_CONTENIDO_A_REVISAR)
            if radicados_hallados and radicado_exacto not in radicados_hallados:
                radicado_dominante, veces = radicados_hallados.most_common(1)[0]
                contenido_no_corresponde.append((carpeta.name, radicado_exacto, radicado_dominante, veces))

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

    if carpetas_sin_radicado:
        logging.warning(
            "[Sin nombre reconocible] %d carpeta(s) no tienen ningun numero que se parezca a un radicado en "
            "su nombre; revisalas a mano:",
            len(carpetas_sin_radicado),
        )
        for nombre in carpetas_sin_radicado:
            logging.warning("   - %s", nombre)

    if carpetas_vacias:
        logging.warning("[Carpeta vacia] %d carpeta(s) no tienen ningun archivo adentro:", len(carpetas_vacias))
        for nombre, radicado_buscado, zip_encontrado, zip_ubicacion in carpetas_vacias:
            if zip_encontrado:
                logging.warning(
                    "   - '%s': vacia, y encontre un .zip en %s que parece ser el mismo caso: '%s' -- "
                    "revisalo, puede que la extraccion haya fallado (ej. protegido con contrasena) o quede "
                    "pendiente.",
                    nombre, zip_ubicacion, zip_encontrado,
                )
            elif radicado_buscado:
                logging.warning(
                    "   - '%s': vacia, no encontre ningun .zip (ni en Descargas ni en Procesados) con el "
                    "radicado %s. Puede que el zip original ya no exista, o que el radicado de esta carpeta "
                    "sea el equivocado (ver [Contenido no corresponde] mas abajo).",
                    nombre, radicado_buscado,
                )
            else:
                logging.warning("   - '%s': vacia y sin radicado reconocible en el nombre.", nombre)

    with open(ARCHIVO_REPORTE_VACIAS, "w", newline="", encoding="utf-8-sig") as f:
        escritor = csv.writer(f, delimiter=";")
        escritor.writerow(["Carpeta", "Radicado", "Zip pendiente encontrado", "Donde se encontro"])
        for nombre, radicado_buscado, zip_encontrado, zip_ubicacion in carpetas_vacias:
            escritor.writerow([nombre, radicado_buscado or "", zip_encontrado or "", zip_ubicacion or ""])
    if carpetas_vacias:
        logging.info("[Carpeta vacia] Reporte guardado en: %s", ARCHIVO_REPORTE_VACIAS)

    if contenido_no_corresponde:
        logging.warning(
            "[Contenido no corresponde] %d carpeta(s): el radicado del NOMBRE nunca aparece dentro de sus "
            "propios documentos, y en cambio se encontro otro radicado -- revisa si el contenido quedo mal "
            "ubicado (por ejemplo por un zip que se proceso mal antes de esta correccion):",
            len(contenido_no_corresponde),
        )
        for nombre, radicado_esperado, radicado_encontrado, veces in contenido_no_corresponde:
            logging.warning(
                "   - '%s': el nombre dice %s, pero encontre %s en su contenido (%d vez/veces).",
                nombre, radicado_esperado, radicado_encontrado, veces,
            )

    with open(ARCHIVO_REPORTE_CONTENIDO, "w", newline="", encoding="utf-8-sig") as f:
        escritor = csv.writer(f, delimiter=";")
        escritor.writerow(["Carpeta", "Radicado del nombre", "Radicado encontrado en el contenido", "Veces"])
        for nombre, radicado_esperado, radicado_encontrado, veces in contenido_no_corresponde:
            escritor.writerow([nombre, radicado_esperado, radicado_encontrado, veces])
    if contenido_no_corresponde:
        logging.info("[Contenido no corresponde] Reporte guardado en: %s", ARCHIVO_REPORTE_CONTENIDO)

    if duplicados_sin_resolver:
        logging.warning(
            "[Duplicado sin resolver] %d radicado(s) con mas de una carpeta en el disco, pero el radicado "
            "no aparece en el Excel -- no se pudo determinar cual conservar. Revisa a mano:",
            len(duplicados_sin_resolver),
        )
        for radicado, nombres in duplicados_sin_resolver:
            logging.warning("   - Radicado %s: %s", radicado, ", ".join(nombres))

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

    if duplicados_resueltos:
        logging.info(
            "[Duplicados resueltos] %d radicado(s) tenian mas de una carpeta; se conservo la mas completa "
            "y se movieron las demas a '%s' (nada se borro):",
            len(duplicados_resueltos), NOMBRE_CARPETA_DUPLICADOS,
        )
        for radicado, nombre_conservado, movidas in duplicados_resueltos:
            for nombre_movido, nombre_destino in movidas:
                logging.info("   - Radicado %s: se conservo '%s'; se movio '%s' -> '%s/%s'",
                             radicado, nombre_conservado, nombre_movido, NOMBRE_CARPETA_DUPLICADOS, nombre_destino)

    if anidadas_mismo_caso:
        logging.info(
            "[Anidadas resueltas] %d carpeta(s) estaban metidas dentro de otra carpeta del MISMO caso; "
            "se sacaron a '%s' (nada se borro):",
            len(anidadas_mismo_caso), NOMBRE_CARPETA_DUPLICADOS,
        )
        for nombre_padre, nombre_anidada, nombre_destino in anidadas_mismo_caso:
            logging.info("   - '%s' (estaba dentro de '%s') -> '%s/%s'",
                         nombre_anidada, nombre_padre, NOMBRE_CARPETA_DUPLICADOS, nombre_destino)

    if anidadas_otro_caso:
        logging.warning(
            "[Anidadas de otro caso] %d carpeta(s) estaban metidas dentro de la carpeta de OTRO proceso "
            "(radicado distinto); se sacaron al nivel principal del disco para evaluarlas en la proxima corrida:",
            len(anidadas_otro_caso),
        )
        for nombre_padre, nombre_anidada, radicado_anidado, nombre_destino in anidadas_otro_caso:
            logging.warning("   - '%s' (radicado %s, estaba dentro de '%s') -> '%s'",
                            nombre_anidada, radicado_anidado, nombre_padre, nombre_destino)

    # --- Auditoria final: si se aplicaron cambios de verdad, vuelve a leer
    # el disco y confirma que cada carpeta (renombrada o movida) quedo
    # donde se esperaba ---
    renombradas = reporte["renombradas"]
    fallos_auditoria = []
    if (renombradas or movidos_a_auditar) and not MODO_PRUEBA:
        nombres_actuales = {d.name for d in carpeta_raiz.iterdir() if d.is_dir()}
        for nombre_original, nuevo_nombre in renombradas:
            if nuevo_nombre not in nombres_actuales:
                fallos_auditoria.append((nombre_original, nuevo_nombre))
        for nombre_original, ruta_destino in movidos_a_auditar:
            if not ruta_destino.exists():
                fallos_auditoria.append((nombre_original, str(ruta_destino)))

        if fallos_auditoria:
            logging.error(
                "[Seguridad] %d carpeta(s) no quedaron donde se esperaba despues de aplicar los cambios; revisalas a mano:",
                len(fallos_auditoria),
            )
            for nombre_original, esperado in fallos_auditoria:
                logging.error("   - se esperaba '%s' (antes: '%s')", esperado, nombre_original)
        else:
            logging.info("[Seguridad] Revision final: todas las carpetas quedaron donde se esperaba.")

    logging.info("-" * 60)
    logging.info(
        "Resumen: %d %s, %d ya tenian el nombre correcto, %d duplicado(s) resuelto(s) (movidos a %s), "
        "%d carpeta(s) anidada(s) del mismo caso resueltas, %d carpeta(s) anidada(s) de otro caso sacadas, "
        "%d sin carpeta en disco, %d carpetas sin proceso en el Excel, %d conflictos de nombre, "
        "%d posibles coincidencias para revisar, %d grupo(s) duplicado(s) sin poder resolver, "
        "%d carpeta(s) vacia(s), %d carpeta(s) sin nombre reconocible, "
        "%d carpeta(s) con contenido que no corresponde al nombre.",
        len(renombradas),
        "carpetas simuladas (MODO_PRUEBA activo)" if MODO_PRUEBA else "carpetas renombradas",
        reporte["ya_correctas"], len(duplicados_resueltos), NOMBRE_CARPETA_DUPLICADOS,
        len(anidadas_mismo_caso), len(anidadas_otro_caso),
        sin_carpeta_en_disco, len(sin_proceso_en_excel), reporte["conflictos"], len(posibles_coincidencias),
        len(duplicados_sin_resolver), len(carpetas_vacias), len(carpetas_sin_radicado),
        len(contenido_no_corresponde),
    )
    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se renombro ni se movio nada todavia. Revisa el reporte de arriba "
            "y, si se ve bien, cambia MODO_PRUEBA = False al inicio del script y vuelve a correrlo."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
