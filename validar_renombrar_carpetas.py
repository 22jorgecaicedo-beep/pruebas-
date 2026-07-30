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

RADICADO REPETIDO EN EL EXCEL (el mismo radicado aparece en dos o mas
filas, con numeros de proceso DISTINTOS -- ej. proceso 19 y proceso 451
con exactamente el mismo radicado): en vez de excluir esas filas del
cruce, el script DUPLICA la carpeta: conserva/renombra la carpeta
existente con el PRIMER numero de proceso, y crea una copia COMPLETA
(mismo contenido) por cada numero adicional, para que cada numero de
proceso tenga su propia carpeta "numero. radicado". No se borra ni se
modifica el Excel -- se asume que si el radicado esta dos veces a
proposito, cada proceso necesita su copia. Al final se informa
claramente cuales radicados se duplicaron y en que carpetas quedaron.

Estos dos casos (consecutivo con ultimo digito distinto, y radicado
repetido en el Excel) tambien se reconocen COMBINADOS: si una carpeta
suelta en el disco tiene el radicado terminado en un digito distinto
al del Excel, Y ADEMAS ese radicado (el del Excel) esta repetido con
varios numeros de proceso, el script corrige el ultimo digito Y
duplica la carpeta para cada numero -- antes esa combinacion no se
reconocia por ninguno de los dos mecanismos (la tolerancia de
consecutivo solo buscaba entre los radicados SIN repetir en el Excel),
y la carpeta se quedaba sin cruzar con nada.

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
  - Carpetas VACIAS (sin ningun archivo real adentro -- no cuentan
    "basura" que crea Windows solo, como desktop.ini o Thumbs.db): se
    revisa si hay un
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

Carpeta de entrada ADICIONAL (CARPETA_ENTRADA_ADICIONAL, ej. una entrega
masiva de expedientes que todavia no se paso a la raiz del disco): si
existe, se revisa antes que la raiz. Los procesos que ya esten en la raiz
se dejan ahi (se informa); los que tengan radicado valido en el Excel y
AUN NO esten en la raiz se MUEVEN a la raiz para que el resto de esta
misma corrida los termine de nombrar; los que no tengan proceso en el
Excel se dejan donde estan y se informan aparte.

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

El reporte de "procesos faltantes" (ARCHIVO_REPORTE_FALTANTES) incluye,
si la columna DEMANDADO existe en el Excel, el demandado de cada
proceso -- lo usa despues buscar_faltantes_en_drive.py para no
confundir dos procesos DISTINTOS que comparten el mismo radicado corto
o cuenta pero van contra demandados diferentes (ej. "2024-00139 CONTRA
RIONEGRO" no es lo mismo que "2024-00139 CONTRA BOLIVAR"). Si esa
columna no existe, el resto del reporte funciona igual, solo queda
vacia esa parte.
"""

import csv
import logging
import os
import re
import shutil
import zipfile
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


def encontrar_disco_por_etiqueta(etiqueta_buscada: str):
    """
    Busca, entre TODAS las unidades conectadas (A: a Z:), la que tenga como
    ETIQUETA DE VOLUMEN (el nombre del disco -- se ve en "Este equipo" y en
    Propiedades del disco) el texto 'etiqueta_buscada' (sin importar
    mayusculas/minusculas ni espacios de mas), y devuelve su ruta actual
    (ej. "D:/"). Esto es a proposito para NO depender de una letra de
    unidad fija: Windows puede asignarle una letra distinta al mismo disco
    duro externo cada vez que se conecta (por ejemplo D: una vez y E: otra,
    segun que otros dispositivos esten conectados), y una letra fija en el
    codigo se desincroniza tarde o temprano.
    Devuelve None si no corre en Windows, o si no encuentra ningun disco
    con esa etiqueta conectado en este momento (por ejemplo si el disco
    esta desconectado).
    """
    if os.name != "nt":
        return None
    try:
        import ctypes
        buffer_etiqueta = ctypes.create_unicode_buffer(261)
        objetivo = etiqueta_buscada.strip().upper()
        for letra in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            raiz = f"{letra}:\\"
            if not os.path.exists(raiz):
                continue
            ok = ctypes.windll.kernel32.GetVolumeInformationW(
                ctypes.c_wchar_p(raiz), buffer_etiqueta, ctypes.sizeof(buffer_etiqueta),
                None, None, None, None, 0,
            )
            if ok and buffer_etiqueta.value.strip().upper() == objetivo:
                return f"{letra}:/"
    except Exception:
        pass
    return None


# ============================= CONFIGURACION =============================

# Ruta al informe de Excel (.xlsx o .xlsm).
RUTA_EXCEL = r"C:\Users\User\Documents\RELACION_595_PROCESOS_ESSA_1S.xlsx"

# Nombre de la hoja/pestaña donde estan los procesos a cruzar.
HOJA_EXCEL = "RELACION PROCESOS"

# Fila donde estan los encabezados de columna (la fila que dice "No.",
# "RADICADO", etc). Los datos empiezan en la fila siguiente.
FILA_ENCABEZADO = 1

# Nombres exactos de las columnas a usar (tal como aparecen en el encabezado).
COLUMNA_NO = "No."
COLUMNA_RADICADO = "RADICADO"

# Columnas que solo se usan para el reporte de "cuantos procesos me hacen
# falta por agregar al disco" (ARCHIVO_REPORTE_FALTANTES), NO para el
# cruce/renombrado -- si alguna de estas columnas no existe, ese reporte
# en particular simplemente se omite (sin afectar el resto del script).
COLUMNA_ESTADO = "ESTADO"
COLUMNA_CUENTA = "CUENTA"
COLUMNA_JUZGADO = "JUZGADO"

# Nombre del DEMANDADO -- se usa para no confundir dos procesos DISTINTOS
# que comparten el mismo radicado corto/cuenta (ej. "2024-00139 CONTRA
# RIONEGRO" no es lo mismo que "2024-00139 CONTRA BOLIVAR"). A
# diferencia de ESTADO/CUENTA/JUZGADO, si esta columna no existe el
# resto del reporte de faltantes NO se ve afectado -- solo se pierde
# esta validacion extra de demandado en buscar_faltantes_en_drive.py.
COLUMNA_DEMANDADO = "DEMANDADO"

# El reporte de "procesos faltantes" SOLO incluye procesos de estos
# estados procesales (los demas estados se ignoran por completo para
# ese reporte -- ni se cuentan ni aparecen en el CSV).
ESTADOS_A_CONTAR = ["ACTIVO", "ACTIVOS CON TITULOS", "SUSPENDIDO", "REORGANIZACION"]

# Nombre (etiqueta de volumen) de tu disco duro externo, tal como aparece en
# "Este equipo" y en Propiedades del disco. El script busca ese disco por
# su NOMBRE entre todas las unidades conectadas y usa la letra que
# encuentre en ese momento -- asi no importa si Windows le asigna D:, E:,
# o cualquier otra letra la proxima vez que lo conectes.
ETIQUETA_DISCO_EXTERNO = "OSCAL"

# Letra de respaldo, SOLO por si el disco no se encuentra por su nombre
# (ej. esta desconectado, o no estas en Windows). En circunstancias
# normales no se deberia llegar a usar esta; si el log dice que se esta
# usando, el disco no se encontro por nombre y hay que revisar por que.
CARPETA_PROCESOS_RESPALDO = r"E:/"

_disco_detectado = encontrar_disco_por_etiqueta(ETIQUETA_DISCO_EXTERNO)

# Carpeta del disco duro donde estan las carpetas de cada proceso (las que
# hoy tienen solo el radicado de 23 digitos como nombre). Si las carpetas
# estan dentro de otra carpeta ahi (no directo en la raiz del disco),
# agrega esa carpeta aqui, ej: CARPETA_PROCESOS += "Procesos".
CARPETA_PROCESOS = _disco_detectado or CARPETA_PROCESOS_RESPALDO

# Carpeta ADICIONAL (opcional) donde a veces caen procesos ya extraidos que
# todavia no se han pasado a la raiz del disco (ej. una entrega masiva de
# expedientes). Si existe, el script la revisa ademas de la raiz: los
# procesos que YA esten en la raiz se dejan donde estan (se informa), y
# los que tengan radicado valido en el Excel pero AUN NO esten en la raiz
# se MUEVEN a la raiz del disco para que el resto del script los termine
# de nombrar en esta misma corrida. Si la ruta no existe, este paso se
# omite sin problema -- no hace falta comentarlo ni nada.
CARPETA_ENTRADA_ADICIONAL = os.path.join(CARPETA_PROCESOS, "PROCESOS LAUE", "ENTREGA EXPEDIENTE ESSA")

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

# Un PDF mas pesado que esto (en MB) NO se abre para leer su contenido --
# se salta directo, como si no se hubiera podido leer. Un PDF con la
# tabla de referencias cruzadas (xref) dañada obliga a pypdf a escanear
# el archivo COMPLETO byte por byte para reconstruirla (se ve en el log
# como una fila tras otra de "Ignoring wrong pointing object"); en un
# escaneo pesado de cientos de MB eso puede tardar minutos por un solo
# archivo y dejar el script "pegado" sin ningun aviso de que sigue
# trabajando. Este limite evita ese caso -- no afecta los PDF normales
# (la inmensa mayoria pesa unos pocos MB), solo a los pocos casos
# extremos que se quedarian trabados.
MAX_MB_PDF_PARA_CONTENIDO = 20

# True: no renombra ni mueve nada, solo muestra/registra que haria
# (recomendado la primera vez). False: aplica los cambios de verdad.
MODO_PRUEBA = True

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "validar_renombrar_carpetas.log")
ARCHIVO_REPORTE_VACIAS = os.path.join(os.path.dirname(__file__), "carpetas_vacias.csv")
ARCHIVO_REPORTE_CONTENIDO = os.path.join(os.path.dirname(__file__), "contenido_no_corresponde.csv")
ARCHIVO_REPORTE_FALTANTES = os.path.join(os.path.dirname(__file__), "procesos_faltantes_en_disco.csv")

# Carpeta donde se mueven (nunca se borran) las copias duplicadas sobrantes.
NOMBRE_CARPETA_DUPLICADOS = "Duplicados_para_revisar"

# Carpetas que NO son de un proceso y hay que ignorar siempre al escanear
# CARPETA_PROCESOS: carpetas propias de Windows que existen en la raiz de
# cualquier disco, y la carpeta temporal que usa procesos_juridicos.py
# mientras extrae un zip (si CARPETA_DESTINO de ese script es la misma
# CARPETA_PROCESOS de aqui, como es lo normal).
CARPETAS_A_IGNORAR = {
    NOMBRE_CARPETA_DUPLICADOS,
    "System Volume Information",
    "$RECYCLE.BIN",
    "_tmp_extraccion",
}

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


def _escribir_csv_tolerante(ruta, encabezado, filas, etiqueta_log) -> bool:
    """
    Escribe un reporte CSV; si el archivo esta abierto en Excel u otro
    programa (PermissionError), no crashea el script -- avisa con
    claridad y sigue. Devuelve True si se pudo guardar.
    """
    try:
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            escritor = csv.writer(f, delimiter=";")
            escritor.writerow(encabezado)
            for fila in filas:
                escritor.writerow(fila)
        return True
    except PermissionError:
        logging.error(
            "[%s] No se pudo guardar %s -- probablemente el archivo esta abierto en Excel u otro programa. "
            "Cierralo y vuelve a correr el script para actualizarlo.",
            etiqueta_log, ruta,
        )
        return False


def encontrar_columna(encabezados, nombre_buscado):
    for idx, valor in enumerate(encabezados, start=1):
        if valor and str(valor).strip().lower() == nombre_buscado.strip().lower():
            return idx
    raise ValueError(
        f"No se encontro la columna '{nombre_buscado}' en la fila {FILA_ENCABEZADO} de '{HOJA_EXCEL}'."
    )


def leer_datos_faltantes_por_radicado():
    """
    Lee el Excel UNA VEZ MAS, solo para el reporte de "procesos faltantes
    por agregar al disco" (ARCHIVO_REPORTE_FALTANTES) -- esto es aparte
    del cruce normal. Para cada radicado de 23 digitos, guarda su ESTADO
    PROCESAL, CUENTA y JUZGADO (columnas COLUMNA_ESTADO, COLUMNA_CUENTA,
    COLUMNA_JUZGADO). Devuelve {radicado: {"estado":.., "cuenta":..,
    "juzgado":.., "demandado":..}}. Si alguna de las columnas ESTADO,
    CUENTA o JUZGADO no existe en el Excel, devuelve un diccionario
    vacio (ese reporte simplemente se omite, sin error, y el resto del
    script sigue igual). La columna DEMANDADO es opcional aparte: si no
    existe, "demandado" queda vacio para todas las filas, pero el resto
    del reporte funciona igual.
    """
    wb = openpyxl.load_workbook(RUTA_EXCEL, data_only=True)
    if HOJA_EXCEL not in wb.sheetnames:
        return {}
    ws = wb[HOJA_EXCEL]
    encabezados = [ws.cell(row=FILA_ENCABEZADO, column=c).value for c in range(1, ws.max_column + 1)]
    try:
        col_rad = encontrar_columna(encabezados, COLUMNA_RADICADO)
        col_estado = encontrar_columna(encabezados, COLUMNA_ESTADO)
        col_cuenta = encontrar_columna(encabezados, COLUMNA_CUENTA)
        col_juzgado = encontrar_columna(encabezados, COLUMNA_JUZGADO)
    except ValueError:
        return {}
    try:
        col_demandado = encontrar_columna(encabezados, COLUMNA_DEMANDADO)
    except ValueError:
        col_demandado = None

    datos_por_radicado = {}
    for fila in range(FILA_ENCABEZADO + 1, ws.max_row + 1):
        radicado_crudo = ws.cell(row=fila, column=col_rad).value
        if radicado_crudo is None:
            continue
        radicado = re.sub(r"[\s\-]", "", str(radicado_crudo).strip())
        if not (radicado.isdigit() and len(radicado) == 23):
            continue
        estado_crudo = ws.cell(row=fila, column=col_estado).value
        cuenta_crudo = ws.cell(row=fila, column=col_cuenta).value
        juzgado_crudo = ws.cell(row=fila, column=col_juzgado).value
        demandado_crudo = ws.cell(row=fila, column=col_demandado).value if col_demandado else None
        datos_por_radicado[radicado] = {
            "estado": str(estado_crudo).strip() if estado_crudo is not None else "",
            "cuenta": str(cuenta_crudo).strip() if cuenta_crudo is not None else "",
            "juzgado": str(juzgado_crudo).strip() if juzgado_crudo is not None else "",
            "demandado": str(demandado_crudo).strip() if demandado_crudo is not None else "",
        }
    return datos_por_radicado


def leer_demandados_por_radicado():
    """
    Lee el Excel para obtener el DEMANDADO de CADA proceso con radicado
    de 23 digitos (sin filtrar por estado -- a diferencia de
    leer_datos_faltantes_por_radicado, que solo cubre los que faltan en
    el disco). Devuelve {radicado: demandado}. Si la columna DEMANDADO
    no existe, devuelve un diccionario vacio.
    """
    wb = openpyxl.load_workbook(RUTA_EXCEL, data_only=True)
    if HOJA_EXCEL not in wb.sheetnames:
        return {}
    ws = wb[HOJA_EXCEL]
    encabezados = [ws.cell(row=FILA_ENCABEZADO, column=c).value for c in range(1, ws.max_column + 1)]
    try:
        col_rad = encontrar_columna(encabezados, COLUMNA_RADICADO)
        col_demandado = encontrar_columna(encabezados, COLUMNA_DEMANDADO)
    except ValueError:
        return {}

    demandados_por_radicado = {}
    for fila in range(FILA_ENCABEZADO + 1, ws.max_row + 1):
        radicado_crudo = ws.cell(row=fila, column=col_rad).value
        if radicado_crudo is None:
            continue
        radicado = re.sub(r"[\s\-]", "", str(radicado_crudo).strip())
        if not (radicado.isdigit() and len(radicado) == 23):
            continue
        demandado_crudo = ws.cell(row=fila, column=col_demandado).value
        demandado = str(demandado_crudo).strip() if demandado_crudo is not None else ""
        if demandado:
            demandados_por_radicado[radicado] = demandado
    return demandados_por_radicado


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


def quitar_repetidos(filas, silencioso=False, permitir_duplicados_radicado=False):
    """
    Excluye del cruce cualquier NUMERO de proceso que aparezca en mas de una
    fila (eso siempre es ambiguo: no hay forma segura de saber a cual
    carpeta corresponde cada fila).

    Para RADICADOS repetidos (mismo radicado en dos o mas filas, con
    numeros de proceso distintos):
      - permitir_duplicados_radicado=False (comportamiento de siempre):
        esas filas tambien se excluyen del cruce.
      - permitir_duplicados_radicado=True: esas filas NO se excluyen; en
        vez de eso se devuelven aparte, en un diccionario
        {radicado: [(fila, numero), ...]}, para que el llamador las trate
        como carpetas a DUPLICAR (una copia por cada numero de proceso)
        en vez de como un conflicto.

    Devuelve la lista de filas sin repetidos si permitir_duplicados_radicado
    es False, o (filas_sin_repetidos, radicados_duplicados) si es True.
    """
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
        if not permitir_duplicados_radicado:
            for radicado in radicados_repetidos:
                logging.warning(
                    "[Excel] Radicado %s aparece en varias filas (%s); esas filas se omiten del cruce.",
                    radicado, filas_por_radicado[radicado],
                )

    filas_limpias = [
        (fila, numero, radicado)
        for fila, numero, radicado in filas
        if numero not in numeros_repetidos
        and (permitir_duplicados_radicado or radicado not in radicados_repetidos)
    ]

    if not permitir_duplicados_radicado:
        return filas_limpias

    radicados_duplicados = {}
    filas_normales = []
    for fila, numero, radicado in filas_limpias:
        if radicado in radicados_repetidos:
            radicados_duplicados.setdefault(radicado, []).append((fila, numero))
        else:
            filas_normales.append((fila, numero, radicado))

    if not silencioso:
        for radicado, entradas in radicados_duplicados.items():
            logging.warning(
                "[Excel] Radicado %s aparece en %d filas con numeros de proceso distintos (%s); se "
                "duplicara la carpeta para que cada numero tenga su propia copia.",
                radicado, len(entradas), [numero for _fila, numero in entradas],
            )

    return filas_normales, radicados_duplicados


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


# Archivos "basura" que Windows crea solo (no cuentan como contenido real
# al decidir si una carpeta esta vacia).
ARCHIVOS_A_IGNORAR_AL_CONTAR = {"desktop.ini", "thumbs.db", ".ds_store"}


def _ruta_larga_segura(ruta: str) -> str:
    """En Windows, antepone el prefijo especial para evitar el limite clasico de 260 caracteres por ruta."""
    if os.name == "nt":
        ruta_abs = os.path.abspath(ruta)
        if not ruta_abs.startswith("\\\\?\\"):
            return "\\\\?\\" + ruta_abs
    return ruta


def _sanear_nombre(nombre: str) -> str:
    """Quita caracteres invalidos para nombres de carpeta/archivo en Windows (incluidos espacios/puntos al final)."""
    nombre = re.sub(r'[<>:"/\\|?*]', "_", nombre).strip(" .")
    return nombre or "SinNombre"


def _ruta_zip_saneada(destino_normalizado: str, nombre_miembro: str) -> str:
    """
    Construye la ruta de destino para un miembro de un zip, saneando
    CADA segmento de la ruta (no solo el nombre final). Un nombre de
    carpeta/archivo con espacios o puntos al final es valido DENTRO de
    un zip, pero Windows lo maneja mal incluso con el prefijo de ruta
    larga (\\\\?\\) al leerlo despues -- mejor nunca dejar que llegue a
    crearse asi en el disco.
    """
    segmentos = [s for s in nombre_miembro.replace("\\", "/").split("/") if s not in ("", ".", "..")]
    if not segmentos:
        return destino_normalizado
    return os.path.join(destino_normalizado, *(_sanear_nombre(s) for s in segmentos))


def _copy2_ruta_segura(origen, destino, *, follow_symlinks=True):
    """Como shutil.copy2, pero pasando cada ruta por _ruta_larga_segura -- se usa como copy_function de shutil.copytree."""
    shutil.copy2(_ruta_larga_segura(str(origen)), _ruta_larga_segura(str(destino)), follow_symlinks=follow_symlinks)


def contar_archivos(carpeta: Path) -> int:
    """Cuenta cuantos archivos de VERDAD (no basura de Windows, no carpetas) hay dentro de una carpeta, recursivamente."""
    try:
        return sum(
            1 for p in carpeta.rglob("*")
            if p.is_file() and p.name.lower() not in ARCHIVOS_A_IGNORAR_AL_CONTAR
        )
    except OSError:
        return 0


def buscar_donante_en_duplicados(carpeta_duplicados: Path, radicado: str):
    """
    Busca dentro de NOMBRE_CARPETA_DUPLICADOS (un solo nivel) una carpeta
    cuyo radicado EXACTO coincida y que tenga al menos un archivo adentro
    -- para poder rellenar una carpeta VACIA de la raiz que sea del mismo
    caso. Si hay mas de una candidata, devuelve la que tenga MAS archivos.
    Devuelve (carpeta, cantidad_de_archivos), o None si no hay ninguna con
    contenido.
    """
    if not carpeta_duplicados.exists():
        return None
    candidatas = []
    try:
        for hijo in carpeta_duplicados.iterdir():
            if hijo.is_dir() and radicado_de_nombre_carpeta(hijo.name) == radicado:
                archivos = contar_archivos(hijo)
                if archivos > 0:
                    candidatas.append((hijo, archivos))
    except OSError:
        return None
    if not candidatas:
        return None
    candidatas.sort(key=lambda par: par[1], reverse=True)
    return candidatas[0]


def _texto_de_pdf(ruta: Path) -> str:
    if PdfReader is None:
        return ""
    try:
        if ruta.stat().st_size > MAX_MB_PDF_PARA_CONTENIDO * 1024 * 1024:
            logging.info(
                "   (se omite el contenido de '%s': pesa mas de %d MB -- probablemente un escaneo pesado con la "
                "tabla de referencias dañada, que se demoraria mucho en leer; se sigue sin abrirlo)",
                ruta.name, MAX_MB_PDF_PARA_CONTENIDO,
            )
            return ""
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


def buscar_carpetas_de_proceso(carpeta_base: Path):
    """
    Recorre RECURSIVAMENTE carpeta_base y devuelve [(carpeta, radicado), ...]
    para cada carpeta que tenga un radicado EXACTO de 23 digitos en su
    nombre. En cuanto encuentra una asi, NO sigue bajando dentro de ella
    (se asume que todo lo de adentro es contenido propio de ese caso, no
    otros procesos sueltos) -- solo sigue bajando por carpetas
    "organizadoras" sin radicado propio en el nombre (ej. carpetas por
    año o por tipo de tramite).
    """
    encontradas = []
    try:
        hijos = [h for h in carpeta_base.iterdir() if h.is_dir()]
    except OSError:
        return encontradas
    for hijo in hijos:
        radicado = radicado_de_nombre_carpeta(hijo.name)
        if radicado:
            encontradas.append((hijo, radicado))
        else:
            encontradas.extend(buscar_carpetas_de_proceso(hijo))
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


def aplanar_carpeta_anidada_unica(carpeta: Path, maximo_niveles: int = 5) -> None:
    """
    Si 'carpeta' termina con un UNICO elemento adentro y ese elemento es a
    su vez una carpeta (tipico cuando el zip original comprimia una sola
    carpeta -- por ejemplo al exportar una carpeta de Google Drive -- con
    un nombre o numero de proceso VIEJO o distinto al de la carpeta real
    de destino), sube todo lo de esa subcarpeta un nivel y borra la
    subcarpeta ya vacia. Asi los documentos quedan directo dentro de
    'carpeta', en vez de metidos en una subcarpeta con un numero
    equivocado. Repite varias veces por si hay mas de un nivel asi anidado.
    """
    for _ in range(maximo_niveles):
        try:
            hijos = [
                h for h in carpeta.iterdir()
                if not (h.is_file() and h.name.lower() in ARCHIVOS_A_IGNORAR_AL_CONTAR)
            ]
        except OSError:
            return
        if len(hijos) != 1 or not hijos[0].is_dir():
            return
        subcarpeta = hijos[0]
        for elemento in list(subcarpeta.iterdir()):
            destino = carpeta / elemento.name
            if not destino.exists():
                try:
                    shutil.move(_ruta_larga_segura(str(elemento)), _ruta_larga_segura(str(destino)))
                except OSError as error:
                    logging.warning(
                        "   (no se pudo subir '%s' un nivel al aplanar '%s' -- probablemente la ruta es "
                        "demasiado larga para Windows, o hay un problema de permisos/antivirus; se omite y se "
                        "sigue con el resto: %s)",
                        elemento.name, carpeta, error,
                    )
        try:
            subcarpeta.rmdir()
        except OSError:
            return


def extraer_zip_en_carpeta(ruta_zip: Path, carpeta_destino: Path):
    """
    Extrae 'ruta_zip' directo dentro de 'carpeta_destino' (una carpeta
    VACIA que ya existe). Si el zip trae todo metido dentro de una sola
    carpeta de primer nivel (tipico de exportar una carpeta de Google
    Drive, con un nombre o numero de proceso VIEJO o distinto), esa
    carpeta se "aplana" para que los documentos queden directo dentro de
    carpeta_destino. Tolera archivos individuales corruptos dentro del
    zip -- si uno falla, lo salta y sigue con el resto, en vez de fallar
    por completo. NUNCA borra ni mueve el zip original: se puede volver a
    intentar mas veces si hace falta. Devuelve (archivos_extraidos, archivos_fallidos).
    """
    extraidos = 0
    fallidos = 0
    destino_normalizado = os.path.normpath(os.path.abspath(str(carpeta_destino)))
    with zipfile.ZipFile(ruta_zip) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            ruta_destino = os.path.normpath(_ruta_zip_saneada(destino_normalizado, info.filename))
            if not ruta_destino.startswith(destino_normalizado):
                fallidos += 1
                continue
            ruta_destino_segura = _ruta_larga_segura(ruta_destino)
            try:
                os.makedirs(os.path.dirname(ruta_destino_segura), exist_ok=True)
                with zf.open(info) as origen, open(ruta_destino_segura, "wb") as destino:
                    shutil.copyfileobj(origen, destino)
                extraidos += 1
            except Exception:
                fallidos += 1
    aplanar_carpeta_anidada_unica(carpeta_destino)
    return extraidos, fallidos


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
    if _disco_detectado:
        logging.info(
            "[Disco] Se encontro el disco '%s' conectado como %s; se usa esa ruta.",
            ETIQUETA_DISCO_EXTERNO, CARPETA_PROCESOS,
        )
    else:
        logging.warning(
            "[Disco] No se encontro ningun disco llamado '%s' conectado ahorita; se usa la ruta de "
            "respaldo %s, que puede estar desactualizada. Verifica que el disco externo este conectado "
            "y que su nombre sea exactamente '%s' (click derecho sobre el disco en 'Este equipo' -> "
            "Propiedades -> el campo de arriba con el nombre).",
            ETIQUETA_DISCO_EXTERNO, CARPETA_PROCESOS, ETIQUETA_DISCO_EXTERNO,
        )

    # --- Doble lectura independiente del Excel, para confirmar que el
    # resultado es estable antes de tocar ninguna carpeta ---
    validas_1, casi_validas_1 = leer_filas_excel(silencioso=False)
    validas_2, casi_validas_2 = leer_filas_excel(silencioso=True)

    procesos_1, radicados_duplicar_1 = quitar_repetidos(validas_1, silencioso=False, permitir_duplicados_radicado=True)
    procesos_2, radicados_duplicar_2 = quitar_repetidos(validas_2, silencioso=True, permitir_duplicados_radicado=True)
    procesos_casi_validos = quitar_repetidos(casi_validas_1, silencioso=True)

    dict_1 = {radicado: numero for _f, numero, radicado in procesos_1}
    dict_2 = {radicado: numero for _f, numero, radicado in procesos_2}
    dup_1 = {r: sorted(numero for _fila, numero in entradas) for r, entradas in radicados_duplicar_1.items()}
    dup_2 = {r: sorted(numero for _fila, numero in entradas) for r, entradas in radicados_duplicar_2.items()}
    if dict_1 != dict_2 or dup_1 != dup_2:
        logging.error(
            "[Seguridad] La primera y la segunda lectura del Excel NO coinciden (¿se esta editando el "
            "archivo justo ahora?). Por seguridad no se toca ninguna carpeta. Cierra el Excel si lo "
            "tienes abierto editando, y vuelve a correr el script."
        )
        return

    procesos = procesos_1
    radicados_para_duplicar = radicados_duplicar_1
    logging.info(
        "Excel: %d proceso(s) con radicado valido y sin repetir, confirmados en dos lecturas independientes.",
        len(procesos),
    )
    if radicados_para_duplicar:
        logging.info(
            "Excel: %d radicado(s) aparecen en mas de una fila con numeros de proceso distintos; se "
            "duplicara la carpeta correspondiente para cada numero (ver [Duplicado en Excel] mas abajo).",
            len(radicados_para_duplicar),
        )

    por_radicado = {radicado: (numero, fila) for fila, numero, radicado in procesos}

    carpeta_raiz = Path(CARPETA_PROCESOS)
    carpeta_duplicados = carpeta_raiz / NOMBRE_CARPETA_DUPLICADOS

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
    duplicados_por_excel = []  # (radicado, [numeros], [nombres_de_carpeta_resultantes])
    movidos_a_auditar = []  # (nombre_original, ruta_destino) para la revision final

    # --- Carpeta de entrada adicional (ej. "PROCESOS LAUE/ENTREGA
    # EXPEDIENTE ESSA"): procesos ya extraidos que todavia no se han
    # pasado a la raiz del disco. Los que YA tengan carpeta en la raiz se
    # dejan donde estan (se informa al final); los que tengan radicado
    # valido en el Excel (directo, o duplicado en el Excel con varios
    # numeros) y AUN NO esten en la raiz se MUEVEN a la raiz para que el
    # resto de esta misma corrida los termine de nombrar; los que no
    # tengan proceso en el Excel se dejan donde estan y se informan
    # aparte. Si la ruta no existe, este paso se omite sin problema.
    laue_movidas = []               # (nombre_original, radicado)
    laue_ya_en_disco = []           # (nombre_original, radicado)
    laue_sin_proceso_en_excel = []  # (nombre_original, radicado)
    carpeta_entrada_adicional = Path(CARPETA_ENTRADA_ADICIONAL)
    if carpeta_entrada_adicional.exists():
        radicados_ya_en_raiz = set()
        for d in carpeta_raiz.iterdir():
            if d.is_dir() and d.name not in CARPETAS_A_IGNORAR:
                r = radicado_de_nombre_carpeta(d.name)
                if r:
                    radicados_ya_en_raiz.add(r)

        for carpeta_laue, radicado_laue in buscar_carpetas_de_proceso(carpeta_entrada_adicional):
            if radicado_laue in radicados_ya_en_raiz:
                laue_ya_en_disco.append((carpeta_laue.name, radicado_laue))
                continue

            if radicado_laue not in por_radicado and radicado_laue not in radicados_para_duplicar:
                laue_sin_proceso_en_excel.append((carpeta_laue.name, radicado_laue))
                continue

            destino_laue = ruta_libre(carpeta_raiz, carpeta_laue.name)
            if MODO_PRUEBA:
                logging.info(
                    "[SIMULACION-LAUE] '%s' (dentro de %s) se moveria a la raiz del disco para "
                    "terminar de nombrarse en esta misma corrida.",
                    carpeta_laue.name, CARPETA_ENTRADA_ADICIONAL,
                )
            else:
                carpeta_laue.rename(destino_laue)
                logging.info(
                    "[LAUE] '%s' (dentro de %s) se movio a la raiz del disco como '%s' para terminar "
                    "de nombrarse en esta misma corrida.",
                    carpeta_laue.name, CARPETA_ENTRADA_ADICIONAL, destino_laue.name,
                )
                # No se agrega a movidos_a_auditar: esta carpeta todavia le
                # falta el paso de renombrado normal (le agrega "numero. "
                # al nombre) mas abajo en esta misma corrida, asi que este
                # nombre intermedio (solo el radicado) no es el final --
                # ese renombrado se audita solo mas abajo, via 'reporte'.
            laue_movidas.append((carpeta_laue.name, radicado_laue))
            radicados_ya_en_raiz.add(radicado_laue)

    carpetas = [d for d in carpeta_raiz.iterdir() if d.is_dir() and d.name not in CARPETAS_A_IGNORAR]

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

    candidatos_cercanos = procesos_casi_validos + [
        (fila, numero, radicado) for fila, numero, radicado in procesos
    ]

    def renombrar_o_consolidar_con_existente(carpeta, numero, radicado_final, radicado_original, reporte):
        """
        Como intentar_renombrar_carpeta, pero si el nombre destino ya
        existe como una carpeta DISTINTA -- lo tipico cuando una
        correccion de "[Consecutivo]" apunta a un nombre que ya tiene su
        propia carpeta (las dos instancias del mismo caso, cada una en su
        carpeta aparte) -- no se deja como conflicto sin resolver: se
        trata igual que un duplicado normal. Se conserva bajo el nombre
        correcto la que tenga MAS ARCHIVOS adentro, y la otra se mueve
        (nunca se borra) a Duplicados_para_revisar.
        """
        nuevo_nombre = f"{numero}. {radicado_final}"
        destino = carpeta.parent / nuevo_nombre

        if not (destino.exists() and destino.resolve() != carpeta.resolve()):
            return intentar_renombrar_carpeta(carpeta, numero, radicado_final, radicado_original, reporte)

        archivos_carpeta = contar_archivos(carpeta)
        archivos_destino = contar_archivos(destino)

        if archivos_carpeta > archivos_destino:
            # La carpeta con el radicado SIN corregir es la mas completa:
            # se mueve la que ya tenia el nombre correcto (menos completa)
            # a Duplicados, y se renombra esta a la que corresponde.
            carpeta_sobra, hay_que_renombrar = destino, True
        else:
            # La que ya tenia el nombre correcto es la mas completa (o
            # empatan): se mueve la del radicado sin corregir, y no hay
            # nada mas que renombrar.
            carpeta_sobra, hay_que_renombrar = carpeta, False

        destino_dup = ruta_libre(carpeta_duplicados, carpeta_sobra.name)
        if MODO_PRUEBA:
            logging.info(
                "[SIMULACION-Consecutivo duplicado] '%s' y '%s' son el mismo caso con distinto consecutivo "
                "(una ya tenia el nombre correcto, la otra suelta); se conservaria la mas completa como "
                "'%s', y se moveria '%s' a '%s/%s'.",
                carpeta.name, destino.name, nuevo_nombre, carpeta_sobra.name,
                NOMBRE_CARPETA_DUPLICADOS, destino_dup.name,
            )
        else:
            carpeta_duplicados.mkdir(parents=True, exist_ok=True)
            carpeta_sobra.rename(destino_dup)
            logging.info(
                "[Consecutivo duplicado] '%s' y '%s' eran el mismo caso con distinto consecutivo (una ya "
                "tenia el nombre correcto, la otra suelta); se conservo la mas completa como '%s', y se "
                "movio '%s' a '%s/%s'.",
                carpeta.name, destino.name, nuevo_nombre, carpeta_sobra.name,
                NOMBRE_CARPETA_DUPLICADOS, destino_dup.name,
            )
            movidos_a_auditar.append((carpeta_sobra.name, destino_dup))

        duplicados_resueltos.append((radicado_final, nuevo_nombre, [(carpeta_sobra.name, destino_dup.name)]))

        if hay_que_renombrar:
            return intentar_renombrar_carpeta(carpeta, numero, radicado_final, radicado_original, reporte)
        reporte["ya_correctas"] += 1
        return True

    for radicado_en_carpeta, lista_carpetas in grupos_por_radicado.items():
        # --- Radicado que aparece en el Excel en mas de una fila, con
        # numeros de proceso DISTINTOS: se duplica la carpeta en vez de
        # tratarlo como conflicto (ver quitar_repetidos). ---
        entradas_duplicar = radicados_para_duplicar.get(radicado_en_carpeta)
        radicado_duplicar_final = radicado_en_carpeta

        if not entradas_duplicar:
            # El radicado EXACTO de la carpeta no esta duplicado en el
            # Excel -- pero puede que SI lo este bajo el mismo caso con
            # el ultimo digito distinto (el "consecutivo" de
            # instancia/reparto, ver mismo_radicado_salvo_ultimo_digito).
            # Sin esto, una carpeta asi nunca encontraba su duplicado en
            # el Excel: ni por coincidencia exacta (el radicado no es
            # igual), ni por la tolerancia normal de consecutivo (esta
            # busca solo entre los radicados SIN duplicar, y este SI
            # esta duplicado, por eso quedaba excluido de ahi).
            for radicado_excel_candidato, entradas_candidatas in radicados_para_duplicar.items():
                if mismo_radicado_salvo_ultimo_digito(radicado_en_carpeta, radicado_excel_candidato):
                    entradas_duplicar = entradas_candidatas
                    radicado_duplicar_final = radicado_excel_candidato
                    break

        if entradas_duplicar:
            numeros_destino = sorted(numero for _fila, numero in entradas_duplicar)
            radicados_encontrados_en_disco.add(radicado_duplicar_final)

            if radicado_duplicar_final != radicado_en_carpeta:
                logging.info(
                    "[Consecutivo] Carpeta '%s': el radicado termina distinto al del informe, que ademas "
                    "tiene este mismo caso duplicado con los procesos %s; se corrige a '%s'.",
                    radicado_en_carpeta, numeros_destino, radicado_duplicar_final,
                )

            # Si ademas hay mas de una carpeta en el disco con este mismo
            # radicado, primero se consolida (se conserva la mas completa)
            # para tener una sola carpeta de origen a partir de la cual duplicar.
            conteos = [(carpeta, contar_archivos(carpeta)) for carpeta in lista_carpetas]
            conteos.sort(key=lambda par: par[1], reverse=True)
            carpeta_origen, archivos_origen = conteos[0]
            otras = conteos[1:]

            primer_numero = numeros_destino[0]
            nombre_origen_final = f"{primer_numero}. {radicado_duplicar_final}"

            movidas = []
            for carpeta_extra, archivos_extra in otras:
                destino_dup = ruta_libre(carpeta_duplicados, carpeta_extra.name)
                if MODO_PRUEBA:
                    logging.info(
                        "[SIMULACION-Duplicado] '%s' (%d archivo(s)) se moveria a '%s/%s' -- se conserva "
                        "'%s' (%d archivo(s)) como origen para duplicar segun el Excel.",
                        carpeta_extra.name, archivos_extra, NOMBRE_CARPETA_DUPLICADOS, destino_dup.name,
                        carpeta_origen.name, archivos_origen,
                    )
                else:
                    carpeta_duplicados.mkdir(parents=True, exist_ok=True)
                    carpeta_extra.rename(destino_dup)
                    logging.info(
                        "[Duplicado] '%s' (%d archivo(s)) se movio a '%s/%s' -- se conserva '%s' (%d "
                        "archivo(s)) como origen para duplicar segun el Excel.",
                        carpeta_extra.name, archivos_extra, NOMBRE_CARPETA_DUPLICADOS, destino_dup.name,
                        carpeta_origen.name, archivos_origen,
                    )
                    movidos_a_auditar.append((carpeta_extra.name, destino_dup))
                movidas.append((carpeta_extra.name, destino_dup.name))
            if movidas:
                duplicados_resueltos.append((radicado_duplicar_final, nombre_origen_final, movidas))

            renombro_ok = intentar_renombrar_carpeta(
                carpeta_origen, primer_numero, radicado_duplicar_final, radicado_en_carpeta, reporte
            )
            if not renombro_ok:
                # El conflicto ya quedo registrado por intentar_renombrar_carpeta;
                # sin la carpeta origen en su lugar no es seguro duplicarla.
                continue

            ruta_origen_para_copiar = carpeta_origen if MODO_PRUEBA else (carpeta_origen.parent / nombre_origen_final)
            nombres_resultantes = [nombre_origen_final]

            for numero_extra in numeros_destino[1:]:
                nombre_copia = f"{numero_extra}. {radicado_duplicar_final}"
                destino_copia = ruta_libre(carpeta_raiz, nombre_copia)
                if MODO_PRUEBA:
                    logging.info(
                        "[SIMULACION-Duplicado en Excel] Se crearia una copia de '%s' como '%s' (el radicado "
                        "%s aparece %d veces en el Excel, con los procesos %s).",
                        ruta_origen_para_copiar.name, destino_copia.name, radicado_duplicar_final,
                        len(numeros_destino), numeros_destino,
                    )
                else:
                    try:
                        shutil.copytree(ruta_origen_para_copiar, destino_copia, copy_function=_copy2_ruta_segura)
                    except (shutil.Error, OSError) as error:
                        logging.error(
                            "[Duplicado en Excel] No se pudo copiar '%s' como '%s' -- alguno de sus archivos "
                            "probablemente tiene una ruta demasiado larga para Windows, o hay un problema de "
                            "permisos/antivirus. Se omite esta copia (revisa y duplica esta a mano si hace "
                            "falta): %s",
                            ruta_origen_para_copiar.name, destino_copia.name, error,
                        )
                        try:
                            shutil.rmtree(_ruta_larga_segura(str(destino_copia)))
                        except OSError as error_borrar:
                            logging.warning(
                                "   (la copia parcial '%s' quedo en el disco, no se pudo borrar: %s)",
                                destino_copia, error_borrar,
                            )
                        continue
                    logging.info(
                        "[Duplicado en Excel] Se creo una copia de '%s' como '%s' (radicado %s duplicado en "
                        "el Excel con los procesos %s).",
                        ruta_origen_para_copiar.name, destino_copia.name, radicado_duplicar_final, numeros_destino,
                    )
                    movidos_a_auditar.append((f"copia de '{ruta_origen_para_copiar.name}'", destino_copia))
                nombres_resultantes.append(destino_copia.name)

            duplicados_por_excel.append((radicado_duplicar_final, numeros_destino, nombres_resultantes))
            continue

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
            renombrar_o_consolidar_con_existente(carpeta, numero, radicado_final, radicado_en_carpeta, reporte)
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

        renombrar_o_consolidar_con_existente(carpeta_conservar, numero, radicado_final, radicado_en_carpeta, reporte)
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
        d for d in carpeta_raiz.iterdir() if d.is_dir() and d.name not in CARPETAS_A_IGNORAR
    ]
    for carpeta_padre in carpetas_nivel_superior_ahora:
        radicado_padre = radicado_de_nombre_carpeta(carpeta_padre.name)
        if not radicado_padre:
            continue

        for hijo, radicado_hijo in subcarpetas_con_radicado(carpeta_padre):
            # Mismo caso tambien si solo difieren en el ultimo digito
            # (consecutivo de instancia/reparto, ej. termina en 0 o en 1) --
            # eso NO es un proceso distinto, ver mismo_radicado_salvo_ultimo_digito.
            es_exacto = radicado_hijo == radicado_padre
            es_consecutivo = not es_exacto and mismo_radicado_salvo_ultimo_digito(radicado_hijo, radicado_padre)
            if es_exacto or es_consecutivo:
                motivo = "mismo radicado" if es_exacto else "mismo radicado salvo el ultimo digito (consecutivo)"
                destino = ruta_libre(carpeta_duplicados, hijo.name)
                if MODO_PRUEBA:
                    logging.info(
                        "[SIMULACION-Anidada] '%s' esta metida dentro de '%s' y es una copia del MISMO caso "
                        "(%s); se moveria a '%s/%s'.",
                        hijo.name, carpeta_padre.name, motivo, NOMBRE_CARPETA_DUPLICADOS, destino.name,
                    )
                else:
                    carpeta_duplicados.mkdir(parents=True, exist_ok=True)
                    hijo.rename(destino)
                    logging.info(
                        "[Anidada] '%s' estaba metida dentro de '%s' (%s); se movio a '%s/%s'.",
                        hijo.name, carpeta_padre.name, motivo, NOMBRE_CARPETA_DUPLICADOS, destino.name,
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
    carpetas_rellenadas_desde_duplicados = []  # (nombre, nombre_donante, archivos_copiados)
    carpetas_rellenadas_desde_zip = []  # (nombre, zip, ubicacion, archivos_extraidos, archivos_fallidos)
    carpetas_aplanadas_historicas = []  # (nombre, nombre_subcarpeta_vieja)

    carpetas_finales = [d for d in carpeta_raiz.iterdir() if d.is_dir() and d.name not in CARPETAS_A_IGNORAR]
    for carpeta in carpetas_finales:
        radicado_exacto = radicado_de_nombre_carpeta(carpeta.name)
        radicado_actual = radicado_exacto or radicado_cercano_de_nombre_carpeta(carpeta.name)
        if not radicado_actual:
            carpetas_sin_radicado.append(carpeta.name)

        numero_archivos = contar_archivos(carpeta)
        if numero_archivos == 0:
            # Antes de darla por vacia, revisa si hay una carpeta con el
            # MISMO radicado en Duplicados_para_revisar que si tenga
            # documentos (de una consolidacion anterior) -- si la hay, se
            # le copian los archivos (nunca se borra el donante, por si
            # algo sale mal).
            donante_info = None
            if radicado_actual:
                donante_info = buscar_donante_en_duplicados(carpeta_duplicados, radicado_actual)

            if donante_info:
                donante, archivos_donante = donante_info
                if MODO_PRUEBA:
                    logging.info(
                        "[SIMULACION-Vacia] '%s' esta vacia; se rellenaria con los %d archivo(s) de la "
                        "carpeta duplicada '%s/%s' (mismo radicado %s).",
                        carpeta.name, archivos_donante, NOMBRE_CARPETA_DUPLICADOS, donante.name, radicado_actual,
                    )
                else:
                    try:
                        shutil.copytree(donante, carpeta, dirs_exist_ok=True, copy_function=_copy2_ruta_segura)
                    except (shutil.Error, OSError) as error:
                        logging.error(
                            "[Vacia] No se pudo rellenar '%s' desde '%s/%s' -- alguno de sus archivos "
                            "probablemente tiene una ruta demasiado larga para Windows, o hay un problema de "
                            "permisos/antivirus. Puede haber quedado parcialmente copiada; revisala a mano: %s",
                            carpeta.name, NOMBRE_CARPETA_DUPLICADOS, donante.name, error,
                        )
                    aplanar_carpeta_anidada_unica(carpeta)
                    numero_archivos = contar_archivos(carpeta)
                    logging.info(
                        "[Vacia-Rellenada] '%s' estaba vacia; se le copiaron los %d archivo(s) de la "
                        "carpeta duplicada '%s/%s' (mismo radicado %s). Esa copia en %s NO se borro, por "
                        "seguridad -- borrala a mano cuando confirmes que todo quedo bien.",
                        carpeta.name, archivos_donante, NOMBRE_CARPETA_DUPLICADOS, donante.name,
                        radicado_actual, NOMBRE_CARPETA_DUPLICADOS,
                    )
                    carpetas_rellenadas_desde_duplicados.append((carpeta.name, donante.name, archivos_donante))

            if numero_archivos == 0:
                zip_encontrado, zip_ubicacion = None, None
                if radicado_actual and carpeta_descargas:
                    resultado = buscar_zip_con_radicado(carpeta_descargas, radicado_actual)
                    if resultado:
                        zip_encontrado, zip_ubicacion = resultado

                if zip_encontrado:
                    subcarpeta = "" if zip_ubicacion == "Descargas" else "Procesados"
                    ruta_zip = carpeta_descargas / subcarpeta / zip_encontrado
                    if MODO_PRUEBA:
                        logging.info(
                            "[SIMULACION-Vacia] '%s' esta vacia; se extraeria el zip pendiente '%s' (en "
                            "%s) directo adentro.",
                            carpeta.name, zip_encontrado, zip_ubicacion,
                        )
                    else:
                        try:
                            extraidos, fallidos = extraer_zip_en_carpeta(ruta_zip, carpeta)
                        except Exception as error:
                            extraidos, fallidos = 0, 0
                            logging.error(
                                "[Error] No se pudo extraer el zip pendiente '%s' (en %s) para '%s': %s",
                                zip_encontrado, zip_ubicacion, carpeta.name, error,
                            )
                        numero_archivos = contar_archivos(carpeta)
                        if numero_archivos > 0:
                            logging.info(
                                "[Vacia-Rellenada] '%s' estaba vacia; se extrajo el zip pendiente '%s' "
                                "(en %s): %d archivo(s) recuperados%s. El zip original NO se borro.",
                                carpeta.name, zip_encontrado, zip_ubicacion, extraidos,
                                f" ({fallidos} no se pudieron extraer)" if fallidos else "",
                            )
                            carpetas_rellenadas_desde_zip.append(
                                (carpeta.name, zip_encontrado, zip_ubicacion, extraidos, fallidos)
                            )

                if numero_archivos == 0:
                    carpetas_vacias.append((carpeta.name, radicado_actual, zip_encontrado, zip_ubicacion))
        else:
            # Detecta y corrige carpetas que ya tienen contenido pero que
            # quedaron con TODO metido dentro de una sola subcarpeta con
            # un numero/nombre viejo o distinto -- herencia de
            # extracciones o copias de hace tiempo, de antes de que esto
            # se corrigiera en el momento de extraer. Se revisa siempre,
            # sin importar VALIDAR_CONTENIDO_CONTRA_NOMBRE.
            if radicado_actual:
                hijos_de_primer_nivel = [
                    h for h in carpeta.iterdir()
                    if not (h.is_file() and h.name.lower() in ARCHIVOS_A_IGNORAR_AL_CONTAR)
                ]
                if len(hijos_de_primer_nivel) == 1 and hijos_de_primer_nivel[0].is_dir():
                    nombre_subcarpeta_vieja = hijos_de_primer_nivel[0].name
                    if MODO_PRUEBA:
                        logging.info(
                            "[SIMULACION-Anidado historico] '%s' tiene TODO su contenido metido dentro de "
                            "una sola subcarpeta '%s'; se aplanaria para que quede directo adentro.",
                            carpeta.name, nombre_subcarpeta_vieja,
                        )
                    else:
                        aplanar_carpeta_anidada_unica(carpeta)
                        logging.info(
                            "[Anidado historico] '%s' tenia TODO su contenido metido dentro de una sola "
                            "subcarpeta vieja '%s'; se aplano para que quede directo adentro.",
                            carpeta.name, nombre_subcarpeta_vieja,
                        )
                        carpetas_aplanadas_historicas.append((carpeta.name, nombre_subcarpeta_vieja))

            if VALIDAR_CONTENIDO_CONTRA_NOMBRE and radicado_exacto:
                radicados_hallados = radicados_encontrados_en_carpeta(carpeta, MAX_ARCHIVOS_CONTENIDO_A_REVISAR)
                if radicados_hallados and radicado_exacto not in radicados_hallados:
                    radicado_dominante, veces = radicados_hallados.most_common(1)[0]
                    contenido_no_corresponde.append((carpeta.name, radicado_exacto, radicado_dominante, veces))

    faltantes = []  # (fila, numero, radicado)
    for fila, numero, radicado in procesos:
        if radicado not in radicados_encontrados_en_disco:
            faltantes.append((fila, numero, radicado))
            logging.warning(
                "[Sin carpeta] Proceso %s (fila %s del Excel, radicado %s) no tiene carpeta correspondiente en %s.",
                numero, fila, radicado, CARPETA_PROCESOS,
            )
    sin_carpeta_en_disco = len(faltantes)

    # --- Reporte filtrado: SOLO procesos ACTIVO / ACTIVOS CON TITULOS /
    # SUSPENDIDO / REORGANIZACION (ESTADOS_A_CONTAR) que faltan en el
    # disco -- ningun otro estado se cuenta ni aparece en el CSV. Se lee
    # el Excel aparte para esto -- si alguna columna no existe, este
    # reporte se omite sin afectar el resto del script. ---
    datos_por_radicado = leer_datos_faltantes_por_radicado()
    if datos_por_radicado and faltantes:
        conteo_por_estado = {estado: 0 for estado in ESTADOS_A_CONTAR}
        filas_reporte = []  # (numero, cuenta, radicado, juzgado, demandado)
        for fila, numero, radicado in faltantes:
            datos = datos_por_radicado.get(radicado)
            if not datos or datos["estado"] not in conteo_por_estado:
                continue
            conteo_por_estado[datos["estado"]] += 1
            filas_reporte.append((numero, datos["cuenta"], radicado, datos["juzgado"], datos["demandado"]))

        total = sum(conteo_por_estado.values())
        logging.warning(
            "[Faltan por agregar] %d proceso(s) de los estados %s no tienen carpeta en el disco:",
            total, ", ".join(ESTADOS_A_CONTAR),
        )
        for estado in ESTADOS_A_CONTAR:
            logging.warning("   - %s: %d", estado, conteo_por_estado[estado])

        filas_reporte.sort(key=lambda t: t[0])
        if _escribir_csv_tolerante(
            ARCHIVO_REPORTE_FALTANTES, ["No.", "Cuenta", "Radicado", "Juzgado", "Demandado"], filas_reporte,
            "Faltan por agregar",
        ):
            logging.info("[Faltan por agregar] Reporte guardado en: %s", ARCHIVO_REPORTE_FALTANTES)

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

    if carpetas_rellenadas_desde_duplicados:
        logging.info(
            "[Vacia-Rellenada] %d carpeta(s) estaban vacias y se rellenaron con los documentos de su "
            "copia duplicada (mismo radicado). Las copias originales NO se borraron, quedaron en '%s' -- "
            "borralas a mano cuando confirmes que todo quedo bien:",
            len(carpetas_rellenadas_desde_duplicados), NOMBRE_CARPETA_DUPLICADOS,
        )
        for nombre, nombre_donante, archivos in carpetas_rellenadas_desde_duplicados:
            logging.info(
                "   - '%s' se relleno con %d archivo(s) de '%s/%s'",
                nombre, archivos, NOMBRE_CARPETA_DUPLICADOS, nombre_donante,
            )

    if carpetas_rellenadas_desde_zip:
        logging.info(
            "[Vacia-Rellenada] %d carpeta(s) estaban vacias y se rellenaron extrayendo un .zip pendiente "
            "que ya estaba en Descargas. El zip original NO se borro:",
            len(carpetas_rellenadas_desde_zip),
        )
        for nombre, zip_nombre, ubicacion, extraidos, fallidos in carpetas_rellenadas_desde_zip:
            logging.info(
                "   - '%s' se relleno extrayendo '%s' (en %s): %d archivo(s)%s",
                nombre, zip_nombre, ubicacion, extraidos,
                f", {fallidos} fallidos" if fallidos else "",
            )

    if carpetas_aplanadas_historicas:
        logging.info(
            "[Anidado historico] %d carpeta(s) tenian TODO su contenido metido dentro de una sola "
            "subcarpeta vieja (herencia de extracciones/copias de antes); se aplanaron para que los "
            "documentos queden directo adentro:",
            len(carpetas_aplanadas_historicas),
        )
        for nombre, nombre_subcarpeta_vieja in carpetas_aplanadas_historicas:
            logging.info("   - '%s' (subcarpeta vieja quitada: '%s')", nombre, nombre_subcarpeta_vieja)

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

    filas_vacias = [
        (nombre, radicado_buscado or "", zip_encontrado or "", zip_ubicacion or "")
        for nombre, radicado_buscado, zip_encontrado, zip_ubicacion in carpetas_vacias
    ]
    if _escribir_csv_tolerante(
        ARCHIVO_REPORTE_VACIAS, ["Carpeta", "Radicado", "Zip pendiente encontrado", "Donde se encontro"],
        filas_vacias, "Carpeta vacia",
    ) and carpetas_vacias:
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

    if _escribir_csv_tolerante(
        ARCHIVO_REPORTE_CONTENIDO, ["Carpeta", "Radicado del nombre", "Radicado encontrado en el contenido", "Veces"],
        contenido_no_corresponde, "Contenido no corresponde",
    ) and contenido_no_corresponde:
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

    if duplicados_por_excel:
        logging.info(
            "[Duplicado en Excel] %d radicado(s) aparecian en el Excel dos o mas veces con numeros de "
            "proceso distintos; se duplico la carpeta para que cada numero tenga la suya:",
            len(duplicados_por_excel),
        )
        for radicado, numeros, nombres in duplicados_por_excel:
            logging.info(
                "   - Radicado %s (procesos %s): carpetas %s",
                radicado, numeros, nombres,
            )

    if laue_movidas or laue_ya_en_disco or laue_sin_proceso_en_excel:
        logging.info(
            "[LAUE] Revision de '%s': %d carpeta(s) se movieron a la raiz del disco, %d ya estaban en el "
            "disco (se dejaron donde estaban), %d no tienen proceso en el Excel (se dejaron donde estaban).",
            CARPETA_ENTRADA_ADICIONAL, len(laue_movidas), len(laue_ya_en_disco), len(laue_sin_proceso_en_excel),
        )
        for nombre, radicado in laue_movidas:
            logging.info("   - Movida a la raiz: '%s' (radicado %s)", nombre, radicado)
        for nombre, radicado in laue_ya_en_disco:
            logging.info("   - Ya estaba en el disco, se dejo en LAUE: '%s' (radicado %s)", nombre, radicado)
        for nombre, radicado in laue_sin_proceso_en_excel:
            logging.warning(
                "   - Sin proceso en el Excel, se dejo en LAUE: '%s' (radicado %s) -- revisa si falta "
                "agregarla al informe.",
                nombre, radicado,
            )

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
        "%d radicado(s) duplicado(s) en el Excel (carpeta duplicada para cada numero), "
        "%d carpeta(s) movidas desde LAUE, %d ya estaban en el disco (se dejaron en LAUE), "
        "%d en LAUE sin proceso en el Excel, "
        "%d carpeta(s) anidada(s) del mismo caso resueltas, %d carpeta(s) anidada(s) de otro caso sacadas, "
        "%d sin carpeta en disco, %d carpetas sin proceso en el Excel, %d conflictos de nombre, "
        "%d posibles coincidencias para revisar, %d grupo(s) duplicado(s) sin poder resolver, "
        "%d carpeta(s) vacia(s) (%d rellenadas desde Duplicados_para_revisar, %d rellenadas extrayendo "
        "un zip pendiente), "
        "%d carpeta(s) sin nombre reconocible, "
        "%d carpeta(s) con contenido que no corresponde al nombre.",
        len(renombradas),
        "carpetas simuladas (MODO_PRUEBA activo)" if MODO_PRUEBA else "carpetas renombradas",
        reporte["ya_correctas"], len(duplicados_resueltos), NOMBRE_CARPETA_DUPLICADOS,
        len(duplicados_por_excel),
        len(laue_movidas), len(laue_ya_en_disco), len(laue_sin_proceso_en_excel),
        len(anidadas_mismo_caso), len(anidadas_otro_caso),
        sin_carpeta_en_disco, len(sin_proceso_en_excel), reporte["conflictos"], len(posibles_coincidencias),
        len(duplicados_sin_resolver), len(carpetas_vacias), len(carpetas_rellenadas_desde_duplicados),
        len(carpetas_rellenadas_desde_zip),
        len(carpetas_sin_radicado),
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
