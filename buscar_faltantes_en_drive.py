"""
Busca en tu Google Drive (y opcionalmente en tu correo de Gmail) los
procesos que el reporte "procesos_faltantes_en_disco.csv" (lo genera
validar_renombrar_carpetas.py) dice que faltan en el disco duro, y
descarga la carpeta correspondiente si la encuentra.

IMPORTANTE: corre primero validar_renombrar_carpetas.py (para que
procesos_faltantes_en_disco.csv este al dia) antes de correr este script.

Para cada proceso faltante, busca por, en este orden:
  1. El RADICADO COMPLETO de 23 digitos. Si encuentra una coincidencia
     exacta, la descarga automatico -- este numero es tan especifico que
     no hay riesgo real de confundirlo con otro caso.
  2. El RADICADO CORTO (ej. "2025-00456" o "2025-456", derivado del año
     y el consecutivo del radicado completo).
  3. El numero de CUENTA (solo si es lo bastante especifica -- ver
     _cuenta_es_valida_para_buscar: una cuenta vacia, "0", o de muy
     pocos digitos NO se busca, porque un termino tan generico haria
     coincidir la busqueda aproximada de Drive con miles de carpetas
     sin relacion de TODO el Drive, dejando el proceso "pegado" un
     buen rato revisando falsos positivos uno por uno).
Las busquedas 2 y 3 son menos confiables (un radicado corto o una cuenta
se puede repetir o coincidir por casualidad con archivos de otro caso);
para esos candidatos, antes de descargar se valida ademas que la
carpeta (o sus archivos) de verdad mencionen ese radicado, para que una
cuenta compartida con procesos de OTRO cliente no traiga la carpeta
equivocada.

REGLA OBLIGATORIA, sin excepcion (aplica a TODO lo que se vaya a
descargar, sea confiable o no, radicado completo o corto, carpeta de
Drive o adjunto de correo): el documento tiene que mencionar a ESSA o
ELECTRIFICADORA DE SANTANDER (como demandante o como demandado -- por
nombre, o abriendo el contenido de sus PDF/DOCX si el nombre no lo
dice). Si no la menciona, NO se descarga, punto -- ni siquiera si el
radicado coincidio exacto. Los candidatos que si pasan (y que ademas
vinieron de una busqueda menos confiable) se descargan automatico
(cada uno en su propia carpeta, sin pisar nada, o fusionados si son
varios del mismo proceso), pero quedan marcados aparte en
faltantes_descargados_a_validar.csv, para que confirmes despues cual de
esas descargas es la correcta y borres a mano las que no correspondan
(el script nunca borra nada solo).

TAMBIEN se valida el DEMANDADO (columna DEMANDADO del Excel, si
existe): el mismo radicado corto o cuenta se puede repetir entre
procesos DISTINTOS que van contra demandados diferentes (ej. "2024-00139
CONTRA RIONEGRO" no es lo mismo que "2024-00139 CONTRA BOLIVAR"). Con
que el nombre del demandado esperado aparezca en CUALQUIER parte del
nombre/asunto/contenido del candidato alcanza para confirmarlo (no hace
falta que este despues de la palabra "CONTRA"). Solo se descarta un
candidato cuando SI trae un "CONTRA <algo>" (la forma habitual de
nombrar expedientes) pero ese "algo" es un demandado DISTINTO al
esperado -- aunque ya haya pasado la validacion de ESSA. Si el
candidato no menciona ni al demandado esperado ni ningun otro "CONTRA
<algo>", esta validacion no bloquea nada (no hay evidencia ni a favor
ni en contra). Ver _demandado_coincide_en_texto.

Si un proceso YA tiene una carpeta en el disco (por ejemplo porque una
corrida anterior ya lo descargo), se omite por completo sin buscar ni
descargar nada -- para no crear carpetas "_2" duplicadas si se vuelve a
correr el script sobre un procesos_faltantes_en_disco.csv desactualizado.

Al empezar, tambien revisa si quedaron carpetas TEMPORALES sueltas de
una corrida anterior que se cerro a la mitad, o que no se pudieron
borrar solas (ej. un archivo bloqueado por el antivirus/OneDrive justo
en ese momento) -- nunca se borran solas, se mueven a
Duplicados_para_revisar para que las revises. Ver
limpiar_carpetas_temporales_huerfanas().

Tambien revisa si ya quedaron carpetas "_2", "_3", etc en el
disco de corridas ANTERIORES a estos filtros (por ejemplo, de antes de
que existiera la validacion de demandante ESSA) -- revalida cada una:
si menciona a ESSA, la fusiona dentro de su carpeta principal; si no,
la mueve a Duplicados_para_revisar (nunca la borra) para que la
revises. Ver consolidar_duplicados_en_disco().

Tambien revisa TODAS las carpetas ya descargadas (de esta corrida o de
cualquier corrida anterior) buscando archivos que mencionen un
radicado corto DISTINTO al de su propia carpeta -- rastro de un bug ya
corregido donde un radicado corto como "2023-24" se confundia con
"2023-244" (de OTRO proceso) por compartir el mismo prefijo numerico.
Esos archivos se mueven a Duplicados_para_revisar (nunca se borran); si
una carpeta queda completamente vacia despues (todo era de otro
proceso), esa carpeta VACIA si se borra, para que se vuelva a buscar
en la proxima corrida. Ver revisar_contaminacion_en_disco().

De la misma forma, tambien revisa TODAS las carpetas ya descargadas
buscando archivos con un "CONTRA <algo>" que no corresponda al
DEMANDADO real del proceso (segun el Excel) -- mismo rastro, pero para
el caso de dos procesos que comparten radicado corto/cuenta y van
contra demandados DISTINTOS. Esos archivos tambien se mueven a
Duplicados_para_revisar (nunca se borran). Si es la CARPETA misma la
que parece tener el demandado equivocado en su nombre, solo se avisa
en el log (no se mueve ni renombra la carpeta sola) para que la
revises a mano. Ver revisar_demandado_en_disco().

Si lo que encuentra es un ARCHIVO suelto (no una carpeta) que coincide,
busca la carpeta que lo contiene. Si esa carpeta contenedora es
"propia" del caso (su nombre menciona el radicado, o la busqueda
encontro directamente la carpeta), descarga la carpeta completa,
asumiendo que ahi esta el resto del expediente. Pero si la carpeta
contenedora es GENERICA -- una carpeta de "informes" o "actuaciones"
que junta documentos de MUCHOS procesos distintos, y el archivo que
coincidio es solo uno mas ahi adentro -- NO descarga la carpeta
completa (traeria folios de otros procesos sin relacion), sino
UNICAMENTE los archivos de esa carpeta que de verdad mencionen este
radicado (y, si aplica, al demandante ESSA).

Al terminar de descargar/fusionar cada proceso, ordena sus documentos
CRONOLOGICAMENTE y les antepone un numero de orden: "1. ", "2. ", etc
(el mas viejo primero). La fecha se busca primero en el NOMBRE del
archivo (varios formatos, incluido "DD MES AAAA" sin la palabra "de",
el mas comun en nombres reales de autos); si no la trae, se abren sus
primeras paginas (PDF/DOCX); y si tampoco hay fecha ahi, como ultimo
recurso se usa la fecha de modificacion del archivo en el disco (solo
si es de mas de un dia atras) -- un documento que llega por adjunto de
correo y no trae fecha propia queda con la fecha del CORREO puesta ahi
(ver _organizar_adjunto_zip), asi que ese es el valor que se usa. Los
que no tengan NINGUNA fecha reconocible quedan al final. Ver
ordenar_y_enumerar_carpeta(). Ademas, al empezar cada corrida, TODAS
las carpetas de proceso que ya existan en el disco (de esta corrida o
de cualquier corrida anterior) tambien se revisan y se ordenan/enumeran
de la misma forma -- no hace falta que el proceso se haya tocado hoy.
Ver ordenar_todas_las_carpetas_en_disco().

Solo se descargan archivos PDF (o archivos nativos de Google -- Doc,
Sheet, Slide -- que Drive exporta como PDF). Cualquier otro tipo de
archivo (Word, Excel, imagenes, etc) que aparezca junto a los PDF en
Drive o dentro de un adjunto .zip de correo se omite y NUNCA se baja
al disco. Ver _es_pdf_o_exportable().

Requiere:
  - credenciales_drive.json: credenciales de OAuth de Google Drive (ver
    README para los pasos de como generarlas en Google Cloud Console).
    La primera vez que corras el script se abre el navegador para
    autorizar el acceso una sola vez; despues queda guardado en
    token_drive.json y no hay que repetirlo.
  - Para buscar tambien en el correo: el mismo credenciales_sgde.txt que
    ya usa procesos_juridicos.py (ver credenciales_sgde.example.txt). Si
    no existe, la busqueda en correo simplemente se omite.

Nunca borra ni sobreescribe nada que ya tengas en el disco: si el nombre
de destino ya existe, elige un nombre libre en vez de pisarlo. Respeta
MODO_PRUEBA (por defecto True): en modo prueba solo BUSCA y te dice que
encontraria/descargaria, sin bajar nada de verdad todavia.
"""

import csv
import datetime
import email
import imaplib
import io
import logging
import os
import re
import shutil
import unicodedata
import uuid
import zipfile
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaIoBaseDownload
except ImportError:
    Credentials = None
    HttpError = Exception

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

import procesos_juridicos as organizador
import validar_renombrar_carpetas as cruce_excel

# ============================= CONFIGURACION =============================

# Carpeta del disco duro donde va todo lo que se descargue (la misma que
# ya usan los otros scripts -- se detecta sola por el nombre del disco).
CARPETA_PROCESOS = cruce_excel.CARPETA_PROCESOS

# Reporte de procesos faltantes (lo genera validar_renombrar_carpetas.py).
ARCHIVO_REPORTE_FALTANTES = cruce_excel.ARCHIVO_REPORTE_FALTANTES

# Credenciales de Google Drive (ver README para como generarlas).
CREDENCIALES_DRIVE = os.path.join(os.path.dirname(__file__), "credenciales_drive.json")
TOKEN_DRIVE = os.path.join(os.path.dirname(__file__), "token_drive.json")
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

# True: tambien busca en tu correo de Gmail (usa las mismas credenciales
# que procesos_juridicos.py, en credenciales_sgde.txt). Si ese archivo no
# existe, la busqueda en correo se omite sola, sin error.
BUSCAR_EN_CORREO = True

ARCHIVO_LOG = os.path.join(os.path.dirname(__file__), "buscar_faltantes_en_drive.log")

# Procesos que se descargaron por una coincidencia MENOS segura
# (radicado corto, cuenta, o un enlace de correo que no traia el
# radicado completo) -- se descargan igual, pero quedan marcados aparte
# aqui para que los confirmes despues.
ARCHIVO_REPORTE_A_VALIDAR = os.path.join(os.path.dirname(__file__), "faltantes_descargados_a_validar.csv")

# True (por defecto): no descarga nada de verdad, solo busca y muestra
# que encontraria. False: descarga de verdad las coincidencias claras
# (radicado completo).
MODO_PRUEBA = True

MIME_CARPETA = "application/vnd.google-apps.folder"
MIME_EXPORTAR = {
    "application/vnd.google-apps.document": (".pdf", "application/pdf"),
    "application/vnd.google-apps.spreadsheet": (".pdf", "application/pdf"),
    "application/vnd.google-apps.presentation": (".pdf", "application/pdf"),
}

# Textos que deben aparecer (en el nombre, o en el contenido de los
# documentos) para confirmar que el demandante del proceso es ESSA --
# se usa para descartar carpetas de OTROS procesos que comparten cuenta
# o radicado corto por casualidad, pero son de un cliente distinto.
TERMINOS_DEMANDANTE_VALIDO = ["essa", "electrificadora de santander"]

# Palabras demasiado genericas del campo DEMANDADO del Excel como para
# usarlas solas para distinguir un proceso de otro (ej. "MUNICIPIO DE
# BARRANCABERMEJA" vs "MUNICIPIO DE RIONEGRO" -- "MUNICIPIO" y "DE" no
# sirven para diferenciarlos, la palabra que SI importa es la ultima).
# Se usan para quedarse solo con las palabras "significativas" del
# nombre del demandado al compararlo contra un "CONTRA <algo>" que
# aparezca en un nombre de carpeta/archivo/correo.
PALABRAS_GENERICAS_DEMANDADO = {
    "municipio", "departamento", "distrito", "alcaldia", "gobernacion",
    "empresa", "sociedad", "compania", "cooperativa", "institucion",
    "educativa", "colegio", "escuela", "fundacion", "corporacion",
    "de", "del", "la", "el", "los", "las", "y", "san", "santa",
    "sa", "esp", "ltda", "s", "a", "eu", "sas", "e", "contra",
}

# Minimo de letras para que una palabra del DEMANDADO cuente como
# "significativa" (evita que iniciales sueltas como "S", "A" pasen el
# filtro de PALABRAS_GENERICAS_DEMANDADO si quedaran mal separadas).
MIN_LETRAS_PALABRA_DEMANDADO = 3

# Busca "CONTRA <texto>" en un nombre de carpeta/archivo/correo -- forma
# muy comun de nombrar expedientes en Colombia (ej. "2024-00139 CONTRA
# RIONEGRO"). Se corta el texto capturado en el primer digito, guion,
# parentesis o punto que aparezca despues, para no arrastrar el resto
# del nombre (fecha, numero de radicado, extension del archivo, etc).
_PATRON_CONTRA = re.compile(r"\bCONTRA\b[\s:.-]*([^0-9(){}\[\].]{1,80})", re.IGNORECASE)


def _normalizar_para_comparar(texto: str) -> str:
    """Mayusculas y sin tildes/diacriticos, para comparar nombres sin depender de como esten escritos."""
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.upper()


def _palabras_significativas(texto: str) -> set:
    """Palabras de 'texto' (min. MIN_LETRAS_PALABRA_DEMANDADO letras) que no son PALABRAS_GENERICAS_DEMANDADO."""
    normalizado = _normalizar_para_comparar(texto)
    palabras = re.findall(r"[A-ZÑ]+", normalizado)
    return {
        p for p in palabras
        if len(p) >= MIN_LETRAS_PALABRA_DEMANDADO and p.lower() not in PALABRAS_GENERICAS_DEMANDADO
    }


def _demandado_coincide_en_texto(texto: str, demandado_esperado: str):
    """
    Compara 'texto' contra el DEMANDADO esperado (del Excel) por sus
    palabras significativas.

    Devuelve:
      - True: el nombre del demandado esperado (o una palabra suya
        suficientemente distintiva) aparece en CUALQUIER parte del
        texto -- no hace falta que este despues de la palabra "CONTRA";
        con que el nombre aparezca alcanza para confirmar.
      - False: el texto SI tiene un "CONTRA <algo>" (la forma habitual
        de nombrar expedientes en Colombia, ej. "2024-00139 CONTRA
        RIONEGRO"), pero ese "algo" NO coincide con ninguna palabra del
        demandado esperado -- fuerte indicio de que es de OTRO proceso
        (ej. "CONTRA RIONEGRO" cuando el proceso es contra BOLIVAR).
      - None: no se encontro el nombre del demandado esperado en ningun
        lado, y tampoco hay un "CONTRA <algo>" que lo contradiga -- no
        hay evidencia ni a favor ni en contra, no se bloquea nada.
    """
    if not demandado_esperado or not demandado_esperado.strip():
        return None
    palabras_esperadas = _palabras_significativas(demandado_esperado)
    if not palabras_esperadas:
        return None

    if _palabras_significativas(texto) & palabras_esperadas:
        return True

    coincidencia = _PATRON_CONTRA.search(texto or "")
    if not coincidencia:
        return None
    palabras_encontradas = _palabras_significativas(coincidencia.group(1))
    if not palabras_encontradas:
        return None
    return False


def _demandado_coincide_en_varios(textos, demandado_esperado):
    """
    Aplica _demandado_coincide_en_texto sobre varios textos (ej. nombre
    de carpeta + nombres de sus archivos de primer nivel) y combina el
    resultado: False (mal, rechazar) si CUALQUIERA de ellos da False;
    si no, True si ALGUNO dio True; si ninguno dio nada, None.
    """
    resultados = [_demandado_coincide_en_texto(t, demandado_esperado) for t in textos]
    if any(r is False for r in resultados):
        return False
    if any(r is True for r in resultados):
        return True
    return None

# Cuantos archivos PDF/DOCX de una carpeta candidata se abren como
# maximo para buscar el demandante en su CONTENIDO (si el nombre de la
# carpeta/archivos no lo dice). Igual que MAX_ARCHIVOS_CONTENIDO_A_REVISAR
# en validar_renombrar_carpetas.py, para no abrir carpetas enteras.
MAX_ARCHIVOS_CONTENIDO_A_REVISAR = 5
EXTENSIONES_CONTENIDO_DRIVE = {".pdf", ".docx"}

# Un PDF mas pesado que esto (en MB) NO se abre para leer su contenido --
# se salta directo. Un PDF con la tabla de referencias cruzadas (xref)
# dañada obliga a pypdf a escanear el archivo COMPLETO byte por byte
# para reconstruirla (se ve en el log como fila tras fila de "Ignoring
# wrong pointing object"); en un escaneo pesado de cientos de MB eso
# puede tardar minutos por un solo archivo y dejar el script "pegado"
# sin ningun aviso de que sigue trabajando. Aplica tanto a PDF locales
# (_fecha_de_contenido) como a los descargados en memoria desde Drive
# (_texto_de_archivo_drive).
MAX_MB_PDF_PARA_CONTENIDO = 20

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


def leer_faltantes():
    """Lee ARCHIVO_REPORTE_FALTANTES (No.;Cuenta;Radicado;Juzgado;Demandado) generado por validar_renombrar_carpetas.py."""
    if not os.path.exists(ARCHIVO_REPORTE_FALTANTES):
        raise RuntimeError(
            f"No existe {ARCHIVO_REPORTE_FALTANTES}. Corre primero validar_renombrar_carpetas.py "
            "para generarlo (con el Excel al dia)."
        )
    faltantes = []
    with open(ARCHIVO_REPORTE_FALTANTES, encoding="utf-8-sig") as f:
        for fila in csv.DictReader(f, delimiter=";"):
            faltantes.append({
                "numero": fila.get("No.", "").strip(),
                "cuenta": fila.get("Cuenta", "").strip(),
                "radicado": fila.get("Radicado", "").strip(),
                "juzgado": fila.get("Juzgado", "").strip(),
                # Columna opcional -- si el CSV es de una version vieja del
                # reporte (sin esta columna) o el Excel no tiene DEMANDADO,
                # queda vacia y simplemente no se hace la validacion extra.
                "demandado": fila.get("Demandado", "").strip(),
            })
    return faltantes


def radicados_cortos(radicado: str):
    """
    Deriva los formatos "cortos" del radicado completo de 23 digitos
    (ej. "2025-00456" y "2025-456"), a partir del año (digitos 13-16) y
    el consecutivo (digitos 17-21) del radicado judicial colombiano.
    """
    if len(radicado) != 23 or not radicado.isdigit():
        return []
    anio = radicado[12:16]
    consecutivo_completo = radicado[16:21]
    consecutivo_sin_ceros = str(int(consecutivo_completo))
    formatos = {f"{anio}-{consecutivo_completo}"}
    formatos.add(f"{anio}-{consecutivo_sin_ceros}")
    return sorted(formatos)


# Cuantos digitos minimo debe tener una cuenta para buscarla en Drive.
MIN_DIGITOS_CUENTA_BUSQUEDA = 4


def _cuenta_es_valida_para_buscar(cuenta: str) -> bool:
    """
    True si 'cuenta' es lo bastante especifica como para buscarla en
    Drive sin arrastrar miles de falsos positivos. Rechaza vacia, "0"
    (o cualquier variante de puros ceros), y cuentas muy cortas -- la
    busqueda aproximada de Drive (ver _nombre_coincide) hace coincidir
    un termino tan generico con casi CUALQUIER nombre que tenga esos
    digitos en cualquier parte, lo que puede hacer que la busqueda de
    un solo proceso revise miles de carpetas sin relacion y se demore
    muchisimo.
    """
    cuenta = (cuenta or "").strip()
    if not cuenta or cuenta.strip("0") == "":
        return False
    return len(cuenta) >= MIN_DIGITOS_CUENTA_BUSQUEDA


# ==================== Google Drive ====================


def autenticar_drive():
    if Credentials is None:
        raise RuntimeError(
            "Faltan las librerias de Google Drive. Instala con: "
            "pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib"
        )
    creds = None
    if os.path.exists(TOKEN_DRIVE):
        creds = Credentials.from_authorized_user_file(TOKEN_DRIVE, DRIVE_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENCIALES_DRIVE):
                raise RuntimeError(
                    f"No existe {CREDENCIALES_DRIVE}. Sigue los pasos del README para generar las "
                    "credenciales de Google Drive en Google Cloud Console."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENCIALES_DRIVE, DRIVE_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_DRIVE, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
    return build("drive", "v3", credentials=creds)


def _nombre_coincide(nombre: str, termino: str) -> bool:
    """
    Confirma que 'termino' de verdad aparezca (como texto, sin importar
    mayus/minus) dentro de 'nombre'. Hace falta porque el operador
    "contains" de Google Drive NO busca el texto exacto -- hace
    coincidencia por PREFIJOS DE PALABRA (ej. buscar "2014-26" tambien
    trae carpetas como "26 JULIO" o "26 ENERO", porque alguna palabra
    del nombre empieza por "26"). Sin este filtro, un termino corto
    (radicado corto o cuenta) trae una cantidad enorme de falsos
    positivos de toda la unidad de Drive, sin relacion con el caso.

    Ademas, si 'termino' empieza o termina en un DIGITO, una coincidencia
    NO cuenta si justo al lado (antes o despues) hay otro digito -- sin
    esto, buscar el radicado corto "2023-24" tambien "coincidiria" con
    "2023-244" o "2023-248" (son en realidad radicados DISTINTOS que
    solo comparten el mismo prefijo numerico), mezclando el contenido
    de un proceso con el de otro.
    """
    nombre = (nombre or "").lower()
    termino = termino.lower()
    if not termino:
        return False
    inicio = 0
    while True:
        indice = nombre.find(termino, inicio)
        if indice == -1:
            return False
        caracter_antes = nombre[indice - 1] if indice > 0 else ""
        indice_despues = indice + len(termino)
        caracter_despues = nombre[indice_despues] if indice_despues < len(nombre) else ""
        prefijo_de_numero_mayor = termino[0].isdigit() and caracter_antes.isdigit()
        sufijo_de_numero_mayor = termino[-1].isdigit() and caracter_despues.isdigit()
        if not prefijo_de_numero_mayor and not sufijo_de_numero_mayor:
            return True
        inicio = indice + 1


def buscar_en_drive(servicio, termino: str):
    """Busca en Drive archivos/carpetas cuyo NOMBRE contenga 'termino' DE VERDAD (ver _nombre_coincide). Devuelve [{id, name, mimeType, parents}, ...]."""
    termino_escapado = termino.replace("\\", "\\\\").replace("'", "\\'")
    consulta = f"name contains '{termino_escapado}' and trashed = false"
    resultados = []
    page_token = None
    while True:
        respuesta = servicio.files().list(
            q=consulta, spaces="drive",
            fields="nextPageToken, files(id, name, mimeType, parents)",
            pageToken=page_token,
        ).execute()
        resultados.extend(respuesta.get("files", []))
        page_token = respuesta.get("nextPageToken")
        if not page_token:
            break
    # Drive ya devolvio algunos falsos positivos por su busqueda
    # aproximada (ver _nombre_coincide) -- se filtran aca antes de
    # devolverlos, para no procesarlos/descargarlos mas adelante.
    return [item for item in resultados if _nombre_coincide(item["name"], termino)]


_ID_RAIZ_DRIVE = {}  # cache por servicio: id(servicio) -> id de la carpeta raiz ("Mi unidad")


def _id_raiz_drive(servicio):
    clave = id(servicio)
    if clave not in _ID_RAIZ_DRIVE:
        _ID_RAIZ_DRIVE[clave] = servicio.files().get(fileId="root", fields="id").execute()["id"]
    return _ID_RAIZ_DRIVE[clave]


def carpeta_contenedora(servicio, item):
    """
    Si 'item' ya es una carpeta, la devuelve tal cual; si es un archivo
    suelto, busca y devuelve SU carpeta contenedora. Si el archivo esta
    directo en la raiz de Drive (sin ninguna carpeta contenedora real),
    devuelve None -- NUNCA se debe tratar "Mi unidad" (la raiz completa
    de Drive) como si fuera la carpeta de un caso, o se intentaria
    descargar TODO el Drive.
    """
    if item.get("mimeType") == MIME_CARPETA:
        return item
    padres = item.get("parents") or []
    if not padres:
        return None
    try:
        carpeta = servicio.files().get(fileId=padres[0], fields="id, name, mimeType, parents").execute()
    except HttpError:
        return None
    if carpeta.get("id") == _id_raiz_drive(servicio):
        logging.warning(
            "   (se omite '%s': esta directo en la raiz de tu Drive, sin una carpeta de caso real que "
            "la contenga -- nunca se descarga 'Mi unidad' completa)",
            item.get("name"),
        )
        return None
    return carpeta


def _descargar_archivo_binario(servicio, file_id: str, ruta_local: Path):
    request = servicio.files().get_media(fileId=file_id)
    # _ruta_larga_segura antepone el prefijo especial de Windows para
    # rutas largas -- sin esto, un archivo con nombre/ruta muy larga
    # (frecuente en Drive, ej. oficios con el nombre completo del
    # juzgado y las partes) falla con FileNotFoundError al abrirlo,
    # aunque la ruta "se vea bien".
    with open(organizador._ruta_larga_segura(str(ruta_local)), "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        listo = False
        while not listo:
            _, listo = downloader.next_chunk()


def _exportar_google_doc(servicio, file_id: str, mime_type: str, ruta_local: Path):
    extension, mime_exportar = MIME_EXPORTAR.get(mime_type, (None, None))
    if not mime_exportar:
        logging.warning("   (se omite '%s': tipo de Google no soportado para exportar)", ruta_local.name)
        return
    ruta_final = ruta_local if ruta_local.suffix == extension else ruta_local.with_suffix(ruta_local.suffix + extension)
    request = servicio.files().export_media(fileId=file_id, mimeType=mime_exportar)
    with open(organizador._ruta_larga_segura(str(ruta_final)), "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        listo = False
        while not listo:
            _, listo = downloader.next_chunk()


def descargar_carpeta_drive(servicio, folder_id: str, destino: Path) -> int:
    """
    Descarga recursivamente el contenido de la carpeta de Drive
    'folder_id' dentro de 'destino'. Devuelve cuantos archivos se
    descargaron. Si un archivo puntual falla (ruta demasiado larga,
    permisos, antivirus, error de red puntual, etc) se salta con una
    advertencia y se sigue con el resto -- un solo archivo problematico
    nunca debe abortar la descarga completa de un proceso, ni mucho
    menos el resto de la corrida.
    """
    destino.mkdir(parents=True, exist_ok=True)
    descargados = 0
    page_token = None
    while True:
        respuesta = servicio.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
        ).execute()
        for item in respuesta.get("files", []):
            nombre_seguro = organizador.sanear_nombre(item["name"])
            ruta_local = destino / nombre_seguro
            try:
                if item["mimeType"] == MIME_CARPETA:
                    descargados += descargar_carpeta_drive(servicio, item["id"], ruta_local)
                elif item["mimeType"] in MIME_EXPORTAR:
                    _exportar_google_doc(servicio, item["id"], item["mimeType"], ruta_local)
                    descargados += 1
                elif item["mimeType"].startswith("application/vnd.google-apps."):
                    logging.warning("   (se omite '%s': tipo de Google no descargable directo)", item["name"])
                elif not _es_pdf_o_exportable(item):
                    logging.info("   (se omite '%s': no es un PDF, solo se descargan PDF)", item["name"])
                else:
                    _descargar_archivo_binario(servicio, item["id"], ruta_local)
                    descargados += 1
            except Exception as error:
                logging.warning(
                    "   (no se pudo descargar '%s' -- probablemente la ruta es demasiado larga para Windows, o "
                    "hay un problema de permisos/antivirus/red; se omite y se sigue con el resto: %s)",
                    item["name"], error,
                )
        page_token = respuesta.get("nextPageToken")
        if not page_token:
            break
    return descargados


def _descargar_archivos_sueltos(servicio, archivos, destino: Path) -> int:
    """
    Como descargar_carpeta_drive, pero para una lista puntual de
    ARCHIVOS (no toda una carpeta) -- se usa cuando la coincidencia con
    el radicado vino de un archivo suelto dentro de una carpeta
    GENERICA (ver _carpeta_es_dedicada_al_caso), para no arrastrar el
    resto de esa carpeta (que puede tener folios de otros procesos).
    """
    destino.mkdir(parents=True, exist_ok=True)
    descargados = 0
    for item in archivos:
        nombre_seguro = organizador.sanear_nombre(item["name"])
        ruta_local = destino / nombre_seguro
        try:
            if item["mimeType"] in MIME_EXPORTAR:
                _exportar_google_doc(servicio, item["id"], item["mimeType"], ruta_local)
                descargados += 1
            elif item["mimeType"].startswith("application/vnd.google-apps."):
                logging.warning("   (se omite '%s': tipo de Google no descargable directo)", item["name"])
            elif not _es_pdf_o_exportable(item):
                logging.info("   (se omite '%s': no es un PDF, solo se descargan PDF)", item["name"])
            else:
                _descargar_archivo_binario(servicio, item["id"], ruta_local)
                descargados += 1
        except Exception as error:
            logging.warning(
                "   (no se pudo descargar '%s' -- probablemente la ruta es demasiado larga para Windows, o hay "
                "un problema de permisos/antivirus/red; se omite y se sigue con el resto: %s)",
                item["name"], error,
            )
    return descargados


def enlace_de(item) -> str:
    if item.get("mimeType") == MIME_CARPETA:
        return f"https://drive.google.com/drive/folders/{item['id']}"
    return f"https://drive.google.com/file/d/{item['id']}"


# ==================== Correo (Gmail) ====================


def _decodificar_asunto(asunto_crudo: str) -> str:
    partes = decode_header(asunto_crudo or "")
    return "".join(
        parte.decode(codificacion or "utf-8", errors="ignore") if isinstance(parte, bytes) else parte
        for parte, codificacion in partes
    )


def _fecha_del_correo(mensaje):
    """
    Fecha (datetime.date) del encabezado "Date" del correo, o None si no
    trae uno reconocible. Se usa como ULTIMO recurso para ordenar un
    documento que llego por correo y no tiene ninguna fecha reconocible
    en su propio nombre ni contenido -- ver _fecha_de_mtime.
    """
    try:
        return parsedate_to_datetime(mensaje.get("Date", "")).date()
    except (TypeError, ValueError):
        return None


def buscar_en_correo(usuario: str, app_password: str, termino: str):
    """
    Busca en TODO el correo (no solo la bandeja de entrada) mensajes que
    mencionen 'termino', usando la busqueda propia de Gmail (X-GM-RAW --
    lo mismo que escribirlo en la barra de busqueda de Gmail). Devuelve
    [(asunto, {enlaces_de_drive}, [(nombre_zip, bytes), ...], fecha_correo), ...]
    (fecha_correo es un datetime.date, o None si el correo no trae un
    encabezado "Date" reconocible).
    """
    resultados = []
    with imaplib.IMAP4_SSL("imap.gmail.com") as mail:
        mail.login(usuario, app_password)
        mail.select('"[Gmail]/All Mail"', readonly=True)
        typ, datos = mail.search(None, "X-GM-RAW", f'"{termino}"')
        if typ != "OK" or not datos or not datos[0]:
            return resultados
        for id_correo in datos[0].split():
            typ, msg_datos = mail.fetch(id_correo, "(RFC822)")
            if typ != "OK" or not msg_datos or not msg_datos[0]:
                continue
            mensaje = email.message_from_bytes(msg_datos[0][1])
            asunto = _decodificar_asunto(mensaje.get("Subject", ""))
            fecha_correo = _fecha_del_correo(mensaje)
            enlaces_drive = set()
            adjuntos_zip = []
            for parte in mensaje.walk():
                tipo_contenido = parte.get_content_type()
                if tipo_contenido in ("text/html", "text/plain"):
                    texto_bruto = parte.get_payload(decode=True)
                    if texto_bruto:
                        texto = texto_bruto.decode(parte.get_content_charset() or "utf-8", errors="ignore")
                        enlaces_drive.update(re.findall(r"https://(?:drive|docs)\.google\.com/\S+", texto))
                nombre_adjunto = parte.get_filename()
                if nombre_adjunto and nombre_adjunto.lower().endswith(".zip"):
                    contenido = parte.get_payload(decode=True)
                    if contenido:
                        adjuntos_zip.append((organizador.sanear_nombre(nombre_adjunto), contenido))
            resultados.append((asunto, enlaces_drive, adjuntos_zip, fecha_correo))
    return resultados


def id_de_enlace_drive(url: str):
    """Extrae (id, tipo) de un enlace de Google Drive/Docs, o (None, None) si no se reconoce el formato."""
    m = re.search(r"/folders/([a-zA-Z0-9_-]{10,})", url)
    if m:
        return m.group(1), "carpeta"
    m = re.search(r"/file/d/([a-zA-Z0-9_-]{10,})", url)
    if m:
        return m.group(1), "archivo"
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]{10,})", url)
    if m:
        return m.group(1), "desconocido"
    return None, None


# ==================== Logica principal ====================


def _carpeta_corresponde_al_radicado(servicio, carpeta, radicado: str) -> bool:
    """
    Confirma que 'carpeta' de verdad tenga que ver con ESTE radicado (no
    solo con la cuenta o el radicado corto que trajo la busqueda) --
    revisa si el radicado (completo, o alguno de sus formatos cortos)
    aparece en el NOMBRE de la carpeta, o en el nombre de alguno de sus
    archivos/subcarpetas de primer nivel. Sin esto, una cuenta que se
    repite en documentos de procesos DISTINTOS (misma cliente, casos
    diferentes a lo largo de los años) traeria la carpeta equivocada.
    """
    terminos = [radicado] + radicados_cortos(radicado)
    if any(_nombre_coincide(carpeta.get("name", ""), t) for t in terminos):
        return True
    try:
        respuesta = servicio.files().list(
            q=f"'{carpeta['id']}' in parents and trashed = false",
            fields="files(name)",
        ).execute()
    except HttpError:
        return False
    return any(
        _nombre_coincide(archivo.get("name", ""), t)
        for archivo in respuesta.get("files", [])
        for t in terminos
    )


def _carpeta_es_dedicada_al_caso(item, carpeta, radicado: str) -> bool:
    """
    True si 'carpeta' es de verdad la carpeta PROPIA de este proceso --
    porque la busqueda encontro DIRECTAMENTE esa carpeta (item ya es la
    carpeta), o porque el NOMBRE de la carpeta contenedora menciona el
    radicado. False si 'carpeta' es solo el padre de un ARCHIVO suelto
    que coincidio, pero la carpeta EN SI no tiene nada en su nombre que
    la relacione con este caso -- en ese caso puede ser una carpeta
    GENERICA compartida por muchos procesos (ej. "02. INFORME 2", "24
    ENERO", carpetas de actuaciones/informes por fecha), y no es seguro
    descargarla completa: tendria folios de otros procesos mezclados.
    """
    if item.get("id") == carpeta.get("id"):
        return True
    terminos = [radicado] + radicados_cortos(radicado)
    return any(_nombre_coincide(carpeta.get("name", ""), t) for t in terminos)


def _es_pdf_o_exportable(item) -> bool:
    """True si 'item' es un PDF de verdad, o un tipo de Google (Doc/Sheet/Slide) que se exporta como PDF -- ver MIME_EXPORTAR. Solo se descargan estos: nada mas."""
    if item.get("mimeType") in MIME_EXPORTAR:
        return True
    return Path(item.get("name", "")).suffix.lower() == ".pdf"


def _archivos_relacionados(servicio, carpeta, radicado: str):
    """Archivos PDF (no subcarpetas, no otros tipos) de primer nivel de 'carpeta' cuyo NOMBRE mencione el radicado (completo o corto)."""
    terminos = [radicado] + radicados_cortos(radicado)
    try:
        respuesta = servicio.files().list(
            q=f"'{carpeta['id']}' in parents and trashed = false",
            fields="files(id, name, mimeType)",
        ).execute()
    except HttpError:
        return []
    return [
        archivo for archivo in respuesta.get("files", [])
        if archivo.get("mimeType") != MIME_CARPETA
        and _es_pdf_o_exportable(archivo)
        and any(_nombre_coincide(archivo.get("name", ""), t) for t in terminos)
    ]


def _archivos_tienen_demandante_valido(servicio, archivos) -> bool:
    """Version de _carpeta_tiene_demandante_valido para una lista puntual de archivos (ver _archivos_relacionados)."""
    if any(_nombre_coincide(a.get("name", ""), t) for a in archivos for t in TERMINOS_DEMANDANTE_VALIDO):
        return True
    candidatos_contenido = [
        a for a in archivos
        if Path(a.get("name", "")).suffix.lower() in EXTENSIONES_CONTENIDO_DRIVE
    ]
    for archivo in candidatos_contenido[:MAX_ARCHIVOS_CONTENIDO_A_REVISAR]:
        texto = _texto_de_archivo_drive(servicio, archivo)
        if any(_nombre_coincide(texto, t) for t in TERMINOS_DEMANDANTE_VALIDO):
            return True
    return False


def _texto_de_archivo_drive(servicio, archivo) -> str:
    """Descarga (en memoria, sin guardar en disco) un PDF/DOCX de Drive y devuelve su texto, o "" si falla."""
    nombre = archivo.get("name", "")
    extension = Path(nombre).suffix.lower()
    if extension not in EXTENSIONES_CONTENIDO_DRIVE:
        return ""
    try:
        request = servicio.files().get_media(fileId=archivo["id"])
        buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)
        listo = False
        while not listo:
            _, listo = downloader.next_chunk()
        buffer.seek(0)
    except Exception:
        return ""
    try:
        if extension == ".pdf":
            if PdfReader is None:
                return ""
            if buffer.getbuffer().nbytes > MAX_MB_PDF_PARA_CONTENIDO * 1024 * 1024:
                logging.info(
                    "   (se omite el contenido de '%s': pesa mas de %d MB -- probablemente un escaneo pesado "
                    "con la tabla de referencias dañada, que se demoraria mucho en leer; se sigue sin abrirlo)",
                    nombre, MAX_MB_PDF_PARA_CONTENIDO,
                )
                return ""
            lector = PdfReader(buffer)
            return "\n".join((pagina.extract_text() or "") for pagina in lector.pages)
        if docx is None:
            return ""
        documento = docx.Document(buffer)
        return "\n".join(p.text for p in documento.paragraphs)
    except Exception:
        return ""


def _carpeta_tiene_demandante_valido(servicio, carpeta) -> bool:
    """
    Confirma que el demandante del proceso de 'carpeta' sea ESSA/
    Electrificadora de Santander (ver TERMINOS_DEMANDANTE_VALIDO) --
    revisa primero el nombre de la carpeta y de sus archivos de primer
    nivel (rapido); si ninguno lo dice, abre el contenido de hasta
    MAX_ARCHIVOS_CONTENIDO_A_REVISAR PDF/DOCX como muestra. Sin este
    filtro, una cuenta/radicado corto compartido con procesos de OTRO
    cliente traeria carpetas que no son de ESSA.
    """
    if any(_nombre_coincide(carpeta.get("name", ""), t) for t in TERMINOS_DEMANDANTE_VALIDO):
        return True
    try:
        respuesta = servicio.files().list(
            q=f"'{carpeta['id']}' in parents and trashed = false",
            fields="files(id, name, mimeType)",
        ).execute()
    except HttpError:
        return False
    archivos = respuesta.get("files", [])
    if any(
        _nombre_coincide(archivo.get("name", ""), t)
        for archivo in archivos
        for t in TERMINOS_DEMANDANTE_VALIDO
    ):
        return True

    candidatos_contenido = [
        a for a in archivos
        if Path(a.get("name", "")).suffix.lower() in EXTENSIONES_CONTENIDO_DRIVE
    ]
    for archivo in candidatos_contenido[:MAX_ARCHIVOS_CONTENIDO_A_REVISAR]:
        texto = _texto_de_archivo_drive(servicio, archivo)
        if any(_nombre_coincide(texto, t) for t in TERMINOS_DEMANDANTE_VALIDO):
            return True
    return False


def _carpeta_local_tiene_demandante_valido(carpeta: Path) -> bool:
    """
    Version LOCAL (en el disco, no en Drive) de _carpeta_tiene_demandante_valido
    -- la usa consolidar_duplicados_en_disco() para revalidar carpetas
    "_2", "_3", etc que quedaron de corridas ANTERIORES a que este
    filtro existiera.
    """
    if any(_nombre_coincide(carpeta.name, t) for t in TERMINOS_DEMANDANTE_VALIDO):
        return True
    try:
        archivos = [h for h in carpeta.rglob("*") if h.is_file()]
    except OSError:
        return False
    if any(_nombre_coincide(a.name, t) for a in archivos for t in TERMINOS_DEMANDANTE_VALIDO):
        return True

    candidatos_contenido = [a for a in archivos if a.suffix.lower() in EXTENSIONES_CONTENIDO_DRIVE]
    for archivo in candidatos_contenido[:MAX_ARCHIVOS_CONTENIDO_A_REVISAR]:
        texto = cruce_excel._texto_de_pdf(archivo) if archivo.suffix.lower() == ".pdf" else cruce_excel._texto_de_docx(archivo)
        if any(_nombre_coincide(texto, t) for t in TERMINOS_DEMANDANTE_VALIDO):
            return True
    return False


def _carpeta_local_demandado_coincide(carpeta: Path, demandado: str):
    """
    Version LOCAL (en el disco, no en Drive) de _carpeta_demandado_coincide
    -- revisa el nombre de 'carpeta' y de todos sus archivos buscando un
    "CONTRA <algo>" que contradiga el DEMANDADO esperado. Ver
    _demandado_coincide_en_texto para el significado de True/False/None.
    """
    try:
        nombres = [carpeta.name] + [h.name for h in carpeta.rglob("*") if h.is_file()]
    except OSError:
        return None
    return _demandado_coincide_en_varios(nombres, demandado)


def _zip_tiene_demandante_valido_por_nombre(contenido: bytes) -> bool:
    """
    Revision RAPIDA (solo nombres de archivo dentro del zip, sin
    extraer nada a disco) para MODO_PRUEBA -- ver TERMINOS_DEMANDANTE_VALIDO.
    No reemplaza la revision completa (nombre + contenido de PDF/DOCX)
    que se hace al descargar de verdad.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(contenido)) as zf:
            nombres = zf.namelist()
    except Exception:
        return False
    return any(_nombre_coincide(Path(n).name, t) for n in nombres for t in TERMINOS_DEMANDANTE_VALIDO)


def limpiar_carpetas_temporales_huerfanas():
    """
    Al empezar, revisa si quedaron carpetas TEMPORALES sueltas de una
    corrida anterior que se cerro a la mitad, o que no se pudieron
    borrar solas (ej. un archivo adentro bloqueado por el antivirus o
    por OneDrive -- ver la advertencia que dejan _fusionar_sin_perder_nada
    y _organizar_adjunto_zip cuando eso pasa). Nunca se sabe con certeza
    si TODO su contenido ya quedo copiado en su carpeta final, asi que
    NUNCA se borran solas: se mueven a Duplicados_para_revisar para que
    las revises tu.

    Revisa dos lugares:
      - Carpetas "_tmp_fusion_...", "_tmp_extraccion_correo_..." sueltas
        directo en CARPETA_PROCESOS (las crea este mismo script al
        fusionar candidatos de Drive o adjuntos de correo).
      - Subcarpetas sueltas DENTRO de "_tmp_extraccion" (la usa
        procesos_juridicos.py para extraer los zips que bajas a mano) --
        esa carpeta en si NO se toca, es de uso permanente; solo lo que
        haya quedado adentro de una corrida interrumpida.

    Respeta MODO_PRUEBA.
    """
    carpeta_procesos = Path(CARPETA_PROCESOS)
    if not carpeta_procesos.exists():
        return

    candidatas = []
    try:
        candidatas += [
            h for h in carpeta_procesos.iterdir()
            if h.is_dir() and (h.name.startswith("_tmp_fusion_") or h.name.startswith("_tmp_extraccion_correo_"))
        ]
    except OSError:
        pass

    carpeta_temp_manual = Path(organizador.CARPETA_TEMP_MANUAL)
    if carpeta_temp_manual.exists():
        try:
            candidatas += [h for h in carpeta_temp_manual.iterdir() if h.is_dir()]
        except OSError:
            pass

    if not candidatas:
        return

    if MODO_PRUEBA:
        for carpeta in candidatas:
            logging.info(
                "[SIMULACION -- Temporales] '%s' parece una carpeta temporal que quedo de una corrida "
                "anterior -- se moveria a %s.",
                carpeta, cruce_excel.NOMBRE_CARPETA_DUPLICADOS,
            )
        return

    movidas = 0
    for carpeta in candidatas:
        carpeta_duplicados = carpeta_procesos / cruce_excel.NOMBRE_CARPETA_DUPLICADOS
        carpeta_duplicados.mkdir(parents=True, exist_ok=True)
        destino = cruce_excel.ruta_libre(carpeta_duplicados, f"temporal - {carpeta.name}")
        try:
            shutil.move(organizador._ruta_larga_segura(str(carpeta)), organizador._ruta_larga_segura(str(destino)))
        except OSError as error:
            logging.warning("   (no se pudo mover la carpeta temporal '%s': %s)", carpeta, error)
            continue
        movidas += 1
        logging.info(
            "[Temporales] '%s' parecia una carpeta temporal que quedo de una corrida anterior -- se movio a "
            "'%s/%s' para que la revises (puede que su contenido ya este copiado en la carpeta final del "
            "proceso, o puede que no -- revisala antes de borrarla).",
            carpeta, cruce_excel.NOMBRE_CARPETA_DUPLICADOS, destino.name,
        )

    if movidas:
        logging.info(
            "[Temporales] %d carpeta(s) temporal(es) sueltas se movieron a %s para que las revises.",
            movidas, cruce_excel.NOMBRE_CARPETA_DUPLICADOS,
        )


_PATRON_SUFIJO_DUPLICADO = re.compile(r"_\d+$")


def consolidar_duplicados_en_disco():
    """
    Antes de buscar nada nuevo, revisa si en CARPETA_PROCESOS ya quedaron
    carpetas "_2", "_3", etc para un mismo radicado -- rastros de
    corridas ANTERIORES a que este script fusionara los candidatos
    validos en una sola carpeta (o a que existiera el filtro de
    demandante ESSA). Para cada grupo de carpetas que comparten
    radicado Y tiene ADEMAS al menos una carpeta con sufijo "_N":
      - la carpeta "principal" (la que NO termina en "_N") se queda
        como destino de la fusion;
      - cada carpeta "_N" que SI menciona a ESSA/Electrificadora de
        Santander (por nombre o contenido, ver
        _carpeta_local_tiene_demandante_valido) se fusiona dentro de la
        principal sin perder archivos con nombres repetidos (ver
        _fusionar_sin_perder_nada);
      - cada carpeta "_N" que NO pasa esa validacion se mueve, tal
        cual, a Duplicados_para_revisar (NUNCA se borra) -- probablemente
        es ruido de otro proceso que compartia cuenta o radicado corto.

    IMPORTANTE: si NINGUNA carpeta del grupo tiene sufijo "_N" (son
    todas "numero. radicado" con numeros de proceso DISTINTOS y
    legitimos), o hay MAS DE UNA carpeta sin sufijo, el grupo NO se
    toca -- eso no es un duplicado accidental, es el mismo radicado
    repetido con varios numeros de proceso en el Excel (ver
    "[Duplicado en Excel]" en validar_renombrar_carpetas.py, que a
    proposito deja una carpeta separada por cada numero); fusionarlas
    destruiria esa separacion intencional.

    Respeta MODO_PRUEBA (solo avisa que haria, sin tocar nada).
    """
    carpeta_procesos = Path(CARPETA_PROCESOS)
    if not carpeta_procesos.exists():
        return

    try:
        hijos = [
            h for h in carpeta_procesos.iterdir()
            if h.is_dir() and h.name != cruce_excel.NOMBRE_CARPETA_DUPLICADOS
        ]
    except OSError:
        return

    grupos = {}
    for hijo in hijos:
        radicado = cruce_excel.radicado_de_nombre_carpeta(hijo.name)
        if radicado:
            grupos.setdefault(radicado, []).append(hijo)

    consolidados = 0
    movidos_a_revisar = 0
    for radicado, carpetas in grupos.items():
        if len(carpetas) < 2:
            continue

        con_sufijo = [c for c in carpetas if _PATRON_SUFIJO_DUPLICADO.search(c.name)]
        sin_sufijo = [c for c in carpetas if c not in con_sufijo]

        if not con_sufijo:
            # NINGUNA carpeta del grupo tiene el sufijo "_N" tipico de
            # un duplicado viejo -- son carpetas con NUMEROS DE PROCESO
            # legitimos y DISTINTOS para el mismo radicado (el Excel
            # tiene ese radicado repetido con varios numeros -- ver
            # "[Duplicado en Excel]" en validar_renombrar_carpetas.py,
            # que a proposito crea una carpeta separada POR CADA
            # numero). No es el caso que esta funcion debe resolver:
            # fusionarlas destruiria esa separacion a proposito.
            continue

        if len(sin_sufijo) > 1:
            # Mas de una carpeta SIN sufijo para el mismo radicado
            # (ademas de alguna con sufijo) -- no hay forma segura de
            # saber cual de las "sin sufijo" es la "principal" de la
            # que si tiene sufijo, asi que no se adivina: se deja todo
            # el grupo intacto para que lo revises a mano.
            continue

        principal = sin_sufijo[0] if sin_sufijo else max(carpetas, key=cruce_excel.contar_archivos)

        for carpeta in con_sufijo:
            if carpeta == principal:
                continue
            valido = _carpeta_local_tiene_demandante_valido(carpeta)

            if MODO_PRUEBA:
                logging.info(
                    "[SIMULACION -- Consolidar] '%s' (radicado %s): %s a ESSA/Electrificadora de Santander -- "
                    "se %s.",
                    carpeta.name, radicado, "menciona" if valido else "no se encontro",
                    f"fusionaria en '{principal.name}'" if valido else f"moveria a {cruce_excel.NOMBRE_CARPETA_DUPLICADOS}",
                )
                continue

            if valido:
                _fusionar_sin_perder_nada(carpeta, principal)
                consolidados += 1
                logging.info(
                    "[Consolidado] '%s' se fusiono dentro de '%s' (radicado %s).",
                    carpeta.name, principal.name, radicado,
                )
            else:
                carpeta_duplicados = carpeta_procesos / cruce_excel.NOMBRE_CARPETA_DUPLICADOS
                carpeta_duplicados.mkdir(parents=True, exist_ok=True)
                destino_dup = cruce_excel.ruta_libre(carpeta_duplicados, carpeta.name)
                try:
                    shutil.move(organizador._ruta_larga_segura(str(carpeta)), organizador._ruta_larga_segura(str(destino_dup)))
                except OSError as error:
                    logging.warning(
                        "   (no se pudo mover '%s' a %s -- se deja donde estaba: %s)",
                        carpeta.name, cruce_excel.NOMBRE_CARPETA_DUPLICADOS, error,
                    )
                    continue
                movidos_a_revisar += 1
                logging.info(
                    "[Revisar] '%s' (radicado %s) no menciona a ESSA/Electrificadora de Santander -- se movio a "
                    "%s/%s en vez de fusionarla (probablemente es de otro proceso que comparte cuenta/radicado "
                    "corto).",
                    carpeta.name, radicado, cruce_excel.NOMBRE_CARPETA_DUPLICADOS, destino_dup.name,
                )

    if consolidados or movidos_a_revisar:
        logging.info(
            "[Consolidar] %d carpeta(s) duplicada(s) de corridas anteriores se fusionaron en su carpeta "
            "principal, %d se movieron a %s para que las revises.",
            consolidados, movidos_a_revisar, cruce_excel.NOMBRE_CARPETA_DUPLICADOS,
        )


_PATRON_RADICADO_CORTO_EN_NOMBRE = re.compile(r"(?<!\d)(\d{4})-(\d{2,5})(?!\d)")


def _radicados_cortos_mencionados(nombre: str):
    """Todos los radicados CORTOS ("AAAA-N...") que aparecen en 'nombre', con limites claros de digito -- ver _PATRON_RADICADO_CORTO_EN_NOMBRE."""
    return {f"{m.group(1)}-{m.group(2)}" for m in _PATRON_RADICADO_CORTO_EN_NOMBRE.finditer(nombre)}


def revisar_contaminacion_en_disco(carpetas=None):
    """
    Revisa TODAS las carpetas ya descargadas en CARPETA_PROCESOS (no
    solo las de esta corrida), buscando archivos cuyo NOMBRE mencione
    un radicado corto DISTINTO al de la carpeta que los contiene --
    rastro del bug ya corregido en _nombre_coincide (donde un radicado
    corto como "2023-24" se confundia con "2023-244" o "2023-248", de
    OTRO proceso, y terminaba mezclado en la carpeta equivocada).

    Un archivo se marca como sospechoso SOLO si menciona OTRO radicado
    corto y NUNCA menciona el radicado propio de la carpeta (si
    menciona ambos, se asume que es un documento legitimo que solo
    cita un caso relacionado, y no se toca). Los archivos sospechosos
    se MUEVEN a Duplicados_para_revisar (nunca se borran) para que los
    revises. Si al sacarlos una carpeta queda COMPLETAMENTE vacia
    (todo su contenido era de otro proceso), esa carpeta VACIA si se
    borra -- no hay nada real que perder, y asi la proxima corrida de
    este script la vuelve a buscar de cero, con el filtro ya corregido.

    Solo revisa los NOMBRES de archivo (no abre el contenido) -- rapido
    incluso con cientos de carpetas. Respeta MODO_PRUEBA.

    Por defecto ('carpetas=None') revisa TODAS las carpetas de
    CARPETA_PROCESOS. Si se pasa 'carpetas' (lista de Path), solo
    revisa esas -- lo usa validar_procesos_faltantes.py para revisar
    unicamente las carpetas de los procesos de la lista de faltantes,
    sin tener que recorrer todo el disco.
    """
    carpeta_procesos = Path(CARPETA_PROCESOS)
    if carpetas is None:
        if not carpeta_procesos.exists():
            return
        try:
            carpetas = [
                h for h in carpeta_procesos.iterdir()
                if h.is_dir() and h.name != cruce_excel.NOMBRE_CARPETA_DUPLICADOS
            ]
        except OSError:
            return

    archivos_sospechosos = 0
    carpetas_vaciadas = 0
    for carpeta in carpetas:
        radicado_carpeta = cruce_excel.radicado_de_nombre_carpeta(carpeta.name)
        if not radicado_carpeta:
            continue
        cortos_propios = set(radicados_cortos(radicado_carpeta))
        if not cortos_propios:
            continue

        try:
            archivos = [a for a in carpeta.rglob("*") if a.is_file()]
        except OSError:
            continue

        for archivo in archivos:
            mencionados = _radicados_cortos_mencionados(archivo.name)
            if not mencionados or (mencionados & cortos_propios):
                continue
            ajenos = sorted(mencionados)

            if MODO_PRUEBA:
                logging.info(
                    "[SIMULACION -- Contaminacion] '%s' (dentro de '%s', radicado %s) menciona %s pero no su "
                    "propio radicado -- parece de OTRO proceso; se moveria a %s.",
                    archivo.name, carpeta.name, radicado_carpeta, ", ".join(ajenos), cruce_excel.NOMBRE_CARPETA_DUPLICADOS,
                )
                continue

            carpeta_dup = carpeta_procesos / cruce_excel.NOMBRE_CARPETA_DUPLICADOS / f"{carpeta.name} - posible contenido de otro proceso"
            carpeta_dup.mkdir(parents=True, exist_ok=True)
            destino = cruce_excel.ruta_libre(carpeta_dup, archivo.name)
            try:
                shutil.move(organizador._ruta_larga_segura(str(archivo)), organizador._ruta_larga_segura(str(destino)))
            except OSError as error:
                logging.warning("   (no se pudo mover '%s' de '%s': %s)", archivo.name, carpeta.name, error)
                continue
            archivos_sospechosos += 1
            logging.info(
                "[Contaminacion] '%s' (dentro de '%s', radicado %s) menciona %s pero no su propio radicado -- "
                "parece de OTRO proceso; se movio a '%s/%s'.",
                archivo.name, carpeta.name, radicado_carpeta, ", ".join(ajenos), cruce_excel.NOMBRE_CARPETA_DUPLICADOS, destino.name,
            )

        if not MODO_PRUEBA and cruce_excel.contar_archivos(carpeta) == 0:
            try:
                shutil.rmtree(organizador._ruta_larga_segura(str(carpeta)))
                carpetas_vaciadas += 1
                logging.info(
                    "[Contaminacion] '%s' quedo completamente vacia (todo lo que tenia era de otros procesos) "
                    "-- se borro la carpeta vacia para que la proxima corrida la vuelva a buscar de cero.",
                    carpeta.name,
                )
            except OSError as error:
                logging.warning("   (no se pudo borrar la carpeta vacia '%s': %s)", carpeta.name, error)

    if archivos_sospechosos or carpetas_vaciadas:
        logging.info(
            "[Contaminacion] %d archivo(s) que parecian de otro proceso se movieron a %s para que los "
            "revises; %d carpeta(s) quedaron completamente vacias y se borraron (se van a volver a buscar).",
            archivos_sospechosos, cruce_excel.NOMBRE_CARPETA_DUPLICADOS, carpetas_vaciadas,
        )


def revisar_demandado_en_disco(carpetas=None):
    """
    Revisa TODAS las carpetas ya descargadas en CARPETA_PROCESOS (no
    solo las de esta corrida), buscando archivos cuyo nombre tenga un
    "CONTRA <algo>" que NO corresponda al DEMANDADO real del proceso
    segun el Excel (ver _demandado_coincide_en_texto) -- rastro de una
    carpeta que mezclo contenido de OTRO proceso con el mismo radicado
    corto o cuenta, pero un demandado DISTINTO (ej. "2024-00139 CONTRA
    RIONEGRO" mezclado por error con documentos de "2024-00139 CONTRA
    BOLIVAR").

    Un archivo se marca como sospechoso SOLO si tiene un "CONTRA <algo>"
    que no coincide con ninguna palabra significativa del demandado
    esperado -- si el archivo no menciona ningun "CONTRA", no hay
    evidencia y no se toca. Los archivos sospechosos se MUEVEN a
    Duplicados_para_revisar (nunca se borran), igual que
    revisar_contaminacion_en_disco.

    Si el nombre de la CARPETA misma (no solo un archivo adentro) tiene
    un "CONTRA <algo>" que no corresponde al demandado esperado, solo
    se AVISA en el log -- no se mueve ni renombra la carpeta sola,
    porque podria significar que quedo cruzada con el radicado
    equivocado desde el principio, y eso conviene revisarlo a mano en
    vez de adivinar.

    Si el Excel no tiene la columna DEMANDADO (o RUTA_EXCEL no esta
    configurado/disponible), esta revision simplemente se omite sin
    afectar el resto del script. Respeta MODO_PRUEBA.

    Por defecto ('carpetas=None') revisa TODAS las carpetas de
    CARPETA_PROCESOS. Si se pasa 'carpetas' (lista de Path), solo
    revisa esas -- lo usa validar_procesos_faltantes.py para revisar
    unicamente las carpetas de los procesos de la lista de faltantes,
    sin tener que recorrer todo el disco.
    """
    carpeta_procesos = Path(CARPETA_PROCESOS)
    if not cruce_excel.RUTA_EXCEL or not os.path.exists(cruce_excel.RUTA_EXCEL):
        return
    try:
        demandados_por_radicado = cruce_excel.leer_demandados_por_radicado()
    except Exception:
        logging.exception("[Demandado] No se pudo leer el Excel para cruzar el demandado de cada proceso.")
        return
    if not demandados_por_radicado:
        return

    if carpetas is None:
        if not carpeta_procesos.exists():
            return
        try:
            carpetas = [
                h for h in carpeta_procesos.iterdir()
                if h.is_dir() and h.name != cruce_excel.NOMBRE_CARPETA_DUPLICADOS
            ]
        except OSError:
            return

    archivos_sospechosos = 0
    carpetas_avisadas = 0
    for carpeta in carpetas:
        radicado_carpeta = cruce_excel.radicado_de_nombre_carpeta(carpeta.name)
        if not radicado_carpeta:
            continue
        demandado_esperado = demandados_por_radicado.get(radicado_carpeta)
        if not demandado_esperado:
            continue

        if _demandado_coincide_en_texto(carpeta.name, demandado_esperado) is False:
            carpetas_avisadas += 1
            logging.warning(
                "[Demandado] '%s' parece ser CONTRA otro demandado distinto a '%s' (el que tiene el Excel para "
                "este radicado) -- revisa a mano si esta carpeta quedo cruzada con el proceso equivocado. No se "
                "movio ni renombro nada solo.",
                carpeta.name, demandado_esperado,
            )

        try:
            archivos = [a for a in carpeta.rglob("*") if a.is_file()]
        except OSError:
            continue

        for archivo in archivos:
            if _demandado_coincide_en_texto(archivo.name, demandado_esperado) is not False:
                continue

            if MODO_PRUEBA:
                logging.info(
                    "[SIMULACION -- Demandado] '%s' (dentro de '%s') parece ser CONTRA otro demandado distinto "
                    "a '%s' -- se moveria a %s.",
                    archivo.name, carpeta.name, demandado_esperado, cruce_excel.NOMBRE_CARPETA_DUPLICADOS,
                )
                continue

            carpeta_dup = carpeta_procesos / cruce_excel.NOMBRE_CARPETA_DUPLICADOS / f"{carpeta.name} - posible contenido de otro proceso"
            carpeta_dup.mkdir(parents=True, exist_ok=True)
            destino = cruce_excel.ruta_libre(carpeta_dup, archivo.name)
            try:
                shutil.move(organizador._ruta_larga_segura(str(archivo)), organizador._ruta_larga_segura(str(destino)))
            except OSError as error:
                logging.warning("   (no se pudo mover '%s' de '%s': %s)", archivo.name, carpeta.name, error)
                continue
            archivos_sospechosos += 1
            logging.info(
                "[Demandado] '%s' (dentro de '%s') parece ser CONTRA otro demandado distinto a '%s' -- se "
                "movio a '%s/%s'.",
                archivo.name, carpeta.name, demandado_esperado, cruce_excel.NOMBRE_CARPETA_DUPLICADOS, destino.name,
            )

        if not MODO_PRUEBA and cruce_excel.contar_archivos(carpeta) == 0:
            try:
                shutil.rmtree(organizador._ruta_larga_segura(str(carpeta)))
                logging.info(
                    "[Demandado] '%s' quedo completamente vacia (todo lo que tenia era de otro demandado) -- se "
                    "borro la carpeta vacia para que la proxima corrida la vuelva a buscar de cero.",
                    carpeta.name,
                )
            except OSError as error:
                logging.warning("   (no se pudo borrar la carpeta vacia '%s': %s)", carpeta.name, error)

    if archivos_sospechosos or carpetas_avisadas:
        logging.info(
            "[Demandado] %d archivo(s) que parecian ser de un demandado distinto se movieron a %s para que los "
            "revises; %d carpeta(s) completas quedaron marcadas para revision manual (su propio nombre no "
            "corresponde al demandado del Excel).",
            archivos_sospechosos, cruce_excel.NOMBRE_CARPETA_DUPLICADOS, carpetas_avisadas,
        )


# ==================== Orden cronologico de documentos ====================

_MESES = {
    "enero": 1, "ene": 1, "jan": 1, "january": 1,
    "febrero": 2, "feb": 2, "february": 2,
    "marzo": 3, "mar": 3, "march": 3,
    "abril": 4, "abr": 4, "apr": 4, "april": 4,
    "mayo": 5, "may": 5,
    "junio": 6, "jun": 6, "june": 6,
    "julio": 7, "jul": 7, "july": 7,
    "agosto": 8, "ago": 8, "aug": 8, "august": 8,
    "septiembre": 9, "setiembre": 9, "sep": 9, "sept": 9, "september": 9,
    "octubre": 10, "oct": 10, "october": 10,
    "noviembre": 11, "nov": 11, "november": 11,
    "diciembre": 12, "dic": 12, "dec": 12, "december": 12,
}

# Distintos formatos de fecha que aparecen en la vida real en los
# nombres de archivo de Drive y en el contenido de los documentos.
# Todos exigen el AÑO completo (4 digitos) -- una fecha sin año (ej.
# "24 ENERO") es ambigua entre distintos años y no sirve para ordenar.
_PATRON_FECHA_ISO = re.compile(r"(?<!\d)(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?!\d)")
_PATRON_FECHA_DMY = re.compile(r"(?<!\d)(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?!\d)")
_PATRON_FECHA_DIA_DE_MES_DE_ANIO = re.compile(
    r"(?<!\w)(\d{1,2})\s+de\s+([a-zA-Záéíóúñ]+)\s+de\s+(\d{4})(?!\d)", re.IGNORECASE
)
# Igual que la anterior, pero SIN la palabra "de" -- ej. "29 AGOSTO
# 2023", "15 ENERO 2024". Es, con mucho, el formato mas comun en los
# nombres de auto/providencia de la vida real (mas que con "de"), y
# antes NO se reconocia -- esos archivos se quedaban sin fecha y el
# orden final terminaba mezclado sin secuencia logica.
_PATRON_FECHA_DIA_MES_ANIO = re.compile(
    r"(?<!\w)(\d{1,2})\s+([a-zA-Záéíóúñ]{3,9})\.?\s+(\d{4})(?!\d)", re.IGNORECASE
)
_PATRON_FECHA_MES_DIA_ANIO = re.compile(
    r"(?<!\w)([a-zA-Záéíóúñ]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})(?!\d)", re.IGNORECASE
)


def _fecha_valida(anio: int, mes: int, dia: int):
    try:
        return datetime.date(anio, mes, dia)
    except ValueError:
        return None


def _fecha_en_texto(texto: str):
    """
    Busca la PRIMERA fecha reconocible en 'texto' (nombre de archivo, o
    contenido de un documento), probando varios formatos comunes:
    ISO (2023-07-24), DD/MM/AAAA, "24 de julio de 2023", "24 JULIO
    2023" (sin "de" -- el formato mas comun en nombres de auto reales),
    y "Jul 24 2023". Devuelve un datetime.date, o None si no encuentra
    ninguna.
    """
    if not texto:
        return None

    m = _PATRON_FECHA_ISO.search(texto)
    if m:
        fecha = _fecha_valida(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if fecha:
            return fecha

    m = _PATRON_FECHA_DMY.search(texto)
    if m:
        fecha = _fecha_valida(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        if fecha:
            return fecha

    m = _PATRON_FECHA_DIA_DE_MES_DE_ANIO.search(texto)
    if m:
        mes = _MESES.get(m.group(2).lower())
        if mes:
            fecha = _fecha_valida(int(m.group(3)), mes, int(m.group(1)))
            if fecha:
                return fecha

    m = _PATRON_FECHA_DIA_MES_ANIO.search(texto)
    if m:
        mes = _MESES.get(m.group(2).lower())
        if mes:
            fecha = _fecha_valida(int(m.group(3)), mes, int(m.group(1)))
            if fecha:
                return fecha

    m = _PATRON_FECHA_MES_DIA_ANIO.search(texto)
    if m:
        mes = _MESES.get(m.group(1).lower())
        if mes:
            fecha = _fecha_valida(int(m.group(3)), mes, int(m.group(2)))
            if fecha:
                return fecha

    return None


_PATRON_PREFIJO_ORDEN = re.compile(r"^\d+\.\s+")


def _quitar_prefijo_orden(nombre: str) -> str:
    """Quita un prefijo "N. " que le haya puesto una corrida ANTERIOR de ordenar_y_enumerar_carpeta(), para recalcular el orden desde cero."""
    return _PATRON_PREFIJO_ORDEN.sub("", nombre, count=1)


# La fecha de un documento casi siempre esta en el encabezado -- para
# buscarla en el CONTENIDO alcanza con las primeras paginas; leer un
# PDF COMPLETO (puede tener decenas/cientos de paginas, sobre todo si
# es un escaneo) seria demasiado lento para una carpeta con muchos
# archivos sin fecha en el nombre.
MAX_PAGINAS_CONTENIDO_PARA_FECHA = 2

# Cuantos archivos, como maximo, se abren para buscarles la fecha en
# el CONTENIDO dentro de una misma carpeta -- el resto, si tampoco
# tiene fecha en el nombre, se deja al final sin abrir nada. Sin este
# limite, una carpeta grande (fusionada de varios candidatos, ej. 70+
# archivos) podia dejar el script "pegado" abriendo PDF tras PDF.
MAX_ARCHIVOS_CONTENIDO_PARA_ORDENAR = 20


def _fecha_de_nombre(ruta: Path):
    """Fecha de 'ruta' buscada solo en su NOMBRE (sin el prefijo de orden de una corrida anterior, si lo tenia)."""
    return _fecha_en_texto(_quitar_prefijo_orden(ruta.name))


def _fecha_de_contenido(ruta: Path):
    """Fecha de 'ruta' buscada en las primeras paginas de su contenido (solo PDF/DOCX). Ver MAX_PAGINAS_CONTENIDO_PARA_FECHA."""
    if ruta.suffix.lower() == ".pdf":
        if PdfReader is None:
            return None
        try:
            if ruta.stat().st_size > MAX_MB_PDF_PARA_CONTENIDO * 1024 * 1024:
                logging.info(
                    "   (se omite el contenido de '%s' para buscarle fecha: pesa mas de %d MB -- probablemente "
                    "un escaneo pesado con la tabla de referencias dañada, que se demoraria mucho en leer)",
                    ruta.name, MAX_MB_PDF_PARA_CONTENIDO,
                )
                return None
            lector = PdfReader(str(ruta))
            texto = "\n".join((pagina.extract_text() or "") for pagina in lector.pages[:MAX_PAGINAS_CONTENIDO_PARA_FECHA])
        except Exception:
            return None
    elif ruta.suffix.lower() == ".docx":
        texto = cruce_excel._texto_de_docx(ruta)
    else:
        return None
    return _fecha_en_texto(texto)


# Si un archivo no tiene fecha ni en el nombre ni en el contenido, este
# es el ULTIMO recurso antes de dejarlo sin fecha: su propia fecha de
# modificacion en el disco. Solo es una pista real si alguien la puso
# ahi a proposito -- ver _organizar_adjunto_zip, que le pone la fecha
# del CORREO a los archivos de un adjunto que no traen fecha propia.
# Un archivo recien extraido/copiado (Drive, zip local, etc) tiene un
# mtime de "ahora mismo"; para no confundir eso con una fecha real, se
# exige que sea de mas de un dia atras.
def _fecha_de_mtime(ruta: Path):
    """Fecha de modificacion de 'ruta' en el disco, o None si es muy reciente (probablemente solo la hora de extraccion/copia, no una fecha real)."""
    try:
        mtime = datetime.date.fromtimestamp(ruta.stat().st_mtime)
    except OSError:
        return None
    if (datetime.date.today() - mtime).days < 1:
        return None
    return mtime


def ordenar_y_enumerar_carpeta(carpeta: Path) -> int:
    """
    Dentro de 'carpeta' (y cada una de sus subcarpetas, cada una por su
    cuenta -- nunca se mueven archivos de una subcarpeta a otra), ordena
    los archivos por FECHA (primero el nombre, ver _fecha_de_nombre; si
    no hay, hasta MAX_ARCHIVOS_CONTENIDO_PARA_ORDENAR archivos por
    carpeta pueden abrirse para buscarla en sus primeras paginas, ver
    _fecha_de_contenido; y si tampoco hay, como ultimo recurso la fecha
    de modificacion del archivo en el disco, ver _fecha_de_mtime -- solo
    sirve si alguien la puso ahi a proposito, ej. la fecha del correo en
    _organizar_adjunto_zip) de mas viejo a mas nuevo, y les antepone un
    numero de orden: "1. ", "2. ", etc. Los archivos sin NINGUNA fecha
    reconocible (o que superaron el limite de contenido) quedan al
    final, en el orden en que ya estaban.

    Si se corre varias veces sobre la misma carpeta, primero quita
    cualquier prefijo de orden que le haya puesto una corrida anterior
    y recalcula todo desde cero -- asi, si despues llega un documento
    mas viejo que los demas (ej. por una fusion posterior), el orden se
    corrige solo.

    Nunca borra nada; si el nombre final de un archivo coincidiera con
    el de otro, se le agrega un sufijo libre en vez de pisarlo. Devuelve
    cuantos archivos se renombraron en total (contando subcarpetas).
    """
    try:
        hijos = list(carpeta.iterdir())
    except OSError:
        return 0

    subcarpetas = [h for h in hijos if h.is_dir()]
    archivos = [
        h for h in hijos
        if h.is_file() and h.name.lower() not in cruce_excel.ARCHIVOS_A_IGNORAR_AL_CONTAR
    ]

    con_fecha = []
    sin_fecha = []
    archivos_contenido_revisados = 0
    for archivo in archivos:
        fecha = _fecha_de_nombre(archivo)
        if not fecha and archivos_contenido_revisados < MAX_ARCHIVOS_CONTENIDO_PARA_ORDENAR:
            fecha = _fecha_de_contenido(archivo)
            archivos_contenido_revisados += 1
        if not fecha:
            fecha = _fecha_de_mtime(archivo)
        (con_fecha if fecha else sin_fecha).append((fecha, archivo))
    con_fecha.sort(key=lambda par: par[0])
    orden_final = [archivo for _fecha, archivo in con_fecha] + [archivo for _fecha, archivo in sin_fecha]

    nombres_finales = [
        f"{indice}. {_quitar_prefijo_orden(archivo.name)}"
        for indice, archivo in enumerate(orden_final, start=1)
    ]

    renombrados = 0
    if any(archivo.name != nombre for archivo, nombre in zip(orden_final, nombres_finales)):
        # Se renombra primero a nombres TEMPORALES unicos, y de ahi a los
        # finales -- para que el nombre final de un archivo nunca choque
        # con el nombre ORIGINAL (todavia sin renombrar) de otro.
        temporales = [
            archivo.rename(archivo.with_name(f"__orden_tmp_{uuid.uuid4().hex}{archivo.suffix}"))
            for archivo in orden_final
        ]
        for archivo_original, temporal, nombre_final in zip(orden_final, temporales, nombres_finales):
            if archivo_original.name != nombre_final:
                renombrados += 1
            destino = temporal.with_name(nombre_final)
            if destino.exists():
                destino = _ruta_archivo_libre(temporal.parent, nombre_final)
            temporal.rename(destino)

    for subcarpeta in subcarpetas:
        renombrados += ordenar_y_enumerar_carpeta(subcarpeta)

    return renombrados


def ordenar_todas_las_carpetas_en_disco(carpetas=None):
    """
    Ordena cronologicamente TODAS las carpetas de proceso que ya haya
    en CARPETA_PROCESOS -- no solo las que se acaban de descargar o
    fusionar en esta corrida (ver ordenar_y_enumerar_carpeta). Asi, una
    carpeta que ya estaba en el disco de antes (de una corrida anterior
    a que existiera este orden, o descargada por otro medio) tambien
    queda numerada. Respeta MODO_PRUEBA (no toca nada si esta activo).

    Por defecto ('carpetas=None') ordena TODAS las carpetas de
    CARPETA_PROCESOS. Si se pasa 'carpetas' (lista de Path), solo
    ordena esas -- lo usa validar_procesos_faltantes.py para ordenar
    unicamente las carpetas de los procesos de la lista de faltantes,
    sin tener que recorrer todo el disco.
    """
    if MODO_PRUEBA:
        return
    carpeta_procesos = Path(CARPETA_PROCESOS)
    if carpetas is None:
        if not carpeta_procesos.exists():
            return
        try:
            carpetas = [
                h for h in carpeta_procesos.iterdir()
                if h.is_dir() and h.name != cruce_excel.NOMBRE_CARPETA_DUPLICADOS
            ]
        except OSError:
            return

    total_renombrados = 0
    for carpeta in carpetas:
        if not cruce_excel.radicado_de_nombre_carpeta(carpeta.name):
            continue
        try:
            total_renombrados += ordenar_y_enumerar_carpeta(carpeta)
        except OSError as error:
            logging.warning("   (no se pudo ordenar '%s': %s)", carpeta.name, error)

    if total_renombrados:
        logging.info(
            "[Orden] %d documento(s), entre todas las carpetas de proceso del disco, se ordenaron "
            "cronologicamente y se enumeraron (1., 2., ...).",
            total_renombrados,
        )


def _radicado_ya_en_disco(radicado: str) -> bool:
    """
    True si ya existe una carpeta en CARPETA_PROCESOS para este radicado
    -- se revisa ANTES de buscar/descargar nada, para no crear una
    carpeta "_2" duplicada si el script se corre de nuevo sobre un
    procesos_faltantes_en_disco.csv que ya quedo desactualizado (porque
    una corrida anterior ya trajo ese proceso).
    """
    carpeta_procesos = Path(CARPETA_PROCESOS)
    if not carpeta_procesos.exists():
        return False
    try:
        for hijo in carpeta_procesos.iterdir():
            if hijo.is_dir() and cruce_excel.radicado_de_nombre_carpeta(hijo.name) == radicado:
                return True
    except OSError:
        pass
    return False


def _ruta_archivo_libre(carpeta_padre: Path, nombre_archivo: str) -> Path:
    """Como ruta_libre, pero respetando la EXTENSION del archivo (ej. "informe_2.pdf", no "informe.pdf_2")."""
    ruta = carpeta_padre / nombre_archivo
    if not ruta.exists():
        return ruta
    base = Path(nombre_archivo).stem
    extension = Path(nombre_archivo).suffix
    contador = 2
    while True:
        candidato = carpeta_padre / f"{base}_{contador}{extension}"
        if not candidato.exists():
            return candidato
        contador += 1


def _fusionar_sin_perder_nada(origen, destino) -> int:
    """
    Copia TODO el contenido de 'origen' (carpeta temporal recien
    descargada de OTRO candidato de Drive, distinto del primero, para
    este mismo proceso) DENTRO de 'destino' (donde ya quedo el primer
    candidato) -- pero SIN reemplazar nunca un archivo que ya exista.

    A diferencia de organizador.fusionar_carpeta_en_destino (pensada
    para cuando se vuelve a procesar el MISMO caso/zip actualizado, y
    ahi si tiene sentido que el archivo mas reciente reemplace al
    viejo), aca cada candidato es una carpeta de Drive DISTINTA -- dos
    candidatos distintos pueden traer, por pura coincidencia, un
    archivo o subcarpeta con el mismo nombre (ej. dos "PRINCIPAL", o
    dos "01. INFORME 1") sin ser el mismo documento. Si se reemplazara
    en ese caso, se perderia contenido real. En vez de eso, si el
    nombre ya existe, el archivo que llega se guarda con un sufijo
    libre (ver _ruta_archivo_libre) para quedarse con AMBOS. Si un
    archivo puntual falla al copiarlo (ruta demasiado larga para
    Windows, permisos, antivirus), se salta con una advertencia y se
    sigue con el resto. Al terminar, borra 'origen' (era temporal).
    Devuelve cuantos archivos se copiaron.
    """
    origen = Path(origen)
    destino = Path(destino)
    copiados = 0
    for ruta in sorted(origen.rglob("*")):
        if ruta.is_dir():
            continue
        relativo = ruta.relative_to(origen)
        destino_archivo = destino / relativo
        try:
            destino_archivo.parent.mkdir(parents=True, exist_ok=True)
            if destino_archivo.exists():
                destino_archivo = _ruta_archivo_libre(destino_archivo.parent, destino_archivo.name)
            shutil.copy2(organizador._ruta_larga_segura(str(ruta)), organizador._ruta_larga_segura(str(destino_archivo)))
            copiados += 1
        except OSError as error:
            logging.warning(
                "   (no se pudo fusionar '%s' en '%s' -- probablemente la ruta es demasiado larga para "
                "Windows, o hay un problema de permisos/antivirus; se omite y se sigue con el resto: %s)",
                relativo, destino, error,
            )
    try:
        shutil.rmtree(organizador._ruta_larga_segura(str(origen)))
    except OSError as error:
        logging.warning(
            "   (ya se fusiono todo lo que se pudo de '%s' en '%s', pero no se pudo borrar la carpeta "
            "temporal '%s' -- probablemente un archivo adentro esta bloqueado por el antivirus o alguna "
            "sincronizacion (OneDrive, etc). Puedes borrarla a mano; su contenido ya quedo copiado: %s)",
            origen, destino, origen, error,
        )
    return copiados


def _destino_compartido(numero: str, radicado: str, contexto: dict):
    """
    Devuelve la carpeta de destino asignada a ESTE proceso en esta
    corrida (la crea la primera vez que se llama, y las llamadas
    siguientes reciben la MISMA ruta) -- y si dice si esta es la
    PRIMERA vez que se asigna (True) o si ya existia de un candidato
    anterior (False, hay que fusionar en vez de crear una carpeta
    nueva). Devuelve (destino, es_el_primero).
    """
    si_es_el_primero = contexto["destino"] is None
    if si_es_el_primero:
        contexto["destino"] = cruce_excel.ruta_libre(Path(CARPETA_PROCESOS), f"{numero}. {radicado}")
    return contexto["destino"], si_es_el_primero


def _carpeta_demandado_coincide(servicio, carpeta, demandado):
    """
    Version de _demandado_coincide_en_varios para una carpeta de Drive:
    revisa el nombre de 'carpeta' y los nombres de sus archivos de
    primer nivel, buscando un "CONTRA <algo>" que contradiga el
    DEMANDADO esperado del proceso (ver _demandado_coincide_en_texto).
    """
    textos = [carpeta.get("name", "")]
    try:
        respuesta = servicio.files().list(
            q=f"'{carpeta['id']}' in parents and trashed = false",
            fields="files(name)",
        ).execute()
        textos += [a.get("name", "") for a in respuesta.get("files", [])]
    except HttpError:
        pass
    return _demandado_coincide_en_varios(textos, demandado)


def descargar_coincidencia(servicio, item, numero: str, radicado: str, motivo: str, confiable: bool,
                            descargas_a_validar: list, contexto: dict, demandado: str = "") -> bool:
    """
    Descarga (o simula) la carpeta de 'item' dentro de CARPETA_PROCESOS,
    como 'numero. radicado'. 'motivo' describe como se encontro (ej.
    "radicado completo", "radicado corto: 2025-456", "cuenta: 1000111").
    Si 'confiable' es False (coincidencia por radicado corto, cuenta, o
    enlace de correo sin radicado completo), se agrega a
    'descargas_a_validar' para el reporte aparte.

    'contexto' es un dict COMPARTIDO por todos los candidatos de ESTE
    MISMO proceso en esta corrida (ver procesar_faltante), con:
      - "ya_descargados": set de ID (de carpeta, o de archivo suelto)
        ya bajados -- la busqueda por radicado corto/cuenta/correo
        puede encontrar el MISMO elemento varias veces (ej. por dos
        formatos distintos del radicado corto); si ya se descargo, se
        omite en vez de volver a bajar los mismos archivos otra vez.
      - "destino": la carpeta de destino ya asignada en el disco para
        este proceso (None hasta el primer candidato). Si un SEGUNDO
        candidato (otra carpeta/archivo de Drive distinto, que tambien
        paso las validaciones de radicado y demandante) aparece para el
        mismo proceso, su contenido se FUSIONA dentro de esa misma
        carpeta en vez de crear "_2", "_3", etc -- asi, si el
        expediente esta repartido en varias carpetas de Drive (ej. una
        con el "poder" y otra con el "expediente"), todo termina junto
        en UNA sola carpeta en el disco.

    Si 'item' era un ARCHIVO suelto que vive dentro de una carpeta
    GENERICA (una carpeta cuyo propio nombre no tiene nada que ver con
    este radicado -- ver _carpeta_es_dedicada_al_caso -- tipico de
    carpetas de "informes" o "actuaciones" que juntan documentos de
    MUCHOS procesos distintos), NO se descarga esa carpeta completa
    (traeria folios de otros casos mezclados): solo se bajan, de esa
    carpeta, los archivos de primer nivel que de verdad mencionen este
    radicado.

    'demandado' (si se conoce, del Excel) se usa para no confundir dos
    procesos DISTINTOS que comparten el mismo radicado corto o cuenta
    (ej. "2024-00139 CONTRA RIONEGRO" no es lo mismo que "2024-00139
    CONTRA BOLIVAR") -- ver _demandado_coincide_en_texto.

    Devuelve True si quedo lista (o se simulo, o ya estaba descargada de
    una busqueda anterior).
    """
    carpeta = carpeta_contenedora(servicio, item)
    if not carpeta:
        logging.warning(
            "[Sin coincidencia] Proceso %s (radicado %s): se encontro '%s' pero no se pudo determinar su "
            "carpeta contenedora en Drive.",
            numero, radicado, item.get("name"),
        )
        return False

    if not confiable and not _carpeta_corresponde_al_radicado(servicio, carpeta, radicado):
        logging.info(
            "   (se omite '%s': coincide por %s, pero ni ella ni sus archivos mencionan el radicado %s -- "
            "probablemente es de OTRO proceso que comparte la misma cuenta/año)",
            carpeta["name"], motivo, radicado,
        )
        return False

    dedicada = _carpeta_es_dedicada_al_caso(item, carpeta, radicado)
    archivos_sueltos = None

    if not dedicada:
        # El archivo que coincidio vive en una carpeta GENERICA (su
        # nombre no tiene nada que ver con este radicado) -- se acota
        # la descarga solo a los archivos de esa carpeta que de verdad
        # lo mencionen, para no arrastrar el resto (folios de otros
        # procesos).
        archivos_sueltos = _archivos_relacionados(servicio, carpeta, radicado)
        if not archivos_sueltos:
            logging.info(
                "   (se omite '%s': el archivo que coincidio esta dentro de una carpeta generica sin relacion "
                "directa con el radicado %s, y no se encontraron mas archivos ahi que lo mencionen -- se evita "
                "bajar la carpeta completa)",
                carpeta["name"], radicado,
            )
            return False
        archivos_sin_demandado_ajeno = [
            a for a in archivos_sueltos
            if _demandado_coincide_en_texto(a.get("name", ""), demandado) is not False
        ]
        if not archivos_sin_demandado_ajeno:
            logging.info(
                "   (se omite '%s': los archivos que mencionan el radicado %s parecen ser CONTRA otro "
                "demandado distinto a '%s' -- probablemente OTRO proceso con el mismo radicado corto/cuenta)",
                carpeta["name"], radicado, demandado,
            )
            return False
        archivos_sueltos = archivos_sin_demandado_ajeno
        if not _archivos_tienen_demandante_valido(servicio, archivos_sueltos):
            logging.info(
                "   (se omite '%s': los archivos que mencionan el radicado %s no mencionan a ESSA/"
                "Electrificadora de Santander -- se evita bajar la carpeta generica completa)",
                carpeta["name"], radicado,
            )
            return False
    else:
        if _carpeta_demandado_coincide(servicio, carpeta, demandado) is False:
            logging.info(
                "   (se omite '%s': coincide por %s, pero parece ser CONTRA otro demandado distinto a '%s' -- "
                "probablemente OTRO proceso con el mismo radicado corto/cuenta)",
                carpeta["name"], motivo, demandado,
            )
            return False
        if not _carpeta_tiene_demandante_valido(servicio, carpeta):
            logging.info(
                "   (se omite '%s': coincide por %s, pero no se encontro a ESSA/Electrificadora de Santander "
                "como demandante o demandado -- regla obligatoria, sin excepcion aunque el radicado sea exacto)",
                carpeta["name"], motivo,
            )
            return False

    if dedicada:
        if carpeta["id"] in contexto["ya_descargados"]:
            logging.info(
                "   (la carpeta '%s' ya se habia descargado para el proceso %s por otra busqueda; se omite duplicado)",
                carpeta["name"], numero,
            )
            return True
        contexto["ya_descargados"].add(carpeta["id"])
    else:
        archivos_sueltos = [a for a in archivos_sueltos if a["id"] not in contexto["ya_descargados"]]
        if not archivos_sueltos:
            logging.info(
                "   (los archivos de '%s' relacionados con el proceso %s ya se habian descargado por otra "
                "busqueda; se omite duplicado)",
                carpeta["name"], numero,
            )
            return True
        contexto["ya_descargados"].update(a["id"] for a in archivos_sueltos)

    destino, es_el_primero = _destino_compartido(numero, radicado, contexto)
    etiqueta = "" if confiable else " -- A VALIDAR (coincidencia no exacta)"

    if MODO_PRUEBA:
        if dedicada:
            verbo = "se descargaria" if es_el_primero else "se fusionaria (junto con lo ya encontrado)"
            logging.info(
                "[SIMULACION%s] Proceso %s (radicado %s, %s): %s la carpeta de Drive '%s' (%s) en '%s'.",
                etiqueta, numero, radicado, motivo, verbo, carpeta["name"], enlace_de(carpeta), destino.name,
            )
            nombre_origen = carpeta["name"]
        else:
            verbo = "se descargarian" if es_el_primero else "se fusionarian (junto con lo ya encontrado)"
            logging.info(
                "[SIMULACION%s] Proceso %s (radicado %s, %s): %s SOLO %d archivo(s) que mencionan el radicado, "
                "de la carpeta generica '%s' (%s), en '%s' -- no la carpeta completa, para no traer folios de "
                "otros procesos.",
                etiqueta, numero, radicado, motivo, verbo, len(archivos_sueltos), carpeta["name"],
                enlace_de(carpeta), destino.name,
            )
            nombre_origen = f"{carpeta['name']} (solo {len(archivos_sueltos)} archivo(s) relacionados)"
        if not confiable:
            descargas_a_validar.append((numero, radicado, motivo, nombre_origen, enlace_de(carpeta), destino.name))
        return True

    if dedicada:
        if es_el_primero:
            archivos = descargar_carpeta_drive(servicio, carpeta["id"], destino)
            cruce_excel.aplanar_carpeta_anidada_unica(destino)
            accion = "Descargado"
        else:
            temporal = destino.parent / f"_tmp_fusion_{carpeta['id']}"
            archivos = descargar_carpeta_drive(servicio, carpeta["id"], temporal)
            cruce_excel.aplanar_carpeta_anidada_unica(temporal)
            _fusionar_sin_perder_nada(temporal, destino)
            accion = "Fusionado"
        nombre_origen = carpeta["name"]
    else:
        if es_el_primero:
            archivos = _descargar_archivos_sueltos(servicio, archivos_sueltos, destino)
            accion = "Descargado"
        else:
            temporal = destino.parent / f"_tmp_fusion_archivos_{carpeta['id']}_{id(archivos_sueltos)}"
            archivos = _descargar_archivos_sueltos(servicio, archivos_sueltos, temporal)
            _fusionar_sin_perder_nada(temporal, destino)
            accion = "Fusionado"
        nombre_origen = f"{carpeta['name']} (solo {len(archivos_sueltos)} archivo(s) relacionados, no la carpeta completa)"

    logging.info(
        "[%s%s] Proceso %s (radicado %s, %s): '%s' (%s) -> '%s' (%d archivo(s)).",
        accion, etiqueta, numero, radicado, motivo, nombre_origen, enlace_de(carpeta), destino.name, archivos,
    )
    if not confiable:
        descargas_a_validar.append((numero, radicado, motivo, nombre_origen, enlace_de(carpeta), destino.name))
    return True


def _finalizar_orden(numero: str, contexto: dict):
    """
    Al terminar de procesar un proceso, si de verdad se descargo algo
    en esta corrida (contexto["destino"] esta asignado y no estamos en
    MODO_PRUEBA), ordena cronologicamente y enumera sus documentos --
    ver ordenar_y_enumerar_carpeta().
    """
    if MODO_PRUEBA or contexto["destino"] is None:
        return
    try:
        renombrados = ordenar_y_enumerar_carpeta(contexto["destino"])
        if renombrados:
            logging.info(
                "[Orden] Proceso %s: %d documento(s) de '%s' se ordenaron cronologicamente y se enumeraron "
                "(1., 2., ...).",
                numero, renombrados, contexto["destino"].name,
            )
    except OSError as error:
        logging.warning("[Orden] Proceso %s: no se pudo ordenar '%s': %s", numero, contexto["destino"].name, error)


def procesar_faltante(servicio, credenciales_correo, fila, descargas_a_validar: list):
    numero, cuenta, radicado, juzgado = fila["numero"], fila["cuenta"], fila["radicado"], fila["juzgado"]
    demandado = fila.get("demandado", "")

    if radicado and _radicado_ya_en_disco(radicado):
        logging.info(
            "[Ya en disco] Proceso %s (radicado %s): ya existe una carpeta para este radicado en %s -- se omite "
            "(seguramente ya se habia descargado en una corrida anterior).",
            numero, radicado, CARPETA_PROCESOS,
        )
        return

    # Contexto compartido para ESTE proceso en esta corrida (ver
    # descargar_coincidencia): dedup de carpetas de Drive ya bajadas, y
    # la carpeta de destino ya asignada en el disco (para fusionar ahi
    # los candidatos siguientes en vez de crear "_2", "_3", etc).
    contexto = {"ya_descargados": set(), "destino": None}

    if servicio and radicado:
        coincidencias = buscar_en_drive(servicio, radicado)
        carpetas = [c for c in coincidencias if c["mimeType"] == MIME_CARPETA]
        objetivo = carpetas[0] if carpetas else (coincidencias[0] if coincidencias else None)
        if objetivo and descargar_coincidencia(
            servicio, objetivo, numero, radicado, "radicado completo", True, descargas_a_validar, contexto,
            demandado,
        ):
            _finalizar_orden(numero, contexto)
            return

    # No hubo coincidencia exacta: se descargan los candidatos que
    # aparezcan por radicado corto o cuenta -- todos dentro de la MISMA
    # carpeta de destino (el primero la crea, los siguientes se
    # fusionan ahi, ver descargar_coincidencia), marcados para validar
    # despues. descargar_coincidencia ya verifica que la carpeta
    # candidata realmente mencione este radicado (ver
    # _carpeta_corresponde_al_radicado) y que el demandante sea ESSA
    # (ver _carpeta_tiene_demandante_valido) antes de bajar nada -- asi
    # una cuenta compartida entre varios procesos (o con otro cliente)
    # no trae la carpeta de OTRO caso.
    if servicio:
        for corto in radicados_cortos(radicado):
            for c in buscar_en_drive(servicio, corto):
                descargar_coincidencia(
                    servicio, c, numero, radicado, f"radicado corto: {corto}", False, descargas_a_validar, contexto,
                    demandado,
                )

    if servicio and cuenta:
        if _cuenta_es_valida_para_buscar(cuenta):
            for c in buscar_en_drive(servicio, cuenta):
                descargar_coincidencia(
                    servicio, c, numero, radicado, f"cuenta: {cuenta}", False, descargas_a_validar, contexto,
                    demandado,
                )
        else:
            logging.info(
                "   (no se busca por cuenta '%s' para el proceso %s: es vacia, cero, o demasiado corta -- "
                "buscarla traeria miles de falsos positivos de todo Drive)",
                cuenta, numero,
            )

    if credenciales_correo and BUSCAR_EN_CORREO:
        usuario, app_password = credenciales_correo
        terminos = [radicado] + radicados_cortos(radicado) + ([cuenta] if _cuenta_es_valida_para_buscar(cuenta) else [])
        for termino in terminos:
            if not termino:
                continue
            try:
                correos = buscar_en_correo(usuario, app_password, termino)
            except Exception as error:
                logging.error("[Correo] Fallo buscando '%s': %s", termino, error)
                continue
            confiable = termino == radicado
            for asunto, enlaces, adjuntos, fecha_correo in correos:
                for enlace in enlaces:
                    id_enlace, _tipo = id_de_enlace_drive(enlace)
                    if id_enlace and servicio:
                        try:
                            item = servicio.files().get(fileId=id_enlace, fields="id, name, mimeType, parents").execute()
                        except HttpError as error:
                            logging.warning("[Correo] No se pudo abrir el enlace de Drive en '%s': %s", asunto, error)
                            continue
                        descargar_coincidencia(
                            servicio, item, numero, radicado, f"correo ({termino}): {asunto}", confiable,
                            descargas_a_validar, contexto, demandado,
                        )
                for nombre_zip, contenido in adjuntos:
                    _organizar_adjunto_zip(
                        numero, radicado, asunto, termino, nombre_zip, contenido, confiable, descargas_a_validar,
                        contexto, demandado, fecha_correo,
                    )

    _finalizar_orden(numero, contexto)


def _organizar_adjunto_zip(numero, radicado, asunto, termino, nombre_zip, contenido: bytes, confiable: bool,
                            descargas_a_validar: list, contexto: dict, demandado: str = "", fecha_correo=None):
    motivo = "radicado completo" if confiable else f"correo ({termino}): {asunto}"
    etiqueta = "" if confiable else " -- A VALIDAR (coincidencia no exacta)"

    if _demandado_coincide_en_texto(f"{asunto} {nombre_zip}", demandado) is False:
        logging.info(
            "   (se omite el adjunto '%s' del correo '%s': el asunto/nombre parece ser CONTRA otro demandado "
            "distinto a '%s' -- probablemente OTRO proceso con el mismo radicado corto/cuenta)",
            nombre_zip, asunto, demandado,
        )
        return

    if MODO_PRUEBA:
        if not _zip_tiene_demandante_valido_por_nombre(contenido):
            logging.info(
                "   (el adjunto '%s' del correo '%s' no menciona a ESSA/Electrificadora de Santander en los "
                "nombres de sus archivos -- se confirmara con el contenido cuando MODO_PRUEBA este en False; "
                "si sigue sin mencionarla, no se descargara -- regla obligatoria)",
                nombre_zip, asunto,
            )
            return
        destino, es_el_primero = _destino_compartido(numero, radicado, contexto)
        verbo = "se extraeria" if es_el_primero else "se fusionaria (junto con lo ya encontrado)"
        logging.info(
            "[SIMULACION%s] Proceso %s (radicado %s, %s): %s el adjunto '%s' del correo '%s' en '%s'.",
            etiqueta, numero, radicado, motivo, verbo, nombre_zip, asunto, destino.name,
        )
        if not confiable:
            descargas_a_validar.append((numero, radicado, motivo, nombre_zip, "(adjunto de correo)", destino.name))
        return

    # Se extrae primero a una carpeta TEMPORAL (nunca directo a la
    # carpeta del proceso) para poder revisar su contenido y confirmar
    # que de verdad mencione a ESSA/Electrificadora de Santander antes
    # de agregarlo -- regla obligatoria, sin excepcion.
    carpeta_procesos = Path(CARPETA_PROCESOS)
    ruta_zip_temp = carpeta_procesos / f"_tmp_{nombre_zip}"
    ruta_zip_temp.write_bytes(contenido)
    temporal = carpeta_procesos / f"_tmp_extraccion_correo_{Path(nombre_zip).stem}_{id(contenido)}"
    try:
        extraidos, fallidos = cruce_excel.extraer_zip_en_carpeta(ruta_zip_temp, temporal)

        # Solo interesan los PDF -- se borran (del TEMPORAL, nunca de una
        # carpeta ya organizada) los que no lo sean, antes de decidir
        # nada mas.
        no_pdf_omitidos = 0
        for archivo_extraido in temporal.rglob("*") if temporal.exists() else []:
            if archivo_extraido.is_file() and archivo_extraido.suffix.lower() != ".pdf":
                archivo_extraido.unlink(missing_ok=True)
                no_pdf_omitidos += 1
        if no_pdf_omitidos:
            extraidos -= no_pdf_omitidos
            logging.info(
                "   (se omitieron %d archivo(s) del adjunto '%s' por no ser PDF -- solo se descargan PDF)",
                no_pdf_omitidos, nombre_zip,
            )

        if extraidos <= 0:
            logging.warning(
                "[Correo] El adjunto '%s' del correo '%s' no dejo ningun archivo PDF -- se omite.",
                nombre_zip, asunto,
            )
            try:
                shutil.rmtree(organizador._ruta_larga_segura(str(temporal)))
            except OSError as error:
                logging.warning("   (no se pudo borrar la carpeta temporal '%s': %s)", temporal, error)
            return

        # Un documento que llega por correo casi nunca trae una fecha
        # reconocible en su propio nombre de archivo -- si no la tiene,
        # se le pone como fecha de modificacion la del CORREO (mejor
        # pista real que no tener ninguna), para que
        # ordenar_y_enumerar_carpeta la use como ultimo recurso (ver
        # _fecha_de_mtime). Si el archivo SI trae su propia fecha en el
        # nombre, se deja tal cual -- esa es mas confiable.
        if fecha_correo:
            timestamp = datetime.datetime.combine(fecha_correo, datetime.time.min).timestamp()
            for archivo_extraido in temporal.rglob("*") if temporal.exists() else []:
                if archivo_extraido.is_file() and not _fecha_de_nombre(archivo_extraido):
                    try:
                        os.utime(archivo_extraido, (timestamp, timestamp))
                    except OSError:
                        pass

        def _mover_adjunto_a_duplicados():
            carpeta_duplicados = carpeta_procesos / cruce_excel.NOMBRE_CARPETA_DUPLICADOS
            carpeta_duplicados.mkdir(parents=True, exist_ok=True)
            destino_dup = cruce_excel.ruta_libre(carpeta_duplicados, f"{numero}. {radicado} - correo {Path(nombre_zip).stem}")
            try:
                shutil.move(organizador._ruta_larga_segura(str(temporal)), organizador._ruta_larga_segura(str(destino_dup)))
            except OSError as error:
                logging.warning(
                    "   (no se pudo mover el adjunto extraido de '%s' a %s -- se deja en %s: %s)",
                    nombre_zip, cruce_excel.NOMBRE_CARPETA_DUPLICADOS, temporal, error,
                )

        if _carpeta_local_demandado_coincide(temporal, demandado) is False:
            logging.info(
                "   (se omite el adjunto '%s' del correo '%s': su contenido parece ser CONTRA otro demandado "
                "distinto a '%s' -- probablemente OTRO proceso con el mismo radicado corto/cuenta)",
                nombre_zip, asunto, demandado,
            )
            _mover_adjunto_a_duplicados()
            return

        if not _carpeta_local_tiene_demandante_valido(temporal):
            logging.info(
                "   (se omite el adjunto '%s' del correo '%s': no se encontro a ESSA/Electrificadora de "
                "Santander como demandante o demandado -- regla obligatoria, sin excepcion)",
                nombre_zip, asunto,
            )
            _mover_adjunto_a_duplicados()
            return

        destino, es_el_primero = _destino_compartido(numero, radicado, contexto)
        accion = "Descargado" if es_el_primero else "Fusionado"
        _fusionar_sin_perder_nada(temporal, destino)
        logging.info(
            "[%s%s] Proceso %s (radicado %s, %s): adjunto '%s' del correo '%s' -> '%s' (%d archivo(s)%s).",
            accion, etiqueta, numero, radicado, motivo, nombre_zip, asunto, destino.name, extraidos,
            f", {fallidos} fallidos" if fallidos else "",
        )
        if not confiable:
            descargas_a_validar.append((numero, radicado, motivo, nombre_zip, "(adjunto de correo)", destino.name))
    finally:
        ruta_zip_temp.unlink(missing_ok=True)


def procesar():
    limpiar_carpetas_temporales_huerfanas()
    consolidar_duplicados_en_disco()
    revisar_contaminacion_en_disco()
    revisar_demandado_en_disco()
    ordenar_todas_las_carpetas_en_disco()

    faltantes = leer_faltantes()
    logging.info("Procesos faltantes a buscar: %d", len(faltantes))

    servicio = None
    try:
        servicio = autenticar_drive()
    except Exception as error:
        logging.error("[Drive] No se pudo conectar con Google Drive, se omite esa busqueda: %s", error)

    credenciales_correo = None
    if BUSCAR_EN_CORREO:
        credenciales_correo = organizador.leer_credenciales()
        if not credenciales_correo:
            logging.warning(
                "[Correo] No hay %s (o le faltan datos); se omite la busqueda en correo.",
                organizador.ARCHIVO_CREDENCIALES,
            )

    if not servicio and not credenciales_correo:
        logging.error("No hay ni Drive ni correo configurados -- no hay donde buscar. Revisa las credenciales.")
        return

    descargas_a_validar = []
    for fila in faltantes:
        if not fila["radicado"]:
            continue
        try:
            procesar_faltante(servicio, credenciales_correo, fila, descargas_a_validar)
        except Exception as error:
            logging.error(
                "[Error] Proceso %s (radicado %s) fallo con un error inesperado y se salta -- se sigue con el "
                "resto de la lista: %s",
                fila["numero"], fila["radicado"], error,
            )

    if descargas_a_validar:
        with open(ARCHIVO_REPORTE_A_VALIDAR, "w", newline="", encoding="utf-8-sig") as f:
            escritor = csv.writer(f, delimiter=";")
            escritor.writerow(["No.", "Radicado", "Encontrado por", "Nombre en Drive/correo", "Enlace", "Carpeta descargada"])
            for fila_descarga in descargas_a_validar:
                escritor.writerow(fila_descarga)
        logging.info(
            "[Revisar] %d carpeta(s) se descargaron por una coincidencia MENOS segura (radicado corto/cuenta/"
            "enlace de correo); confirma cuales son correctas en %s -- las que no correspondan, borralas a "
            "mano (el script nunca borra nada solo).",
            len(descargas_a_validar), ARCHIVO_REPORTE_A_VALIDAR,
        )

    if MODO_PRUEBA:
        logging.info(
            "MODO_PRUEBA esta activo: no se descargo nada de verdad todavia. Revisa el reporte de arriba "
            "y, si se ve bien, cambia MODO_PRUEBA = False al inicio del script y vuelve a correrlo."
        )


def main():
    configurar_logging()
    procesar()


if __name__ == "__main__":
    main()
