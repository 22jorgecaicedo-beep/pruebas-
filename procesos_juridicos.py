"""
Herramienta unica para procesos juridicos: correo -> descarga -> organizado.

Hace dos cosas en paralelo, dentro de un solo programa:

  A) Vigila tu Gmail buscando correos de
     "notificacionessgde@cendoj.ramajudicial.gov.co" que avisan que un
     juzgado te comparte un expediente. Para cada uno:
       1. Abre el link del correo con un navegador automatizado.
       2. Escribe tu correo en el formulario de validacion.
       3. Espera el segundo correo con el "token" de 6 digitos y lo escribe.
       4. Descarga cada elemento de la tabla "Elementos Compartidos" (si una
          fila no tiene flecha de descarga, sino solo un icono de carpeta,
          entra a ella y descarga archivo por archivo, reconstruyendo la
          misma estructura de subcarpetas).
       5. Descomprime/organiza el resultado directo en el disco duro, en una
          carpeta nombrada con el numero de expediente (23 digitos), que ya
          viene confirmado por el propio correo/portal.

  B) Vigila tu carpeta de Descargas por si alguna vez bajas un zip de un
     proceso a mano (por ejemplo desde otro sistema). En ese caso lo
     extrae, busca el numero de radicado dentro de los PDF/DOCX (o en el
     nombre de archivo) y organiza la carpeta igual que la parte A.

En ambos casos, si configuraste `RUTA_EXCEL` en validar_renombrar_carpetas.py
(el informe de procesos), la carpeta nueva se cruza automaticamente contra
ese informe: si el radicado ya aparece ahi, la carpeta se nombra
"numero. radicado" en vez de solo el radicado. Si el radicado todavia no
esta en el informe (por ejemplo porque el Excel no se ha actualizado), la
carpeta se deja solo con el radicado como siempre, y mas tarde puedes
correr validar_renombrar_carpetas.py para completar el nombre cuando el
informe ya lo tenga.

Antes de usarla, edita la seccion CONFIGURACION mas abajo y crea
`credenciales_sgde.txt` (ver credenciales_sgde.example.txt) con tu correo
y una Contrasena de aplicacion de Gmail.

Si `credenciales_sgde.txt` no existe, la herramienta sigue funcionando en
modo "solo organizador manual" (parte B), avisando en el log que la
vigilancia de correo esta desactivada.
"""

import datetime
import email
import imaplib
import logging
import os
import re
import shutil
import threading
import time
import zipfile
from email.header import decode_header
from pathlib import Path

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

try:
    from pypdf import PdfReader
except ImportError:
    from PyPDF2 import PdfReader

import docx

# Cruce opcional con el informe de Excel: reutiliza la MISMA configuracion
# (RUTA_EXCEL, HOJA_EXCEL, etc) que ya tengas en validar_renombrar_carpetas.py,
# para no repetirla en dos archivos. Si ese script no esta al lado de este,
# o falta openpyxl, el cruce simplemente queda desactivado.
try:
    import validar_renombrar_carpetas as cruce_excel
except ImportError:
    cruce_excel = None

# ============================= CONFIGURACION =============================

# --- General ---

# Carpeta donde el navegador guarda los .zip descargados manualmente, y a
# donde caen tambien los archivos que descarga el portal SGDE.
CARPETA_DESCARGAS = r"C:\Users\Francy\Downloads"

# Carpeta en el disco duro donde se organizan los procesos ya extraidos.
CARPETA_DESTINO = r"E:/"

ARCHIVO_LOG = os.path.join(CARPETA_DESTINO, "procesos_juridicos.log")

# --- Organizador manual (parte B) ---

# Patrones para reconocer el numero de radicado dentro del texto de un zip
# descargado a mano (cuando no viene de un correo SGDE que ya lo confirma).
PATRONES_RADICADO = [
    r"\b\d{5}[\s\-]?\d{2}[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{4}[\s\-]?\d{5}[\s\-]?\d{2}\b",
    r"\b\d{23}\b",
]
EXTENSIONES_A_REVISAR = {".pdf", ".docx"}
ESPERA_ESTABILIDAD_SEGUNDOS = 3
INTERVALO_CHEQUEO_SEGUNDOS = 1
MAX_INTENTOS_ESTABILIDAD = 120

# --- Correo + portal SGDE (parte A) ---

ARCHIVO_CREDENCIALES = os.path.join(os.path.dirname(__file__), "credenciales_sgde.txt")
REMITENTE_SGDE = "notificacionessgde@cendoj.ramajudicial.gov.co"
ASUNTO_COMPARTIDO = "Se le ha compartido información de proceso judicial"
ASUNTO_TOKEN = "Token de validación de acceso a información de proceso judicial"

INTERVALO_REVISION_SEGUNDOS = 60
ESPERA_MAXIMA_TOKEN_SEGUNDOS = 90
INTERVALO_CHEQUEO_TOKEN_SEGUNDOS = 3

# Mostrar el navegador mientras trabaja. Deja True mientras pruebas por
# primera vez; pasalo a False cuando ya confies en que funciona bien.
NAVEGADOR_VISIBLE = True

ARCHIVO_PROCESADOS = os.path.join(os.path.dirname(__file__), "expedientes_procesados.txt")
CARPETA_TEMP_DESCARGAS = os.path.join(os.path.dirname(__file__), "_tmp_descargas_sgde")

CARPETA_TEMP_MANUAL = os.path.join(CARPETA_DESTINO, "_tmp_extraccion")

INTENTOS_POR_DESCARGA = 2
TIMEOUT_DESCARGA_MS = 30000

# ===========================================================================


def configurar_logging():
    Path(CARPETA_DESTINO).mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def sanear_nombre(nombre: str) -> str:
    """Quita caracteres invalidos para nombres de carpeta/archivo en Windows."""
    nombre = re.sub(r'[<>:"/\\|?*]', "_", nombre).strip(" .")
    return nombre or "SinNombre"


def ruta_destino_disponible(carpeta_padre: str, nombre: str) -> str:
    destino = os.path.join(carpeta_padre, nombre)
    contador = 2
    while os.path.exists(destino):
        destino = os.path.join(carpeta_padre, f"{nombre}_{contador}")
        contador += 1
    return destino


# ==================== Cruce opcional con el informe de Excel ==================

_CACHE_INFORME = {"mtime": None, "por_radicado": {}}


def _radicado_a_numero_proceso(radicado: str):
    """
    Busca el radicado en el informe de Excel configurado en
    validar_renombrar_carpetas.py y devuelve su numero de proceso, o None
    si el informe no esta configurado/disponible o el radicado todavia no
    aparece ahi. Recarga el Excel solo cuando cambio en disco, para no
    releerlo en cada carpeta si llegan varias seguidas.
    """
    if cruce_excel is None:
        return None
    ruta = cruce_excel.RUTA_EXCEL
    if not ruta or not os.path.exists(ruta):
        return None

    mtime_actual = os.path.getmtime(ruta)
    if _CACHE_INFORME["mtime"] != mtime_actual:
        try:
            filas = cruce_excel.quitar_repetidos(cruce_excel.leer_filas_excel())
            _CACHE_INFORME["por_radicado"] = {radicado_fila: numero for _fila, numero, radicado_fila in filas}
            _CACHE_INFORME["mtime"] = mtime_actual
            logging.info("[Informe] Leido %s (%d procesos) para cruzar radicados.", ruta, len(filas))
        except Exception:
            logging.exception("[Informe] No se pudo leer %s para cruzar el radicado.", ruta)
            return None

    return _CACHE_INFORME["por_radicado"].get(radicado)


def nombre_carpeta_con_numero_proceso(radicado: str) -> str:
    """Devuelve 'numero. radicado' si el radicado ya esta en el informe, o solo el radicado si no."""
    numero = _radicado_a_numero_proceso(radicado)
    if numero is not None:
        return f"{numero}. {radicado}"

    if cruce_excel is not None and cruce_excel.RUTA_EXCEL and os.path.exists(cruce_excel.RUTA_EXCEL):
        logging.warning(
            "[Informe] El radicado %s todavia no aparece en el informe de Excel; la carpeta queda solo "
            "con el radicado. Corre validar_renombrar_carpetas.py mas tarde cuando el informe lo tenga.",
            radicado,
        )
    return radicado


# ======================= PARTE B: organizador manual ======================


def esperar_descarga_completa(ruta_zip: str) -> bool:
    tamano_anterior = -1
    segundos_estable = 0
    intentos = 0

    while intentos < MAX_INTENTOS_ESTABILIDAD:
        if not os.path.exists(ruta_zip):
            return False
        tamano_actual = os.path.getsize(ruta_zip)
        if tamano_actual == tamano_anterior and tamano_actual > 0:
            segundos_estable += INTERVALO_CHEQUEO_SEGUNDOS
        else:
            segundos_estable = 0
        tamano_anterior = tamano_actual

        if segundos_estable >= ESPERA_ESTABILIDAD_SEGUNDOS and zipfile.is_zipfile(ruta_zip):
            return True
        time.sleep(INTERVALO_CHEQUEO_SEGUNDOS)
        intentos += 1

    return False


def _ruta_larga_segura(ruta: str) -> str:
    """En Windows, antepone el prefijo especial para evitar el limite clasico de 260 caracteres por ruta."""
    if os.name == "nt":
        ruta_abs = os.path.abspath(ruta)
        if not ruta_abs.startswith("\\\\?\\"):
            return "\\\\?\\" + ruta_abs
    return ruta


def _extraer_zip_tolerante(ruta_zip: str, destino_extraccion: str):
    """
    Extrae un zip archivo por archivo. Si uno esta protegido con
    contrasena, corrupto, o su ruta es demasiado larga para Windows, lo
    salta con una advertencia en el log y sigue con el resto, en vez de
    abortar la extraccion completa por un solo archivo problematico.
    """
    destino_normalizado = os.path.normpath(os.path.abspath(destino_extraccion))
    with zipfile.ZipFile(ruta_zip, "r") as zf:
        for miembro in zf.infolist():
            ruta_destino = os.path.normpath(os.path.join(destino_normalizado, miembro.filename))
            if not ruta_destino.startswith(destino_normalizado):
                logging.warning("Ruta sospechosa dentro de %s, se omite: %s", os.path.basename(ruta_zip), miembro.filename)
                continue

            ruta_destino_segura = _ruta_larga_segura(ruta_destino)
            try:
                if miembro.is_dir():
                    os.makedirs(ruta_destino_segura, exist_ok=True)
                    continue
                os.makedirs(os.path.dirname(ruta_destino_segura), exist_ok=True)
                with zf.open(miembro) as origen, open(ruta_destino_segura, "wb") as destino:
                    shutil.copyfileobj(origen, destino)
            except RuntimeError:
                logging.warning(
                    "'%s' dentro de %s esta protegido con contrasena; se omite ese archivo.",
                    miembro.filename, os.path.basename(ruta_zip),
                )
            except OSError as exc:
                logging.warning(
                    "No se pudo extraer '%s' de %s (ruta probablemente muy larga para Windows): %s",
                    miembro.filename, os.path.basename(ruta_zip), exc,
                )


def extraer_zip(ruta_zip: str, carpeta_temp: str) -> str:
    nombre_base = sanear_nombre(Path(ruta_zip).stem)
    destino_extraccion = os.path.join(carpeta_temp, nombre_base)
    contador = 1
    while os.path.exists(destino_extraccion):
        destino_extraccion = os.path.join(carpeta_temp, f"{nombre_base}_{contador}")
        contador += 1

    _extraer_zip_tolerante(ruta_zip, destino_extraccion)

    return destino_extraccion


def texto_de_pdf(ruta_pdf: str) -> str:
    try:
        lector = PdfReader(ruta_pdf)
        return "\n".join((pagina.extract_text() or "") for pagina in lector.pages)
    except Exception as exc:
        logging.warning("No se pudo leer PDF %s: %s", ruta_pdf, exc)
        return ""


def texto_de_docx(ruta_docx: str) -> str:
    try:
        documento = docx.Document(ruta_docx)
        return "\n".join(p.text for p in documento.paragraphs)
    except Exception as exc:
        logging.warning("No se pudo leer DOCX %s: %s", ruta_docx, exc)
        return ""


def buscar_radicado(carpeta_extraida: str):
    for raiz, _dirs, archivos in os.walk(carpeta_extraida):
        for nombre_archivo in archivos:
            ruta = os.path.join(raiz, nombre_archivo)
            extension = Path(nombre_archivo).suffix.lower()

            for patron in PATRONES_RADICADO:
                m = re.search(patron, nombre_archivo)
                if m:
                    return re.sub(r"[\s\-]", "", m.group(0))

            if extension not in EXTENSIONES_A_REVISAR:
                continue
            texto = texto_de_pdf(ruta) if extension == ".pdf" else texto_de_docx(ruta)

            for patron in PATRONES_RADICADO:
                m = re.search(patron, texto)
                if m:
                    return re.sub(r"[\s\-]", "", m.group(0))
    return None


def mover_zip_a_procesados(ruta_zip: str):
    carpeta_procesados = os.path.join(os.path.dirname(ruta_zip), "Procesados")
    Path(carpeta_procesados).mkdir(exist_ok=True)
    destino = ruta_destino_disponible(carpeta_procesados, os.path.basename(ruta_zip))
    shutil.move(ruta_zip, destino)


def procesar_zip_manual(ruta_zip: str):
    nombre_zip = os.path.basename(ruta_zip)
    logging.info("[Manual] Nuevo zip detectado: %s", nombre_zip)

    if not esperar_descarga_completa(ruta_zip):
        logging.error("[Manual] La descarga de %s nunca se completo o no es un zip valido. Se omite.", nombre_zip)
        return

    try:
        carpeta_extraida = extraer_zip(ruta_zip, CARPETA_TEMP_MANUAL)
    except zipfile.BadZipFile:
        logging.error("[Manual] %s no es un zip valido. Se omite.", nombre_zip)
        return

    radicado = buscar_radicado(carpeta_extraida)
    if radicado:
        nombre_final = sanear_nombre(nombre_carpeta_con_numero_proceso(radicado))
        logging.info("[Manual] Radicado encontrado para %s: %s", nombre_zip, radicado)
    else:
        nombre_final = sanear_nombre(Path(ruta_zip).stem)
        logging.warning("[Manual] No se encontro radicado en %s. Se usara: %s", nombre_zip, nombre_final)

    destino_final = ruta_destino_disponible(CARPETA_DESTINO, nombre_final)
    shutil.move(carpeta_extraida, destino_final)
    logging.info("[Manual] Proceso organizado en: %s", destino_final)

    mover_zip_a_procesados(ruta_zip)


class ManejadorDescargasManual(FileSystemEventHandler):
    def __init__(self):
        self.en_proceso = set()

    def _manejar(self, ruta_zip: str):
        if ruta_zip in self.en_proceso:
            return
        self.en_proceso.add(ruta_zip)
        try:
            procesar_zip_manual(ruta_zip)
        except Exception:
            logging.exception("[Manual] Error inesperado procesando %s", ruta_zip)
        finally:
            self.en_proceso.discard(ruta_zip)

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".zip"):
            self._manejar(event.src_path)

    def on_moved(self, event):
        if not event.is_directory and event.dest_path.lower().endswith(".zip"):
            self._manejar(event.dest_path)


def procesar_zips_manuales_existentes():
    """
    Al iniciar, solo procesa los .zip de HOY que ya esten en Descargas
    (para no reprocesar años de descargas viejas cada vez que arrancas el
    programa). Los zips nuevos que aparezcan mientras el programa esta
    corriendo se procesan en tiempo real sin importar la fecha.
    """
    hoy = datetime.date.today()
    for nombre in os.listdir(CARPETA_DESCARGAS):
        if not nombre.lower().endswith(".zip"):
            continue
        ruta = os.path.join(CARPETA_DESCARGAS, nombre)
        fecha_modificacion = datetime.date.fromtimestamp(os.path.getmtime(ruta))
        if fecha_modificacion != hoy:
            continue
        try:
            procesar_zip_manual(ruta)
        except Exception:
            logging.exception("[Manual] Error inesperado procesando %s", ruta)


def iniciar_vigilancia_manual():
    Path(CARPETA_TEMP_MANUAL).mkdir(parents=True, exist_ok=True)
    logging.info("[Manual] Procesando zips ya existentes en %s ...", CARPETA_DESCARGAS)
    procesar_zips_manuales_existentes()

    logging.info("[Manual] Vigilando %s por si descargas algo a mano...", CARPETA_DESCARGAS)
    observador = Observer()
    observador.schedule(ManejadorDescargasManual(), CARPETA_DESCARGAS, recursive=False)
    observador.start()
    return observador


# ==================== PARTE A: correo + portal SGDE ========================


def leer_credenciales():
    if not os.path.exists(ARCHIVO_CREDENCIALES):
        return None
    datos = {}
    with open(ARCHIVO_CREDENCIALES, encoding="utf-8") as f:
        for linea in f:
            if "=" in linea and not linea.strip().startswith("#"):
                clave, _, valor = linea.partition("=")
                datos[clave.strip()] = valor.strip()
    if "GMAIL_USUARIO" not in datos or "GMAIL_APP_PASSWORD" not in datos:
        return None
    return datos["GMAIL_USUARIO"], datos["GMAIL_APP_PASSWORD"]


def cargar_procesados() -> set:
    if not os.path.exists(ARCHIVO_PROCESADOS):
        return set()
    with open(ARCHIVO_PROCESADOS, encoding="utf-8") as f:
        return {linea.strip() for linea in f if linea.strip()}


def marcar_procesado(expediente: str):
    with open(ARCHIVO_PROCESADOS, "a", encoding="utf-8") as f:
        f.write(expediente + "\n")


def _decodificar(valor) -> str:
    partes = decode_header(valor)
    return "".join(
        parte.decode(codificacion or "utf-8") if isinstance(parte, bytes) else parte
        for parte, codificacion in partes
    )


def _texto_del_correo(msg) -> str:
    if msg.is_multipart():
        partes = []
        for parte in msg.walk():
            if parte.get_content_type() == "text/plain":
                partes.append(parte.get_payload(decode=True).decode(errors="ignore"))
        if partes:
            return "\n".join(partes)
        for parte in msg.walk():
            if parte.get_content_type() == "text/html":
                return parte.get_payload(decode=True).decode(errors="ignore")
        return ""
    return msg.get_payload(decode=True).decode(errors="ignore")


def conectar_gmail(usuario: str, app_password: str) -> imaplib.IMAP4_SSL:
    conexion = imaplib.IMAP4_SSL("imap.gmail.com")
    conexion.login(usuario, app_password)
    return conexion


def _fecha_imap(dias_atras: int) -> str:
    fecha = datetime.date.today() - datetime.timedelta(days=dias_atras)
    return fecha.strftime("%d-%b-%Y")


# Cuantos dias hacia atras buscar correos del SGDE. Con cuentas que llevan
# años recibiendo estos correos, buscar "desde siempre" descarga cientos o
# miles de mensajes y hace que Gmail corte la conexion (socket error: EOF).
DIAS_ATRAS_BUSQUEDA_CORREO = 3

# Palabras clave sin tildes (para que el SEARCH de IMAP las acepte sin
# problemas de codificacion) usadas para filtrar el asunto DIRECTAMENTE en
# el servidor de Gmail, en vez de descargar el asunto de cada correo uno
# por uno desde Python. Con bandejas muy activas, revisar correo por correo
# es tan lento que Gmail llega a cortar la conexion a mitad de camino.
PALABRA_CLAVE_IMAP_COMPARTIDO = "compartido"
PALABRA_CLAVE_IMAP_TOKEN = "Token de validaci"


def buscar_correos(conexion, asunto_contiene: str, palabra_clave_imap: str):
    conexion.select("INBOX")
    fecha_desde = _fecha_imap(DIAS_ATRAS_BUSQUEDA_CORREO)
    criterio = f'(FROM "{REMITENTE_SGDE}" SINCE {fecha_desde} SUBJECT "{palabra_clave_imap}")'
    estado, datos = conexion.search(None, criterio)
    if estado != "OK":
        return []

    mensajes = []
    for num in datos[0].split():
        estado, datos_msg = conexion.fetch(num, "(RFC822)")
        if estado != "OK":
            continue
        msg = email.message_from_bytes(datos_msg[0][1])
        asunto = _decodificar(msg.get("Subject", ""))
        # El SEARCH del servidor ya filtro por la palabra clave; esta
        # comparacion local (con tildes) es solo una confirmacion extra.
        if asunto_contiene.lower() in asunto.lower():
            mensajes.append(msg)
    return mensajes


def extraer_expediente_y_link(texto: str):
    expediente = re.search(r"Expediente\s*:?\s*(\d{10,})", texto)
    link = re.search(r"https://siugj-sgde\.ramajudicial\.gov\.co\S+", texto)
    if expediente and link:
        return expediente.group(1), link.group(0).rstrip(".,)")
    return None, None


def extraer_token(texto: str):
    m = re.search(r"token de acceso\s*:?\s*\**\s*(\d{6})", texto, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"\b\d{6}\b", texto)
    return m.group(0) if m else None


def esperar_token(conexion, expediente: str):
    limite = time.time() + ESPERA_MAXIMA_TOKEN_SEGUNDOS
    while time.time() < limite:
        mensajes = buscar_correos(conexion, ASUNTO_TOKEN, PALABRA_CLAVE_IMAP_TOKEN)
        for msg in reversed(mensajes):
            asunto = _decodificar(msg.get("Subject", ""))
            if expediente in asunto:
                token = extraer_token(_texto_del_correo(msg))
                if token:
                    return token
        time.sleep(INTERVALO_CHEQUEO_TOKEN_SEGUNDOS)
    return None


def _nombre_fila(fila) -> str:
    try:
        return fila.locator("td").nth(0).inner_text().strip()
    except Exception:
        return "elemento"


def _celda_tiene_control_descarga(celda) -> bool:
    try:
        return celda.locator("svg, img, a, button, [role='button'], [class*='download']").count() > 0
    except Exception:
        return False


def _fila_es_carpeta(fila) -> bool:
    try:
        celda_nombre = fila.locator("td").nth(0)
        return celda_nombre.locator("svg, img, [class*='folder']").count() > 0
    except Exception:
        return False


def _descargar_celda(pagina, celda, carpeta_local: str, nombre_base: str) -> bool:
    for intento in range(1, INTENTOS_POR_DESCARGA + 1):
        try:
            with pagina.expect_download(timeout=TIMEOUT_DESCARGA_MS) as info_descarga:
                celda.locator("svg, img, a, button, [role='button']").first.click()
            descarga = info_descarga.value
            nombre_archivo = descarga.suggested_filename or sanear_nombre(nombre_base)
            ruta_destino = os.path.join(carpeta_local, nombre_archivo)
            descarga.save_as(ruta_destino)
            logging.info("[SGDE] Descargado: %s", ruta_destino)
            return True
        except Exception as exc:
            logging.warning(
                "[SGDE] Intento %s/%s fallido descargando '%s': %s", intento, INTENTOS_POR_DESCARGA, nombre_base, exc
            )
    return False


def descargar_elementos_de_tabla(pagina, carpeta_local: str) -> int:
    pagina.wait_for_selector("table:visible tbody tr", timeout=15000)
    descargas_totales = 0

    while True:
        filas = pagina.locator("table:visible tbody tr")
        total_filas = filas.count()

        for i in range(total_filas):
            fila = filas.nth(i)
            nombre = _nombre_fila(fila)
            celda_acciones = fila.locator("td").nth(3)
            celda_anexos = fila.locator("td").nth(4) if fila.locator("td").count() > 4 else None

            if _celda_tiene_control_descarga(celda_acciones):
                if _descargar_celda(pagina, celda_acciones, carpeta_local, nombre):
                    descargas_totales += 1

            elif _fila_es_carpeta(fila):
                logging.info("[SGDE] '%s' no tiene descarga directa; entrando a la carpeta...", nombre)
                subcarpeta = os.path.join(carpeta_local, sanear_nombre(nombre))
                os.makedirs(subcarpeta, exist_ok=True)
                try:
                    fila.locator("td").nth(0).click()
                    descargas_totales += descargar_elementos_de_tabla(pagina, subcarpeta)
                finally:
                    boton_regresar = pagina.get_by_role("button", name=re.compile("regresar", re.IGNORECASE))
                    if boton_regresar.count() > 0:
                        boton_regresar.first.click()
                    else:
                        pagina.go_back()
                    pagina.wait_for_selector("table:visible tbody tr", timeout=15000)
                    filas = pagina.locator("table:visible tbody tr")

            else:
                logging.warning(
                    "[SGDE] No se encontro forma de descargar '%s' (sin flecha ni carpeta). Revisa manualmente.",
                    nombre,
                )

            if celda_anexos is not None and _celda_tiene_control_descarga(celda_anexos):
                _descargar_celda(pagina, celda_anexos, carpeta_local, f"{nombre}_anexo")

        boton_siguiente = pagina.locator("button[aria-label='Next Page']")
        if boton_siguiente.count() > 0 and boton_siguiente.first.is_enabled():
            boton_siguiente.first.click()
            pagina.wait_for_timeout(500)
        else:
            break

    return descargas_totales


def organizar_descarga_sgde(carpeta_temp: str, expediente: str) -> str:
    """
    Deja en CARPETA_DESTINO una carpeta ya descomprimida, nombrada con el
    numero de expediente (23 digitos), ya sea:
      - descomprimiendo el unico .zip descargado (caso normal, flecha
        disponible), o
      - moviendo tal cual la estructura reconstruida archivo por archivo
        (caso de carpetas sin flecha).
    """
    nombre_final = sanear_nombre(nombre_carpeta_con_numero_proceso(expediente))
    destino_final = ruta_destino_disponible(CARPETA_DESTINO, nombre_final)
    contenidos = os.listdir(carpeta_temp)

    if len(contenidos) == 1 and contenidos[0].lower().endswith(".zip"):
        ruta_zip = os.path.join(carpeta_temp, contenidos[0])
        _extraer_zip_tolerante(ruta_zip, destino_final)
    else:
        shutil.move(carpeta_temp, destino_final)

    return destino_final


def _esperar_resultado_envio_correo(pagina, timeout_ms: int = 8000):
    """
    Tras darle 'Enviar' al correo, el portal hace una de dos cosas: muestra
    el campo para escribir el token (exito), o un dialogo de error (correo
    no coincide). Esta funcion espera activamente a que aparezca cualquiera
    de los dos, y devuelve el texto del error si lo hay, o None si tuvo
    exito (o si ninguno aparecio dentro del tiempo dado).
    """
    limite = time.time() + (timeout_ms / 1000)
    while time.time() < limite:
        if pagina.get_by_placeholder("Token de autenticación").count() > 0:
            return None
        error_loc = pagina.locator("text=/no coinciden/i")
        if error_loc.count() > 0:
            try:
                return error_loc.first.inner_text()
            except Exception:
                return "El portal mostro un mensaje de error tras enviar el correo."
        pagina.wait_for_timeout(300)
    return None


def descargar_expediente(pagina, correo_usuario: str, link: str, expediente: str, conexion_imap):
    logging.info("[SGDE] Abriendo portal para expediente %s", expediente)
    pagina.goto(link, wait_until="networkidle")

    pagina.get_by_placeholder("Correo Electrónico").fill(correo_usuario.strip())
    pagina.get_by_role("button", name=re.compile("enviar", re.IGNORECASE)).click()

    error_portal = _esperar_resultado_envio_correo(pagina)
    if error_portal:
        raise RuntimeError(
            f"El portal SGDE rechazo el correo para el expediente {expediente}: '{error_portal}'. "
            "Verifica que el correo en credenciales_sgde.txt sea exactamente el que el juzgado "
            "registro para este expediente (mayusculas, espacios, dominio)."
        )

    logging.info("[SGDE] Esperando el correo con el token para %s...", expediente)
    token = esperar_token(conexion_imap, expediente)
    if not token:
        raise RuntimeError(f"No llego el correo con el token para el expediente {expediente} a tiempo.")

    pagina.get_by_placeholder("Token de autenticación").fill(token)
    pagina.get_by_role("button", name=re.compile("validar", re.IGNORECASE)).click()
    pagina.wait_for_selector("text=Elementos Compartidos")

    carpeta_temp = os.path.join(CARPETA_TEMP_DESCARGAS, expediente)
    if os.path.exists(carpeta_temp):
        shutil.rmtree(carpeta_temp)
    os.makedirs(carpeta_temp, exist_ok=True)

    try:
        descargas_totales = descargar_elementos_de_tabla(pagina, carpeta_temp)
        if descargas_totales == 0:
            raise RuntimeError(
                f"No se descargo ningun archivo para el expediente {expediente}. "
                "Es posible que la pagina haya cambiado de estructura; revisa manualmente."
            )
        destino_final = organizar_descarga_sgde(carpeta_temp, expediente)
        logging.info("[SGDE] Proceso organizado en: %s", destino_final)
    finally:
        if os.path.exists(carpeta_temp):
            shutil.rmtree(carpeta_temp, ignore_errors=True)


def procesar_expedientes_nuevos(usuario: str, app_password: str, navegador):
    procesados = cargar_procesados()
    conexion = conectar_gmail(usuario, app_password)
    try:
        mensajes = buscar_correos(conexion, ASUNTO_COMPARTIDO, PALABRA_CLAVE_IMAP_COMPARTIDO)
        for msg in mensajes:
            texto = _texto_del_correo(msg)
            expediente, link = extraer_expediente_y_link(texto)
            if not expediente or not link:
                continue
            if expediente in procesados:
                continue

            logging.info("[SGDE] Expediente nuevo detectado: %s", expediente)
            contexto = navegador.new_context(accept_downloads=True)
            pagina = contexto.new_page()
            try:
                descargar_expediente(pagina, usuario, link, expediente, conexion)
                marcar_procesado(expediente)
            except Exception:
                logging.exception("[SGDE] Fallo procesando el expediente %s", expediente)
            finally:
                contexto.close()
    finally:
        conexion.logout()


def iniciar_vigilancia_correo(usuario: str, app_password: str):
    from playwright.sync_api import sync_playwright

    Path(CARPETA_TEMP_DESCARGAS).mkdir(parents=True, exist_ok=True)
    logging.info("[SGDE] Iniciando vigilancia de correo para %s...", usuario)
    logging.info("[SGDE] Correo exacto que se va a usar en el portal (revisa que no tenga espacios raros): %r", usuario.strip())
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=not NAVEGADOR_VISIBLE)
        try:
            while True:
                try:
                    procesar_expedientes_nuevos(usuario, app_password, navegador)
                except Exception:
                    logging.exception("[SGDE] Error revisando correos nuevos")
                time.sleep(INTERVALO_REVISION_SEGUNDOS)
        finally:
            navegador.close()


# ================================== MAIN ===================================


def main():
    configurar_logging()
    Path(CARPETA_DESCARGAS).mkdir(parents=True, exist_ok=True)
    Path(CARPETA_DESTINO).mkdir(parents=True, exist_ok=True)

    observador_manual = iniciar_vigilancia_manual()

    credenciales = leer_credenciales()
    if not credenciales:
        logging.warning(
            "No hay %s (o le faltan datos). La vigilancia automatica de correo "
            "queda desactivada; solo se organizaran los zips que descargues a mano.",
            ARCHIVO_CREDENCIALES,
        )
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
    else:
        usuario, app_password = credenciales
        try:
            iniciar_vigilancia_correo(usuario, app_password)
        except KeyboardInterrupt:
            pass

    observador_manual.stop()
    observador_manual.join()
    logging.info("Detenido por el usuario.")


if __name__ == "__main__":
    main()
