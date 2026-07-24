"""
Organizador automatico de ZIP de procesos juridicos.

Vigila una carpeta de descargas, y cada vez que llega un .zip nuevo (o al
iniciar, para los que ya esten ahi de hoy):

  1. Espera a que la descarga termine.
  2. Lo extrae (tolerando archivos individuales protegidos con contrasena
     o con rutas demasiado largas para Windows, sin abortar todo el zip
     por uno solo).
  3. Busca dentro de los PDF/DOCX (o en los nombres de archivo) un numero
     de radicado judicial.
  4. Renombra la carpeta extraida con ese radicado (o con el nombre del
     zip si no encuentra ninguno) y la mueve al disco duro destino.
  5. Mueve el .zip original a una subcarpeta "Procesados" (no lo borra).
"""

import datetime
import logging
import os
import re
import shutil
import time
import zipfile
from pathlib import Path

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

try:
    from pypdf import PdfReader
except ImportError:
    from PyPDF2 import PdfReader

import docx

# ============================= CONFIGURACION =============================

# Carpeta donde el navegador guarda los .zip descargados.
CARPETA_DESCARGAS = r"C:\Users\Francy\Downloads"

# Carpeta en el disco duro donde se organizan los procesos ya extraidos.
CARPETA_DESTINO = r"E:/"

ARCHIVO_LOG = os.path.join(CARPETA_DESTINO, "procesos_juridicos.log")

# Patrones para reconocer el numero de radicado dentro del contenido del
# zip. Por defecto reconoce el radicado judicial colombiano de 23 digitos.
PATRONES_RADICADO = [
    r"\b\d{5}[\s\-]?\d{2}[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{4}[\s\-]?\d{5}[\s\-]?\d{2}\b",
    r"\b\d{23}\b",
]
EXTENSIONES_A_REVISAR = {".pdf", ".docx"}
ESPERA_ESTABILIDAD_SEGUNDOS = 3
INTERVALO_CHEQUEO_SEGUNDOS = 1
MAX_INTENTOS_ESTABILIDAD = 120

CARPETA_TEMP_MANUAL = os.path.join(CARPETA_DESTINO, "_tmp_extraccion")

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
    logging.info("Nuevo zip detectado: %s", nombre_zip)

    if not esperar_descarga_completa(ruta_zip):
        logging.error("La descarga de %s nunca se completo o no es un zip valido. Se omite.", nombre_zip)
        return

    try:
        carpeta_extraida = extraer_zip(ruta_zip, CARPETA_TEMP_MANUAL)
    except zipfile.BadZipFile:
        logging.error("%s no es un zip valido. Se omite.", nombre_zip)
        return

    radicado = buscar_radicado(carpeta_extraida)
    if radicado:
        nombre_final = sanear_nombre(radicado)
        logging.info("Radicado encontrado para %s: %s", nombre_zip, radicado)
    else:
        nombre_final = sanear_nombre(Path(ruta_zip).stem)
        logging.warning("No se encontro radicado en %s. Se usara: %s", nombre_zip, nombre_final)

    destino_final = ruta_destino_disponible(CARPETA_DESTINO, nombre_final)
    shutil.move(carpeta_extraida, destino_final)
    logging.info("Proceso organizado en: %s", destino_final)

    mover_zip_a_procesados(ruta_zip)


class ManejadorDescargas(FileSystemEventHandler):
    def __init__(self):
        self.en_proceso = set()

    def _manejar(self, ruta_zip: str):
        if ruta_zip in self.en_proceso:
            return
        self.en_proceso.add(ruta_zip)
        try:
            procesar_zip_manual(ruta_zip)
        except Exception:
            logging.exception("Error inesperado procesando %s", ruta_zip)
        finally:
            self.en_proceso.discard(ruta_zip)

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".zip"):
            self._manejar(event.src_path)

    def on_moved(self, event):
        if not event.is_directory and event.dest_path.lower().endswith(".zip"):
            self._manejar(event.dest_path)


def procesar_zips_existentes():
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
            logging.exception("Error inesperado procesando %s", ruta)


def main():
    configurar_logging()
    Path(CARPETA_DESCARGAS).mkdir(parents=True, exist_ok=True)
    Path(CARPETA_TEMP_MANUAL).mkdir(parents=True, exist_ok=True)

    logging.info("Procesando zips de hoy ya existentes en %s ...", CARPETA_DESCARGAS)
    procesar_zips_existentes()

    logging.info("Vigilando %s en busca de nuevos zips (Ctrl+C para salir)...", CARPETA_DESCARGAS)
    observador = Observer()
    observador.schedule(ManejadorDescargas(), CARPETA_DESCARGAS, recursive=False)
    observador.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observador.stop()
        logging.info("Detenido por el usuario.")
    observador.join()


if __name__ == "__main__":
    main()
