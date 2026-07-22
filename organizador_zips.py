"""
Organizador automatico de ZIP de procesos juridicos.

Vigila una carpeta de descargas, y cada vez que llega un .zip nuevo:
  1. Espera a que la descarga termine.
  2. Lo extrae.
  3. Busca dentro de los documentos (PDF/DOCX) un numero de radicado.
  4. Renombra la carpeta extraida con ese radicado (o con el nombre del
     zip si no encuentra ninguno) y la mueve al disco duro destino.
  5. Mueve el .zip original a una subcarpeta "Procesados" (no lo borra).

Antes de usarlo, edita la seccion CONFIGURACION mas abajo.
"""

import logging
import os
import re
import shutil
import time
import zipfile
from pathlib import Path
from typing import Optional

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

try:
    from pypdf import PdfReader
except ImportError:
    from PyPDF2 import PdfReader

import docx

# ============================= CONFIGURACION =============================

# Carpeta donde el navegador guarda los .zip descargados.
CARPETA_DESCARGAS = r"C:\Users\TuUsuario\Downloads"

# Carpeta en el disco duro donde se deben organizar los procesos.
CARPETA_DESTINO = r"D:\ProcesosJuridicos"

# Patrones para reconocer el numero de radicado dentro del texto.
# Por defecto: radicado judicial colombiano de 23 digitos, con o sin
# guiones/espacios entre grupos. Agrega mas patrones si tus documentos
# usan otro formato.
PATRONES_RADICADO = [
    r"\b\d{5}[\s\-]?\d{2}[\s\-]?\d{2}[\s\-]?\d{3}[\s\-]?\d{4}[\s\-]?\d{5}[\s\-]?\d{2}\b",  # 23 digitos agrupados
    r"\b\d{23}\b",  # 23 digitos seguidos
]

# Extensiones de archivo dentro del zip donde se debe buscar el radicado.
EXTENSIONES_A_REVISAR = {".pdf", ".docx"}

# Segundos que debe permanecer estable el tamano del archivo para
# considerar que la descarga termino.
ESPERA_ESTABILIDAD_SEGUNDOS = 3
INTERVALO_CHEQUEO_SEGUNDOS = 1
MAX_INTENTOS_ESTABILIDAD = 120  # ~2 minutos maximo esperando la descarga

ARCHIVO_LOG = os.path.join(CARPETA_DESTINO, "organizador_zips.log")

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
    """Quita caracteres invalidos para nombres de carpeta en Windows."""
    nombre = re.sub(r'[<>:"/\\|?*]', "_", nombre).strip(" .")
    return nombre or "SinNombre"


def esperar_descarga_completa(ruta_zip: str) -> bool:
    """Espera hasta que el archivo deje de crecer y sea un zip valido."""
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

        if segundos_estable >= ESPERA_ESTABILIDAD_SEGUNDOS:
            if zipfile.is_zipfile(ruta_zip):
                return True
            # Tamano estable pero aun no es un zip valido: seguir esperando un poco.
        time.sleep(INTERVALO_CHEQUEO_SEGUNDOS)
        intentos += 1

    return False


def extraer_zip(ruta_zip: str, carpeta_temp: str) -> str:
    nombre_base = sanear_nombre(Path(ruta_zip).stem)
    destino_extraccion = os.path.join(carpeta_temp, nombre_base)
    contador = 1
    while os.path.exists(destino_extraccion):
        destino_extraccion = os.path.join(carpeta_temp, f"{nombre_base}_{contador}")
        contador += 1

    with zipfile.ZipFile(ruta_zip, "r") as zf:
        zf.extractall(destino_extraccion)

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


def buscar_radicado(carpeta_extraida: str) -> Optional[str]:
    for raiz, _dirs, archivos in os.walk(carpeta_extraida):
        for nombre_archivo in archivos:
            ruta = os.path.join(raiz, nombre_archivo)
            extension = Path(nombre_archivo).suffix.lower()

            # 1) probar contra el nombre del propio archivo
            for patron in PATRONES_RADICADO:
                m = re.search(patron, nombre_archivo)
                if m:
                    return re.sub(r"[\s\-]", "", m.group(0))

            if extension not in EXTENSIONES_A_REVISAR:
                continue

            if extension == ".pdf":
                texto = texto_de_pdf(ruta)
            elif extension == ".docx":
                texto = texto_de_docx(ruta)
            else:
                continue

            for patron in PATRONES_RADICADO:
                m = re.search(patron, texto)
                if m:
                    return re.sub(r"[\s\-]", "", m.group(0))

    return None


def ruta_destino_disponible(carpeta_padre: str, nombre: str) -> str:
    destino = os.path.join(carpeta_padre, nombre)
    contador = 2
    while os.path.exists(destino):
        destino = os.path.join(carpeta_padre, f"{nombre}_{contador}")
        contador += 1
    return destino


def mover_zip_a_procesados(ruta_zip: str):
    carpeta_procesados = os.path.join(os.path.dirname(ruta_zip), "Procesados")
    Path(carpeta_procesados).mkdir(exist_ok=True)
    destino = ruta_destino_disponible(carpeta_procesados, os.path.basename(ruta_zip))
    shutil.move(ruta_zip, destino)


def procesar_zip(ruta_zip: str, carpeta_temp: str):
    nombre_zip = os.path.basename(ruta_zip)
    logging.info("Nuevo zip detectado: %s", nombre_zip)

    if not esperar_descarga_completa(ruta_zip):
        logging.error("La descarga de %s nunca se completo o no es un zip valido. Se omite.", nombre_zip)
        return

    try:
        carpeta_extraida = extraer_zip(ruta_zip, carpeta_temp)
    except zipfile.BadZipFile:
        logging.error("%s no es un zip valido. Se omite.", nombre_zip)
        return

    radicado = buscar_radicado(carpeta_extraida)
    if radicado:
        nombre_final = sanear_nombre(radicado)
        logging.info("Radicado encontrado para %s: %s", nombre_zip, radicado)
    else:
        nombre_final = sanear_nombre(Path(ruta_zip).stem)
        logging.warning("No se encontro radicado en %s. Se usara el nombre del zip: %s", nombre_zip, nombre_final)

    Path(CARPETA_DESTINO).mkdir(parents=True, exist_ok=True)
    destino_final = ruta_destino_disponible(CARPETA_DESTINO, nombre_final)
    shutil.move(carpeta_extraida, destino_final)
    logging.info("Proceso organizado en: %s", destino_final)

    mover_zip_a_procesados(ruta_zip)


class ManejadorDescargas(FileSystemEventHandler):
    def __init__(self, carpeta_temp: str):
        self.carpeta_temp = carpeta_temp
        self.en_proceso = set()

    def _manejar(self, ruta_zip: str):
        if ruta_zip in self.en_proceso:
            return
        self.en_proceso.add(ruta_zip)
        try:
            procesar_zip(ruta_zip, self.carpeta_temp)
        except Exception:
            logging.exception("Error inesperado procesando %s", ruta_zip)
        finally:
            self.en_proceso.discard(ruta_zip)

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".zip"):
            self._manejar(event.src_path)

    def on_moved(self, event):
        # Algunos navegadores renombran el archivo temporal a .zip al terminar.
        if not event.is_directory and event.dest_path.lower().endswith(".zip"):
            self._manejar(event.dest_path)


def procesar_zips_existentes(carpeta_temp: str):
    for nombre in os.listdir(CARPETA_DESCARGAS):
        if nombre.lower().endswith(".zip"):
            ruta = os.path.join(CARPETA_DESCARGAS, nombre)
            try:
                procesar_zip(ruta, carpeta_temp)
            except Exception:
                logging.exception("Error inesperado procesando %s", ruta)


def main():
    configurar_logging()
    Path(CARPETA_DESCARGAS).mkdir(parents=True, exist_ok=True)
    carpeta_temp = os.path.join(CARPETA_DESTINO, "_tmp_extraccion")
    Path(carpeta_temp).mkdir(parents=True, exist_ok=True)

    logging.info("Procesando zips ya existentes en %s ...", CARPETA_DESCARGAS)
    procesar_zips_existentes(carpeta_temp)

    logging.info("Vigilando %s en busca de nuevos zips (Ctrl+C para salir)...", CARPETA_DESCARGAS)
    observador = Observer()
    observador.schedule(ManejadorDescargas(carpeta_temp), CARPETA_DESCARGAS, recursive=False)
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
