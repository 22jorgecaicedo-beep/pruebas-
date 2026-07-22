"""
Descargador automatico de procesos compartidos por el SGDE (Rama Judicial).

Vigila tu bandeja de Gmail buscando correos de
"notificacionessgde@cendoj.ramajudicial.gov.co" que avisan que un juzgado
te comparte un expediente. Cuando encuentra uno nuevo:

  1. Abre el link del correo con un navegador automatizado.
  2. Escribe tu correo en el formulario de validacion.
  3. Espera el segundo correo con el "token" de 6 digitos y lo escribe.
  4. Descarga cada elemento de la tabla "Elementos Compartidos".

Los archivos descargados caen en la misma carpeta que vigila
`organizador_zips.py`, asi que ese script los extrae y organiza solo.

IMPORTANTE:
  - Las credenciales de Gmail se leen de `credenciales_sgde.txt` (NO de
    este archivo), para que nunca queden guardadas en el repositorio.
  - Necesitas una "Contrasena de aplicacion" de Google, no tu contrasena
    normal. Instrucciones en el README.
  - Este script depende de que el sitio web mantenga la misma estructura
    de formularios que se ve en las capturas de pantalla. Si el juzgado
    o la Rama Judicial cambian el diseno de la pagina, los selectores de
    Playwright de mas abajo pueden necesitar ajuste.
"""

import email
import imaplib
import logging
import os
import re
import time
from email.header import decode_header
from pathlib import Path

from playwright.sync_api import sync_playwright

import organizador_zips as organizador

# ============================= CONFIGURACION =============================

ARCHIVO_CREDENCIALES = os.path.join(os.path.dirname(__file__), "credenciales_sgde.txt")

# A donde caen los archivos descargados (debe ser la misma carpeta que
# vigila organizador_zips.py).
CARPETA_DESCARGAS = organizador.CARPETA_DESCARGAS

REMITENTE_SGDE = "notificacionessgde@cendoj.ramajudicial.gov.co"
ASUNTO_COMPARTIDO = "Se le ha compartido información de proceso judicial"
ASUNTO_TOKEN = "Token de validación de acceso a información de proceso judicial"

# Cada cuanto revisar la bandeja en busca de procesos nuevos (segundos).
INTERVALO_REVISION_SEGUNDOS = 60

# Cuanto esperar (maximo) a que llegue el correo con el token, tras pedirlo.
ESPERA_MAXIMA_TOKEN_SEGUNDOS = 90
INTERVALO_CHEQUEO_TOKEN_SEGUNDOS = 3

# Mostrar el navegador mientras trabaja (util para revisar la primera vez).
# Ponlo en True mientras pruebas, y en False cuando ya confies en que funciona.
NAVEGADOR_VISIBLE = True

ARCHIVO_PROCESADOS = os.path.join(os.path.dirname(__file__), "expedientes_procesados.txt")
ARCHIVO_LOG = os.path.join(organizador.CARPETA_DESTINO, "descargador_sgde.log")

# ===========================================================================


def configurar_logging():
    Path(organizador.CARPETA_DESTINO).mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(ARCHIVO_LOG, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def leer_credenciales():
    if not os.path.exists(ARCHIVO_CREDENCIALES):
        raise SystemExit(
            f"No encuentro {ARCHIVO_CREDENCIALES}. Copia "
            "credenciales_sgde.example.txt, renombralo a credenciales_sgde.txt "
            "y pon ahi tu correo y tu contrasena de aplicacion de Gmail."
        )
    datos = {}
    with open(ARCHIVO_CREDENCIALES, encoding="utf-8") as f:
        for linea in f:
            if "=" in linea and not linea.strip().startswith("#"):
                clave, _, valor = linea.partition("=")
                datos[clave.strip()] = valor.strip()
    if "GMAIL_USUARIO" not in datos or "GMAIL_APP_PASSWORD" not in datos:
        raise SystemExit(f"{ARCHIVO_CREDENCIALES} debe tener GMAIL_USUARIO y GMAIL_APP_PASSWORD.")
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


def buscar_correos(conexion, asunto_contiene: str):
    """Devuelve la lista de mensajes del remitente SGDE cuyo asunto contiene el texto dado."""
    conexion.select("INBOX")
    criterio = f'(FROM "{REMITENTE_SGDE}")'
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
        mensajes = buscar_correos(conexion, ASUNTO_TOKEN)
        for msg in reversed(mensajes):  # el mas reciente primero
            asunto = _decodificar(msg.get("Subject", ""))
            if expediente in asunto:
                token = extraer_token(_texto_del_correo(msg))
                if token:
                    return token
        time.sleep(INTERVALO_CHEQUEO_TOKEN_SEGUNDOS)
    return None


def descargar_expediente(pagina, correo_usuario: str, link: str, expediente: str, conexion_imap):
    logging.info("Abriendo portal para expediente %s", expediente)
    pagina.goto(link, wait_until="networkidle")

    campo_correo = pagina.get_by_placeholder("Correo Electrónico")
    campo_correo.fill(correo_usuario)
    pagina.get_by_role("button", name=re.compile("enviar", re.IGNORECASE)).click()

    logging.info("Esperando el correo con el token para %s...", expediente)
    token = esperar_token(conexion_imap, expediente)
    if not token:
        raise RuntimeError(f"No llego el correo con el token para el expediente {expediente} a tiempo.")

    campo_token = pagina.get_by_placeholder("Token de autenticación")
    campo_token.fill(token)
    pagina.get_by_role("button", name=re.compile("validar", re.IGNORECASE)).click()
    pagina.wait_for_selector("text=Elementos Compartidos")

    flechas_descarga = pagina.locator("table [class*=download], table svg, table a").all()
    descargas_realizadas = 0
    filas = pagina.locator("table tbody tr")
    total_filas = filas.count()
    for i in range(total_filas):
        fila = filas.nth(i)
        try:
            with pagina.expect_download(timeout=30000) as info_descarga:
                fila.locator("td:last-child, td >> nth=-1").locator("*").first.click()
            descarga = info_descarga.value
            nombre_sugerido = descarga.suggested_filename or f"{expediente}_{i}.zip"
            ruta_destino = os.path.join(CARPETA_DESCARGAS, nombre_sugerido)
            descarga.save_as(ruta_destino)
            logging.info("Descargado: %s", ruta_destino)
            descargas_realizadas += 1
        except Exception as exc:
            logging.warning("No se pudo descargar la fila %s del expediente %s: %s", i, expediente, exc)

    if descargas_realizadas == 0:
        raise RuntimeError(
            f"No se descargo ningun archivo para el expediente {expediente}. "
            "Es posible que la pagina haya cambiado de estructura; revisa manualmente."
        )


def procesar_expedientes_nuevos(usuario: str, app_password: str, navegador):
    procesados = cargar_procesados()
    conexion = conectar_gmail(usuario, app_password)
    try:
        mensajes = buscar_correos(conexion, ASUNTO_COMPARTIDO)
        for msg in mensajes:
            texto = _texto_del_correo(msg)
            expediente, link = extraer_expediente_y_link(texto)
            if not expediente or not link:
                continue
            if expediente in procesados:
                continue

            logging.info("Expediente nuevo detectado: %s", expediente)
            contexto = navegador.new_context(accept_downloads=True)
            pagina = contexto.new_page()
            try:
                descargar_expediente(pagina, usuario, link, expediente, conexion)
                marcar_procesado(expediente)
            except Exception:
                logging.exception("Fallo procesando el expediente %s", expediente)
            finally:
                contexto.close()
    finally:
        conexion.logout()


def main():
    configurar_logging()
    usuario, app_password = leer_credenciales()
    Path(CARPETA_DESCARGAS).mkdir(parents=True, exist_ok=True)

    logging.info("Iniciando vigilancia de correos SGDE para %s...", usuario)
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=not NAVEGADOR_VISIBLE)
        try:
            while True:
                try:
                    procesar_expedientes_nuevos(usuario, app_password, navegador)
                except Exception:
                    logging.exception("Error revisando correos nuevos")
                time.sleep(INTERVALO_REVISION_SEGUNDOS)
        except KeyboardInterrupt:
            logging.info("Detenido por el usuario.")
        finally:
            navegador.close()


if __name__ == "__main__":
    main()
