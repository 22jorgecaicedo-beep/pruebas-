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
  3. El numero de CUENTA.
Las busquedas 2 y 3 son menos confiables (un radicado corto o una cuenta
se puede repetir o coincidir por casualidad con archivos de otro caso),
pero TAMBIEN se descargan automatico (cada candidato en su propia
carpeta, sin pisar nada). La diferencia es que quedan marcadas aparte,
en faltantes_descargados_a_validar.csv, para que confirmes despues cual
de esas descargas es la correcta y borres a mano las que no correspondan
(el script nunca borra nada solo).

Si lo que encuentra es un ARCHIVO suelto (no una carpeta) que coincide,
busca la carpeta que lo contiene y descarga esa carpeta completa (no
solo el archivo), asumiendo que ahi esta el resto del expediente.

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
import email
import imaplib
import logging
import os
import re
from email.header import decode_header
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
    """Lee ARCHIVO_REPORTE_FALTANTES (No.;Cuenta;Radicado;Juzgado) generado por validar_renombrar_carpetas.py."""
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
    """
    return termino.lower() in (nombre or "").lower()


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
    with open(ruta_local, "wb") as f:
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
    with open(ruta_final, "wb") as f:
        downloader = MediaIoBaseDownload(f, request)
        listo = False
        while not listo:
            _, listo = downloader.next_chunk()


def descargar_carpeta_drive(servicio, folder_id: str, destino: Path) -> int:
    """Descarga recursivamente el contenido de la carpeta de Drive 'folder_id' dentro de 'destino'. Devuelve cuantos archivos se descargaron."""
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
            if item["mimeType"] == MIME_CARPETA:
                descargados += descargar_carpeta_drive(servicio, item["id"], ruta_local)
            elif item["mimeType"] in MIME_EXPORTAR:
                _exportar_google_doc(servicio, item["id"], item["mimeType"], ruta_local)
                descargados += 1
            elif item["mimeType"].startswith("application/vnd.google-apps."):
                logging.warning("   (se omite '%s': tipo de Google no descargable directo)", item["name"])
            else:
                _descargar_archivo_binario(servicio, item["id"], ruta_local)
                descargados += 1
        page_token = respuesta.get("nextPageToken")
        if not page_token:
            break
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


def buscar_en_correo(usuario: str, app_password: str, termino: str):
    """
    Busca en TODO el correo (no solo la bandeja de entrada) mensajes que
    mencionen 'termino', usando la busqueda propia de Gmail (X-GM-RAW --
    lo mismo que escribirlo en la barra de busqueda de Gmail). Devuelve
    [(asunto, {enlaces_de_drive}, [(nombre_zip, bytes), ...]), ...].
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
            resultados.append((asunto, enlaces_drive, adjuntos_zip))
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


def descargar_coincidencia(servicio, item, numero: str, radicado: str, motivo: str, confiable: bool,
                            descargas_a_validar: list, ya_descargados: set) -> bool:
    """
    Descarga (o simula) la carpeta de 'item' como 'numero. radicado' (o
    'numero. radicado_2', etc, si ya se descargo otro candidato para
    este mismo proceso) dentro de CARPETA_PROCESOS. 'motivo' describe
    como se encontro (ej. "radicado completo", "radicado corto: 2025-456",
    "cuenta: 1000111"). Si 'confiable' es False (coincidencia por
    radicado corto, cuenta, o enlace de correo sin radicado completo),
    se agrega a 'descargas_a_validar' para el reporte aparte.

    'ya_descargados' es un set con los ID de carpeta ya bajados para
    ESTE MISMO proceso en esta misma corrida -- la busqueda por radicado
    corto/cuenta/correo puede encontrar la MISMA carpeta de Drive varias
    veces (ej. por dos formatos distintos del radicado corto, ya que la
    busqueda de Drive es por texto aproximado, no exacta); si la carpeta
    ya se descargo, se omite en vez de volver a bajar los mismos
    archivos otra vez. Devuelve True si quedo lista (o se simulo, o ya
    estaba descargada de una busqueda anterior).
    """
    carpeta = carpeta_contenedora(servicio, item)
    if not carpeta:
        logging.warning(
            "[Sin coincidencia] Proceso %s (radicado %s): se encontro '%s' pero no se pudo determinar su "
            "carpeta contenedora en Drive.",
            numero, radicado, item.get("name"),
        )
        return False

    if carpeta["id"] in ya_descargados:
        logging.info(
            "   (la carpeta '%s' ya se habia descargado para el proceso %s por otra busqueda; se omite duplicado)",
            carpeta["name"], numero,
        )
        return True
    ya_descargados.add(carpeta["id"])

    nombre_destino = f"{numero}. {radicado}"
    destino = cruce_excel.ruta_libre(Path(CARPETA_PROCESOS), nombre_destino)
    etiqueta = "" if confiable else " -- A VALIDAR (coincidencia no exacta)"

    if MODO_PRUEBA:
        logging.info(
            "[SIMULACION%s] Proceso %s (radicado %s, %s): se descargaria la carpeta de Drive '%s' (%s) como '%s'.",
            etiqueta, numero, radicado, motivo, carpeta["name"], enlace_de(carpeta), destino.name,
        )
        if not confiable:
            descargas_a_validar.append((numero, radicado, motivo, carpeta["name"], enlace_de(carpeta), destino.name))
        return True

    archivos = descargar_carpeta_drive(servicio, carpeta["id"], destino)
    cruce_excel.aplanar_carpeta_anidada_unica(destino)
    logging.info(
        "[Descargado%s] Proceso %s (radicado %s, %s): '%s' (%s) -> '%s' (%d archivo(s)).",
        etiqueta, numero, radicado, motivo, carpeta["name"], enlace_de(carpeta), destino.name, archivos,
    )
    if not confiable:
        descargas_a_validar.append((numero, radicado, motivo, carpeta["name"], enlace_de(carpeta), destino.name))
    return True


def procesar_faltante(servicio, credenciales_correo, fila, descargas_a_validar: list):
    numero, cuenta, radicado, juzgado = fila["numero"], fila["cuenta"], fila["radicado"], fila["juzgado"]

    # Carpetas de Drive ya descargadas para ESTE proceso en esta corrida
    # (para no bajar la misma carpeta dos veces si varias busquedas la
    # encuentran -- ver descargar_coincidencia).
    ya_descargados = set()

    if servicio and radicado:
        coincidencias = buscar_en_drive(servicio, radicado)
        carpetas = [c for c in coincidencias if c["mimeType"] == MIME_CARPETA]
        objetivo = carpetas[0] if carpetas else (coincidencias[0] if coincidencias else None)
        if objetivo and descargar_coincidencia(
            servicio, objetivo, numero, radicado, "radicado completo", True, descargas_a_validar, ya_descargados
        ):
            return

    # No hubo coincidencia exacta: se descargan TODOS los candidatos que
    # aparezcan por radicado corto o cuenta (cada uno en su propia
    # carpeta, sin pisar nada), pero marcados para validar despues.
    if servicio:
        for corto in radicados_cortos(radicado):
            for c in buscar_en_drive(servicio, corto):
                descargar_coincidencia(
                    servicio, c, numero, radicado, f"radicado corto: {corto}", False, descargas_a_validar, ya_descargados
                )

    if servicio and cuenta:
        for c in buscar_en_drive(servicio, cuenta):
            descargar_coincidencia(
                servicio, c, numero, radicado, f"cuenta: {cuenta}", False, descargas_a_validar, ya_descargados
            )

    if credenciales_correo and BUSCAR_EN_CORREO:
        usuario, app_password = credenciales_correo
        terminos = [radicado] + radicados_cortos(radicado) + ([cuenta] if cuenta else [])
        for termino in terminos:
            if not termino:
                continue
            try:
                correos = buscar_en_correo(usuario, app_password, termino)
            except Exception as error:
                logging.error("[Correo] Fallo buscando '%s': %s", termino, error)
                continue
            confiable = termino == radicado
            for asunto, enlaces, adjuntos in correos:
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
                            descargas_a_validar, ya_descargados,
                        )
                for nombre_zip, contenido in adjuntos:
                    _organizar_adjunto_zip(numero, radicado, asunto, termino, nombre_zip, contenido, confiable, descargas_a_validar)


def _organizar_adjunto_zip(numero, radicado, asunto, termino, nombre_zip, contenido: bytes, confiable: bool,
                            descargas_a_validar: list):
    nombre_destino = f"{numero}. {radicado}"
    destino = cruce_excel.ruta_libre(Path(CARPETA_PROCESOS), nombre_destino)
    motivo = "radicado completo" if confiable else f"correo ({termino}): {asunto}"
    etiqueta = "" if confiable else " -- A VALIDAR (coincidencia no exacta)"

    if MODO_PRUEBA:
        logging.info(
            "[SIMULACION%s] Proceso %s (radicado %s, %s): se extraeria el adjunto '%s' del correo '%s' como '%s'.",
            etiqueta, numero, radicado, motivo, nombre_zip, asunto, destino.name,
        )
        if not confiable:
            descargas_a_validar.append((numero, radicado, motivo, nombre_zip, "(adjunto de correo)", destino.name))
        return

    destino.mkdir(parents=True, exist_ok=True)
    ruta_zip_temp = destino.parent / f"_tmp_{nombre_zip}"
    ruta_zip_temp.write_bytes(contenido)
    try:
        extraidos, fallidos = cruce_excel.extraer_zip_en_carpeta(ruta_zip_temp, destino)
        logging.info(
            "[Descargado%s] Proceso %s (radicado %s, %s): adjunto '%s' del correo '%s' -> '%s' (%d archivo(s)%s).",
            etiqueta, numero, radicado, motivo, nombre_zip, asunto, destino.name, extraidos,
            f", {fallidos} fallidos" if fallidos else "",
        )
        if not confiable:
            descargas_a_validar.append((numero, radicado, motivo, nombre_zip, "(adjunto de correo)", destino.name))
    finally:
        ruta_zip_temp.unlink(missing_ok=True)


def procesar():
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
        procesar_faltante(servicio, credenciales_correo, fila, descargas_a_validar)

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
