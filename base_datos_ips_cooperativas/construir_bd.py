"""
Base de datos de IPS y cooperativas por departamento.

Descarga los registros oficiales publicados en datos abiertos de Colombia
(www.datos.gov.co) y arma una base de datos con:

  - Todas las IPS (Instituciones Prestadoras de Servicios de Salud)
    inscritas en el REPS del Ministerio de Salud, con todas sus sedes.
  - Todas las cooperativas (y precooperativas) del listado de entidades
    del sector solidario de la Supersolidaria.
  - Las IPS que a su vez son cooperativas (cruce por NIT con la
    Supersolidaria, o por razon social).

para Antioquia, Atlantico, Bolivar, Cundinamarca y Norte de Santander.

Resultado (en la carpeta "salida"):
  - ips_cooperativas.sqlite : base de datos SQLite con todas las tablas.
  - IPS_y_Cooperativas.xlsx : el mismo contenido en Excel, una hoja por tabla.
  - csv/                    : cada tabla en CSV (separado por ";").
  - fuentes/                : los archivos originales descargados.
"""

import argparse
import csv
import datetime
import http.client
import io
import re
import sqlite3
import sys
import time
import unicodedata
import urllib.error
import urllib.request
from pathlib import Path

# ============================= CONFIGURACION =============================

# Departamentos a incluir: nombre -> codigo DANE.
DEPARTAMENTOS = {
    "ANTIOQUIA": "05",
    "ATLÁNTICO": "08",
    "BOLÍVAR": "13",
    "CUNDINAMARCA": "25",
    "NORTE DE SANTANDER": "54",
}

# Bogota D.C. es un distrito aparte de Cundinamarca en todos los registros
# oficiales; solo se incluye si se ejecuta con --incluir-bogota.
BOGOTA = ("BOGOTÁ D.C.", "11")

FUENTES = {
    "reps": {
        "titulo": "Registro Especial de Prestadores y Sedes de Servicios "
                  "de Salud (REPS) - Ministerio de Salud",
        "dataset": "c36g-9fc2",
    },
    "solidarias": {
        "titulo": "Listado de Entidades del Sector Solidario - "
                  "Superintendencia de la Economía Solidaria",
        "dataset": "kg2d-yfyg",
    },
}
URL_DESCARGA = ("https://www.datos.gov.co/api/views/{dataset}/rows.csv"
                "?accessType=DOWNLOAD")

CARPETA_SALIDA = Path(__file__).resolve().parent / "salida"
NOMBRE_SQLITE = "ips_cooperativas.sqlite"
NOMBRE_EXCEL = "IPS_y_Cooperativas.xlsx"
NOMBRE_CORREOS = "Correos_IPS.xlsx"

INTENTOS_DESCARGA = 4
TIMEOUT_DESCARGA_SEGUNDOS = 300

# ===========================================================================

# Cooperativa, cooperativas, precooperativa, organismo cooperativo, "COOP."
# No reconoce "cooperativismo", para no confundir las instituciones
# auxiliares del cooperativismo con cooperativas.
PATRON_COOPERATIVA = re.compile(r"\b(PRE)?COOPERATIV(A|AS|O|OS)\b|\bCOOP\b")

# Tipos de entidad de la Supersolidaria que son cooperativas (multiactivas,
# especializadas, integrales, de trabajo asociado, precooperativas...) y
# tipos que nunca lo son, aunque su nombre lo diga.
TIPOS_COOPERATIVOS = re.compile(
    r"MULTIACTIVA|ESPECIALIZADA (DE|SIN)|INTEGRAL|TRABAJO ASOCIADO|"
    r"PRECOOPERATIVA|APORTES Y CREDITO|COOPERATIV(A|AS|O|OS)\b|"
    r"CARACTER ECONOMICO")
TIPOS_NO_COOPERATIVOS = re.compile(r"FONDO|MUTUAL")

# Un reporte a la Supersolidaria en los ultimos 12 meses del listado indica
# que la entidad sigue activa.
MESES_REPORTE_ACTIVO = 12

PATRON_IPS = re.compile(r"\bIPS\b|\bINSTITUCION(ES)? PRESTADORA")

MAX_FILAS_EXCEL = 1_048_575

# Correos dentro de un campo del REPS, que puede traer varios separados por
# espacios, comas, barras o guiones ("a@x.com - b@y.com", "a@x.com-b@y.com").
PATRON_CORREO = re.compile(
    r"[a-z0-9._%+\-]+@(?:[a-z0-9](?:[a-z0-9\-]*[a-z0-9])?\.)+[a-z]+(?![a-z0-9])")
CORREO_VALIDO = re.compile(
    r"^[a-z0-9_%+\-]+(\.[a-z0-9_%+\-]+)*"
    r"@(?:[a-z0-9](?:[a-z0-9\-]*[a-z0-9])?\.)+[a-z]{2,}$")

# Errores de digitacion evidentes en dominios de correo gratuito.
DOMINIOS_CORREGIDOS = {
    "gmail.con": "gmail.com", "gmail.c": "gmail.com", "gmail.cm": "gmail.com",
    "gmail.om": "gmail.com", "gmail.co": "gmail.com", "gmail.com.co": "gmail.com",
    "gmial.com": "gmail.com", "gamil.com": "gmail.com", "gmai.com": "gmail.com",
    "hotmail.con": "hotmail.com", "hotmail.c": "hotmail.com",
    "hotmail.cm": "hotmail.com", "hotmail.co": "hotmail.com",
    "hotmailo.com": "hotmail.com", "hotmal.com": "hotmail.com",
    "hotmial.com": "hotmail.com", "hormail.com": "hotmail.com",
    "yahoo.con": "yahoo.com", "outlook.con": "outlook.com",
}


# ------------------------------- Texto -------------------------------------

def sin_tildes(texto):
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def normalizar(texto):
    """Mayusculas, sin tildes y con espacios simples."""
    return re.sub(r"\s+", " ", sin_tildes(str(texto or ""))).strip().upper()


def reparar_texto(texto):
    """Repara texto UTF-8 mal decodificado: 'BOLÃ\x8dVAR' -> 'BOLÍVAR'."""
    if "Ã" not in texto and "Â" not in texto:
        return texto
    try:
        datos = bytes(ord(c) if ord(c) < 256 else c.encode("cp1252")[0]
                      for c in texto)
        return datos.decode("utf-8")
    except (UnicodeError, ValueError):
        return texto


def nombre_columna(encabezado):
    """'NombrePrestador' -> 'nombre_prestador', 'RAZÓN SOCIAL' -> 'razon_social'."""
    texto = sin_tildes(str(encabezado or "").strip())
    texto = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", texto)
    texto = re.sub(r"[^A-Za-z0-9]+", "_", texto).strip("_").lower()
    if texto and texto[0].isdigit():
        texto = "c_" + texto
    return texto or "columna"


def clave(columna):
    """Nombre de columna reducido a letras y numeros, para buscar columnas
    sin depender de guiones bajos, tildes ni mayusculas."""
    return re.sub(r"[^a-z0-9]", "", normalizar(columna).lower())


def unicos(nombres):
    """Agrega _2, _3... a los nombres de columna repetidos."""
    resultado, usados = [], set()
    for nombre in nombres:
        candidato, n = nombre, 1
        while candidato in usados:
            n += 1
            candidato = f"{nombre}_{n}"
        usados.add(candidato)
        resultado.append(candidato)
    return resultado


def normalizar_nit(valor):
    """'891-500-074-3', '891.500.074-3', '8915000743' y '891500074'
    -> '891500074'."""
    texto = str(valor or "").strip()
    if re.fullmatch(r"\d+\.0+", texto):
        texto = texto.split(".")[0]
    digitos = re.sub(r"\D", "", texto)
    # Quita el digito de verificacion: va separado por un guion al final, o
    # pegado a un NIT de persona juridica (que empieza por 8 o 9).
    if re.search(r"\d\s*-\s*\d$", texto) or (
            len(digitos) == 10 and digitos[0] in "89"):
        digitos = digitos[:-1]
    return digitos.lstrip("0")


def clave_fecha(valor):
    """Convierte una fecha de corte en una tupla comparable (anio, mes, dia)."""
    texto = str(valor or "").strip()
    m = re.match(r"(\d{4})[-/](\d{1,2})(?:[-/](\d{1,2}))?", texto)
    if m:
        return int(m[1]), int(m[2]), int(m[3] or 0)
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", texto)
    if m:
        a, b, anio = int(m[1]), int(m[2]), int(m[3])
        # datos.gov.co exporta las fechas como MM/DD/AAAA.
        mes, dia = (b, a) if a > 12 else (a, b)
        return anio, mes, dia
    m = re.fullmatch(r"(\d{4})(\d{2})?(\d{2})?", texto)
    if m:
        return int(m[1]), int(m[2] or 0), int(m[3] or 0)
    m = re.search(r"\b([A-Za-z]{3})[a-z]*\.? (\d{1,2}),? (\d{4})", texto)
    if m and normalizar(m[1]) in MESES:
        return int(m[3]), MESES[normalizar(m[1])], int(m[2])
    return 0, 0, 0


MESES = {"JAN": 1, "ENE": 1, "FEB": 2, "MAR": 3, "APR": 4, "ABR": 4,
         "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8, "AGO": 8, "SEP": 9,
         "OCT": 10, "NOV": 11, "DEC": 12, "DIC": 12}


def fecha_iso(clave_de_fecha):
    anio, mes, dia = clave_de_fecha
    if not anio:
        return ""
    return f"{anio:04d}-{mes:02d}-{dia:02d}" if dia else f"{anio:04d}-{mes:02d}"


def restar_meses(clave_de_fecha, meses):
    anio, mes, dia = clave_de_fecha
    total = anio * 12 + (mes - 1) - meses
    return total // 12, total % 12 + 1, dia


# ---------------------------- Departamentos --------------------------------

def construir_objetivos(incluir_bogota=False):
    objetivos = dict(DEPARTAMENTOS)
    if incluir_bogota:
        objetivos[BOGOTA[0]] = BOGOTA[1]
    return objetivos


def departamento_de(valor, objetivos):
    """Devuelve el departamento de `objetivos` al que corresponde `valor`
    (nombre o codigo DANE de departamento o municipio), o None."""
    texto = re.sub(r"[^A-Z0-9 ]", " ", normalizar(valor))
    letras = re.sub(r"\s+", " ", re.sub(r"\d+", " ", texto)).strip()

    if not letras:
        digitos = texto.replace(" ", "")
        if not digitos:
            return None
        if len(digitos) <= 2:
            codigo = digitos.zfill(2)
        elif len(digitos) in (4, 5):
            codigo = digitos.zfill(5)[:2]
        elif len(digitos) in (7, 8):
            codigo = digitos.zfill(8)[:2]
        else:
            return None
        for nombre, cod in objetivos.items():
            if cod == codigo:
                return nombre
        return None

    letras = re.sub(r"^(DEPARTAMENTO|DPTO|DEPTO)( DEL?)? ", "", letras)
    if "SANTANDER" in letras and re.match(r"N(ORTE|TE)?\b", letras):
        letras = "NORTE DE SANTANDER"
    elif "BOGOTA" in letras or "DISTRITO CAPITAL" in letras:
        letras = normalizar(BOGOTA[0])
    for nombre in objetivos:
        if normalizar(nombre).replace(".", " ").split() == letras.split() \
                or normalizar(nombre) == letras:
            return nombre
    return None


# ------------------------------ Columnas -----------------------------------

def columnas(encabezados, incluye, excluye=()):
    """Columnas cuyo nombre contiene alguno de `incluye` y ninguno de `excluye`."""
    return [c for c in encabezados
            if any(p in clave(c) for p in incluye)
            and not any(p in clave(c) for p in excluye)]


def buscar_columna(encabezados, patrones, excluye=()):
    """La columna que mejor coincide con `patrones` (en orden de prioridad):
    primero coincidencias exactas, luego parciales."""
    candidatos = [c for c in encabezados
                  if not any(e in clave(c) for e in excluye)]
    for patron in patrones:
        for c in candidatos:
            if clave(c) == patron:
                return c
    for patron in patrones:
        for c in candidatos:
            if patron in clave(c):
                return c
    return None


def columnas_de_ubicacion(encabezados, patrones):
    """Columnas de departamento o municipio. Si el registro trae la ubicacion
    de la sede, se usa solo esa: es donde realmente se prestan los servicios."""
    todas = columnas(encabezados, patrones)
    de_sede = [c for c in todas if "sede" in clave(c)]
    return de_sede or todas


PATRONES_DEPARTAMENTO = ("depa", "dpto", "depto")
PATRONES_MUNICIPIO = ("munic", "mpio", "ciudad")


def columna_nit(encabezados):
    return buscar_columna(
        encabezados,
        ("nit", "numeroidentificacion", "nroidentificacion",
         "numidentificacion", "numerodocumento", "identificacion"),
        ("tipo", "digito", "represent"))


def columna_nombre(encabezados):
    """La razon social de la entidad (no la de la sede ni la del
    representante legal)."""
    return buscar_columna(
        encabezados,
        ("nombreprestador", "razonsocial", "nombreentidad", "nombre",
         "entidad", "sigla"),
        ("represent", "gerente", "contador", "revisor", "contacto", "depa",
         "munic", "ciudad", "sede", "tipo", "clase", "codigo"))


def columna_corte(encabezados):
    """Fecha de corte o del ultimo reporte de cada registro."""
    return buscar_columna(
        encabezados,
        ("fechacorte", "fechaultirepo", "fechaultimoreporte", "ultimoreporte",
         "fechareporte", "corte", "periodo"))


def corte_de_los_datos(encabezados, filas):
    """La fecha de corte mas reciente del registro, para la hoja Fuentes."""
    col = columna_corte(encabezados)
    valores = {f[col] for f in filas if f[col]} if col else set()
    if not valores:
        return ""
    mas_reciente = max(valores, key=clave_fecha)
    return fecha_iso(clave_fecha(mas_reciente)) or mas_reciente


def ubicar(fila, cols_depto, cols_muni, objetivos):
    """Devuelve (departamento objetivo o None, municipio)."""
    departamento = None
    if cols_depto:
        for c in cols_depto:
            departamento = departamento_de(fila.get(c), objetivos)
            if departamento:
                break
    else:
        # Sin columna de departamento solo sirve el codigo DANE del
        # municipio: hay municipios llamados igual que un departamento
        # (Bolivar, en Cauca, Santander y Valle).
        for c in cols_muni:
            if fila.get(c, "").strip().isdigit():
                departamento = departamento_de(fila[c], objetivos)
                if departamento:
                    break
    municipio = ""
    for c in cols_muni:
        valor = fila.get(c, "")
        if re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]", valor):
            municipio = normalizar(valor)
            break
    return departamento, municipio


# --------------------------- Lectura de archivos ---------------------------

def leer_tabla(ruta):
    """Lee un CSV o XLSX y devuelve (columnas, filas como diccionarios)."""
    ruta = Path(ruta)
    if ruta.suffix.lower() in (".xlsx", ".xlsm"):
        encabezados, filas = leer_xlsx(ruta)
    else:
        encabezados, filas = leer_csv(ruta)
    cols = unicos([nombre_columna(reparar_texto(h)) for h in encabezados])
    datos = []
    for fila in filas:
        fila = [reparar_texto(v) for v in fila] + [""] * (len(cols) - len(fila))
        datos.append(dict(zip(cols, fila)))
    return cols, datos


def leer_csv(ruta):
    csv.field_size_limit(min(sys.maxsize, 2**31 - 1))
    try:
        texto = Path(ruta).read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        texto = Path(ruta).read_text(encoding="latin-1")
    primera_linea = texto.split("\n", 1)[0]
    separador = max(",;\t|", key=primera_linea.count)
    lector = csv.reader(io.StringIO(texto), delimiter=separador)
    encabezados = [h.strip() for h in next(lector, [])]
    filas = [[v.strip() for v in fila] for fila in lector
             if any(v.strip() for v in fila)]
    return encabezados, filas


def leer_xlsx(ruta):
    from openpyxl import load_workbook

    def texto(valor):
        if valor is None:
            return ""
        if isinstance(valor, float) and valor.is_integer():
            valor = int(valor)
        return str(valor).strip()

    libro = load_workbook(ruta, read_only=True, data_only=True)
    encabezados, filas = None, []
    for fila in libro.worksheets[0].iter_rows(values_only=True):
        valores = [texto(v) for v in fila]
        if encabezados is None:
            # Salta titulos o filas en blanco antes de los encabezados.
            if sum(1 for v in valores if v) >= 3:
                encabezados = valores
        elif any(valores):
            filas.append(valores)
    libro.close()
    return encabezados or [], filas


# -------------------------------- Descarga ---------------------------------

def descargar(dataset, destino, app_token=None):
    url = URL_DESCARGA.format(dataset=dataset)
    cabeceras = {"User-Agent": "base-datos-ips-cooperativas/1.0"}
    if app_token:
        cabeceras["X-App-Token"] = app_token
    parcial = destino.with_name(destino.name + ".parcial")
    for intento in range(1, INTENTOS_DESCARGA + 1):
        try:
            peticion = urllib.request.Request(url, headers=cabeceras)
            with urllib.request.urlopen(
                    peticion, timeout=TIMEOUT_DESCARGA_SEGUNDOS) as respuesta, \
                    open(parcial, "wb") as archivo:
                total, aviso = 0, 5 * 2**20
                while bloque := respuesta.read(2**20):
                    archivo.write(bloque)
                    total += len(bloque)
                    if total >= aviso:
                        print(f"    {total / 2**20:.0f} MB descargados...",
                              flush=True)
                        aviso += 5 * 2**20
            parcial.replace(destino)
            return url
        except (OSError, http.client.HTTPException) as error:
            print(f"    Intento {intento} de {INTENTOS_DESCARGA} fallido: {error}")
            if intento < INTENTOS_DESCARGA:
                time.sleep(2 ** intento)
    parcial.unlink(missing_ok=True)
    raise RuntimeError(f"no se pudo descargar {url}")


def obtener_fuente(nombre, archivo_local, carpeta_fuentes, app_token):
    """Devuelve (ruta del archivo, de donde salio)."""
    if archivo_local:
        ruta = Path(archivo_local)
        if not ruta.exists():
            raise RuntimeError(f"no existe el archivo {ruta}")
        return ruta, f"Archivo local: {ruta.name}"

    dataset = FUENTES[nombre]["dataset"]
    destino = carpeta_fuentes / f"{nombre}_{dataset}.csv"
    print(f"  Descargando {FUENTES[nombre]['titulo']}...", flush=True)
    try:
        url = descargar(dataset, destino, app_token)
        return destino, url
    except RuntimeError as error:
        if not destino.exists():
            raise
        fecha = datetime.datetime.fromtimestamp(destino.stat().st_mtime)
        print(f"  AVISO: {error}. Se usa la copia descargada el "
              f"{fecha:%Y-%m-%d %H:%M}.")
        return destino, f"Copia descargada el {fecha:%Y-%m-%d %H:%M}"


# ------------------------------ Procesamiento ------------------------------

def procesar_solidarias(encabezados, filas, objetivos):
    """Devuelve (entidades en los departamentos, NITs de las cooperativas de
    todo el pais)."""
    cols_depto = columnas_de_ubicacion(encabezados, PATRONES_DEPARTAMENTO)
    cols_muni = columnas_de_ubicacion(encabezados, PATRONES_MUNICIPIO)
    if not cols_depto and not cols_muni:
        raise RuntimeError(
            "el listado de la Supersolidaria no tiene columna de departamento "
            f"ni de municipio. Columnas encontradas: {', '.join(encabezados)}")
    col_nit = columna_nit(encabezados)
    col_nombre = columna_nombre(encabezados)
    cols_tipo = columnas(
        encabezados, ("tipo", "clase", "naturaleza", "organizacion"),
        ("identific", "document", "nit", "reporte", "nivel"))
    col_corte = columna_corte(encabezados)

    # El listado trae un registro por cada reporte de la entidad; deja el
    # mas reciente.
    if col_nit and col_corte:
        ultimas, sin_nit = {}, []
        for fila in filas:
            nit = normalizar_nit(fila[col_nit])
            if not nit:
                sin_nit.append(fila)
            elif nit not in ultimas or (clave_fecha(fila[col_corte])
                                        >= clave_fecha(ultimas[nit][col_corte])):
                ultimas[nit] = fila
        filas = list(ultimas.values()) + sin_nit

    if col_corte:
        ultimo_corte = max((clave_fecha(f[col_corte]) for f in filas),
                           default=(0, 0, 0))
        corte_activo = restar_meses(ultimo_corte, MESES_REPORTE_ACTIVO)

    nits_cooperativas = set()
    entidades = []
    for fila in filas:
        nombre = normalizar(fila.get(col_nombre, "")) if col_nombre else ""
        tipo = normalizar(" ".join(fila[c] for c in cols_tipo))
        cooperativa = not TIPOS_NO_COOPERATIVOS.search(tipo) and bool(
            TIPOS_COOPERATIVOS.search(tipo) or PATRON_COOPERATIVA.search(nombre))
        if cooperativa and col_nit:
            nit = normalizar_nit(fila[col_nit])
            if nit:
                nits_cooperativas.add(nit)

        departamento, municipio = ubicar(fila, cols_depto, cols_muni,
                                         objetivos)
        if departamento:
            calculadas = {
                "bd_departamento": departamento,
                "bd_municipio": municipio,
                "bd_es_cooperativa": "SI" if cooperativa else "NO",
            }
            if col_corte:
                reporte = clave_fecha(fila[col_corte])
                calculadas["bd_ultimo_reporte"] = fecha_iso(reporte)
                calculadas["bd_reporta_actualmente"] = (
                    "SI" if reporte[0] and reporte >= corte_activo else "NO")
            entidades.append({**calculadas, **fila})
    return entidades, nits_cooperativas


def procesar_reps(encabezados, filas, objetivos, nits_cooperativas):
    """Devuelve las sedes de prestadores ubicadas en los departamentos."""
    cols_depto = columnas_de_ubicacion(encabezados, PATRONES_DEPARTAMENTO)
    cols_muni = columnas_de_ubicacion(encabezados, PATRONES_MUNICIPIO)
    if not cols_depto and not cols_muni:
        raise RuntimeError(
            "el REPS no tiene columna de departamento ni de municipio. "
            f"Columnas encontradas: {', '.join(encabezados)}")
    col_clase = buscar_columna(encabezados, ("claseprestador", "clase"))
    if not col_clase:
        print("  AVISO: el REPS no trae la clase de prestador; todos los "
              "prestadores se tratan como IPS.")
    col_nit = columna_nit(encabezados)
    col_nombre = columna_nombre(encabezados)

    sedes = []
    for fila in filas:
        departamento, municipio = ubicar(fila, cols_depto, cols_muni,
                                         objetivos)
        if not departamento:
            continue

        if col_clase:
            clase = normalizar(fila[col_clase])
            ips = bool(PATRON_IPS.search(clase)) or clase == "1"
        else:
            ips = True

        nit = normalizar_nit(fila[col_nit]) if col_nit else ""
        nombre = normalizar(fila.get(col_nombre, "")) if col_nombre else ""
        if nit and nit in nits_cooperativas:
            criterio = "NIT registrado como cooperativa en la Supersolidaria"
        elif PATRON_COOPERATIVA.search(nombre):
            criterio = "Razón social de cooperativa"
        else:
            criterio = ""

        sedes.append({
            "bd_departamento": departamento,
            "bd_municipio": municipio,
            "bd_es_ips": "SI" if ips else "NO",
            "bd_es_cooperativa": "SI" if criterio else "NO",
            "bd_criterio_cooperativa": criterio,
            **fila,
        })
    return sedes


def identificador_ips(encabezados):
    """Funcion que da la clave de la IPS de una sede: su NIT. El REPS inscribe
    a una misma IPS una vez por cada municipio o distrito donde presta
    servicios, cada vez con otro codigo de prestador, asi que el codigo no
    sirve para reconocerla."""
    col_codigo = buscar_columna(encabezados, ("codigoprestador", "codprestador"))
    col_nit = columna_nit(encabezados)
    col_nombre = columna_nombre(encabezados)

    def identificador(sede):
        return ((col_nit and normalizar_nit(sede[col_nit]))
                or (col_codigo and sede[col_codigo])
                or (col_nombre and normalizar(sede[col_nombre]))
                or id(sede))
    return identificador


def agrupar_ips(encabezados, sedes):
    """Una fila por IPS, con todas sus sedes en los departamentos. Los datos
    son los de la sede principal de su inscripcion con mas sedes."""
    col_codigo = buscar_columna(encabezados, ("codigoprestador", "codprestador"))
    col_principal = buscar_columna(encabezados, ("sedeprincipal", "principal"))
    col_sede = buscar_columna(encabezados, ("codigohabilitacionsede", "codigosede"))
    identificador = identificador_ips(encabezados)

    def es_principal(sede):
        if col_principal:
            return normalizar(sede[col_principal]) in ("SI", "S", "1", "TRUE", "X")
        # El REPS numera la sede principal con el codigo del prestador + "01".
        return bool(col_codigo and col_sede
                    and sede[col_sede] == sede[col_codigo] + "01")

    grupos = {}
    for sede in sedes:
        if sede["bd_es_ips"] == "SI":
            grupos.setdefault(identificador(sede), []).append(sede)

    ips = []
    for grupo in grupos.values():
        inscripciones = {}
        for sede in grupo:
            inscripciones.setdefault(sede[col_codigo] if col_codigo else "", []).append(sede)
        codigo = min(inscripciones, key=lambda c: (-len(inscripciones[c]), c))
        candidatas = inscripciones[codigo]
        if col_sede:
            candidatas.sort(key=lambda sede: sede[col_sede])
        principal = next((sede for sede in candidatas if es_principal(sede)),
                         candidatas[0])

        fila = dict(principal)
        fila["bd_departamentos"] = ", ".join(
            sorted({sede["bd_departamento"] for sede in grupo}, key=normalizar))
        fila["bd_num_sedes"] = len(grupo)
        fila["bd_inscripciones_reps"] = len(inscripciones)
        cooperativa = next((sede for sede in grupo
                            if sede["bd_es_cooperativa"] == "SI"), None)
        if cooperativa:
            fila["bd_es_cooperativa"] = "SI"
            fila["bd_criterio_cooperativa"] = cooperativa["bd_criterio_cooperativa"]
        ips.append(fila)
    return ips


def extraer_correos(valor):
    """Correos validos de un campo de correo del REPS, como (correo,
    observacion). Separa los que vienen juntos, quita tildes y corrige
    errores evidentes de dominio, dejando anotado lo que cambio."""
    original = re.sub(r"\s*@\s*", "@", str(valor or "").strip().lower())
    texto = sin_tildes(original)
    correos = []
    for candidato in PATRON_CORREO.findall(texto):
        usuario, dominio = candidato.rsplit("@", 1)
        usuario = usuario.strip(".-")
        notas = []
        if candidato not in original:
            notas.append("se le quitó una tilde")
        if dominio in DOMINIOS_CORREGIDOS:
            notas.append(f"dominio corregido, decía {dominio}")
            dominio = DOMINIOS_CORREGIDOS[dominio]
        correo = f"{usuario}@{dominio}"
        if CORREO_VALIDO.match(correo) and correo not in (c for c, _ in correos):
            correos.append((correo, "; ".join(notas)))
    return correos


def lista_de_correos(encabezados, sedes_ips):
    """Un registro por correo, sin repetidos, para envios masivos, y los
    campos de correo que no se pudieron leer."""
    cols_correo = columnas(encabezados, ("email", "correo"))
    col_nit = columna_nit(encabezados)
    col_nombre = columna_nombre(encabezados)
    identificador = identificador_ips(encabezados)
    cooperativas = {identificador(s) for s in sedes_ips
                    if s["bd_es_cooperativa"] == "SI"}

    correos, ilegibles, ips_con_correo = {}, [], set()
    for sede in sedes_ips:
        ips = identificador(sede)
        nombre = sede.get(col_nombre, "") if col_nombre else ""
        nit = sede.get(col_nit, "") if col_nit else ""
        for col in cols_correo:
            valor = sede.get(col, "").strip()
            encontrados = extraer_correos(valor)
            if valor and not encontrados:
                ilegibles.append((ips, nombre, nit, sede, col, valor))
            for correo, nota in encontrados:
                ips_con_correo.add(ips)
                datos = correos.setdefault(correo, {
                    "ips": {}, "departamentos": set(), "municipios": set(),
                    "principal": False, "notas": set()})
                datos["ips"].setdefault(ips, (nombre, nit))
                datos["departamentos"].add(sede["bd_departamento"])
                datos["municipios"].add(sede["bd_municipio"])
                datos["principal"] |= "sede" not in clave(col)
                if nota:
                    datos["notas"].add(nota)

    filas = []
    for correo, datos in correos.items():
        filas.append({
            "Correo": correo,
            "IPS": " / ".join(nombre for nombre, _ in datos["ips"].values()),
            "NIT": " / ".join(nit for _, nit in datos["ips"].values()),
            "Es cooperativa": "SI" if cooperativas & datos["ips"].keys() else "NO",
            "Departamentos": ", ".join(sorted(datos["departamentos"], key=normalizar)),
            "Municipios": ", ".join(sorted(datos["municipios"])),
            "Tipo de correo": ("Principal de la IPS" if datos["principal"]
                               else "De una sede"),
            "Observación": "; ".join(sorted(datos["notas"])),
        })
    filas.sort(key=lambda f: (normalizar(f["Departamentos"]), normalizar(f["IPS"]),
                              f["Tipo de correo"] != "Principal de la IPS",
                              f["Correo"]))

    revisar = []
    for ips, nombre, nit, sede, col, valor in ilegibles:
        revisar.append({
            "IPS": nombre, "NIT": nit,
            "Departamento": sede["bd_departamento"],
            "Municipio": sede["bd_municipio"],
            "Campo": col, "Valor en el REPS": valor,
            "La IPS tiene otro correo válido": "SI" if ips in ips_con_correo else "NO",
        })
    return filas, revisar


def ordenar(filas, encabezados):
    """Por departamento, municipio y razon social."""
    col_nombre = columna_nombre(encabezados)
    return sorted(filas, key=lambda f: (
        f["bd_departamento"], f["bd_municipio"],
        normalizar(f.get(col_nombre, "")) if col_nombre else ""))


def resumen(objetivos, sedes, entidades, identificador):
    """Totales por departamento. Una IPS con sedes en varios departamentos
    cuenta en cada uno, pero una sola vez en el total."""
    def contar_ips(departamento=None, **condiciones):
        return len({identificador(f) for f in sedes
                    if f["bd_es_ips"] == "SI"
                    and departamento in (None, f["bd_departamento"])
                    and all(f.get(k) == v for k, v in condiciones.items())})

    filas = []
    for departamento in sorted(objetivos, key=normalizar):
        def contar(lista, **condiciones):
            return sum(1 for f in lista if f["bd_departamento"] == departamento
                       and all(f.get(k) == v for k, v in condiciones.items()))
        filas.append({
            "departamento": departamento,
            "ips": contar_ips(departamento),
            "sedes_de_ips": contar(sedes, bd_es_ips="SI"),
            "ips_cooperativas": contar_ips(departamento, bd_es_cooperativa="SI"),
            "otras_sedes_reps_no_ips": contar(sedes, bd_es_ips="NO"),
            "cooperativas": contar(entidades, bd_es_cooperativa="SI"),
            "cooperativas_que_reportan_actualmente": contar(
                entidades, bd_es_cooperativa="SI", bd_reporta_actualmente="SI"),
            "otras_entidades_solidarias": contar(entidades, bd_es_cooperativa="NO"),
        })
    total = {"departamento": "TOTAL"}
    for clave_total in filas[0]:
        if clave_total != "departamento":
            total[clave_total] = sum(f[clave_total] for f in filas)
    total["ips"] = contar_ips()
    total["ips_cooperativas"] = contar_ips(bd_es_cooperativa="SI")
    return filas + [total]


# -------------------------------- Guardado ---------------------------------

def columnas_de(encabezados_originales, calculadas):
    """Primero las columnas calculadas (bd_...), luego las del registro."""
    return calculadas + [c for c in encabezados_originales if c not in calculadas]


def guardar_sqlite(ruta, tablas):
    temporal = ruta.with_name(ruta.name + ".tmp")
    temporal.unlink(missing_ok=True)
    con = sqlite3.connect(temporal)
    try:
        for nombre, (cols, filas) in tablas.items():
            con.execute(f'CREATE TABLE "{nombre}" ('
                        + ", ".join(f'"{c}"' for c in cols) + ")")
            con.executemany(
                f'INSERT INTO "{nombre}" VALUES ({", ".join("?" * len(cols))})',
                ([f.get(c, "") for c in cols] for f in filas))
            if "bd_departamento" in cols:
                con.execute(f'CREATE INDEX "ix_{nombre}_departamento" '
                            f'ON "{nombre}" (bd_departamento)')
        con.commit()
    finally:
        con.close()
    return reemplazar(temporal, ruta)


def guardar_excel(ruta, hojas):
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    libro = Workbook(write_only=True)
    for titulo, cols, filas in hojas:
        if len(filas) > MAX_FILAS_EXCEL:
            print(f"  AVISO: la hoja {titulo} tiene {len(filas)} filas; Excel "
                  f"solo admite {MAX_FILAS_EXCEL}. El resto esta en el "
                  "SQLite y en el CSV.")
            filas = filas[:MAX_FILAS_EXCEL]
        hoja = libro.create_sheet(title=titulo[:31])
        hoja.freeze_panes = "A2"
        for i, c in enumerate(cols, start=1):
            ancho = max([len(c)] + [len(str(f.get(c, ""))) for f in filas[:300]])
            hoja.column_dimensions[get_column_letter(i)].width = min(ancho + 2, 60)
        if filas:
            hoja.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(filas) + 1}"

        encabezado = []
        for c in cols:
            celda = WriteOnlyCell(hoja, value=c)
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill("solid", fgColor="1F4E78")
            encabezado.append(celda)
        hoja.append(encabezado)

        for fila in filas:
            valores = []
            for c in cols:
                valor = fila.get(c, "")
                if isinstance(valor, str):
                    valor = ILLEGAL_CHARACTERS_RE.sub("", valor)
                    if valor.startswith("="):
                        # Que Excel lo muestre como texto y no como formula.
                        celda = WriteOnlyCell(hoja, value=valor)
                        celda.data_type = "s"
                        valor = celda
                valores.append(valor)
            hoja.append(valores)

    temporal = ruta.with_name("~tmp_" + ruta.name)
    libro.save(temporal)
    return reemplazar(temporal, ruta)


def guardar_csv(carpeta, tablas):
    carpeta.mkdir(parents=True, exist_ok=True)
    for nombre, (cols, filas) in tablas.items():
        temporal = carpeta / f"{nombre}.csv.tmp"
        with open(temporal, "w", encoding="utf-8-sig", newline="") as archivo:
            escritor = csv.writer(archivo, delimiter=";")
            escritor.writerow(cols)
            escritor.writerows([f.get(c, "") for c in cols] for f in filas)
        reemplazar(temporal, carpeta / f"{nombre}.csv")


def reemplazar(temporal, destino):
    """Mueve el archivo nuevo a su lugar. Si el anterior esta abierto (por
    ejemplo en Excel), guarda el nuevo con la fecha en el nombre."""
    try:
        temporal.replace(destino)
        return destino
    except PermissionError:
        alterno = destino.with_name(
            f"{destino.stem}_{datetime.datetime.now():%Y%m%d_%H%M%S}{destino.suffix}")
        temporal.replace(alterno)
        print(f"  AVISO: {destino.name} esta abierto en otro programa; se "
              f"guardo como {alterno.name}.")
        return alterno


# ---------------------------------- Main -----------------------------------

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Construye la base de datos de IPS y cooperativas de "
                    "Antioquia, Atlántico, Bolívar, Cundinamarca y Norte de "
                    "Santander a partir de los registros oficiales.")
    parser.add_argument("--salida", default=str(CARPETA_SALIDA),
                        help="carpeta donde se guarda el resultado")
    parser.add_argument("--incluir-bogota", action="store_true",
                        help="incluir también Bogotá D.C. (es un distrito "
                             "aparte de Cundinamarca)")
    parser.add_argument("--archivo-reps",
                        help="usar un CSV/XLSX del REPS ya descargado en vez "
                             "de descargarlo")
    parser.add_argument("--archivo-solidarias",
                        help="usar un CSV/XLSX del listado de la "
                             "Supersolidaria ya descargado en vez de "
                             "descargarlo")
    parser.add_argument("--app-token",
                        help="token de aplicación de datos.gov.co (opcional, "
                             "evita límites de descarga)")
    args = parser.parse_args(argv)

    salida = Path(args.salida)
    carpeta_fuentes = salida / "fuentes"
    carpeta_fuentes.mkdir(parents=True, exist_ok=True)
    objetivos = construir_objetivos(args.incluir_bogota)
    ahora = datetime.datetime.now()

    print("Departamentos: " + ", ".join(objetivos))
    try:
        print("\n[1/3] Entidades del sector solidario (Supersolidaria)")
        ruta_sol, origen_sol = obtener_fuente(
            "solidarias", args.archivo_solidarias, carpeta_fuentes,
            args.app_token)
        cols_sol, filas_sol = leer_tabla(ruta_sol)
        print(f"  {len(filas_sol)} registros leídos.")
        entidades, nits_cooperativas = procesar_solidarias(
            cols_sol, filas_sol, objetivos)

        print("\n[2/3] Prestadores de servicios de salud (REPS)")
        ruta_reps, origen_reps = obtener_fuente(
            "reps", args.archivo_reps, carpeta_fuentes, args.app_token)
        cols_reps, filas_reps = leer_tabla(ruta_reps)
        print(f"  {len(filas_reps)} registros leídos.")
        sedes = procesar_reps(cols_reps, filas_reps, objetivos,
                              nits_cooperativas)
    except RuntimeError as error:
        print(f"\nERROR: {error}")
        print("Si no hay conexión con www.datos.gov.co, descarga los archivos "
              "a mano (ver README.md) y pásalos con --archivo-reps y "
              "--archivo-solidarias.")
        return 1

    print("\n[3/3] Armando la base de datos")
    ips = agrupar_ips(cols_reps, sedes)
    calc_reps = ["bd_departamento", "bd_municipio", "bd_es_ips",
                 "bd_es_cooperativa", "bd_criterio_cooperativa"]
    cols_sedes = columnas_de(cols_reps, calc_reps)
    cols_ips = columnas_de(cols_reps, calc_reps[:2] + [
        "bd_departamentos", "bd_num_sedes", "bd_inscripciones_reps"]
        + calc_reps[3:])
    calc_sol = ["bd_departamento", "bd_municipio", "bd_es_cooperativa"]
    if columna_corte(cols_sol):
        calc_sol += ["bd_reporta_actualmente", "bd_ultimo_reporte"]
    cols_entidades = columnas_de(cols_sol, calc_sol)

    ips = ordenar(ips, cols_reps)
    sedes = ordenar(sedes, cols_reps)
    entidades = ordenar(entidades, cols_sol)
    cooperativas = [e for e in entidades if e["bd_es_cooperativa"] == "SI"]
    ips_cooperativas = [i for i in ips if i["bd_es_cooperativa"] == "SI"]
    sedes_ips = [s for s in sedes if s["bd_es_ips"] == "SI"]

    filas_resumen = resumen(objetivos, sedes, entidades,
                            identificador_ips(cols_reps))
    cols_resumen = list(filas_resumen[0])
    fuentes = [
        {"fuente": FUENTES["reps"]["titulo"], "dataset": FUENTES["reps"]["dataset"],
         "origen": origen_reps,
         "corte_de_los_datos": corte_de_los_datos(cols_reps, filas_reps),
         "registros_nacionales": len(filas_reps),
         "registros_en_departamentos": len(sedes)},
        {"fuente": FUENTES["solidarias"]["titulo"],
         "dataset": FUENTES["solidarias"]["dataset"], "origen": origen_sol,
         "corte_de_los_datos": corte_de_los_datos(cols_sol, filas_sol),
         "registros_nacionales": len(filas_sol),
         "registros_en_departamentos": len(entidades)},
    ]
    for fuente in fuentes:
        fuente["fecha_construccion"] = f"{ahora:%Y-%m-%d %H:%M}"
    cols_fuentes = list(fuentes[0])

    tablas = {
        "resumen": (cols_resumen, filas_resumen),
        "ips": (cols_ips, ips),
        "ips_sedes": (cols_sedes, sedes_ips),
        "cooperativas": (cols_entidades, cooperativas),
        "ips_cooperativas": (cols_ips, ips_cooperativas),
        "reps_todas_las_sedes": (cols_sedes, sedes),
        "entidades_solidarias": (cols_entidades, entidades),
        "fuentes": (cols_fuentes, fuentes),
    }
    ruta_sqlite = guardar_sqlite(salida / NOMBRE_SQLITE, tablas)
    ruta_excel = guardar_excel(salida / NOMBRE_EXCEL, [
        ("Resumen", cols_resumen, filas_resumen),
        ("IPS", cols_ips, ips),
        ("Sedes IPS", cols_sedes, sedes_ips),
        ("Cooperativas", cols_entidades, cooperativas),
        ("IPS cooperativas", cols_ips, ips_cooperativas),
        ("Fuentes", cols_fuentes, fuentes),
    ])
    guardar_csv(salida / "csv", tablas)
    correos, revisar = lista_de_correos(cols_reps, sedes_ips)
    cols_correos = ["Correo", "IPS", "NIT", "Es cooperativa", "Departamentos",
                    "Municipios", "Tipo de correo", "Observación"]
    ruta_correos = guardar_excel(salida / NOMBRE_CORREOS, [
        ("Correos", cols_correos, correos),
        ("Correos IPS cooperativas", cols_correos,
         [c for c in correos if c["Es cooperativa"] == "SI"]),
        ("Por revisar", ["IPS", "NIT", "Departamento", "Municipio", "Campo",
                         "Valor en el REPS", "La IPS tiene otro correo válido"],
         revisar),
    ])

    print()
    print(f"{'Departamento':<22}{'IPS':>8}{'Sedes IPS':>11}"
          f"{'IPS coop.':>11}{'Cooperativas':>14}{'Coop. activas':>15}")
    for f in filas_resumen:
        print(f"{f['departamento']:<22}{f['ips']:>8}{f['sedes_de_ips']:>11}"
              f"{f['ips_cooperativas']:>11}{f['cooperativas']:>14}"
              f"{f['cooperativas_que_reportan_actualmente']:>15}")
    print(f"\nListo. Resultado en {salida.resolve()}:")
    print(f"  - {ruta_sqlite.name}")
    print(f"  - {ruta_excel.name}")
    print(f"  - {ruta_correos.name} ({len(correos)} correos)")
    print("  - csv/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
