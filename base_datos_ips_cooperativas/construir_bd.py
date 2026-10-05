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

INTENTOS_DESCARGA = 4
TIMEOUT_DESCARGA_SEGUNDOS = 300

# ===========================================================================

# Cooperativa, cooperativas, precooperativa, organismo cooperativo, "COOP."
# No reconoce "cooperativismo", para no confundir las instituciones
# auxiliares del cooperativismo con cooperativas.
PATRON_COOPERATIVA = re.compile(r"\b(PRE)?COOPERATIV(A|AS|O|OS)\b|\bCOOP\b")

PATRON_IPS = re.compile(r"\bIPS\b|\bINSTITUCION(ES)? PRESTADORA")

MAX_FILAS_EXCEL = 1_048_575


# ------------------------------- Texto -------------------------------------

def sin_tildes(texto):
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def normalizar(texto):
    """Mayusculas, sin tildes y con espacios simples."""
    return re.sub(r"\s+", " ", sin_tildes(str(texto or ""))).strip().upper()


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
    """'890.981.234-5', '8909812345' y '890981234' -> '890981234'."""
    texto = str(valor or "").strip()
    if re.fullmatch(r"\d+\.0+", texto):
        texto = texto.split(".")[0]
    digitos = re.sub(r"\D", "", texto.split("-")[0])
    # NIT de persona juridica (empieza por 8 o 9) con el digito de
    # verificacion pegado al final.
    if len(digitos) == 10 and digitos[0] in "89":
        digitos = digitos[:9]
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
    return 0, 0, 0


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
    cols = unicos([nombre_columna(h) for h in encabezados])
    datos = []
    for fila in filas:
        fila = list(fila) + [""] * (len(cols) - len(fila))
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
    col_corte = buscar_columna(encabezados, ("fechacorte", "corte", "periodo"))

    # Si el listado trae varios cortes de la misma entidad, deja el mas reciente.
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

    nits_cooperativas = set()
    entidades = []
    for fila in filas:
        nombre = fila.get(col_nombre, "") if col_nombre else ""
        texto = normalizar(" ".join([fila[c] for c in cols_tipo] + [nombre]))
        cooperativa = bool(PATRON_COOPERATIVA.search(texto))
        if cooperativa and col_nit:
            nit = normalizar_nit(fila[col_nit])
            if nit:
                nits_cooperativas.add(nit)

        departamento, municipio = ubicar(fila, cols_depto, cols_muni,
                                         objetivos)
        if departamento:
            entidades.append({
                "bd_departamento": departamento,
                "bd_municipio": municipio,
                "bd_es_cooperativa": "SI" if cooperativa else "NO",
                **fila,
            })
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


def agrupar_ips(encabezados, sedes):
    """Una fila por IPS y departamento, con el numero de sedes que tiene ahi.
    Usa los datos de la sede principal cuando el REPS la identifica."""
    col_codigo = buscar_columna(encabezados, ("codigoprestador", "codprestador"))
    col_nit = columna_nit(encabezados)
    col_nombre = columna_nombre(encabezados)
    col_principal = buscar_columna(encabezados, ("sedeprincipal", "principal"))

    grupos = {}
    for sede in sedes:
        if sede["bd_es_ips"] != "SI":
            continue
        identificador = (
            (col_codigo and sede[col_codigo])
            or (col_nit and normalizar_nit(sede[col_nit]))
            or (col_nombre and normalizar(sede[col_nombre]))
            or id(sede))
        grupos.setdefault((identificador, sede["bd_departamento"]), []).append(sede)

    ips = []
    for grupo in grupos.values():
        principal = grupo[0]
        if col_principal:
            for sede in grupo:
                if normalizar(sede[col_principal]) in ("SI", "S", "1", "TRUE", "X"):
                    principal = sede
                    break
        fila = dict(principal)
        fila["bd_num_sedes"] = len(grupo)
        ips.append(fila)
    return ips


def ordenar(filas, encabezados):
    """Por departamento, municipio y razon social."""
    col_nombre = columna_nombre(encabezados)
    return sorted(filas, key=lambda f: (
        f["bd_departamento"], f["bd_municipio"],
        normalizar(f.get(col_nombre, "")) if col_nombre else ""))


def resumen(objetivos, ips, sedes, entidades):
    filas = []
    for departamento in sorted(objetivos, key=normalizar):
        def contar(lista, **condiciones):
            return sum(1 for f in lista if f["bd_departamento"] == departamento
                       and all(f.get(k) == v for k, v in condiciones.items()))
        filas.append({
            "departamento": departamento,
            "ips": contar(ips),
            "sedes_de_ips": contar(sedes, bd_es_ips="SI"),
            "ips_cooperativas": contar(ips, bd_es_cooperativa="SI"),
            "otras_sedes_reps_no_ips": contar(sedes, bd_es_ips="NO"),
            "cooperativas": contar(entidades, bd_es_cooperativa="SI"),
            "otras_entidades_solidarias": contar(entidades, bd_es_cooperativa="NO"),
        })
    total = {"departamento": "TOTAL"}
    for clave_total in filas[0]:
        if clave_total != "departamento":
            total[clave_total] = sum(f[clave_total] for f in filas)
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
    cols_ips = columnas_de(cols_reps, calc_reps[:2] + ["bd_num_sedes"]
                           + calc_reps[3:])
    cols_entidades = columnas_de(
        cols_sol, ["bd_departamento", "bd_municipio", "bd_es_cooperativa"])

    ips = ordenar(ips, cols_reps)
    sedes = ordenar(sedes, cols_reps)
    entidades = ordenar(entidades, cols_sol)
    cooperativas = [e for e in entidades if e["bd_es_cooperativa"] == "SI"]
    ips_cooperativas = [i for i in ips if i["bd_es_cooperativa"] == "SI"]
    sedes_ips = [s for s in sedes if s["bd_es_ips"] == "SI"]

    filas_resumen = resumen(objetivos, ips, sedes, entidades)
    cols_resumen = list(filas_resumen[0])
    fuentes = [
        {"fuente": FUENTES["reps"]["titulo"], "dataset": FUENTES["reps"]["dataset"],
         "origen": origen_reps, "registros_nacionales": len(filas_reps),
         "registros_en_departamentos": len(sedes)},
        {"fuente": FUENTES["solidarias"]["titulo"],
         "dataset": FUENTES["solidarias"]["dataset"], "origen": origen_sol,
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

    print()
    print(f"{'Departamento':<22}{'IPS':>8}{'Sedes IPS':>11}"
          f"{'IPS coop.':>11}{'Cooperativas':>14}")
    for f in filas_resumen:
        print(f"{f['departamento']:<22}{f['ips']:>8}{f['sedes_de_ips']:>11}"
              f"{f['ips_cooperativas']:>11}{f['cooperativas']:>14}")
    print(f"\nListo. Resultado en {salida.resolve()}:")
    print(f"  - {ruta_sqlite.name}")
    print(f"  - {ruta_excel.name}")
    print("  - csv/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
