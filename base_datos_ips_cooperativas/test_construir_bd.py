"""Pruebas de construir_bd.py con archivos de ejemplo (no usa internet).

Ejecutar con:  python -m unittest test_construir_bd
"""

import contextlib
import csv
import io
import sqlite3
import tempfile
import unittest
from pathlib import Path

import construir_bd as bd

REPS = [
    ["CodigoPrestador", "NombrePrestador", "CodigoHabilitacionSede",
     "NombreSede", "SedePrincipal", "TipoIdentificacion",
     "NumeroIdentificacion", "ClasePrestador", "NaturalezaJuridica",
     "DepartamentoPrestadorDesc", "MunicipioPrestadorDesc",
     "DepartamentoSedeDesc", "MunicipioSedeDesc", "Direccion", "Telefono"],
    # IPS con dos sedes en Antioquia: debe quedar una sola fila con 2 sedes.
    ["0500100001", "CLINICA SAN JUAN SAS", "050010000102", "SEDE ENVIGADO",
     "NO", "NI", "900111222", "Instituciones Prestadoras de Servicios de Salud - IPS",
     "Privada", "Antioquia", "MEDELLÍN", "Antioquia", "ENVIGADO", "CL 1", "111"],
    ["0500100001", "CLINICA SAN JUAN SAS", "050010000101", "SEDE PRINCIPAL",
     "SI", "NI", "900111222", "Instituciones Prestadoras de Servicios de Salud - IPS",
     "Privada", "Antioquia", "MEDELLÍN", "Antioquia", "MEDELLÍN", "CL 2", "222"],
    # Norte de Santander escrito abreviado.
    ["5400100002", "IPS FRONTERA LTDA", "540010000201", "IPS FRONTERA LTDA",
     "SI", "NI", "800222333-4", "Instituciones Prestadoras de Servicios de Salud - IPS",
     "Privada", "N. de Santander", "CÚCUTA", "N. de Santander", "CÚCUTA", "AV 0", "333"],
    # Santander no es Norte de Santander.
    ["6800100003", "IPS BUCARAMANGA", "680010000301", "IPS BUCARAMANGA",
     "SI", "NI", "900333444", "Instituciones Prestadoras de Servicios de Salud - IPS",
     "Privada", "Santander", "BUCARAMANGA", "Santander", "BUCARAMANGA", "CR 1", "444"],
    # Bogota solo entra con --incluir-bogota.
    ["1100100004", "IPS CAPITAL", "110010000401", "IPS CAPITAL",
     "SI", "NI", "900444555", "Instituciones Prestadoras de Servicios de Salud - IPS",
     "Privada", "Bogotá D.C.", "BOGOTÁ", "Bogotá D.C.", "BOGOTÁ", "CR 7", "555"],
    # Prestador con domicilio en Cundinamarca pero sede en Bogota: la sede
    # esta en Bogota.
    ["2575400005", "IPS SOACHA", "257540000502", "SEDE BOGOTA",
     "NO", "NI", "900555666", "Instituciones Prestadoras de Servicios de Salud - IPS",
     "Privada", "Cundinamarca", "SOACHA", "Bogotá D.C.", "BOGOTÁ", "CL 80", "666"],
    # Profesional independiente: va a reps_todas_las_sedes pero no es IPS.
    ["0800100006", "PEREZ GOMEZ JUAN", "080010000601", "CONSULTORIO",
     "SI", "CC", "72123456", "Profesional Independiente",
     "Privada", "Atlántico", "BARRANQUILLA", "Atlántico", "BARRANQUILLA", "CL 3", "777"],
    # IPS cuyo NIT (con digito de verificacion) es de una cooperativa de la
    # Supersolidaria registrada en otro departamento.
    ["1300100007", "SERVISALUD BOLIVAR", "130010000701", "SERVISALUD",
     "SI", "NI", "890.123.456-7", "Instituciones Prestadoras de Servicios de Salud - IPS",
     "Privada", "Bolívar", "CARTAGENA", "Bolívar", "CARTAGENA", "CL 4", "888"],
    # IPS que es cooperativa por su razon social.
    ["2530700008", "COOPERATIVA DE SALUD DE GIRARDOT", "253070000801", "COOPSALUD",
     "SI", "NI", "900777888", "Instituciones Prestadoras de Servicios de Salud - IPS",
     "Privada", "Cundinamarca", "GIRARDOT", "Cundinamarca", "GIRARDOT", "CL 5", "999"],
]

SOLIDARIAS = [
    ["NIT", "DÍGITO DE VERIFICACIÓN", "RAZÓN SOCIAL", "SIGLA", "TIPO DE ENTIDAD",
     "CÓDIGO DEPARTAMENTO", "DEPARTAMENTO", "MUNICIPIO", "DIRECCIÓN",
     "FECHA DE CORTE"],
    # Dos cortes de la misma cooperativa: queda el mas reciente.
    ["890900111", "1", "COOPERATIVA VIEJA", "", "COOPERATIVA DE AHORRO Y CREDITO",
     "05", "ANTIOQUIA", "MEDELLIN", "CL 10", "01/31/2025 12:00:00 AM"],
    ["890900111", "1", "COOPERATIVA ANTIOQUEÑA", "COOPANT", "COOPERATIVA DE AHORRO Y CREDITO",
     "05", "ANTIOQUIA", "MEDELLIN", "CL 10", "08/31/2026 12:00:00 AM"],
    ["800100200", "2", "FONDO DE EMPLEADOS DEL CARIBE", "FONCARIBE", "FONDO DE EMPLEADOS",
     "08", "ATLANTICO", "BARRANQUILLA", "CL 11", "08/31/2026 12:00:00 AM"],
    ["800100300", "3", "INSTITUTO DEL COOPERATIVISMO", "", "INSTITUCION AUXILIAR DEL COOPERATIVISMO",
     "13", "BOLIVAR", "CARTAGENA", "CL 12", "08/31/2026 12:00:00 AM"],
    ["800100400", "4", "PRECOOPERATIVA DEL MAR", "", "PRECOOPERATIVA",
     "13", "BOLÍVAR", "CARTAGENA", "CL 13", "08/31/2026 12:00:00 AM"],
    ["800100500", "5", "COOPERATIVA DE TRABAJO ASOCIADO NORTE", "", "COOPERATIVA DE TRABAJO ASOCIADO",
     "54", "", "CUCUTA", "CL 14", "08/31/2026 12:00:00 AM"],
    # Cooperativa fuera de los departamentos; su NIT marca a la IPS de Bolivar.
    ["890123456", "7", "COOPERATIVA DEL VALLE", "", "COOPERATIVA MULTIACTIVA",
     "76", "VALLE DEL CAUCA", "CALI", "CL 15", "08/31/2026 12:00:00 AM"],
]


def escribir_csv(ruta, filas):
    with open(ruta, "w", encoding="utf-8", newline="") as archivo:
        csv.writer(archivo).writerows(filas)


class PruebasTexto(unittest.TestCase):
    def test_departamento_de(self):
        objetivos = bd.construir_objetivos()
        casos = {
            "Antioquia": "ANTIOQUIA",
            "ATLANTICO": "ATLÁNTICO",
            "Bolívar": "BOLÍVAR",
            "N. de Santander": "NORTE DE SANTANDER",
            "NORTE SANTANDER": "NORTE DE SANTANDER",
            "Departamento de Cundinamarca": "CUNDINAMARCA",
            "05 - ANTIOQUIA": "ANTIOQUIA",
            "54": "NORTE DE SANTANDER",
            "5": "ANTIOQUIA",
            "13001": "BOLÍVAR",
            "5001": "ANTIOQUIA",
            "Santander": None,
            "Bogotá D.C.": None,
            "Valle del Cauca": None,
            "": None,
        }
        for valor, esperado in casos.items():
            self.assertEqual(bd.departamento_de(valor, objetivos), esperado, valor)
        con_bogota = bd.construir_objetivos(incluir_bogota=True)
        self.assertEqual(bd.departamento_de("BOGOTA DC", con_bogota), "BOGOTÁ D.C.")
        self.assertEqual(bd.departamento_de("11", con_bogota), "BOGOTÁ D.C.")

    def test_normalizar_nit(self):
        for valor in ("890.123.456-7", "8901234567", "890123456", "890123456.0"):
            self.assertEqual(bd.normalizar_nit(valor), "890123456", valor)
        self.assertEqual(bd.normalizar_nit("72123456"), "72123456")
        self.assertEqual(bd.normalizar_nit(""), "")

    def test_patron_cooperativa(self):
        for texto in ("COOPERATIVA DE AHORRO", "PRECOOPERATIVA", "COOP. MEDICA",
                      "ORGANISMO COOPERATIVO DE GRADO SUPERIOR"):
            self.assertTrue(bd.PATRON_COOPERATIVA.search(texto), texto)
        for texto in ("INSTITUCION AUXILIAR DEL COOPERATIVISMO", "COOPSALUD",
                      "FONDO DE EMPLEADOS"):
            self.assertFalse(bd.PATRON_COOPERATIVA.search(texto), texto)

    def test_nombre_columna(self):
        self.assertEqual(bd.nombre_columna("NombrePrestador"), "nombre_prestador")
        self.assertEqual(bd.nombre_columna("RAZÓN SOCIAL"), "razon_social")
        self.assertEqual(bd.nombre_columna("DepartamentoSedeDesc"),
                         "departamento_sede_desc")
        self.assertEqual(bd.unicos(["a", "a", "a_2", "b"]), ["a", "a_2", "a_2_2", "b"])

    def test_clave_fecha(self):
        self.assertGreater(bd.clave_fecha("08/31/2026 12:00:00 AM"),
                           bd.clave_fecha("12/31/2025 12:00:00 AM"))
        self.assertEqual(bd.clave_fecha("2026-08-31T00:00:00.000"), (2026, 8, 31))
        self.assertEqual(bd.clave_fecha("31/08/2026"), (2026, 8, 31))
        self.assertEqual(bd.clave_fecha("202608"), (2026, 8, 0))


class PruebaCompleta(unittest.TestCase):
    def construir(self, *extra):
        carpeta = Path(self.enterContext(tempfile.TemporaryDirectory()))
        escribir_csv(carpeta / "reps.csv", REPS)
        escribir_csv(carpeta / "solidarias.csv", SOLIDARIAS)
        salida = carpeta / "salida"
        with contextlib.redirect_stdout(io.StringIO()):
            codigo = bd.main(["--archivo-reps", str(carpeta / "reps.csv"),
                              "--archivo-solidarias", str(carpeta / "solidarias.csv"),
                              "--salida", str(salida), *extra])
        self.assertEqual(codigo, 0)
        con = sqlite3.connect(salida / bd.NOMBRE_SQLITE)
        self.addCleanup(con.close)
        return salida, con

    def test_ips(self):
        _, con = self.construir()
        filas = con.execute(
            "SELECT bd_departamento, bd_municipio, nombre_prestador, bd_num_sedes, "
            "nombre_sede, bd_es_cooperativa, bd_criterio_cooperativa FROM ips "
            "ORDER BY bd_departamento, nombre_prestador").fetchall()
        self.assertEqual(filas, [
            ("ANTIOQUIA", "MEDELLIN", "CLINICA SAN JUAN SAS", 2, "SEDE PRINCIPAL", "NO", ""),
            ("BOLÍVAR", "CARTAGENA", "SERVISALUD BOLIVAR", 1, "SERVISALUD", "SI",
             "NIT registrado como cooperativa en la Supersolidaria"),
            ("CUNDINAMARCA", "GIRARDOT", "COOPERATIVA DE SALUD DE GIRARDOT", 1,
             "COOPSALUD", "SI", "Razón social de cooperativa"),
            ("NORTE DE SANTANDER", "CUCUTA", "IPS FRONTERA LTDA", 1, "IPS FRONTERA LTDA",
             "NO", ""),
        ])
        self.assertEqual(con.execute("SELECT COUNT(*) FROM ips_sedes").fetchone(), (5,))
        self.assertEqual(con.execute("SELECT COUNT(*) FROM ips_cooperativas").fetchone(), (2,))
        # El profesional independiente queda en el REPS completo, no como IPS.
        self.assertEqual(con.execute(
            "SELECT bd_departamento, bd_es_ips FROM reps_todas_las_sedes "
            "WHERE nombre_prestador = 'PEREZ GOMEZ JUAN'").fetchall(),
            [("ATLÁNTICO", "NO")])

    def test_cooperativas(self):
        _, con = self.construir()
        filas = con.execute(
            "SELECT bd_departamento, razon_social FROM cooperativas "
            "ORDER BY bd_departamento").fetchall()
        self.assertEqual(filas, [
            ("ANTIOQUIA", "COOPERATIVA ANTIOQUEÑA"),
            ("BOLÍVAR", "PRECOOPERATIVA DEL MAR"),
            ("NORTE DE SANTANDER", "COOPERATIVA DE TRABAJO ASOCIADO NORTE"),
        ])
        self.assertEqual(con.execute(
            "SELECT razon_social, bd_es_cooperativa FROM entidades_solidarias "
            "WHERE bd_es_cooperativa = 'NO' ORDER BY razon_social").fetchall(), [
            ("FONDO DE EMPLEADOS DEL CARIBE", "NO"),
            ("INSTITUTO DEL COOPERATIVISMO", "NO"),
        ])

    def test_resumen_y_archivos(self):
        salida, con = self.construir()
        total = con.execute(
            "SELECT ips, sedes_de_ips, ips_cooperativas, otras_sedes_reps_no_ips, "
            "cooperativas, otras_entidades_solidarias FROM resumen "
            "WHERE departamento = 'TOTAL'").fetchone()
        self.assertEqual(total, (4, 5, 2, 1, 3, 2))

        from openpyxl import load_workbook
        libro = load_workbook(salida / bd.NOMBRE_EXCEL)
        self.assertEqual(libro.sheetnames, ["Resumen", "IPS", "Sedes IPS",
                                            "Cooperativas", "IPS cooperativas",
                                            "Fuentes"])
        hoja = libro["IPS"]
        self.assertEqual(hoja.max_row, 5)
        self.assertEqual(hoja.freeze_panes, "A2")
        self.assertTrue(hoja.auto_filter.ref)
        self.assertEqual(hoja["A1"].value, "bd_departamento")
        self.assertTrue((salida / "csv" / "ips.csv").exists())

    def test_incluir_bogota(self):
        _, con = self.construir("--incluir-bogota")
        self.assertEqual(con.execute(
            "SELECT nombre_prestador FROM ips WHERE bd_departamento = 'BOGOTÁ D.C.' "
            "ORDER BY nombre_prestador").fetchall(),
            [("IPS CAPITAL",), ("IPS SOACHA",)])

    def test_lee_xlsx(self):
        from openpyxl import Workbook
        carpeta = Path(self.enterContext(tempfile.TemporaryDirectory()))
        libro = Workbook()
        libro.active.append(["Listado de prestadores"])  # titulo antes de la tabla
        for fila in REPS:
            libro.active.append(fila)
        libro.save(carpeta / "reps.xlsx")
        cols, filas = bd.leer_tabla(carpeta / "reps.xlsx")
        self.assertEqual(cols[:2], ["codigo_prestador", "nombre_prestador"])
        self.assertEqual(len(filas), len(REPS) - 1)


if __name__ == "__main__":
    unittest.main()
