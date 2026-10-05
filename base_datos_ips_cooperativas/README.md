# Base de datos de IPS y cooperativas

Arma una base de datos con **todas las IPS registradas** y **todas las
cooperativas** de:

- Antioquia
- Atlántico
- Bolívar
- Cundinamarca
- Norte de Santander

a partir de los registros oficiales publicados en
[datos.gov.co](https://www.datos.gov.co). Los datos no se escriben a mano:
cada vez que lo ejecutas descarga la versión vigente de los registros, así
que la base siempre queda al día.

| Fuente | Entidad | Dataset |
|---|---|---|
| Registro Especial de Prestadores y Sedes de Servicios de Salud (REPS) | Ministerio de Salud | [`c36g-9fc2`](https://www.datos.gov.co/d/c36g-9fc2) |
| Listado de Entidades del Sector Solidario | Supersolidaria | [`kg2d-yfyg`](https://www.datos.gov.co/d/kg2d-yfyg) |

## La base de datos ya construida

La carpeta [`resultado/`](resultado/) tiene la base de datos ya armada con
los registros oficiales (ver la hoja **Fuentes** para la fecha de corte de
cada registro). Abre `resultado/IPS_y_Cooperativas.xlsx` en Excel.

Para actualizarla más adelante, ejecuta el programa en tu computador (ver
abajo) o, desde GitHub, en la pestaña **Actions** → **Construir base de
datos de IPS y cooperativas** → **Run workflow** (disponible cuando el
cambio esté en la rama principal). GitHub descarga los registros, arma la
base y guarda el resultado en `resultado/`.

## Instalación (Windows)

1. Instala [Python 3.10+](https://www.python.org/downloads/) marcando la
   opción "Add Python to PATH" durante la instalación.
2. Haz doble clic en `construir_bd.bat`. Instala lo necesario (solo
   `openpyxl`), descarga los registros y arma la base de datos.

O desde una terminal en esta carpeta:

```
pip install -r requirements.txt
python construir_bd.py
```

## Qué produce

Todo queda en la carpeta `salida/`:

- `IPS_y_Cooperativas.xlsx`: libro de Excel con las hojas:
  - **Resumen**: cuántas IPS, sedes y cooperativas hay en cada departamento.
  - **IPS**: una fila por IPS (por NIT), sin repetidas, con los datos de su
    sede principal, el total de sedes y los departamentos donde tiene sedes.
  - **Sedes IPS**: todas las sedes de las IPS, con dirección y teléfono de
    cada una.
  - **Cooperativas**: todas las cooperativas y precooperativas, con la
    fecha de su último reporte a la Supersolidaria.
  - **IPS cooperativas**: las IPS que además son cooperativas.
  - **Fuentes**: de dónde salió cada dato y la fecha en que se construyó la
    base.
- `ips_cooperativas.sqlite`: las mismas tablas en una base de datos SQLite
  (se abre con [DB Browser for SQLite](https://sqlitebrowser.org/), Power BI,
  Access, etc.), más dos tablas completas:
  - `reps_todas_las_sedes`: todo el REPS de los departamentos, incluidos
    profesionales independientes, transporte especial de pacientes y
    entidades con objeto social diferente.
  - `entidades_solidarias`: todo el sector solidario de los departamentos,
    incluidos fondos de empleados y asociaciones mutuales.
- `csv/`: cada tabla en un CSV separado por `;` (se abre directo en Excel).
- `fuentes/`: los archivos originales descargados, tal cual.

Las columnas que empiezan por `bd_` las calcula el programa; el resto son
las columnas originales del registro oficial:

| Columna | Significado |
|---|---|
| `bd_departamento` | Departamento donde está la sede o la entidad (en la hoja IPS, el de su sede principal). |
| `bd_municipio` | Municipio donde está la sede o la entidad (en la hoja IPS, el de su sede principal). |
| `bd_departamentos` | Todos los departamentos donde la IPS tiene sedes. |
| `bd_es_ips` | `SI` si la clase de prestador es IPS. |
| `bd_num_sedes` | Número de sedes de la IPS en los cinco departamentos. |
| `bd_inscripciones_reps` | Cuántas inscripciones (códigos de prestador) tiene la IPS en el REPS; ver "Una fila por IPS". |
| `bd_es_cooperativa` | `SI` si la entidad es una cooperativa. |
| `bd_reporta_actualmente` | `SI` si la cooperativa le reportó a la Supersolidaria en los últimos 12 meses del listado; `NO` suele indicar que está inactiva o en liquidación. |
| `bd_ultimo_reporte` | Fecha del último reporte de la entidad a la Supersolidaria. |
| `bd_criterio_cooperativa` | Por qué se marcó la IPS como cooperativa: su NIT está registrado como cooperativa en la Supersolidaria, o su razón social es de cooperativa. |

## Criterios

- **IPS**: prestadores del REPS cuya clase es "Instituciones Prestadoras de
  Servicios de Salud - IPS". Los profesionales independientes, el
  transporte especial de pacientes y las entidades con objeto social
  diferente quedan solo en `reps_todas_las_sedes`.
- **Departamento de una IPS**: el de cada **sede**, no el del domicilio del
  prestador. Una IPS de Bogotá con una sede en Medellín aparece en
  Antioquia con esa sede.
- **Una fila por IPS**: el REPS inscribe a una misma IPS una vez por cada
  municipio o distrito donde presta servicios, cada vez con otro código de
  prestador (Davita S.A.S. tiene 7 inscripciones en estos departamentos).
  La hoja IPS las junta por NIT, así que cada IPS aparece una sola vez. En
  el **Resumen**, una IPS con sedes en varios departamentos cuenta en cada
  uno, pero una sola vez en el total. IPS distintas con el mismo nombre
  (por ejemplo, varios "ESE Hospital San Juan de Dios") tienen NIT
  distinto y se mantienen.
- **Cooperativas**: entidades de la Supersolidaria cuyo tipo es de
  cooperativa (multiactiva, especializada, integral, de trabajo asociado,
  de aportes y crédito, precooperativa, administración pública cooperativa
  u organismo de carácter económico) o cuya razón social dice
  "cooperativa". Los fondos de empleados y las asociaciones mutuales nunca
  cuentan como cooperativas; quedan en `entidades_solidarias` junto con las
  instituciones auxiliares y demás entidades del sector.
- **Un registro por entidad**: el listado de la Supersolidaria trae un
  registro por cada reporte desde 2017; se deja el más reciente de cada
  NIT. Se incluyen todas las cooperativas registradas, también las que ya
  no reportan; la columna `bd_reporta_actualmente` las distingue.
- **IPS cooperativas**: IPS cuyo NIT está registrado como cooperativa en la
  Supersolidaria (en cualquier departamento), o cuya razón social dice
  "cooperativa".
- **Bogotá D.C.** es un distrito aparte de Cundinamarca en los registros
  oficiales, así que no se incluye. Para incluirla:

  ```
  python construir_bd.py --incluir-bogota
  ```

Para cambiar los departamentos, edita `DEPARTAMENTOS` al inicio de
`construir_bd.py`.

## Si la descarga falla

Si www.datos.gov.co no responde, el programa usa la última copia que haya
descargado en `salida/fuentes/`. Si nunca se ha descargado, puedes bajar
los archivos a mano y pasárselos:

1. Abre cada dataset (enlaces de la tabla de arriba), pulsa **Exportar** y
   descarga el **CSV**.
2. Ejecuta:

   ```
   python construir_bd.py --archivo-reps "C:\ruta\reps.csv" --archivo-solidarias "C:\ruta\solidarias.csv"
   ```

También acepta archivos `.xlsx`, por ejemplo una exportación más reciente
hecha desde el portal del REPS
([prestadores.minsalud.gov.co](https://prestadores.minsalud.gov.co/habilitacion/)).

Si descargas muchas veces seguidas y datos.gov.co empieza a limitar las
descargas, crea un token de aplicación gratuito en datos.gov.co y pásalo
con `--app-token TU_TOKEN`.

## Pruebas

```
python -m unittest test_construir_bd
```

Las pruebas usan archivos de ejemplo y no necesitan internet.
