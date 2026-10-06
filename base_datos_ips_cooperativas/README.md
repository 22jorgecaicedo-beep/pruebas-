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
| Relación de IPS públicas y privadas según el nivel de atención y capacidad instalada (para ordenar las IPS por tamaño) | Ministerio de Salud | [`s2ru-bqt6`](https://www.datos.gov.co/d/s2ru-bqt6) |
| Estados financieros de entidades solidarias, cuenta ACTIVO (para ordenar las cooperativas por tamaño) | Supersolidaria | [`tic6-rbue`](https://www.datos.gov.co/d/tic6-rbue) |

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
  - **IPS**: una fila por IPS (por NIT), sin repetidas, de la más grande a
    la más pequeña, con los datos de su sede principal, su capacidad
    instalada, el total de sedes y los departamentos donde tiene sedes.
  - **Sedes IPS**: todas las sedes de las IPS, con dirección y teléfono de
    cada una, en el orden de su IPS por tamaño.
  - **Cooperativas**: todas las cooperativas y precooperativas, de la más
    grande a la más pequeña, con sus activos y la fecha de su último
    reporte a la Supersolidaria.
  - **IPS cooperativas**: las IPS que además son cooperativas.
  - **Fuentes**: de dónde salió cada dato y la fecha en que se construyó la
    base.
- `Ranking_IPS.xlsx`: solo las IPS, de la más grande a la más pequeña, con
  las columnas para leer y contactar (ranking, nombre, NIT, si es
  cooperativa, ubicación, capacidad instalada, naturaleza, dirección,
  teléfono y correo). Una hoja con todas las IPS y otra con las IPS
  cooperativas.
- `Correos_IPS.xlsx`: los correos de las IPS para envíos masivos, un
  registro por correo, sin repetidos:
  - **Correos**: todos los correos de las IPS, empezando por los de las IPS
    más grandes, con el ranking de tamaño de la IPS, su nombre, NIT,
    departamentos, municipios, si es cooperativa y si es el correo
    principal de la IPS o el de una de sus sedes.
  - **Correos IPS cooperativas**: solo los de las IPS que son cooperativas.
  - **Por revisar**: campos de correo del REPS que no traen un correo
    válido (una página web, un correo sin dominio...).

  El REPS a veces trae varios correos en un mismo campo; se separan. Se
  quitan las tildes y se corrigen errores evidentes de dominio
  (`gmail.con` → `gmail.com`); la columna **Observación** dice qué se
  cambió.
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
| `bd_ranking_tamano` | Posición por tamaño: 1 es la IPS (o la cooperativa) más grande. Ver "Orden por tamaño". |
| `bd_departamento` | Departamento donde está la sede o la entidad (en la hoja IPS, el de su sede principal). |
| `bd_municipio` | Municipio donde está la sede o la entidad (en la hoja IPS, el de su sede principal). |
| `bd_departamentos` | Todos los departamentos donde la IPS tiene sedes. |
| `bd_es_ips` | `SI` si la clase de prestador es IPS. |
| `bd_num_sedes` | Número de sedes de la IPS en los cinco departamentos. |
| `bd_inscripciones_reps` | Cuántas inscripciones (códigos de prestador) tiene la IPS en el REPS; ver "Una fila por IPS". |
| `bd_capacidad_instalada` | Total de camas, camillas, consultorios, salas, ambulancias, sillas y unidades móviles de la IPS en los cinco departamentos. Vacío si el registro de capacidad no tiene datos de la IPS. |
| `bd_camas`, `bd_consultorios`, `bd_salas`, `bd_ambulancias` | El detalle de esa capacidad. |
| `bd_nivel_atencion` | Nivel de atención (1, 2 o 3; el 3 es el de mayor complejidad). Solo lo reportan los hospitales públicos. |
| `bd_activos_pesos` | Total de activos de la cooperativa en su último reporte de estados financieros a la Supersolidaria. |
| `bd_fecha_activos` | Año y mes de ese reporte. |
| `bd_es_cooperativa` | `SI` si la entidad es una cooperativa. |
| `bd_reporta_actualmente` | `SI` si la cooperativa le reportó a la Supersolidaria en los últimos 12 meses del listado; `NO` suele indicar que está inactiva o en liquidación. |
| `bd_ultimo_reporte` | Fecha del último reporte de la entidad a la Supersolidaria. |
| `bd_criterio_cooperativa` | Por qué se marcó la IPS como cooperativa: su NIT está registrado como cooperativa en la Supersolidaria, o su razón social es de cooperativa. |

## Orden por tamaño

- **IPS**: de mayor a menor **capacidad instalada** en los cinco
  departamentos (camas, camillas, consultorios, salas, ambulancias, sillas
  y unidades móviles, cada una cuenta como una). A igual capacidad, primero
  la de más camas y luego la de más sedes. El registro de capacidad
  instalada que publica el Ministerio es de noviembre de 2022: las IPS que
  no aparecen en él (casi siempre, las inscritas después) quedan al final,
  ordenadas por número de sedes.
- **Cooperativas**: de mayor a menor **total de activos** en su último
  reporte de estados financieros a la Supersolidaria. Las que ya no le
  reportan (o cuyos activos no se pudieron consultar) quedan al final, por
  nivel de supervisión (la Supersolidaria pone en el nivel 1 a las más
  grandes). Los estados financieros tienen más de 300 millones de filas y
  datos.gov.co solo responde a tiempo si se consultan entidad por entidad,
  así que esta parte toma unos minutos.
- Si alguno de esos dos registros no se puede descargar, la base se arma
  igual, y las IPS se ordenan por número de sedes y las cooperativas por
  nivel de supervisión.

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
