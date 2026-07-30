# Procesos Jurídicos: herramienta única (correo → descarga → organizado)

Un solo programa, `procesos_juridicos.py`, que hace dos cosas a la vez:

1. **Vigila tu Gmail**: cuando un juzgado te comparte un expediente por el
   SGDE (Rama Judicial), entra automáticamente al portal, resuelve el
   token de verificación (leyéndolo de tu correo), descarga todo lo que
   haya en "Elementos Compartidos" —incluidas las carpetas que no tienen
   flecha de descarga directa, entrando a ellas y descargando archivo por
   archivo— y deja el resultado ya descomprimido en tu disco duro, en una
   carpeta nombrada con el número de expediente (23 dígitos).
2. **Vigila tu carpeta de Descargas**: por si alguna vez bajas un zip de
   un proceso a mano (de otro sistema, por ejemplo), lo extrae, busca el
   radicado dentro de los PDF/DOCX y lo organiza igual en el disco duro.

Si configuras `RUTA_EXCEL` en `validar_renombrar_carpetas.py` (ver más
abajo), cada carpeta nueva —de correo o manual— se cruza automáticamente
contra ese informe: si el radicado ya aparece ahí, la carpeta queda
nombrada `"numero. radicado"` desde el momento en que se crea, en vez de
solo el radicado. Si el radicado todavía no está en el informe, la
carpeta se deja solo con el radicado (como antes) y queda registrada en
el log; más tarde puedes correr `validar_renombrar_carpetas.py` para
completar el nombre cuando el informe se actualice.

Todo queda registrado en `procesos_juridicos.log`, dentro de la carpeta
destino.

## Instalación (Windows)

1. Instala [Python 3.10+](https://www.python.org/downloads/) marcando la
   opción "Add Python to PATH" durante la instalación.
2. Abre una terminal (`cmd` o PowerShell) en esta carpeta y ejecuta:

   ```
   pip install -r requirements.txt
   playwright install chromium
   ```

## Configuración

Abre `procesos_juridicos.py` y edita, al inicio del archivo, la sección
`CONFIGURACION`:

- `CARPETA_DESCARGAS`: carpeta donde el navegador guarda los `.zip`
  (por ejemplo `C:\Users\TuUsuario\Downloads`).
- `ETIQUETA_DISCO_EXTERNO`: el NOMBRE de tu disco duro externo (por
  ejemplo `"OSCAL"`, tal como aparece en "Este equipo"). El script busca
  el disco por ese nombre entre todas las unidades conectadas y usa la
  letra que encuentre en ese momento -- así no importa si Windows le
  asigna `D:`, `E:` o cualquier otra letra la próxima vez que lo
  conectes. `CARPETA_DESTINO` se calcula sola a partir de esto; no hace
  falta que edites `CARPETA_DESTINO` directamente. Si el disco no se
  encuentra conectado, se usa `CARPETA_DESTINO_RESPALDO` como respaldo
  (y el log avisa claramente que pasó eso).
- `PATRONES_RADICADO`: usados solo para zips que descargues a mano.
  Reconoce por defecto el radicado judicial colombiano de 23 dígitos.

## Configurar el acceso a tu Gmail (para la vigilancia automática)

El programa necesita leer tu correo (vía IMAP) para conseguir el link de
cada expediente y el token de verificación. **Nunca uses tu contraseña
normal de Gmail** — usa una "Contraseña de aplicación":

1. Activa la verificación en 2 pasos (una sola vez):
   https://myaccount.google.com/security
2. Crea una contraseña de aplicación (una sola vez):
   https://myaccount.google.com/apppasswords
3. En esta carpeta, copia `credenciales_sgde.example.txt` y renómbralo a
   `credenciales_sgde.txt`.
4. Ábrelo y reemplaza los valores por tu correo y la contraseña de
   aplicación de 16 letras.

Una vez configurada, el inicio de sesión con esa contraseña **no pide
código ni confirmación en el teléfono** — el 2FA interactivo solo aplica
a logins manuales desde el navegador, no al uso de contraseñas de
aplicación por IMAP.

`credenciales_sgde.txt` está en `.gitignore`: nunca se sube a git, ni
debes compartirlo con nadie.

Si no creas este archivo, el programa sigue funcionando igual, solo que
sin la vigilancia automática de correo (únicamente organiza lo que
descargues a mano).

### ⚠️ Antes de dejarlo corriendo

Verifica que automatizar el acceso al SGDE con tu cuenta autorizada no
infrinja los términos de uso del portal. El programa solo automatiza
pasos que tú ya estás autorizado a hacer manualmente (abrir el link que
te comparten, escribir tu correo y el token que te llega a ti mismo); no
intenta acceder a expedientes ajenos ni evade ninguna medida de
seguridad.

## Uso

Haz doble clic en `iniciar.bat`. Hace las dos cosas en orden, en la misma
ventana:

1. Corre `validar_renombrar_carpetas.py` una vez (revisa/renombra las
   carpetas que ya existen en el disco contra el informe de Excel) y
   muestra su reporte.
2. Cuando termina, arranca `procesos_juridicos.py` y se queda vigilando
   Descargas (y correo, si configuraste `credenciales_sgde.txt`) de forma
   indefinida.

**Importante**: los dos pasos corren UNO DESPUÉS DEL OTRO, no al mismo
tiempo -- mientras el Paso 1 sigue trabajando (puede tardar varios
minutos si `VALIDAR_CONTENIDO_CONTRA_NOMBRE = True`), la vigilancia de
Descargas todavía no ha arrancado, así que un zip que descargues justo
en ese momento no se procesa solo. Si quieres que la vigilancia esté
corriendo siempre, sin depender de que termine el Paso 1, abre
`vigilar.bat` en una ventana aparte -- corre solo `procesos_juridicos.py`
de una vez, sin esperar nada.

O si prefieres correr cada uno por separado, desde una terminal:

```
python validar_renombrar_carpetas.py
python procesos_juridicos.py
```

La primera vez, deja `NAVEGADOR_VISIBLE = True` (así viene por defecto)
para ver el navegador trabajando y confirmar que todo va bien. Cuando
confíes en que funciona, cámbialo a `False` en `procesos_juridicos.py`
para que corra en segundo plano sin abrir ventanas.

Déjalo abierto; ciérralo con `Ctrl+C` (o cerrando la ventana) cuando
termines.

### Si no quieres dejarlo corriendo de fondo

Por defecto el programa se queda vigilando Descargas en tiempo real (y el
correo, si configuraste `credenciales_sgde.txt`) hasta que lo cierres.

Si prefieres correrlo una vez al día y que termine solo —por ejemplo con
el Programador de tareas de Windows, en vez de dejar una ventana
abierta— abre `procesos_juridicos.py` y cambia:

```
SOLO_PROCESAR_HOY_Y_SALIR = True
```

Con esto, cada vez que lo corras organiza únicamente los `.zip` de los
últimos `DIAS_ATRAS_PROCESAR_EXISTENTES` días (por defecto 14, es decir
las últimas dos semanas) que ya estén en `CARPETA_DESCARGAS` (igual que hace
siempre al arrancar) y termina inmediatamente: no vigila el correo ni se
queda esperando descargas nuevas. Si no tienes `credenciales_sgde.txt`,
el correo ya se salta de por sí; este interruptor es para el otro caso,
cuando tampoco quieres que se quede vigilando Descargas.

Ese mismo `DIAS_ATRAS_PROCESAR_EXISTENTES` controla también, al arrancar
en modo vigilancia normal, cuántos días hacia atrás de zips ya existentes
se procesan (ponlo en `1` si solo quieres los de hoy).

## Ejecutarlo automáticamente al iniciar Windows (opcional)

1. Abre el "Programador de tareas" de Windows.
2. Crea una tarea nueva que se ejecute "Al iniciar sesión".
3. Como acción, apunta a `iniciar.bat` (o a `pythonw.exe
   procesos_juridicos.py` si prefieres que corra sin ventana visible).

## Cómo evita repetir trabajo

- Cada expediente ya descargado por correo se guarda en
  `expedientes_procesados.txt` (no se sube a git). Para volver a
  descargar uno, borra esa línea o el archivo completo.
- Los zips descargados a mano se mueven a una subcarpeta `Procesados`
  dentro de tu carpeta de Descargas (no se borran).
- Si ya existe una carpeta con el mismo nombre en el destino, se agrega
  un sufijo (`_2`, `_3`, ...) en vez de sobrescribir.

## Validar y renombrar carpetas contra el informe de Excel

`validar_renombrar_carpetas.py` es una herramienta aparte (no se ejecuta
junto con `procesos_juridicos.py`). Sirve para cuando ya tienes en el disco
duro carpetas nombradas solo con el radicado de 23 dígitos y quieres
verificar que cada una corresponda a un proceso del informe de Excel,
renombrándola a `"<numero>. <radicado>"` (por ejemplo
`133. 68001400301020180087200`).

El cruce se hace por el **valor exacto del radicado** (columna `RADICADO`
del Excel) contra el radicado que aparece en el nombre de cada carpeta; el
número de proceso sale de la columna `No.` de la misma fila. No importa si
faltan carpetas o filas, cada una se empareja de forma independiente.

1. Instala la dependencia nueva (ya incluida en `requirements.txt`):

   ```
   pip install -r requirements.txt
   ```

2. Abre `validar_renombrar_carpetas.py` y edita, al inicio del archivo, la
   sección `CONFIGURACION`:

   - `RUTA_EXCEL`: ruta al informe (`.xlsx` o `.xlsm`).
   - `HOJA_EXCEL`, `FILA_ENCABEZADO`, `COLUMNA_NO`, `COLUMNA_RADICADO`: en
     qué hoja y fila están los encabezados, y cómo se llaman las columnas
     del número de proceso y del radicado.
   - `ETIQUETA_DISCO_EXTERNO`: el NOMBRE de tu disco duro externo (por
     ejemplo `"OSCAL"`). `CARPETA_PROCESOS` se calcula sola buscando ese
     disco por su nombre entre las unidades conectadas, así no importa
     qué letra (`D:`, `E:`, etc) le asigne Windows esta vez.
   - `CARPETA_ENTRADA_ADICIONAL`: una carpeta aparte (por defecto
     `PROCESOS LAUE/ENTREGA EXPEDIENTE ESSA` dentro del disco) donde a
     veces caen entregas masivas de expedientes ya extraídos que todavía
     no se pasan a la raíz. Si existe, el script mueve a la raíz los que
     tengan radicado válido en el Excel y no estén ya en el disco, deja
     donde están los que ya existan (avisando), y deja aparte (avisando)
     los que no tengan proceso en el Excel. Si la ruta no existe, este
     paso simplemente se omite.
   - `CARPETA_DESCARGAS`: tu carpeta de Descargas (se detecta sola). Se usa
     solo para revisar si una carpeta vacía tiene un `.zip` pendiente de
     extraer ahí.

   Esta misma configuración (`RUTA_EXCEL`, `HOJA_EXCEL`, etc) es la que usa
   `procesos_juridicos.py` para nombrar bien las carpetas nuevas apenas las
   crea — no hay que configurarla dos veces.

3. Ejecuta el script (por defecto corre en `MODO_PRUEBA = True`, así que no
   renombra nada todavía, solo muestra un reporte):

   ```
   python validar_renombrar_carpetas.py
   ```

4. Revisa el reporte en pantalla y en `validar_renombrar_carpetas.log`:
   qué se renombraría, qué carpetas ya tienen el nombre correcto, qué
   procesos del Excel no tienen carpeta en el disco, y qué carpetas del
   disco no aparecen en el Excel (radicados repetidos o con formato
   inválido en el Excel se reportan y se omiten del cruce, para no
   arriesgar un renombrado incorrecto).

   Casos especiales:
   - `[Consecutivo]`: la carpeta y el Excel coinciden en los primeros 22
     dígitos del radicado y solo difieren en el último (el proceso ya
     cambió de instancia, ej. de "...00" a "...01") — esto **sí se
     corrige automático**, dejando el radicado del Excel.
   - `[Duplicado en Excel]`: el mismo radicado aparece en **dos o más
     filas** del Excel con números de proceso **distintos** (ej. proceso
     19 y proceso 451 con exactamente el mismo radicado). En vez de
     excluir esas filas del cruce, el script **duplica la carpeta**:
     conserva/renombra la existente con el primer número, y crea una
     copia completa (mismo contenido) por cada número adicional, para
     que cada número de proceso tenga su propia carpeta
     `"numero. radicado"`. Este caso también se reconoce **combinado**
     con `[Consecutivo]`: si la carpeta suelta tiene el último dígito
     distinto al del Excel, y ese radicado del Excel además está
     repetido con varios números, se corrige el dígito **y** se duplica
     para cada número, todo en la misma pasada.
   - `[POSIBLE COINCIDENCIA]`: la carpeta y el Excel difieren por un
     dígito de más o de menos en **otra** posición — esto **no** se
     corrige solo, se reporta para que lo confirmes a mano (dos procesos
     distintos del mismo juzgado y año pueden compartir casi todos los
     dígitos, así que adivinar mal sería peligroso).
   - `[Duplicado]` / `[Duplicados resueltos]`: cuando el mismo radicado
     aparece en más de una carpeta (típico de descargas repetidas), el
     script **nunca borra nada**. Se queda con la carpeta que tenga más
     archivos adentro (la más completa), la renombra con el número más
     reciente del Excel, y mueve las demás copias — intactas — a una
     carpeta `Duplicados_para_revisar` dentro de tu disco, para que las
     revises y borres a mano si de verdad sobran.
   - `[Duplicado sin resolver]`: hay más de una carpeta con el mismo
     radicado, pero ese radicado no está en el Excel — no se puede saber
     cuál conservar, así que no se toca ninguna; revísalas a mano.
   - `[Anidada]` / `[Anidadas resueltas]` / `[Anidadas de otro caso]`:
     cuando una carpeta de proceso queda metida DENTRO de otra (ej.
     `1014. radicado` adentro de `941. radicado`), también se resuelve
     sola, sin borrar ni fusionar contenido — solo mueve la carpeta
     completa. Si el radicado de la anidada es el mismo que el de la
     carpeta que la contiene, **o solo difiere en el último dígito**
     (el consecutivo de instancia/reparto, ej. termina en 0 o en 1 —
     es el mismo caso, no un proceso distinto), se mueve a
     `Duplicados_para_revisar`. Si es un radicado realmente distinto
     (contenido de otro caso mal ubicado), se saca al nivel principal
     del disco para evaluarla en la próxima corrida.
   - `[Carpeta vacía]`: carpetas sin ningún archivo adentro. El script
     revisa si hay un `.zip` con ese mismo radicado, tanto en
     `CARPETA_DESCARGAS` directamente como en su subcarpeta `Procesados`
     (ahí es donde quedaban los zips que en versiones viejas del programa
     se marcaban como "ya procesados" aunque la extracción hubiera
     fallado por completo) y te lo señala — esto solo se reporta, no se
     extrae ni se toca nada automático. Además de salir en el log, queda
     un reporte aparte en `carpetas_vacias.csv` (carpeta, radicado, el
     zip encontrado, y si estaba en Descargas o en Procesados) para que
     lo revises en Excel sin tener que buscar en el log completo.
   - `[Sin nombre reconocible]`: carpetas que no tienen ningún número que
     se parezca a un radicado en su nombre (antes se ignoraban en
     silencio, ahora se listan para que las revises a mano).
   - `[Contenido no corresponde]`: abre los documentos DENTRO de cada
     carpeta (nombres de archivo, y si hace falta el texto de hasta
     `MAX_ARCHIVOS_CONTENIDO_A_REVISAR` PDF/DOCX) y revisa si el radicado
     que aparece adentro corresponde con el radicado del nombre de la
     carpeta. Si el radicado del nombre nunca aparece en su propio
     contenido pero otro sí aparece claramente, se marca como sospechosa
     de tener contenido de otro caso mal ubicado (esto fue justo lo que
     pasó con algunos zips antes de la corrección del radicado-por-nombre-
     de-zip). Queda además en un reporte aparte:
     `contenido_no_corresponde.csv`. Esta revisión tarda más porque tiene
     que leer documentos de todas las carpetas; puedes desactivarla
     poniendo `VALIDAR_CONTENIDO_CONTRA_NOMBRE = False` si prefieres una
     corrida rápida.

5. Si el reporte se ve bien, cambia `MODO_PRUEBA = False` y vuelve a
   correrlo para aplicar los renombrados de verdad. Puedes correrlo las
   veces que quieras: las carpetas que ya tengan el nombre correcto se
   dejan igual.

## Validar el juzgado contra el portal de Rama Judicial

`validar_juzgados_ramajudicial.py` es otra herramienta aparte. Consulta
cada radicado del informe de Excel en el portal público **Consulta de
Procesos Nacional Unificada** (CPNU,
`consultaprocesos.ramajudicial.gov.co`) y compara el despacho que
reporta el portal contra el juzgado anotado en la columna `JUZGADO` del
Excel, para detectar procesos que ya cambiaron de despacho pero el Excel
todavía no se actualizó.

Cómo hace el cruce: el radicado termina en un "consecutivo" (últimos 2
dígitos) que cambia cuando el proceso pasa a otro despacho. El script
consulta el radicado tal como está en el Excel, y luego prueba el
consecutivo siguiente (+1, +2, ...) mientras el portal siga encontrando
resultado; el despacho del último consecutivo que sí exista es el que se
compara contra la columna `JUZGADO`.

⚠️ **Antes de usarlo, ten en cuenta:**

- Este script no se pudo probar contra el portal real antes de
  entregarlo (el entorno donde se escribió no tiene acceso a ese sitio).
  Corre primero con `SOLO_ESTOS_NUMEROS = [1, 133]` (o los que tú
  elijas) para confirmar que funciona antes de lanzarlo contra todo el
  informe, y avísame qué error sale si algo falla para ajustar el script.
- La comparación de nombres de juzgado es tolerante a diferencias de
  formato (el Excel dice "PRIMERO CIVIL MUNICIPAL...", el portal dice
  "JUZGADO 001 CIVIL MUNICIPAL... (SANTANDER)"), pero no es infalible:
  revisa el CSV completo (no solo la lista de diferencias) si algo se ve
  raro.
- Hace una pausa entre cada consulta para no saturar un portal público
  del Estado; con ~950 procesos, considera correrlo en un par de tandas.
  Si lo interrumpes con `Ctrl+C`, la próxima vez sigue donde iba en vez
  de repetir consultas ya hechas (gracias a
  `validar_juzgados_progreso.txt`).
- No hay que resolver ningún captcha (se confirmó manualmente que la
  consulta por número de radicado no lo pide).

Uso:

1. Edita la sección `CONFIGURACION` al inicio del archivo: `RUTA_EXCEL`,
   `HOJA_EXCEL`, y si quieres, `SOLO_ESTOS_NUMEROS` para una prueba
   chica primero.
2. Corre:

   ```
   python validar_juzgados_ramajudicial.py
   ```

3. Revisa `validar_juzgados_reporte.csv` (el detalle completo de todos
   los procesos consultados) y el resumen final en pantalla / en
   `validar_juzgados_ramajudicial.log` (solo las diferencias y los
   radicados que el portal no encontró).

## Buscar en Google Drive/correo los procesos que faltan en el disco

`buscar_faltantes_en_drive.py` es otra herramienta aparte. Lee
`procesos_faltantes_en_disco.csv` (lo genera `validar_renombrar_carpetas.py`)
y busca cada proceso faltante en tu Google Drive y en tu correo de Gmail,
por este orden:

1. El **radicado completo** de 23 dígitos -- si encuentra una
   coincidencia exacta, la descarga automático (tan específico que no
   hay riesgo real de confundirlo con otro caso).
2. El **radicado corto** (ej. `2025-00456` o `2025-456`).
3. El número de **cuenta** -- solo si es lo bastante específica. Una
   cuenta vacía, `"0"`, o de muy pocos dígitos **no se busca**: como la
   búsqueda de Drive es aproximada, un término tan genérico coincidiría
   con miles de carpetas de todo el Drive sin relación, dejando el
   proceso "pegado" un buen rato revisando falsos positivos uno por
   uno.

Los casos 2 y 3 son menos confiables (un radicado corto o una cuenta
puede coincidir por casualidad con archivos de otro proceso); para
esos, antes de descargar también se valida que la carpeta (o sus
archivos) de verdad mencionen ese radicado.

**Regla obligatoria, sin excepción, para TODO lo que se descargue**
(sea confiable o no, radicado completo o corto, carpeta de Drive o
adjunto de correo): el documento tiene que mencionar a **ESSA** o
**ELECTRIFICADORA DE SANTANDER** (como demandante o como demandado --
por nombre, o abriendo el contenido de sus PDF/DOCX si el nombre no lo
dice). Si no la menciona, no se descarga -- ni siquiera si el radicado
coincidió exacto. Un adjunto de correo que no pase esta regla se
extrae igual, pero a `Duplicados_para_revisar` en vez de a la carpeta
del proceso (nunca se pierde, solo no se mezcla con el caso).

**También se valida el DEMANDADO** (columna `DEMANDADO` del Excel, si
existe): el mismo radicado corto o cuenta se puede repetir entre
procesos **distintos** que van contra demandados diferentes -- por
ejemplo, `"2024-00139 CONTRA RIONEGRO"` no es lo mismo que `"2024-00139
CONTRA BOLIVAR"`. Con que el nombre del demandado esperado aparezca en
**cualquier parte** del nombre/asunto/contenido del candidato alcanza
para confirmarlo -- no hace falta que esté después de la palabra
"CONTRA". Solo se descarta un candidato (una carpeta, un archivo dentro
de una carpeta genérica, o un adjunto de correo) cuando sí trae un
`"CONTRA <algo>"` (la forma habitual de nombrar expedientes) pero ese
"algo" es un demandado **distinto** al esperado -- aunque ya haya
pasado la validación de ESSA. Si el candidato no menciona ni al
demandado esperado ni ningún otro `"CONTRA <algo>"`, esta validación no
bloquea nada (no hay evidencia ni a favor ni en contra).

Los candidatos que sí pasan **también se descargan automático**. Si para el mismo proceso aparece
MÁS de un candidato válido (ej. el expediente está repartido en varias
carpetas de Drive -- una con el "poder", otra con el "expediente"),
todos quedan **fusionados dentro de UNA sola carpeta** en el disco (no
se crean "_2", "_3", etc): el primer candidato crea la carpeta, y cada
candidato siguiente que también pase las validaciones se copia dentro
de esa misma carpeta. Como cada candidato es una carpeta de Drive
DISTINTA, puede traer -- por pura coincidencia -- un archivo o
subcarpeta con el mismo nombre que otro candidato (ej. dos carpetas
"PRINCIPAL" de dos procesos distintos, o dos archivos "01. INFORME 1")
sin ser el mismo documento; en ese caso **nunca se reemplaza uno con el
otro** -- el que llega se guarda con un sufijo (ej. "PRINCIPAL_2") para
quedarse con ambos, y nunca se borra nada que ya estuviera.
La diferencia con el caso 1 es que estos quedan marcados aparte en
`faltantes_descargados_a_validar.csv` (una fila por cada candidato que
se fusionó), para que después confirmes que todo corresponde al mismo
proceso y borres a mano lo que no corresponda (el script nunca borra
nada por su cuenta). Si lo que encuentra es un archivo suelto (no una
carpeta), busca la carpeta que lo contiene. Si esa carpeta es "propia"
del caso (su nombre menciona el radicado, o la búsqueda encontró
directamente la carpeta) descarga la carpeta completa. Pero si es una
carpeta **genérica** -- de "informes" o "actuaciones" que junta
documentos de varios procesos, y el archivo que coincidió es solo uno
más ahí adentro -- **no** descarga la carpeta completa (traería folios
de otros procesos sin relación); solo baja los archivos de esa carpeta
que de verdad mencionen el radicado (y, si aplica, al demandante ESSA).

**Orden cronológico automático:** al terminar de descargar/fusionar
cada proceso, sus documentos quedan ordenados por FECHA y numerados
`1. `, `2. `, etc (el más viejo primero). La fecha se busca, en este
orden:

1. En el NOMBRE del archivo -- reconoce varios formatos: ISO
   (`2023-07-24`), `DD/MM/AAAA`, `"24 de julio de 2023"`, `"24 JULIO
   2023"` (sin la palabra "de" -- el formato más común en los nombres
   reales de autos/providencias) y `"Jul 24, 2023"`.
2. Si el nombre no trae fecha, se abren sus primeras páginas (PDF/DOCX)
   para buscarla ahí -- pero como máximo en 20 archivos por carpeta,
   para no quedarse abriendo PDF tras PDF en una carpeta grande (ej.
   fusionada de varios candidatos, con 70+ archivos).
3. Si tampoco hay fecha en el contenido, como último recurso se usa la
   fecha de modificación del propio archivo en el disco -- pero SOLO si
   es de más de un día atrás (para no confundir la hora en que se
   acaba de copiar/extraer con una fecha real). Un documento que llega
   por **adjunto de correo** y no trae fecha propia queda con la fecha
   del correo puesta como fecha de modificación, así que ese es el
   valor que se usa aquí -- mejor guiarse por cuándo llegó el correo
   que dejarlo sin ningún orden.

Los documentos sin NINGUNA fecha reconocible (ni nombre, ni contenido,
ni fecha de correo) quedan al final, en el orden en que ya estaban. Si
el proceso está repartido en subcarpetas (ej. "principal" y "anexos"),
cada una se ordena por su cuenta, sin mezclar los documentos de una con
los de otra. Si vuelves a correr el script y llega un documento más
viejo que los demás, el orden se recalcula solo. Además de las
carpetas que se acaban de descargar/fusionar en esta corrida, **al
empezar cada corrida también se revisan y ordenan TODAS las carpetas
de proceso que ya existan en el disco** (de esta corrida o de
cualquier corrida anterior) -- no hace falta que el proceso se haya
tocado hoy para que quede numerado.

**Solo se descargan archivos PDF:** cualquier otro tipo de archivo
(Word, Excel, imágenes, etc) que aparezca junto a los PDF -- ya sea
dentro de una carpeta de Drive o de un adjunto `.zip` de correo -- se
omite y **nunca** se baja al disco. La única excepción son los
archivos nativos de Google (Doc, Sheet, Slide), que sí se descargan
pero **exportados como PDF** (Drive los convierte automáticamente al
exportarlos, así que igual terminan como `.pdf` en el disco).

**Si un archivo puntual falla al descargarlo** (ruta demasiado larga
para Windows, permisos, antivirus, un corte de red momentáneo), el
script lo salta con una advertencia en el log y sigue con el resto --
un solo archivo problemático nunca detiene la descarga de ese proceso,
ni mucho menos el resto de la lista. Lo mismo si un proceso entero
falla por algo inesperado: se registra el error y se sigue con el
siguiente, en vez de detener toda la corrida a mitad de camino.

**Carpetas temporales sueltas:** lo primero que hace el script al
empezar es revisar si quedaron carpetas `_tmp_fusion_...`,
`_tmp_extraccion_correo_...` (u otras sueltas dentro de
`_tmp_extraccion`) de una corrida anterior que se cerró a la mitad, o
que no se pudieron borrar solas (por ejemplo, un archivo adentro
bloqueado por el antivirus o por OneDrive justo en ese momento). Como
nunca se sabe con certeza si todo su contenido ya quedó copiado en su
carpeta final, **nunca se borran solas**: se mueven a
`Duplicados_para_revisar` para que las revises tú.

**Limpieza automática de corridas anteriores:** al empezar, el script
también revisa si ya quedaron carpetas "_2", "_3", etc en el disco de
corridas viejas (por ejemplo, de antes de que existiera el filtro de
demandante ESSA). Revalida cada una: si menciona a ESSA (por nombre o
contenido), la fusiona dentro de su carpeta principal (sin perder
archivos con nombres repetidos); si no la menciona, la mueve tal cual
a `Duplicados_para_revisar` -- nunca la borra -- para que la revises a
mano, porque probablemente es ruido de otro proceso que compartía
cuenta o radicado corto. **Importante:** esto solo aplica cuando de
verdad hay una carpeta vieja con sufijo `"_N"` -- si el mismo radicado
aparece con **dos o más números de proceso distintos y legítimos**
(porque el Excel tiene ese radicado repetido con varios números, y
`validar_renombrar_carpetas.py` a propósito dejó una carpeta separada
por cada uno), esas carpetas **no se tocan ni se fusionan entre sí**:
fusionarlas destruiría esa separación intencional.

**Revisión de contaminación entre procesos:** también al empezar,
revisa TODAS las carpetas ya descargadas (de cualquier corrida,
incluidas las de hace tiempo) buscando archivos cuyo nombre mencione
un radicado corto **distinto** al de su propia carpeta -- rastro de un
bug ya corregido donde, por ejemplo, `"2023-24"` se confundía con
`"2023-244"` (otro proceso, solo comparte el prefijo numérico) y
terminaban documentos de un caso mezclados en la carpeta del otro. Un
archivo solo se mueve si menciona OTRO radicado y **nunca** el propio
(si cita ambos, se asume que es un documento legítimo que solo
referencia un caso relacionado, y no se toca). Los archivos
sospechosos se mueven a `Duplicados_para_revisar` -- nunca se borran.
Si al sacarlos una carpeta queda completamente vacía (todo su
contenido era de otro proceso), esa carpeta vacía sí se borra, para
que la próxima corrida la vuelva a buscar desde cero con el filtro ya
corregido.

**Revisión de demandado entre procesos:** de la misma forma, también
al empezar, revisa TODAS las carpetas ya descargadas buscando archivos
con un `"CONTRA <algo>"` que no corresponda al **demandado real** del
proceso según el Excel -- mismo rastro que la revisión de
contaminación, pero para el caso de dos procesos que comparten
radicado corto/cuenta y van contra demandados **distintos**. Esos
archivos también se mueven a `Duplicados_para_revisar` -- nunca se
borran. Si es la **carpeta misma** la que parece tener el demandado
equivocado en su nombre, solo se avisa en el log (no se mueve ni
renombra la carpeta sola) para que la revises a mano -- podría
significar que quedó cruzada con el radicado equivocado desde el
principio. Si el Excel no tiene la columna `DEMANDADO`, esta revisión
simplemente se omite.

### Configurar el acceso a Google Drive (una sola vez)

1. Ve a [console.cloud.google.com](https://console.cloud.google.com/) y
   crea un proyecto nuevo (o usa uno existente) -- es gratis.
2. En el menú, ve a **APIs y servicios → Biblioteca**, busca **Google
   Drive API** y dale **Habilitar**.
3. Ve a **APIs y servicios → Credenciales → Crear credenciales → ID de
   cliente de OAuth**. Si te pide configurar antes la "pantalla de
   consentimiento", elige tipo **Externo**, pon cualquier nombre, y
   agrégate a ti misma como "usuario de prueba" (no hace falta publicarla).
4. Tipo de aplicación: **Aplicación de escritorio**. Créala y descarga el
   JSON.
5. Renombra ese archivo a `credenciales_drive.json` y ponlo en la misma
   carpeta que los demás scripts.
6. La primera vez que corras `buscar_faltantes_en_drive.py`, se abre el
   navegador pidiendo que autorices el acceso con tu cuenta de Google.
   Acepta, y queda guardado `token_drive.json` para las próximas veces
   (no hay que repetir esto).

Para que también busque en tu correo, usa el mismo `credenciales_sgde.txt`
que ya tienes configurado para `procesos_juridicos.py` -- si no existe,
esa búsqueda simplemente se omite.

Uso:

1. Corre primero `iniciar.bat` (o `validar_renombrar_carpetas.py`) para
   que `procesos_faltantes_en_disco.csv` esté al día.
2. Haz doble clic en `buscar_drive.bat` (o corre `python
   buscar_faltantes_en_drive.py` desde una terminal en esa carpeta).
3. Por defecto corre en `MODO_PRUEBA = True` (solo busca y te dice qué
   descargaría). Revisa el log, y cuando confíes en el resultado cambia
   `MODO_PRUEBA = False` para descargar de verdad. Después de correrlo
   revisa `faltantes_descargados_a_validar.csv`: son las carpetas que se
   descargaron por una coincidencia menos segura (radicado corto, cuenta,
   o un enlace de correo sin el radicado completo) -- confírmalas y
   borra a mano las que no correspondan.

Notas sobre las búsquedas por radicado corto/cuenta (las menos
confiables): Google Drive no busca por texto exacto, busca por
*prefijo de palabra* (buscar `"2014-26"` también puede traer `"26
julio"`) -- el script filtra esos falsos positivos solos antes de
descargar nada. Ese mismo filtro también evita que un radicado corto
como `"2023-24"` se confunda con uno DISTINTO que solo comparte el
prefijo (ej. `"2023-244"` o `"2023-248"`) -- si justo al lado de la
coincidencia hay otro dígito, no cuenta como coincidencia real. Además,
antes de descargar cualquier candidato que solo coincidió por radicado
corto o cuenta, se verifican dos cosas:

- Que la carpeta candidata (o alguno de sus archivos) de verdad
  mencione ESE radicado -- así, si el mismo número de cuenta aparece en
  varios procesos distintos del mismo cliente a lo largo de los años,
  solo se descarga la carpeta que en realidad corresponde a este caso.
- Que el demandante sea **ESSA/Electrificadora de Santander** -- se
  revisa primero el nombre de la carpeta y de sus archivos; si ninguno
  lo dice, se abre el contenido de hasta 5 PDF/DOCX como muestra. Así,
  si la cuenta o el radicado corto coincide con un proceso de OTRO
  cliente, esa carpeta no se descarga.

Si un proceso ya tiene una carpeta en el disco (por ejemplo porque una
corrida anterior ya lo descargó), el script lo omite por completo sin
buscar ni descargar nada -- para no crear carpetas "_2" duplicadas si
se vuelve a correr sobre un `procesos_faltantes_en_disco.csv`
desactualizado.

## Validar solo las carpetas de la lista de faltantes

`validar_procesos_faltantes.py` es un script aparte, más liviano, para
cuando solo quieres revisar rápido las carpetas de los procesos que
están en `procesos_faltantes_en_disco.csv` (los que ya bajó
`buscar_faltantes_en_drive.py`, u otros que hayas agregado a mano) --
sin esperar a que se revise el disco completo.

Además de `procesos_faltantes_en_disco.csv`, también cruza contra
`carpetas_vacias.csv` (el otro reporte que genera
`validar_renombrar_carpetas.py`): si una carpeta de la lista ya existe
en el disco pero está **completamente vacía** (sin ningún archivo
adentro), se avisa aparte con claridad -- incluyendo si ese reporte ya
había encontrado un `.zip` pendiente en Descargas para ella -- en vez
de tratarla en silencio como si tuviera contenido para ordenar. Una
carpeta vacía **no se toca** (no hay nada que ordenar ni que borrar en
ella).

Para cada proceso de esa lista que **ya tenga carpeta con contenido**
en el disco:

**Por defecto** (`BORRAR_ARCHIVOS_DE_OTRO_PROCESO = False`), lo
**único** que hace es **ordenar cronológicamente** los documentos que
YA están adentro de cada carpeta, numerándolos `1. `, `2. `, etc -- no
mueve ni borra nada, no fusiona nada, no toca ninguna otra carpeta del
disco.

Si además quieres que **borre** los archivos que parezcan de OTRO
proceso -- porque mencionan un radicado corto distinto al de su propia
carpeta, o porque tienen un `"CONTRA <algo>"` que no corresponde al
demandado real según el Excel -- pon
`BORRAR_ARCHIVOS_DE_OTRO_PROCESO = True` al inicio del script.

**⚠️ A diferencia de TODOS los demás scripts de este proyecto** (que
nunca borran nada, solo mueven a `Duplicados_para_revisar`), con
`BORRAR_ARCHIVOS_DE_OTRO_PROCESO = True` los archivos que no coincidan
se **borran de forma permanente e irreversible** -- no quedan en
`Duplicados_para_revisar`, no se pueden recuperar. Esto se pidió así a
propósito, para no acumular carpetas de revisión manual; revisa con
calma el reporte en `MODO_PRUEBA = True` antes de correrlo con
`MODO_PRUEBA = False`. Lo que sí se mantiene igual que en el resto del
proyecto es que la **carpeta en sí nunca se mueve, renombra, ni se
fusiona con otra** -- solo se borran archivos puntuales adentro.

Los procesos de la lista que **todavía no tengan carpeta** en el disco
se cuentan aparte y se omiten -- este script **no descarga nada** (para
eso está `buscar_faltantes_en_drive.py`). Las carpetas que **no** estén
en la lista de faltantes tampoco se tocan, aunque tengan el mismo tipo
de problemas -- este script es deliberadamente angosto, solo mira la
lista.

Uso:

1. Corre primero `validar_renombrar_carpetas.py` para que
   `procesos_faltantes_en_disco.csv` esté al día.
2. Haz doble clic en `validar_faltantes.bat` (o corre `python
   validar_procesos_faltantes.py` desde una terminal en esa carpeta).
3. Por defecto corre en `MODO_PRUEBA = True` (solo revisa y te dice qué
   ordenaría/borraría). Cambia `MODO_PRUEBA = False` al inicio del
   script para aplicar los cambios de verdad.

Todo queda registrado en `validar_procesos_faltantes.log`.

## Crear carpetas de procesos terminados/remitidos/no iniciados

`validar_renombrar_carpetas.py` solo organiza los procesos en los
estados de `ESTADOS_A_CONTAR` (`ACTIVO`, `ACTIVOS CON TITULOS`,
`SUSPENDIDO`, `REORGANIZACION`). Para los procesos en cualquier OTRO
estado -- terminados (por pago, por auto, por prepago, con contrato,
etc), remitidos (a castigo, a prepago, etc), o que nunca llegaron a
iniciarse -- usa `crear_carpetas_terminados_castigo.py`.

Estos procesos normalmente no tienen (o no importa) un radicado real
para organizar documentos -- son más un registro administrativo que un
expediente judicial activo -- así que la carpeta **no** se nombra
`"numero. radicado"` como el resto del proyecto, sino:

- Si el `ESTADO PROCESAL` empieza con **`TERMINADO`** (por pago, por
  auto, por prepago, con contrato, etc): `"<numero>. <ESTADO PROCESAL
  EXACTO del Excel>"` -- ej. `"123. TERMINADO POR AUTO"`.
- Si el `ESTADO PROCESAL` empieza con **`REMITIDA`** (a castigo, a
  prepago, etc) o con **`NO INICIO`**: **solo el número** -- ej.
  `"145."`.

Los procesos con un `ESTADO PROCESAL` que no encaje en ninguna de esas
dos reglas (ej. `"DESISTIMIENTO DE PRETENSIONES"`, `"DEVUELTA INCURRIO
EN GASTOS"`) se dejan **fuera a propósito** -- no se crea carpeta para
ellos, y quedan listados en el log para que decidas qué hacer.

Si **ya existe** una carpeta en el disco para ese número (cualquier
nombre que empiece por `"<numero>. "` o sea exactamente `"<numero>."`
-- por ejemplo porque el proceso ya se organizó de la forma normal con
su radicado real), **no se crea una nueva**. Este script nunca borra,
renombra ni mueve nada -- solo **crea** carpetas vacías nuevas donde
todavía no exista ninguna para ese número. Correrlo varias veces no
duplica nada.

Uso:

1. Haz doble clic en `crear_carpetas_terminados_castigo.bat` (o corre
   `python crear_carpetas_terminados_castigo.py`).
2. Por defecto corre en `MODO_PRUEBA = True` (solo revisa y te dice qué
   carpetas crearía). Cambia `MODO_PRUEBA = False` al inicio del script
   para crearlas de verdad.

Usa la misma configuración (`RUTA_EXCEL`, `CARPETA_PROCESOS`, etc) que
`validar_renombrar_carpetas.py` -- no hay que configurarla dos veces.
Todo queda registrado en `crear_carpetas_terminados_castigo.log`.

### Por qué el total de carpetas puede no coincidir con el total de filas del Excel

Además de crear, el script audita **todas** las filas del Excel (no
solo las que va a crear) para explicar cualquier diferencia entre el
número de filas del Excel y el número de carpetas en el disco. Cada
número de proceso sin carpeta cae en una de estas categorías, y todas
quedan reportadas en el log con su fila y detalle:

- **Número duplicado en el Excel**: el mismo número de proceso aparece
  en más de una fila (con radicado y/o estado distintos). No se crea
  ni se toca nada -- hay que corregir el Excel a mano, porque no se
  puede adivinar cuál fila es la correcta.
- **Activo/Suspendido/Reorganización con radicado**: ya está
  rastreado en el listado de procesos faltantes
  (`procesos_faltantes_en_disco.csv`) -- a propósito **no** se crea
  una carpeta vacía aquí, porque ese proceso lo organiza
  `buscar_faltantes_en_drive.py` buscando su contenido real. Si
  quieres esa carpeta, corre ese script en vez de este.
- **Activo/Suspendido/Reorganización sin radicado**: el radicado
  todavía no está diligenciado en el Excel -- hay que llenarlo antes
  de poder buscar o crear nada para ese proceso.
- **Sin estado procesal diligenciado**: la fila tiene número de
  proceso pero la columna `ESTADO PROCESAL` está vacía todavía.
- **Estado sin clasificar**: no empieza con `TERMINADO`, `REMITIDA` ni
  `NO INICIO` (ver arriba).

### Deshacer: borrar_carpetas_terminados_castigo.py

Si quieres borrar las carpetas vacías de procesos que no son activos,
usa `borrar_carpetas_terminados_castigo.py` (o su iniciador
`borrar_carpetas_terminados_castigo.bat`):

- Para cada número de proceso que **no** esté en `ESTADOS_A_CONTAR`
  (o sea, cualquiera que no sea `ACTIVO`, `ACTIVOS CON TITULOS`,
  `SUSPENDIDO` o `REORGANIZACION`) **y tenga algún `ESTADO PROCESAL`
  diligenciado**, borra su carpeta si está **completamente vacía** --
  sin importar cómo se llame exactamente (`"numero. ESTADO"`, solo
  `"numero."`, o incluso `"numero. radicado"`) ni si el estado es uno
  de los que clasifica `crear_carpetas_terminados_castigo.py`
  (`TERMINADO*`/`REMITIDA*`/`NO INICIO*`) o uno sin clasificar (ej.
  `DESISTIMIENTO`, `DEVUELTA`).
- **Nunca** toca procesos activos/suspendidos/reorganización, filas
  sin `ESTADO PROCESAL` diligenciado todavía, números duplicados en el
  Excel, ni ninguna carpeta que ya tenga contenido adentro (por
  ejemplo porque le agregaste documentos a mano) -- esas se reportan
  en el log para que decidas a mano.
- ⚠️ Borrar una carpeta es **irreversible**. Respeta `MODO_PRUEBA`
  (por defecto `True`): revisa el log primero, y solo cambia
  `MODO_PRUEBA = False` cuando estés segura de que la lista de
  carpetas a borrar es la correcta.

## Comparar Excel vs disco (solo un reporte de lo que falta o sobra)

`comparar_excel_disco.py` (o su iniciador `comparar_excel_disco.bat`)
compara, por **radicado**, los procesos del Excel contra las carpetas
que existen en el disco, en las **dos direcciones**:

1. Procesos del Excel que **no** tienen ninguna carpeta en el disco --
   reporte `comparar_excel_disco_faltantes.csv` (fila del Excel,
   cuenta, radicado, juzgado, demandado y estado, cuando esas columnas
   existen).
2. Carpetas del disco cuyo radicado **no** aparece en el Excel --
   reporte `comparar_excel_disco_sobran_en_disco.csv` (nombre de la
   carpeta y radicado) -- útil para detectar carpetas de procesos que
   ya no están vigentes o con el radicado mal escrito.

- Dos radicados se consideran el **mismo proceso** si son idénticos, o
  si solo difieren en el último dígito (el consecutivo de
  instancia/reparto, ej. termina en 0 o en 1) -- la misma tolerancia
  que usa `validar_renombrar_carpetas.py`.
- Es un script **de solo lectura**: nunca mueve, renombra, crea ni
  borra nada, solo compara y reporta. Es el más rápido de correr
  cuando solo quieres saber qué falta o qué sobra, sin tocar ninguna
  carpeta.
- Solo necesita la columna `RADICADO` del Excel para comparar -- no
  depende de ninguna columna de número de proceso, así que funciona
  incluso con informes que no tengan una columna `"No."`.
- Usa la misma configuración (`RUTA_EXCEL`, `CARPETA_PROCESOS`, etc)
  de `validar_renombrar_carpetas.py` -- no hay que configurarla dos
  veces.

## Borrar una carpeta que Windows no deja borrar

Si el Explorador de Windows se niega a borrar una carpeta (rutas muy
largas, metadatos raros, o simplemente se traba) usa
`borrar_carpeta_forzado.bat`:

- Arrastra la carpeta problemática sobre el `.bat` (o dale doble clic y
  pega la ruta completa cuando te la pida).
- Te pide confirmación escribiendo `SI` antes de borrar nada.
- Primero vacía el contenido con `robocopy` (que sí sabe manejar rutas
  largas, a diferencia del Explorador) y luego borra la carpeta ya
  vacía.
- ⚠️ Es **irreversible** -- solo úsalo cuando estés segura de que
  quieres borrar esa carpeta completa y todo su contenido.
- Si después de correrlo la carpeta sigue sin borrarse del todo, es
  que algún archivo de adentro está abierto en otro programa (Word,
  Excel, un PDF, el antivirus escaneándola) -- ciérralo e inténtalo de
  nuevo.

## Si algo falla

- El sitio del SGDE puede cambiar de diseño con el tiempo, lo que puede
  romper los selectores que usa Playwright. Si ves errores en el log,
  dime qué cambió en la pantalla y te ayudo a actualizar el script.
- Si un proceso descargado a mano no tiene ningún documento con el
  radicado en el formato esperado, la carpeta se organiza igual usando
  el nombre del zip original, y queda registrado como advertencia en el
  log.
- Si un zip no deja **ningún** archivo al extraerlo (por ejemplo porque
  todo su contenido está protegido con contraseña, las rutas son
  demasiado largas para Windows, o el antivirus puso los archivos en
  cuarentena justo después de extraerlos), el programa avisa con un
  `ERROR` bien visible en el log, **no** crea una carpeta vacía
  disfrazada de "organizada", y **no** mueve el zip a `Procesados` — se
  queda en Descargas para que lo revises o reintentes a mano. Si algunos
  archivos sí se extrajeron pero otros no, la carpeta se organiza igual
  con lo que se pudo, y queda un `WARNING` explicando cuántos fallaron.
- Al extraer un zip, si una carpeta o archivo dentro trae **espacios o
  puntos al final del nombre** (Windows los maneja mal incluso con el
  truco de rutas largas), se le quitan automáticamente al crearlo en el
  disco -- así se evita el error clásico "el sistema no puede encontrar
  la ruta especificada" al leerlo después.
- Si `procesos_faltantes_en_disco.csv` (o los otros reportes CSV) están
  abiertos en Excel cuando corres `validar_renombrar_carpetas.py`, el
  script ya no crashea al intentar guardarlos -- avisa con un `ERROR`
  claro pidiéndote que cierres el archivo, y el resto del análisis
  (que sí se calculó bien) queda visible igual en el log.
- **Importante sobre carpetas vacías de ANTES de esta corrección**: si ya
  tienes carpetas vacías de cuando el programa sí las creaba aunque la
  extracción fallara, `procesos_juridicos.py` **no las va a arreglar
  solo**, aunque subas `DIAS_ATRAS_PROCESAR_EXISTENTES` — el zip
  correspondiente probablemente ya se movió a `Descargas\Procesados`
  (porque el programa viejo lo marcaba como "hecho" sin comprobar), y la
  vigilancia automática nunca mira dentro de esa subcarpeta a propósito
  (para no reprocesar cosas ya hechas de verdad). Corre
  `validar_renombrar_carpetas.py`, revisa `carpetas_vacias.csv` (columna
  "Donde se encontró"): si dice `Descargas/Procesados`, tienes que sacar
  ese zip de ahí a mano (muévelo de vuelta a Descargas) para que se
  vuelva a intentar.
