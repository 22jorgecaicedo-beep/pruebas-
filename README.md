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
- `CARPETA_DESTINO`: carpeta en el disco duro donde se organizarán los
  procesos (por ejemplo `D:\ProcesosJuridicos`).
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
últimos `DIAS_ATRAS_PROCESAR_EXISTENTES` días (por defecto 7, es decir la
última semana) que ya estén en `CARPETA_DESCARGAS` (igual que hace
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
   - `CARPETA_PROCESOS`: carpeta del disco duro donde están las carpetas de
     cada proceso.
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
     carpeta que la contiene, se mueve a `Duplicados_para_revisar`. Si es
     un radicado distinto (contenido de otro caso mal ubicado), se saca
     al nivel principal del disco para evaluarla en la próxima corrida.
   - `[Carpeta vacía]`: carpetas sin ningún archivo adentro. El script
     revisa si hay un `.zip` en `CARPETA_DESCARGAS` con ese mismo
     radicado (por si quedó pendiente de extraer) y te lo señala — esto
     solo se reporta, no se extrae ni se toca nada automático. Además de
     salir en el log, queda un reporte aparte en `carpetas_vacias.csv`
     (carpeta, radicado, y el zip pendiente si lo encontró) para que lo
     revises en Excel sin tener que buscar en el log completo.
   - `[Sin nombre reconocible]`: carpetas que no tienen ningún número que
     se parezca a un radicado en su nombre (antes se ignoraban en
     silencio, ahora se listan para que las revises a mano).

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

## Si algo falla

- El sitio del SGDE puede cambiar de diseño con el tiempo, lo que puede
  romper los selectores que usa Playwright. Si ves errores en el log,
  dime qué cambió en la pantalla y te ayudo a actualizar el script.
- Si un proceso descargado a mano no tiene ningún documento con el
  radicado en el formato esperado, la carpeta se organiza igual usando
  el nombre del zip original, y queda registrado como advertencia en el
  log.
