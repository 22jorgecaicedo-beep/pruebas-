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

Haz doble clic en `iniciar.bat`, o desde una terminal:

```
python procesos_juridicos.py
```

La primera vez, deja `NAVEGADOR_VISIBLE = True` (así viene por defecto)
para ver el navegador trabajando y confirmar que todo va bien. Cuando
confíes en que funciona, cámbialo a `False` en `procesos_juridicos.py`
para que corra en segundo plano sin abrir ventanas.

Déjalo abierto; ciérralo con `Ctrl+C` (o cerrando la ventana) cuando
termines.

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

## Si algo falla

- El sitio del SGDE puede cambiar de diseño con el tiempo, lo que puede
  romper los selectores que usa Playwright. Si ves errores en el log,
  dime qué cambió en la pantalla y te ayudo a actualizar el script.
- Si un proceso descargado a mano no tiene ningún documento con el
  radicado en el formato esperado, la carpeta se organiza igual usando
  el nombre del zip original, y queda registrado como advertencia en el
  log.
