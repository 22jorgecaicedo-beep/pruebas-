# Organizador de ZIP de procesos jurídicos

Este repositorio tiene dos herramientas que se complementan:

1. **`descargador_sgde.py`**: vigila tu Gmail, y cuando un juzgado te
   comparte un proceso a través del SGDE (Rama Judicial), inicia sesión
   automáticamente en el portal, resuelve el token de verificación y
   descarga los archivos a tu carpeta de Descargas.
2. **`organizador_zips.py`**: vigila esa misma carpeta de Descargas y
   organiza cada `.zip` que llegue (manual o automáticamente) en tu disco
   duro.

Puedes usar solo el organizador (si prefieres descargar manualmente), o
las dos juntas para automatizar todo el proceso.

## Parte 1: Organizador de ZIP (`organizador_zips.py`)

Vigila una carpeta de descargas y, cada vez que llega un `.zip` nuevo:

1. Espera a que la descarga termine.
2. Lo extrae.
3. Busca dentro de los PDF/DOCX (o en los nombres de archivo) un número de
   radicado judicial.
4. Renombra la carpeta extraída con ese radicado (o con el nombre del zip
   si no encuentra ninguno) y la mueve al disco duro destino.
5. Mueve el `.zip` original a una subcarpeta `Procesados` dentro de la
   carpeta de descargas (no lo borra).

Todo queda registrado en `organizador_zips.log`, dentro de la carpeta
destino.

## Instalación (Windows)

1. Instala [Python 3.10+](https://www.python.org/downloads/) marcando la
   opción "Add Python to PATH" durante la instalación.
2. Abre una terminal (`cmd` o PowerShell) en esta carpeta y ejecuta:

   ```
   pip install -r requirements.txt
   ```

## Configuración

Abre `organizador_zips.py` y edita, al inicio del archivo, la sección
`CONFIGURACION`:

- `CARPETA_DESCARGAS`: carpeta donde el navegador guarda los `.zip`
  (por ejemplo `C:\Users\TuUsuario\Downloads`).
- `CARPETA_DESTINO`: carpeta en el disco duro donde se organizarán los
  procesos (por ejemplo `D:\ProcesosJuridicos`).
- `PATRONES_RADICADO`: expresiones regulares para reconocer el número de
  proceso. Por defecto reconoce el radicado judicial colombiano de 23
  dígitos, con o sin guiones. Agrega más patrones si tus documentos usan
  otro formato.

## Uso

Haz doble clic en `iniciar_organizador.bat`, o desde una terminal:

```
python organizador_zips.py
```

El programa primero procesa los `.zip` que ya estén en la carpeta de
descargas y luego se queda vigilando en segundo plano. Déjalo abierto
mientras descargas los procesos; ciérralo con `Ctrl+C` (o cerrando la
ventana) cuando termines.

## Ejecutarlo automáticamente al iniciar Windows (opcional)

1. Abre el "Programador de tareas" de Windows.
2. Crea una tarea nueva que se ejecute "Al iniciar sesión".
3. Como acción, apunta a `iniciar_organizador.bat` (o a `pythonw.exe
   organizador_zips.py` si prefieres que corra sin ventana visible).

## Notas del organizador

- Si un proceso no tiene ningún documento con el radicado en el formato
  esperado, la carpeta se organiza igual, usando el nombre del zip
  original, y queda registrado como advertencia en el log.
- Si ya existe una carpeta con el mismo nombre en el destino, se agrega
  un sufijo (`_2`, `_3`, ...) en vez de sobrescribir.

## Parte 2: Descargador automático del SGDE (`descargador_sgde.py`)

Automatiza la descarga de los procesos que los juzgados comparten por
correo a través del portal `siugj-sgde.ramajudicial.gov.co`.

### Antes de usarlo

⚠️ Verifica que automatizar el acceso a este portal con tu cuenta
autorizada no infrinja los términos de uso del SGDE / Rama Judicial.
Este script solo automatiza pasos que tú ya estás autorizado a hacer
manualmente (abrir el link que te comparten, escribir tu correo y el
token que te llega a ti mismo); no intenta acceder a expedientes ajenos
ni evade ninguna medida de seguridad.

### Instalación adicional

Además de `pip install -r requirements.txt` (que ya incluye Playwright),
instala el navegador que usará Playwright:

```
playwright install chromium
```

### Configurar tus credenciales de Gmail

El script necesita leer tu correo para conseguir el link del proceso y
el token, usando IMAP. **Nunca uses tu contraseña normal de Gmail** —
usa una "Contraseña de aplicación":

1. Activa la verificación en 2 pasos: https://myaccount.google.com/security
2. Crea una contraseña de aplicación: https://myaccount.google.com/apppasswords
3. En esta carpeta, copia `credenciales_sgde.example.txt` y renómbralo a
   `credenciales_sgde.txt`.
4. Ábrelo y reemplaza los valores por tu correo y la contraseña de
   aplicación de 16 letras que Google te dio.

`credenciales_sgde.txt` está en `.gitignore`: si en algún momento subes
este proyecto a GitHub, ese archivo con tus credenciales reales **no**
se sube. Nunca lo compartas ni lo subas a ningún lado manualmente.

### Uso

```
python descargador_sgde.py
```

O usa `iniciar_todo.bat` para arrancar el descargador y el organizador
de zips juntos, cada uno en su propia ventana.

La primera vez, deja `NAVEGADOR_VISIBLE = True` (ya viene así) dentro de
`descargador_sgde.py` para ver el navegador trabajando y confirmar que
todo va bien. Cuando confíes en que funciona, puedes ponerlo en `False`
para que corra en segundo plano sin abrir ventanas.

### Cómo evita descargar el mismo proceso dos veces

Cada expediente ya descargado se guarda en `expedientes_procesados.txt`
(tampoco se sube a git). Si necesitas volver a descargar uno, borra esa
línea del archivo o el archivo completo.

### Si algo falla

El sitio del SGDE puede cambiar de diseño con el tiempo, lo que puede
romper los selectores que usa Playwright para encontrar los botones.
Si ves errores en `descargador_sgde.log` (dentro de la carpeta destino),
lo más probable es que necesites ajustar los textos/selectores en
`descargador_sgde.py` para que coincidan con la pantalla actual del
portal. En ese caso, dime qué cambió y te ayudo a actualizarlo.
