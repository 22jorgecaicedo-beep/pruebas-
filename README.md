# Organizador de ZIP de procesos jurídicos

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

## Notas

- Si un proceso no tiene ningún documento con el radicado en el formato
  esperado, la carpeta se organiza igual, usando el nombre del zip
  original, y queda registrado como advertencia en el log.
- Si ya existe una carpeta con el mismo nombre en el destino, se agrega
  un sufijo (`_2`, `_3`, ...) en vez de sobrescribir.
