@echo off
cd /d "%~dp0"

echo ============================================================
echo  Paso 1: revisando/renombrando las carpetas que ya existen
echo  en el disco (validar_renombrar_carpetas.py)...
echo ============================================================
python validar_renombrar_carpetas.py

echo.
echo ============================================================
echo  Paso 2: vigilando Descargas (y correo, si esta configurado)
echo  de forma indefinida (procesos_juridicos.py)...
echo  Cierra esta ventana o Ctrl+C cuando quieras detenerlo.
echo ============================================================
python procesos_juridicos.py

pause
