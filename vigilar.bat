@echo off
cd /d "%~dp0"

echo ============================================================
echo  Vigilando Descargas (y correo, si esta configurado) de forma
echo  indefinida (procesos_juridicos.py) -- sin esperar a que
echo  termine validar_renombrar_carpetas.py.
echo  Cierra esta ventana o Ctrl+C cuando quieras detenerlo.
echo ============================================================
python procesos_juridicos.py

pause
