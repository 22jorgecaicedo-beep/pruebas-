@echo off
cd /d "%~dp0"

echo ============================================================
echo  Validando SOLO las carpetas de la lista de faltantes
echo  (validar_procesos_faltantes.py)...
echo ============================================================
python validar_procesos_faltantes.py

pause
