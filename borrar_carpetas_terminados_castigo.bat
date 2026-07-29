@echo off
cd /d "%~dp0"

echo ============================================================
echo  Borrando carpetas vacias de TERMINADOS/REMITIDOS/NO INICIO
echo  (borrar_carpetas_terminados_castigo.py)...
echo ============================================================
python borrar_carpetas_terminados_castigo.py

pause
