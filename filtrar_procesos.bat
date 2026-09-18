@echo off
cd /d "%~dp0"
python -m pip install --quiet openpyxl
python filtrar_procesos.py %*
pause
