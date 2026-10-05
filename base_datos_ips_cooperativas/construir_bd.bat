@echo off
cd /d "%~dp0"
python -m pip install --quiet -r requirements.txt
python construir_bd.py %*
pause
