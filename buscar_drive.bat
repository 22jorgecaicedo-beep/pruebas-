@echo off
cd /d "%~dp0"

echo ============================================================
echo  Buscando en Google Drive/correo los procesos faltantes
echo  (buscar_faltantes_en_drive.py)...
echo ============================================================
python buscar_faltantes_en_drive.py

pause
