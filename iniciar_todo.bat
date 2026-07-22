@echo off
cd /d "%~dp0"
start "Organizador de zips" cmd /k python organizador_zips.py
start "Descargador SGDE" cmd /k python descargador_sgde.py
