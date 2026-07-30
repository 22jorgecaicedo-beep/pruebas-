@echo off
setlocal enabledelayedexpansion

echo ============================================================
echo  Borrar carpeta atascada (ruta larga o Windows no la borra)
echo ============================================================
echo.

if "%~1"=="" (
    set /p "CARPETA=Arrastra la carpeta a borrar sobre este .bat, o pega aqui su ruta completa y presiona Enter: "
) else (
    set "CARPETA=%~1"
)

if "%CARPETA%"=="" (
    echo No se dio ninguna carpeta. Cerrando.
    pause
    exit /b 1
)

if not exist "%CARPETA%\" (
    echo.
    echo No se encontro la carpeta:
    echo   %CARPETA%
    echo Verifica la ruta ^(o refresca el Explorador con F5 antes de intentar^) y vuelve a correr esto.
    pause
    exit /b 1
)

echo.
echo Se va a borrar PERMANENTEMENTE esta carpeta y TODO su contenido:
echo   %CARPETA%
echo.
set /p "CONFIRMAR=Escribe SI (en mayusculas) para confirmar: "
if /i not "%CONFIRMAR%"=="SI" (
    echo Cancelado, no se borro nada.
    pause
    exit /b 0
)

set "CARPETA_VACIA=%TEMP%\_vacia_para_borrar_%RANDOM%"
mkdir "%CARPETA_VACIA%"

echo.
echo Vaciando el contenido (esto puede tardar si hay muchos archivos)...
robocopy "%CARPETA_VACIA%" "%CARPETA%" /MIR /NFL /NDL /NJH /NJS >nul

echo Borrando la carpeta ya vacia...
rmdir "%CARPETA%" 2>nul
rmdir "%CARPETA_VACIA%" 2>nul

if exist "%CARPETA%\" (
    echo.
    echo No se pudo terminar de borrar. Puede que un archivo de adentro este
    echo abierto en otro programa ^(Word, Excel, un PDF, un antivirus escaneandola,
    echo etc^). Cierra esos programas y vuelve a correr este .bat.
) else (
    echo.
    echo Listo, la carpeta se borro por completo.
)

pause
