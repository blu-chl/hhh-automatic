@echo off
REM Windows: doble clic en este archivo para procesar el Excel de Xpendit que
REM este en esta carpeta. La primera vez prepara un entorno propio (venv\) e
REM instala pandas/openpyxl; las siguientes veces lo reutiliza.
cd /d "%~dp0"

set VENV_DIR=venv
set PY_SYS=

where py >nul 2>nul
if %errorlevel%==0 (
    set PY_SYS=py -3
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set PY_SYS=python
    )
)

if "%PY_SYS%"=="" (
    echo No se encontro Python 3 instalado en este equipo.
    echo Instalalo desde https://www.python.org/downloads/ y marca la casilla "Add python.exe to PATH".
    pause
    exit /b 1
)

REM Si el venv existe pero no funciona en ESTE equipo (por ejemplo, la carpeta
REM se copio desde otra maquina), lo detecta y lo reconstruye solo.
if exist "%VENV_DIR%\Scripts\python.exe" (
    "%VENV_DIR%\Scripts\python.exe" -c "import pandas, openpyxl" >nul 2>nul
    if errorlevel 1 rmdir /s /q "%VENV_DIR%"
)

if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo Primera vez en este equipo: preparando el entorno ^(puede tardar 1-2 minutos^)...
    rmdir /s /q "%VENV_DIR%" 2>nul
    %PY_SYS% -m venv %VENV_DIR%
    if errorlevel 1 (
        echo No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    "%VENV_DIR%\Scripts\python.exe" -m pip install --quiet --upgrade pip
    "%VENV_DIR%\Scripts\python.exe" -m pip install --quiet -r requirements.txt
    if errorlevel 1 (
        echo Fallo la instalacion de pandas/openpyxl. Revisa tu conexion a internet.
        pause
        exit /b 1
    )
    echo Entorno listo.
    echo.
)

"%VENV_DIR%\Scripts\python.exe" xpendit_a_nubatic.py
echo.
pause
