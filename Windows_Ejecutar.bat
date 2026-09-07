@echo off
setlocal enabledelayedexpansion
REM Windows: doble clic en este archivo para procesar el Excel de Xpendit que
REM este en esta carpeta. La primera vez prepara un entorno propio (venv\) e
REM instala pandas/openpyxl; las siguientes veces lo reutiliza.
cd /d "%~dp0"

set VENV_DIR=venv
set PY_SYS=

REM --- Elegir un Python 3 "normal" (no "free-threaded" / 3.13t) ---
REM Las builds free-threaded no tienen wheels precompilados de pandas/numpy
REM en PyPI: pip intenta compilar numpy desde codigo fuente y falla si no
REM hay compilador C instalado. Se prueba primero el default de "py -3" y,
REM si resulta ser free-threaded, se recorren otras versiones instaladas.
where py >nul 2>nul
if %errorlevel%==0 (
    call :probar_py -3
    if "!PY_SYS!"=="" for %%v in (3.13 3.12 3.11 3.10 3.9) do (
        if "!PY_SYS!"=="" call :probar_py -%%v
    )
)

if "%PY_SYS%"=="" (
    where python >nul 2>nul
    if %errorlevel%==0 (
        python -c "import sysconfig,sys; sys.exit(1 if sysconfig.get_config_var('Py_GIL_DISABLED') else 0)" >nul 2>nul
        if !errorlevel!==0 set PY_SYS=python
    )
)

if "%PY_SYS%"=="" (
    echo No se encontro una version estandar ^(no free-threaded^) de Python 3.
    echo Instala Python desde https://www.python.org/downloads/ y marca la
    echo casilla "Add python.exe to PATH".
    echo ^(Si el unico Python instalado es la variante "free-threaded" / 3.13t,
    echo  instala tambien la version normal de Python 3 junto a esa.^)
    pause
    exit /b 1
)
goto :seguir

:probar_py
py %1 -c "import sysconfig,sys; sys.exit(1 if sysconfig.get_config_var('Py_GIL_DISABLED') else 0)" >nul 2>nul
if %errorlevel%==0 set PY_SYS=py %1
exit /b 0

:seguir

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
