#!/bin/bash
# Doble clic en este archivo para procesar el Excel de Xpendit que esté en esta carpeta.
# La primera vez prepara un entorno propio (venv/) e instala pandas/openpyxl.
# Las siguientes veces lo reutiliza, así que arranca mucho más rápido.
cd "$(dirname "$0")"

VENV_DIR="venv"

if ! command -v python3 >/dev/null 2>&1; then
    echo "❌ No se encontró Python 3 instalado en este Mac."
    echo "   Instálalo desde https://www.python.org/downloads/ y vuelve a intentar."
    read -p "Presiona Enter para cerrar..."
    exit 1
fi

# Si el venv existe pero no funciona en ESTE equipo (por ejemplo, la carpeta se
# copió desde otro Mac), lo detecta y lo reconstruye solo.
if [ -x "$VENV_DIR/bin/python3" ]; then
    "$VENV_DIR/bin/python3" -c "import pandas, openpyxl" >/dev/null 2>&1 || rm -rf "$VENV_DIR"
fi

if [ ! -x "$VENV_DIR/bin/python3" ]; then
    echo "🔧 Primera vez en este equipo: preparando el entorno (puede tardar 1-2 minutos)..."
    rm -rf "$VENV_DIR"
    python3 -m venv "$VENV_DIR" || {
        echo "❌ No se pudo crear el entorno virtual."
        read -p "Presiona Enter para cerrar..."
        exit 1
    }
    "$VENV_DIR/bin/python3" -m pip install --quiet --upgrade pip
    "$VENV_DIR/bin/python3" -m pip install --quiet -r requirements.txt || {
        echo "❌ Falló la instalación de pandas/openpyxl (¿hay conexión a internet?)."
        read -p "Presiona Enter para cerrar..."
        exit 1
    }
    echo "✅ Entorno listo."
    echo
fi

"$VENV_DIR/bin/python3" xpendit_a_nubatic.py
echo
read -p "Presiona Enter para cerrar esta ventana..."
