#!/usr/bin/env python3
"""
instalar_formato.py
Detecta automáticamente a qué obra pertenece un archivo "FORMATO EXCEL GASTOS
RENDICION" (la tabla de centros de costo en columnas J:K) y lo instala en
Formatos/ con el nombre que xpendit_a_nubatic.py espera (empieza con el
código de obra), sin necesidad de renombrarlo a mano ni de que Claude lo
inspeccione cada vez.

Cómo detecta la obra:
    Cada obra tiene, en config/obras.json, un centro_costo.codigo "madre"
    (ej: CLH210 -> 25048). Ese código SIEMPRE aparece como una de las filas
    de la tabla J:K del archivo de formato de esa obra (es su propia
    partida de "10. GASTOS GENERALES DIRECTOS"). Si el código madre de una
    sola obra aparece en el archivo, esa es la obra.

Uso:
    python instalar_formato.py "ruta/al/FORMATO EXCEL GASTOS RENDICION.xlsx"

    Sin argumentos, revisa todos los .xlsx sueltos en esta misma carpeta
    (los que no estén ya en Formatos/) y los instala uno por uno.
"""

import argparse
import json
import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook

AQUI = Path(__file__).resolve().parent
CONFIG_DIR = AQUI / "config"
FORMATOS_DIR = AQUI / "Formatos"

# Prefijos de archivos que el propio script genera como SALIDA. Un importador_*
# ya roto (sin template) también trae 1 fila con el código de su obra en J:K —
# si lo dejáramos pasar, lo "instalaríamos" como si fuera el template real y
# el problema nunca se arreglaría. Se excluyen explícitamente por nombre.
_PREFIJOS_SALIDA = ("importador_", "facturas_", "informe_manual_", "reporte consolidado")
MIN_FILAS_TEMPLATE = 3  # header + al menos 2 partidas reales; 1 sola fila = salida rota, no template


def _leer_codigos_jk(ruta_xlsx: Path) -> set:
    """Códigos (columna J) presentes en la tabla de centros de costo del archivo."""
    wb = load_workbook(str(ruta_xlsx), data_only=True, read_only=True)
    ws = wb.active
    codigos = set()
    for (cod,) in ws.iter_rows(min_row=1, max_row=300, min_col=10, max_col=10, values_only=True):
        if cod is None:
            continue
        try:
            codigos.add(int(cod))
        except (TypeError, ValueError):
            continue
    wb.close()
    return codigos


def _cargar_obras() -> dict:
    with open(CONFIG_DIR / "obras.json", encoding="utf-8") as f:
        obras = json.load(f)
    return {
        cod: cfg
        for cod, cfg in obras.items()
        if not cod.startswith("_") and isinstance(cfg.get("centro_costo"), dict)
    }


def detectar_obra(ruta_xlsx: Path, obras: dict):
    """Devuelve (codigo_obra, cfg) si detecta una única obra candidata.
    Lanza ValueError si no hay ninguna, si hay más de una (ambiguo), o si el
    archivo no alcanza a ser una tabla real de centros de costo."""
    nombre_lower = ruta_xlsx.name.lower()
    if nombre_lower.startswith(_PREFIJOS_SALIDA):
        raise ValueError(
            f"'{ruta_xlsx.name}' es un archivo de SALIDA del propio script (importador/facturas/"
            f"informe/reporte), no un formato de referencia — se ignora."
        )

    codigos_archivo = _leer_codigos_jk(ruta_xlsx)
    if len(codigos_archivo) < MIN_FILAS_TEMPLATE:
        raise ValueError(
            f"'{ruta_xlsx.name}' solo trae {len(codigos_archivo)} código(s) en la tabla J:K — "
            f"muy pocos para ser un template real (probablemente es una salida ya rota, no un formato)."
        )

    candidatos = [
        (cod, cfg) for cod, cfg in obras.items() if cfg["centro_costo"].get("codigo") in codigos_archivo
    ]
    if not candidatos:
        raise ValueError(
            f"No pude identificar la obra de '{ruta_xlsx.name}': ninguno de los "
            f"códigos 'madre' de obras.json aparece en su tabla J:K (columna J, filas 1-300)."
        )
    if len(candidatos) > 1:
        nombres = ", ".join(cod for cod, _ in candidatos)
        raise ValueError(
            f"'{ruta_xlsx.name}' calza con más de una obra ({nombres}) — revísalo a mano."
        )
    return candidatos[0]


def instalar(ruta_xlsx: Path, obras: dict, mover: bool = False) -> Path:
    obra_cod, cfg = detectar_obra(ruta_xlsx, obras)
    FORMATOS_DIR.mkdir(exist_ok=True)
    # El nombre de la obra en obras.json ya empieza con su código (ej. "CLH210 PLANTA..."),
    # que es justo lo que _buscar_template() del script principal busca al inicio del nombre.
    destino = FORMATOS_DIR / f"{cfg['nombre']}.xlsx"
    if mover:
        shutil.move(str(ruta_xlsx), str(destino))
    else:
        shutil.copy(str(ruta_xlsx), str(destino))
    print(f"✅ {ruta_xlsx.name}")
    print(f"   obra detectada: {obra_cod} ({cfg['nombre']})")
    print(f"   instalado en:   {destino.relative_to(AQUI)}")
    return destino


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("archivo", nargs="?", default=None, help="Archivo de formato a instalar")
    p.add_argument("--mover", action="store_true", help="Mover en vez de copiar el archivo original")
    args = p.parse_args()

    obras = _cargar_obras()

    if args.archivo:
        objetivo = [Path(args.archivo)]
    else:
        ya_instalados = {f.name for f in FORMATOS_DIR.glob("*.xlsx")} if FORMATOS_DIR.exists() else set()
        objetivo = [
            f for f in AQUI.glob("*.xlsx") if f.name not in ya_instalados and not f.name.startswith("~$")
        ]
        if not objetivo:
            print("No hay archivos .xlsx sueltos para instalar en esta carpeta.")
            sys.exit(0)

    errores = []
    for f in objetivo:
        if not f.exists():
            errores.append(f"{f}: no existe")
            continue
        try:
            instalar(f, obras, mover=args.mover)
        except ValueError as e:
            errores.append(str(e))

    if errores:
        print("\n⚠️  No se pudieron instalar automáticamente:")
        for e in errores:
            print(f"   - {e}")
        sys.exit(1)
