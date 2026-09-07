#!/usr/bin/env python3
"""
xpendit_a_nubatic.py
Convierte directamente el Reporte Consolidado de Xpendit a archivos de
importación Nubatic y planillas de facturas, sin Paso 1 (visión/imágenes).

Todos los gastos del Excel se tratan como aprobados por el solo hecho de
estar incluidos en la exportación de Xpendit.

Genera:
  - importador_{OBRA}_{PERSONA}_{semana}.xlsx  — boletas/comprobantes por persona/obra
  - facturas_{PERSONA}_{semana}.xlsx           — facturas por persona
  - informe_manual_{semana}.xlsx               — folio > 7 car. u obra sin CC

Uso simple (recomendado):
    Copia este script junto al Reporte Consolidado de Xpendit (único .xlsx
    de la carpeta) y ejecútalo sin argumentos. Detecta solo el Excel de
    entrada, calcula la semana desde las fechas del reporte, y deja los
    archivos generados en la misma carpeta:

        python xpendit_a_nubatic.py

Uso explícito (opcional, para override puntual):
    python xpendit_a_nubatic.py \\
        --excel "Reporte Consolidado...xlsx" \\
        --semana 2026-W27 \\
        --config ../config \\
        --dir-output ./salida_2026-W27
"""

import argparse
import json
import re
import shutil
import sys
import warnings
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# La consola de Windows suele usar cp1252/cp850 (no UTF-8), lo que rompe los
# emojis (📄 ✅ ⚠️) usados en los mensajes de progreso. Se fuerza UTF-8 en
# stdout/stderr para que funcione igual en cmd.exe, PowerShell o terminal.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

FOLIO_MAX_LEN = 7
HOJAS_VALIDAS = ["Reembolso", "Tarjeta corporativa", "Fondo recurrente"]


# ─── Auto-detección (uso sin argumentos) ──────────────────────────────────────

_PREFIJOS_SALIDA = ("importador_", "facturas_", "informe_manual_")


def _detectar_excel() -> Path:
    """Busca el único .xlsx a procesar: primero en el directorio actual,
    y si no hay nada ahí, junto al script (por si se ejecuta con doble clic
    desde otra carpeta de trabajo)."""

    def _candidatos(carpeta: Path):
        if not carpeta.exists():
            return []
        return sorted(
            f
            for f in carpeta.glob("*.xlsx")
            if not f.name.startswith("~$") and not f.name.startswith(_PREFIJOS_SALIDA)
        )

    cwd = Path.cwd()
    script_dir = Path(__file__).resolve().parent
    dirs_a_probar = [cwd] if cwd == script_dir else [cwd, script_dir]

    candidatos = []
    for carpeta in dirs_a_probar:
        candidatos = _candidatos(carpeta)
        if candidatos:
            break

    if not candidatos:
        raise SystemExit(
            f"❌ No se encontró ningún Excel para procesar (busqué en {cwd}).\n"
            f"   Coloca el Reporte Consolidado de Xpendit (.xlsx) en esta carpeta, "
            f"o indícalo con --excel."
        )
    if len(candidatos) > 1:
        nombres = "\n".join(f"   - {f.name}" for f in candidatos)
        raise SystemExit(
            f"❌ Hay más de un Excel candidato en {candidatos[0].parent}/:\n{nombres}\n"
            f"   Deja solo el Reporte Consolidado en la carpeta, o indica cuál usar con --excel."
        )
    return candidatos[0]


def _detectar_semana(df: pd.DataFrame) -> str:
    """Calcula la semana ISO (ej: 2026-W27) a partir de la fecha de documento
    más frecuente en el reporte. Si no hay fechas válidas, usa la semana actual."""
    fechas = pd.to_datetime(df.get("FECHA DOCUMENTO"), errors="coerce").dropna()
    if fechas.empty:
        anio, semana, _ = datetime.now().isocalendar()
        return f"{anio}-W{semana:02d}"
    iso = fechas.dt.isocalendar()[["year", "week"]]
    anio, semana = iso.value_counts().idxmax()
    return f"{int(anio)}-W{int(semana):02d}"


def _resolver_config_formatos(carpeta_entrada: Path, config_arg, formatos_arg):
    """Resuelve dónde están obras.json/tipos_gasto.json y los templates:
    1. Ruta explícita (--config / --formatos), si se pasó.
    2. Una carpeta config/ o Formatos/ junto al Excel de entrada (carpeta autocontenida).
    3. La estructura fija del proyecto (xpendit-skill/config, xpendit-skill/Formatos)."""
    proyecto_dir = Path(__file__).resolve().parent.parent  # xpendit-skill/

    if config_arg:
        dir_config = Path(config_arg)
    elif (carpeta_entrada / "config").is_dir():
        dir_config = carpeta_entrada / "config"
    else:
        dir_config = proyecto_dir / "config"

    if formatos_arg:
        dir_formatos = Path(formatos_arg)
    elif (carpeta_entrada / "Formatos").is_dir():
        dir_formatos = carpeta_entrada / "Formatos"
    else:
        dir_formatos = proyecto_dir / "Formatos"

    return dir_config, dir_formatos


# ─── Lectura del Excel Xpendit ────────────────────────────────────────────────


def leer_excel(ruta_excel: str) -> pd.DataFrame:
    xl = pd.ExcelFile(ruta_excel)
    frames = []
    for hoja in HOJAS_VALIDAS:
        if hoja not in xl.sheet_names:
            continue
        df = pd.read_excel(xl, sheet_name=hoja, header=0)
        if "ID GASTO" not in df.columns:
            continue
        df = df[pd.to_numeric(df["ID GASTO"], errors="coerce").notna()].copy()
        if df.empty:
            continue
        df["_hoja"] = hoja
        frames.append(df)
    if not frames:
        raise ValueError("No se encontró ninguna hoja válida en el Excel.")
    return pd.concat(frames, ignore_index=True)


# ─── Normalización de campos ──────────────────────────────────────────────────


def _folio(v) -> str:
    """Convierte el número de documento a string limpio (sin decimales .0).
    Los floats en notación científica se expanden a entero completo."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        f = float(v)
        return str(int(f))
    except (ValueError, TypeError):
        return str(v).strip()


def _folio_excel(folio_str: str):
    """Retorna int si el folio es puramente numérico (Nubatic exige número),
    o el string original si contiene letras."""
    try:
        return int(folio_str)
    except (ValueError, TypeError):
        return folio_str


def _monto(v):
    """Entero desde float de Pandas."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return int(round(float(v)))


def _fecha_ddmmaaaa(v) -> str:
    """String DD-MM-YYYY para informes de texto."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        if isinstance(v, str):
            v = pd.to_datetime(v)
        return v.strftime("%d-%m-%Y")
    except Exception:
        return str(v)[:10]


def _fecha_dt(v):
    """datetime.date para escribir en Excel como serial de fecha (no string)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        if isinstance(v, str):
            v = pd.to_datetime(v)
        return v.date()
    except Exception:
        return None


def _str(v) -> str:
    """Convierte un valor de celda a texto limpio. A diferencia de `v or ""`,
    maneja bien los NaN de pandas: un float NaN es "verdadero" en Python
    (bool(float('nan')) == True), así que `v or ""` lo deja pasar como NaN
    y termina escribiendo el texto literal "nan" en vez de vacío."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v)


def _empleado(fila) -> str:
    nombre = _str(fila.get("NOMBRE")).strip()
    apellido = _str(fila.get("APELLIDO")).strip()
    emp = f"{nombre} {apellido}".strip()
    return emp or "SIN NOMBRE"


def _slug(nombre: str) -> str:
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in nombre)


def _calcular_neto_iva(monto_total, id_cuenta_contable):
    """Calcula monto_neto e IVA desde el total.
    Peajes y otros sin IVA recuperable → iva=0, neto=total.
    Resto → neto = round(total / 1.19)."""
    if not monto_total:
        return None, None
    cuenta = str(int(float(id_cuenta_contable))) if id_cuenta_contable and str(id_cuenta_contable) != "nan" else ""
    if cuenta in (
        "20050",
        "20060",
        "20061",
        "20062",
        "20063",
    ):  # peajes, alim, movil, estac, notaría
        return int(monto_total), 0
    neto = round(monto_total / 1.19)
    iva = int(monto_total) - neto
    return neto, iva


# ─── Clasificación ────────────────────────────────────────────────────────────


def clasificar(df: pd.DataFrame, obras: dict):
    """Clasifica cada gasto en: importador, facturas, folio_largo, sin_cc."""
    codigos_obra = [k for k in obras if not k.startswith("_")]
    importador, facturas, folio_largo, sin_cc = [], [], [], []

    for _, fila in df.iterrows():
        tipo = _str(fila.get("TIPO DE DOCUMENTO")).strip()
        cc_texto = _str(fila.get("CENTRO DE COSTOS"))
        folio = _folio(fila.get("NUMERO DE DOCUMENTO"))
        monto = _monto(fila.get("MONTO EN MONEDA LOCAL"))
        id_cc = _str(fila.get("ID CENTRO DE COSTOS")).strip()
        id_cta = fila.get("ID CUENTA CONTABLE")
        empleado = _empleado(fila)
        fecha_doc = _fecha_ddmmaaaa(fila.get("FECHA DOCUMENTO"))
        fecha_doc_dt = _fecha_dt(fila.get("FECHA DOCUMENTO"))
        rut = _str(fila.get("RUT PROVEEDOR")).strip()
        proveedor = _str(fila.get("NOMBRE PROVEEDOR")).strip()
        id_gasto = str(int(float(fila.get("ID GASTO"))))

        obra_cod = next((c for c in codigos_obra if cc_texto.startswith(c)), None)
        obra_cfg = obras.get(obra_cod, {}) if obra_cod else {}

        g = {
            "id_gasto": id_gasto,
            "empleado": empleado,
            "tipo": tipo,
            "folio": folio,
            "fecha_doc": fecha_doc,
            "fecha_doc_dt": fecha_doc_dt,
            "monto_total": monto,
            "id_cc": id_cc,
            "cc_texto": cc_texto,
            "id_cuenta_contable": id_cta,
            "cuenta_contable": _str(fila.get("CUENTA CONTABLE")).strip(),
            "rut": rut,
            "proveedor": proveedor,
            "descripcion": _str(fila.get("DESCRIPCION")).strip(),
            "_obra": obra_cod,
            "_obra_cfg": obra_cfg,
        }

        # Toda factura (sin excepción de peajes) → solo Excel de facturas, NO al importador.
        if tipo.lower() == "factura":
            neto, iva = _calcular_neto_iva(monto, id_cta)
            g["monto_neto"] = neto
            g["iva"] = iva
            facturas.append(g)
            continue

        # Sin CC configurado → manual
        if not obra_cfg.get("centro_costo"):
            sin_cc.append(g)
            continue

        # Folio largo → número secuencial, igual que el programa original
        # (no bloquea el gasto, se renumera para que calce con Nubatic).
        if len(folio) > FOLIO_MAX_LEN:
            g["folio"] = ""

        importador.append(g)

    return importador, facturas, folio_largo, sin_cc


# ─── Escritura de archivos ─────────────────────────────────────────────────────


def _buscar_template(obra_cod: str, dir_formatos: str) -> Path:
    """Devuelve el template de Nubatic para la obra, o None si no existe."""
    d = Path(dir_formatos)
    if not d.exists():
        return None
    for f in d.glob("*.xlsx"):
        if f.stem.startswith(obra_cod):
            return f
    return None


def _leer_tabla_jk(template_path):
    """Lee la tabla J:K del template (filas 1+) y la devuelve como lista de tuplas."""
    wb = load_workbook(str(template_path), data_only=True, read_only=True)
    ws = wb.active
    tabla = []
    for row in ws.iter_rows(min_row=1, max_row=200, min_col=10, max_col=11, values_only=True):
        cod, nom = row
        if cod is not None or nom is not None:
            tabla.append((cod, nom))
    wb.close()
    return tabla


_METADATA_XLSX = {
    "docProps/app.xml": [
        (rb"<Application>.*?</Application>", b"<Application>Microsoft Excel</Application>"),
        (rb"<AppVersion>.*?</AppVersion>", b"<AppVersion>16.0300</AppVersion>"),
    ],
    "docProps/core.xml": [
        (rb"<dc:creator>.*?</dc:creator>", b"<dc:creator></dc:creator>"),
        (rb"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", b"<cp:lastModifiedBy></cp:lastModifiedBy>"),
    ],
}


def _excel_metadata_nativa(ruta: Path):
    """openpyxl deja su propia firma en los metadatos internos del .xlsx
    (docProps/app.xml -> "...Openpyxl X.X.X", docProps/core.xml -> creador
    "openpyxl"). Algunas plataformas de importacion (ERPs) validan esos
    metadatos y rechazan el archivo si no "parece" venir de Excel real; abrir
    el archivo en Excel y volver a guardarlo lo arregla porque Excel
    reescribe esos campos. Esto hace lo mismo sin depender de abrir Excel."""
    tmp = ruta.with_name(ruta.name + ".tmp")
    with zipfile.ZipFile(ruta, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            for patron, nuevo in _METADATA_XLSX.get(item.filename, []):
                data = re.sub(patron, nuevo, data)
            zout.writestr(item, data)
    tmp.replace(ruta)


def _escribir_importador(gastos, empleado, obra, obra_cfg, tipos, ruta, template_path=None):
    cc = obra_cfg.get("centro_costo", {})
    cc_cod = cc.get("codigo")
    cc_nombre = cc.get("nombre", "")
    try:
        cc_cod_val = int(cc_cod)
    except (TypeError, ValueError):
        cc_cod_val = cc_cod

    tipos_tabla = tipos.get("tipos", {})

    # Tabla J:K: leer del template si existe, o usar solo el CC de la obra
    if template_path and Path(template_path).exists():
        tabla_jk = _leer_tabla_jk(template_path)
    else:
        tabla_jk = [("ID", "CENTRO COSTO"), (cc_cod_val, cc_nombre)]

    # Tabla M:N: tipos de gasto globales (siempre desde config)
    tabla_mn = [("ID", "TIPO DE GASTO")] + [(int(cod), info["nombre"]) for cod, info in tipos_tabla.items()]

    # ── Crear archivo desde cero (evita problemas de caché de fórmulas del template) ──
    wb = Workbook()
    ws = wb.active
    ws.title = "Worksheet"

    HDR = "1F4E79"
    headers = [
        "Tipo Docto.",
        "Fecha Docto",
        "Nº Docto.",
        "Valor",
        "ID CENTRO COSTO",
        "CENTRO COSTO",
        "ID TIPO GASTO",
        "TIPO GASTO",
        "    ",
        "ID",
        "CENTRO COSTO",
        "    ",
        "ID",
        "TIPO DE GASTO",
    ]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, name="Arial", size=10, color="FFFFFF" if c <= 8 else "000000")
        if c <= 8:
            cell.fill = PatternFill("solid", fgColor=HDR)

    # Tabla J:K (incluye header)
    for idx, (cod, nom) in enumerate(tabla_jk, 1):
        ws.cell(row=idx, column=10, value=cod)
        ws.cell(row=idx, column=11, value=nom)

    # Tabla M:N (incluye header)
    for idx, (cod, nom) in enumerate(tabla_mn, 1):
        ws.cell(row=idx, column=13, value=cod)
        ws.cell(row=idx, column=14, value=nom)

    anchos = [12, 12, 20, 12, 16, 32, 13, 24, 4, 7, 36, 4, 7, 26]
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a

    # ── Datos: F y H como valores literales (Nubatic no recalcula fórmulas) ──────
    folio_seq = 0
    for i, g in enumerate(gastos, 2):
        ws.cell(row=i, column=1, value="BOLETA")
        cell_f = ws.cell(row=i, column=2, value=g["fecha_doc_dt"])
        cell_f.number_format = "DD-MM-YYYY"
        folio = g["folio"]
        if not folio:
            folio_seq += 1
            folio = str(folio_seq)
        ws.cell(row=i, column=3, value=_folio_excel(folio))
        ws.cell(row=i, column=4, value=g["monto_total"])
        ws.cell(row=i, column=5, value=cc_cod_val)
        ws.cell(row=i, column=6, value=cc_nombre)
        try:
            cci = int(float(g["id_cuenta_contable"]))
        except (TypeError, ValueError):
            cci = g["id_cuenta_contable"]
        ws.cell(row=i, column=7, value=cci)
        ws.cell(row=i, column=8, value=tipos_tabla.get(str(cci), {}).get("nombre", ""))

    wb.save(str(ruta))
    _excel_metadata_nativa(ruta)
    total = sum(g["monto_total"] or 0 for g in gastos)
    print(f"   📄 {ruta.name}  ({len(gastos)} comprobantes · ${total:,})")


def _escribir_facturas(facturas, empleado, ruta):
    wb = Workbook()
    ws = wb.active
    ws.title = (empleado or "SIN NOMBRE")[:31]

    HDR = "8A2C2C"
    headers = [
        "RUT",
        "Proveedor",
        "Nº Documento",
        "Monto Neto",
        "IVA",
        "Monto Total",
        "Fecha",
        "Centro de Costo",
    ]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.fill = PatternFill("solid", fgColor=HDR)
        cell.alignment = Alignment(horizontal="center")

    for i, g in enumerate(facturas, 2):
        ws.cell(row=i, column=1, value=g["rut"])
        ws.cell(row=i, column=2, value=g["proveedor"])
        ws.cell(row=i, column=3, value=g["folio"])
        for c_idx, campo in enumerate(("monto_neto", "iva", "monto_total"), 4):
            cell = ws.cell(row=i, column=c_idx, value=g.get(campo))
            cell.number_format = "#,##0"
        cell_f = ws.cell(row=i, column=7, value=g["fecha_doc_dt"])
        cell_f.number_format = "DD-MM-YYYY"
        ws.cell(row=i, column=8, value=g["cc_texto"])

    anchos = [16, 30, 14, 13, 11, 13, 12, 26]
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    n = len(facturas)
    if n > 0:
        fila_tot = n + 2
        ws.cell(row=fila_tot, column=2, value="TOTAL").font = Font(bold=True, name="Arial", size=10)
        for col in (4, 5, 6):
            letra = get_column_letter(col)
            cell = ws.cell(row=fila_tot, column=col, value=f"=SUM({letra}2:{letra}{n + 1})")
            cell.number_format = "#,##0"
            cell.font = Font(bold=True, name="Arial", size=10)

    wb.save(ruta)
    _excel_metadata_nativa(ruta)
    total = sum(g["monto_total"] or 0 for g in facturas)
    print(f"   📄 {ruta.name}  ({n} facturas · ${total:,})")


def _escribir_manual(folio_largo, sin_cc, ruta):
    if not folio_largo and not sin_cc:
        return
    wb = Workbook()
    primera = True
    HDR = "5C5C5C"
    cols = [
        "ID Gasto",
        "Empleado",
        "Obra/CC",
        "Tipo Doc",
        "Proveedor",
        "RUT",
        "Folio",
        "Fecha Doc",
        "Monto",
        "Cuenta Contable",
        "Motivo",
    ]

    def _hoja(titulo, gastos, motivo):
        nonlocal primera
        ws = wb.active if primera else wb.create_sheet()
        primera = False
        ws.title = titulo[:31]
        for c, h in enumerate(cols, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
            cell.fill = PatternFill("solid", fgColor=HDR)
        for i, g in enumerate(gastos, 2):
            vals = [
                g["id_gasto"],
                g["empleado"],
                g["cc_texto"],
                g["tipo"],
                g["proveedor"],
                g["rut"],
                g["folio"],
                g["fecha_doc"],
                g["monto_total"],
                g["cuenta_contable"],
                motivo,
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=i, column=c, value=v)
                if c == 9:
                    cell.number_format = "#,##0"
        anchos = [10, 22, 22, 12, 28, 16, 14, 11, 11, 18, 18]
        for i, a in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(i)].width = a
        ws.freeze_panes = "A2"

    if sin_cc:
        _hoja("Sin centro de costo", sin_cc, "Obra sin CC configurado")
    if folio_largo:
        _hoja(
            f"Folio > {FOLIO_MAX_LEN} caracteres",
            folio_largo,
            "Folio demasiado largo para Nubatic",
        )

    wb.save(ruta)
    _excel_metadata_nativa(ruta)
    print(f"   📄 {ruta.name}  ({len(sin_cc)} sin CC · {len(folio_largo)} folio largo)")


# ─── Orquestador ──────────────────────────────────────────────────────────────


def generar(
    ruta_excel: str = None,
    semana: str = None,
    dir_config: str = None,
    dir_output: str = None,
    dir_formatos: str = None,
):
    """Todos los parámetros son opcionales. Sin argumentos, el script:
    - busca el único .xlsx en la carpeta actual (o junto al script),
    - calcula la semana desde las fechas del reporte,
    - usa config/ y Formatos/ junto al Excel si existen, o los del proyecto,
    - y deja la salida en la misma carpeta del Excel de entrada."""
    ruta_excel_p = Path(ruta_excel) if ruta_excel else _detectar_excel()
    carpeta_entrada = ruta_excel_p.resolve().parent

    dir_config_r, dir_formatos_r = _resolver_config_formatos(carpeta_entrada, dir_config, dir_formatos)

    with open(dir_config_r / "obras.json", encoding="utf-8") as f:
        obras = json.load(f)
    with open(dir_config_r / "tipos_gasto.json", encoding="utf-8") as f:
        tipos = json.load(f)

    dir_out = Path(dir_output) if dir_output else carpeta_entrada
    dir_out.mkdir(parents=True, exist_ok=True)

    df = leer_excel(str(ruta_excel_p))
    if not semana:
        semana = _detectar_semana(df)
    importador, facturas, folio_largo, sin_cc = clasificar(df, obras)

    print(f"\n{'=' * 60}")
    print(f"  Exportando a Nubatic — semana {semana}")
    print(f"{'=' * 60}\n")
    print(f"  Total gastos leídos:         {len(df)}")
    print(f"    · Importador Nubatic:      {len(importador)}")
    print(f"    · Facturas (por persona):  {len(facturas)}")
    print(f"    · Folio > {FOLIO_MAX_LEN} car. (manual): {len(folio_largo)}")
    print(f"    · Sin CC configurado:      {len(sin_cc)}\n")

    # 1. Importador por persona por obra
    por_persona_obra = {}
    for g in importador:
        key = (g["empleado"], g["_obra"])
        por_persona_obra.setdefault(key, []).append(g)

    for (persona, obra), gastos_grupo in sorted(por_persona_obra.items()):
        template_path = _buscar_template(obra, dir_formatos_r)
        if template_path:
            print(f"   🗂  Usando template: {template_path.name}")
        slug = _slug(persona)
        ruta = dir_out / f"importador_{obra}_{slug}_{semana}.xlsx"
        _escribir_importador(gastos_grupo, persona, obra, obras[obra], tipos, ruta, template_path)

    if not por_persona_obra:
        print("   ⚠️  Sin gastos para importador Nubatic (ver informe_manual)")

    # 2. Facturas por persona
    por_persona_factura = {}
    for g in facturas:
        por_persona_factura.setdefault(g["empleado"], []).append(g)

    for persona, fac in sorted(por_persona_factura.items()):
        slug = _slug(persona)
        ruta = dir_out / f"facturas_{slug}_{semana}.xlsx"
        _escribir_facturas(fac, persona, ruta)

    # 3. Informe manual
    _escribir_manual(folio_largo, sin_cc, dir_out / f"informe_manual_{semana}.xlsx")

    print(f"\n✅ Archivos en: {dir_out.resolve()}/\n")


# ─── CLI ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    p = argparse.ArgumentParser(
        description="Xpendit → Nubatic: genera importador y facturas directo desde el "
        "Reporte Consolidado de Xpendit, sin procesamiento de imágenes. "
        "Ejecutado sin argumentos, detecta todo automáticamente (ver docstring del archivo)."
    )
    p.add_argument(
        "--excel",
        default=None,
        help="Reporte Consolidado Xpendit (.xlsx). Por defecto, se busca el único .xlsx en la carpeta.",
    )
    p.add_argument(
        "--semana",
        default=None,
        help="Semana ISO, ej: 2026-W27. Por defecto, se calcula desde las fechas del Excel.",
    )
    p.add_argument(
        "--config",
        default=None,
        help="Directorio de configuración (obras.json, tipos_gasto.json). "
        "Por defecto, config/ junto al Excel si existe, o el de xpendit-skill/config.",
    )
    p.add_argument(
        "--dir-output",
        default=None,
        help="Carpeta donde guardar los archivos generados. Por defecto, la misma carpeta del Excel.",
    )
    p.add_argument(
        "--formatos",
        default=None,
        help="Carpeta con los templates Nubatic por obra. "
        "Por defecto, Formatos/ junto al Excel si existe, o el de xpendit-skill/Formatos.",
    )
    args = p.parse_args()

    generar(args.excel, args.semana, args.config, args.dir_output, args.formatos)
