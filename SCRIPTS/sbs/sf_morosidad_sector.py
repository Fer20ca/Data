import os
import re
import calendar
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List

import pandas as pd

# -------- CONFIG --------
BASE_DIR = Path(r"C:\Users\José Estrada\OneDrive - ABC Capital\Web Scraping\SBS\SF_Data") #cambiar la ruta a la carpeta con tu base de datos creada con el script previo
WANTED_SHEET = "Morosidad x SE"
HEADER_ROW = 5
DATA_START_ORIG = 9
DATA_END_ORIG = 24

OUT_XLSX = BASE_DIR / "Morosidad_Sectorial.xlsx"

ABREV_TO_MONTH = {
    "en": 1, "fe": 2, "ma": 3, "ab": 4, "my": 5, "jn": 6,
    "jl": 7, "ag": 8, "se": 9, "oc":10, "no":11, "di":12
}
FILE_REGEX = re.compile(r"^SF-([a-z]{2})(\d{4})\.xls(x)?$", re.IGNORECASE)

TARGET_COLS = [
    "Banca Múltiple",
    "Empresas Financieras",
    "Cajas Municipales",
    "Cajas Rurales de Ahorro y Crédito",
    "EDPYMEs",
    "Agrobanco",
    "Total",
]

# -------- utilidades --------
def norm_simple(s: str) -> str:
    if s is None:
        return ""
    return (str(s).lower().strip()
              .replace("á","a").replace("é","e").replace("í","i")
              .replace("ó","o").replace("ú","u")
              .replace("\n"," "))

def parse_period(fname: str) -> Optional[Dict[str, object]]:
    m = FILE_REGEX.match(fname)
    if not m:
        return None
    ab = m.group(1).lower()
    year = int(m.group(2))
    month = ABREV_TO_MONTH.get(ab)
    if not month:
        return None
    last_day = calendar.monthrange(year, month)[1]
    return {"year": year, "month": month, "date": datetime(year, month, last_day)}

def engines_to_try(path: Path):
    ext = path.suffix.lower()
    if ext == ".xlsx":
        return ["openpyxl"]
    if ext == ".xls":
        # algunos .xls son realmente .xlsx → probamos ambos
        return ["xlrd", "openpyxl"]
    return ["openpyxl", "xlrd"]

def read_excel_safe(path: Path, sheet_name: str, header=None, dtype=None):
    """Lee una hoja probando engines en orden."""
    for eng in engines_to_try(path):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, header=header, dtype=dtype, engine=eng)
        except Exception:
            continue
    return None

def normalize_sheet_name(s: str) -> str:
    return re.sub(r"\s+"," ", norm_simple(s))

def find_sheet(path: Path, wanted: str) -> Optional[str]:
    """Encuentra la hoja objetivo con varios fallbacks."""
    wn = normalize_sheet_name(wanted)
    for eng in engines_to_try(path):
        try:
            xls = pd.ExcelFile(path, engine=eng)
            # exacto
            for s in xls.sheet_names:
                if normalize_sheet_name(s) == wn:
                    return s
            # fallback por palabras
            for s in xls.sheet_names:
                ns = normalize_sheet_name(s)
                if "morosidad" in ns and ("se" in ns or "sector" in ns):
                    return s
            # última opción: primera hoja
            if xls.sheet_names:
                return xls.sheet_names[0]
        except Exception:
            continue
    return None

def map_present_to_standard(columns: List[str]) -> Dict[str, str]:
    """Mapea encabezados reales -> estándar para las columnas de entidades."""
    real_norm = {norm_simple(c): c for c in columns}
    patterns = {
        "Banca Múltiple": ["banca multiple"],
        "Empresas Financieras": ["empresas financieras", "empresas  financieras"],
        "Cajas Municipales": ["cajas municipales"],
        "Cajas Rurales de Ahorro y Crédito": [
            "cajas rurales de ahorro y credito", "caja rurales de ahorro y credito",
            "cajas rurales de ahorro y credito (miles)"
        ],
        "EDPYMEs": ["edpymes", "edpyme"],
        "Agrobanco": ["agrobanco"],
        "Total": ["total","total (miles)","total general"],
    }
    mapping = {}
    for std_name, keys in patterns.items():
        for k in keys:
            if k in real_norm:
                mapping[real_norm[k]] = std_name
                break
    return mapping

# -------- limpieza por archivo --------
def clean_one(path: Path) -> Optional[pd.DataFrame]:
    meta = parse_period(path.name)
    if not meta:
        print(f"[SKIP] Nombre no reconocido: {path.name}")
        return None

    sheet = find_sheet(path, WANTED_SHEET)
    if not sheet:
        print(f"[ERROR] Hoja no encontrada en {path.name}")
        return None

    df = read_excel_safe(path, sheet_name=sheet, header=HEADER_ROW, dtype=str)
    if df is None:
        print(f"[ERROR] No se pudo leer '{sheet}' en {path.name}")
        return None

    # Fila original clamp 9..24
    df["__orig_row__"] = df.index + 7
    df = df[(df["__orig_row__"] >= DATA_START_ORIG) & (df["__orig_row__"] <= DATA_END_ORIG)].copy()

    # Normalizar columnas
    df.columns = [re.sub(r"\s+"," ", str(c)).strip() for c in df.columns]

    # Detectar columna Sector
    col_sector = next((c for c in df.columns if "sector" in norm_simple(c)), None)
    if not col_sector:
        print(f"[WARN] Sin columna 'Sector' en {path.name}")
        return None

    # Limpiar filas de notas
    df[col_sector] = df[col_sector].astype(str).str.strip()
    mask_notes = (
        df[col_sector].str.match(r"^\s*(?:\*+|\d+/)", na=False) |
        df[col_sector].str.contains(r"(?:fuente|p[aá]gina)", case=False, na=False)
    )
    df = df[~mask_notes].copy()
    df = df[df[col_sector].notna() & (df[col_sector].str.strip()!="")]

    # Mapear columnas presentes -> estándar
    col_map = map_present_to_standard(list(df.columns))
    present_real = list(col_map.keys())
    if not present_real:
        print(f"[WARN] No se hallaron columnas de entidades en {path.name}")
        return None

    keep = [col_sector] + present_real
    df = df[keep].copy().rename(columns=col_map)

    # Numerificar
    for c in TARGET_COLS:
        if c in df.columns:
            df[c] = (df[c].astype(str)
                            .str.replace("%","", regex=False)
                            .str.replace(r"\s+","", regex=True)
                            .replace({"-": pd.NA, "–": pd.NA, "—": pd.NA, "nan": pd.NA, "None": pd.NA}))
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Metadatos
    df.insert(0, "year", meta["year"])
    df.insert(1, "date", meta["date"])

    if col_sector != "Sector Económico":
        df = df.rename(columns={col_sector: "Sector Económico"})

    # ---- Formato TIDY ----
    value_cols = [c for c in TARGET_COLS if c in df.columns]
    tidy = df.melt(
        id_vars=["year", "date", "Sector Económico"],
        value_vars=value_cols,
        var_name="Entidad",
        value_name="morosidad"
    )
    tidy = tidy.dropna(subset=["morosidad"]).reset_index(drop=True)

    # Orden sugerido
    entidad_order = [c for c in TARGET_COLS if c in value_cols]
    tidy["Entidad"] = pd.Categorical(tidy["Entidad"], categories=entidad_order, ordered=True)
    tidy = tidy.sort_values(["year","date","Sector Económico","Entidad"]).reset_index(drop=True)

    return tidy

# -------- pipeline --------
def build_db(base: Path) -> pd.DataFrame:
    files = [base / f for f in os.listdir(base) if f.lower().endswith((".xls",".xlsx"))]
    files.sort()
    frames: List[pd.DataFrame] = []
    for f in files:
        out = clean_one(f)
        if out is not None and not out.empty:
            frames.append(out)
    if not frames:
        raise RuntimeError("No se pudo construir la base; revisa archivos.")
    db = pd.concat(frames, ignore_index=True, sort=False)
    return db

def main():
    if not BASE_DIR.exists():
        raise FileNotFoundError(f"No existe carpeta: {BASE_DIR}")
    db = build_db(BASE_DIR)
    db.to_excel(OUT_XLSX, index=False)
    print(f"Listo.\n- {OUT_XLSX}")

if __name__ == "__main__":
    main()
