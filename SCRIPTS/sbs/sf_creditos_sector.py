import os
import re
import calendar
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List

import pandas as pd

# -------- CONFIG --------
BASE_DIR = Path(r"C:\Users\José Estrada\OneDrive - ABC Capital\Web Scraping\SBS\SF_Data") #cambiar la ruta a la carpeta con tu base de datos creada con el script previo
WANTED_SHEET = "Créditos x SE"
OUT_XLSX = BASE_DIR / "Creditos_Sectorial.xlsx" #nombre de archivo 

# Abreviaturas de mes del nombre de archivo -> número de mes
ABREV_TO_MONTH = {
    "en": 1, "fe": 2, "ma": 3, "ab": 4, "my": 5, "jn": 6,
    "jl": 7, "ag": 8, "se": 9, "oc":10, "no":11, "di":12
}
FILE_REGEX = re.compile(r"^SF-([a-z]{2})(\d{4})\.xls(x)?$", re.IGNORECASE)

# Columnas estándar (entidades)
TARGET_COLS_STD = [
    "Banca Múltiple",
    "Empresas Financieras",
    "Cajas Municipales",
    "Cajas Rurales de Ahorro y Crédito",
    "EDPYMEs",
    "Agrobanco",
    "Total",
]

# -------- utilidades --------
def norm_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    for a,b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u")]:
        s = s.replace(a,b)
    return re.sub(r"\s+"," ", s)

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
        # algunos .xls en realidad son .xlsx → probamos ambos
        return ["xlrd", "openpyxl"]
    return ["openpyxl", "xlrd"]

def read_excel_safe(path: Path, sheet_name=None, header=None, dtype=None):
    for eng in engines_to_try(path):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, header=header, dtype=dtype, engine=eng)
        except Exception:
            continue
    return None

def find_sheet_flexible(path: Path, wanted: str) -> Optional[str]:
    wanted_norm = norm_text(wanted)
    for eng in engines_to_try(path):
        try:
            xls = pd.ExcelFile(path, engine=eng)
            for s in xls.sheet_names:
                if norm_text(s) == wanted_norm:
                    return s
            for s in xls.sheet_names:
                ns = norm_text(s)
                if ("credito" in ns or "creditos" in ns) and ("se" in ns or "sector" in ns or "economico" in ns):
                    return s
            if xls.sheet_names:
                return xls.sheet_names[0]
        except Exception:
            continue
    return None

def find_header_row(df: pd.DataFrame, search_limit: int = 60) -> Optional[int]:
    target = "sector economico"
    n = min(len(df), search_limit)
    for i in range(n):
        row = df.iloc[i, :]
        for val in row:
            if norm_text(val) == target:
                return i
    return None

def map_columns_to_targets(cols: List[str]) -> Dict[str, str]:
    real_cols_norm = {norm_text(c): c for c in cols}
    mapping: Dict[str, str] = {}
    patterns = {
        "Banca Múltiple": ["banca multiple", "banca multiple."],
        "Empresas Financieras": ["empresas financieras", "empresas  financieras", "empresas\nfinancieras"],
        "Cajas Municipales": ["cajas municipales", "cajas\nmunicipales"],
        "Cajas Rurales de Ahorro y Crédito": [
            "cajas rurales de ahorro y credito", "caja rurales de ahorro y credito",
            "cajas rurales\nde ahorro y credito", "cajas rurales de ahorro y credito."
        ],
        "EDPYMEs": ["edpymes", "edpyme", "edpym es"],
        "Agrobanco": ["agrobanco"],
        "Total": ["total", "total (miles)", "total general", "total,"],
    }
    for std_name, keys in patterns.items():
        for k in keys:
            if k in real_cols_norm:
                mapping[real_cols_norm[k]] = std_name
                break
    return mapping

def clean_numbers(series: pd.Series) -> pd.Series:
    return (series.astype(str)
                  .str.replace(r"[\s]", "", regex=True)
                  .str.replace(",", "", regex=False)   # miles ES
                  .str.replace("%", "", regex=False)
                  .replace({"-": pd.NA, "–": pd.NA, "—": pd.NA, "nan": pd.NA, "None": pd.NA})
                  .pipe(pd.to_numeric, errors="coerce"))
0
# -------- procesamiento por archivo --------
def clean_one(path: Path) -> Optional[pd.DataFrame]:
    meta = parse_period(path.name)
    if not meta:
        print(f"[SKIP] Nombre no reconocido: {path.name}")
        return None

    sheet = find_sheet_flexible(path, WANTED_SHEET)
    if not sheet:
        print(f"[ERROR] No se pudo hallar hoja en {path.name}")
        return None

    raw = read_excel_safe(path, sheet_name=sheet, header=None, dtype=str)
    if raw is None:
        print(f"[ERROR] No se pudo leer '{sheet}' en {path.name}")
        return None

    header_row = find_header_row(raw)
    if header_row is None:
        print(f"[WARN] No se halló encabezado 'Sector Económico' en {path.name}.")
        return None

    # Encabezados + datos
    cols = raw.iloc[header_row].astype(str).tolist()
    cols = [re.sub(r"\s+", " ", c).strip() for c in cols]
    data = raw.iloc[header_row+1:].copy()
    data.columns = cols

    # Columna Sector
    sector_col = None
    for c in data.columns:
        nt = norm_text(c)
        if nt == "sector economico" or ("sector" in nt and "economico" in nt):
            sector_col = c
            break
    if not sector_col:
        print(f"[WARN] No se encontró 'Sector Económico' en {path.name}.")
        return None

    # Cortar antes de "Créditos Corporativos..." si existe
    mask_end = data[sector_col].astype(str).str.contains(r"^cr[eé]ditos?\s+corporativos", case=False, na=False)
    if mask_end.any():
        end_idx = mask_end.idxmax()
        data = data.loc[: end_idx - 1].copy()

    # Limpiar filas no-dato
    data[sector_col] = data[sector_col].astype(str).str.strip()
    mask_notes = (
        data[sector_col].str.match(r"^\s*(\*+|\d+/)", na=False) |
        data[sector_col].str.contains(r"(?:fuente|p[aá]gina)", case=False, na=False)
    )
    data = data[~mask_notes].copy()
    data = data[(data[sector_col].notna()) & (data[sector_col]!="")]

    # Mapear columnas de entidades
    col_map = map_columns_to_targets(list(data.columns))
    keep_real_cols = [sector_col] + list(col_map.keys())
    data = data[keep_real_cols].copy().rename(columns=col_map)

    # Numerificar
    for c in [x for x in data.columns if x != sector_col and x != "Sector Económico"]:
        data[c] = clean_numbers(data[c])

    # Renombrar sector
    if sector_col != "Sector Económico":
        data = data.rename(columns={sector_col: "Sector Económico"})

    # Metadatos
    data.insert(0, "year", meta["year"])
    data.insert(1, "date", meta["date"])

    # ---- TIDY ----
    value_cols = [c for c in TARGET_COLS_STD if c in data.columns]
    tidy = data.melt(
        id_vars=["year", "date", "Sector Económico"],
        value_vars=value_cols,
        var_name="Entidad",
        value_name="monto",
    )
    tidy = tidy.dropna(subset=["monto"]).reset_index(drop=True)

    # Orden sugerido
    entidad_order = [c for c in TARGET_COLS_STD if c in value_cols]
    tidy["Entidad"] = pd.Categorical(tidy["Entidad"], categories=entidad_order, ordered=True)
    tidy = tidy.sort_values(["year","date","Sector Económico","Entidad"]).reset_index(drop=True)

    return tidy

# -------- pipeline --------
def build_db(base: Path) -> pd.DataFrame:
    files = [base / f for f in os.listdir(base) if f.lower().endswith((".xls",".xlsx"))]
    files = [f for f in files if FILE_REGEX.match(f.name)]  # solo SF-*.xls[x]
    files.sort()
    frames: List[pd.DataFrame] = []
    for f in files:
        out = clean_one(f)
        if out is not None and not out.empty:
            frames.append(out)
    if not frames:
        raise RuntimeError("No se pudo construir la base; revisa archivos/hojas.")
    return pd.concat(frames, ignore_index=True, sort=False)

def main():
    if not BASE_DIR.exists():
        raise FileNotFoundError(f"No existe carpeta: {BASE_DIR}")
    db = build_db(BASE_DIR)
    db.to_excel(OUT_XLSX, index=False)
    print(f"Listo.\n- {OUT_XLSX}")

if __name__ == "__main__":
    main()