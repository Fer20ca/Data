#INPUTS DE FECHA AL FINAL
import argparse
import datetime as dt
import io
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List

try:
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
except Exception:
    print("ERROR: Necesitas instalar 'requests' (pip install requests)")
    raise

SPANISH_MONTH_FOLDER = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Setiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

SPANISH_MONTH_ABBREV = {
    1: "en",
    2: "fe",
    3: "ma",
    4: "ab",
    5: "my",
    6: "jn",
    7: "jl",
    8: "ag",
    9: "se",
    10: "oc",
    11: "no",
    12: "di",
}


@dataclass
class Period:
    year: int
    month: int

    @property
    def folder_name(self) -> str:
        return SPANISH_MONTH_FOLDER[self.month]

    @property
    def abbrev(self) -> str:
        return SPANISH_MONTH_ABBREV[self.month]

    def __str__(self) -> str:
        return f"{self.folder_name} {self.year}"


def month_range(start: str, end: str) -> List[Period]:
    """Genera Periods desde start YYYY-MM hasta end YYYY-MM (inclusive)."""
    y1, m1 = [int(x) for x in start.split("-")]
    y2, m2 = [int(x) for x in end.split("-")]
    d1 = dt.date(y1, m1, 1)
    d2 = dt.date(y2, m2, 1)
    if d1 > d2: 
        raise ValueError("El inicio no puede ser posterior al fin.")
    periods: List[Period] = []
    y, m = y1, m1
    while True:
        periods.append(Period(year=y, month=m))
        if (y, m) == (y2, m2):
            break
        # avanzar un mes
        if m == 12:
            y += 1
            m = 1
        else:
            m += 1
    return periods


def build_url(p: Period, serie: str) -> str:
    # https://intranet2.sbs.gob.pe/estadistica/financiera/{YYYY}/{Mes}/SF-{serie}-{abbrev}{YYYY}.ZIP
    return (
        f"https://intranet2.sbs.gob.pe/estadistica/financiera/"
        f"{p.year}/{p.folder_name}/SF-{serie}-{p.abbrev}{p.year}.ZIP"
    )


def make_session() -> requests.Session:
    session = requests.Session()
    retries = Retry(
        total=3,
        backoff_factor=0.8,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD"],
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retries, pool_connections=10, pool_maxsize=10)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update(
        {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) SBS-Downloader/1.1"}
    )
    return session


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def safe_write(path: Path, data: bytes) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_bytes(data)
    tmp.replace(path)


def download_zip(session: requests.Session, url: str, zip_path: Path) -> bool:
    """Descarga el ZIP si no existe. True si descargó, False si ya existía."""
    if zip_path.exists():
        return False
    r = session.get(url, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code} para {url}")
    safe_write(zip_path, r.content)
    return True


def extract_and_rename(zip_path: Path, out_dir: Path, rename_pattern: str, p: Period) -> List[Path]:
    """
    Extrae el ZIP al out_dir y renombra cualquier archivo extraído según rename_pattern.
    rename_pattern tokens: {year}, {month_num:02d}, {month_name}, {abbrev}
    Retorna rutas finales (renombradas).
    """
    result: List[Path] = []
    with zipfile.ZipFile(zip_path, "r") as zf:
        for member in zf.namelist():
            member_name = Path(member).name
            # Nombre final deseado
            final_name = rename_pattern.format(
                year=p.year,
                month_num=p.month,
                month_name=SPANISH_MONTH_FOLDER[p.month],
                abbrev=p.abbrev,
            )
            final_path = out_dir / final_name

            with zf.open(member, "r") as src:
                data = src.read()

            # Si ya existe, sobreescribir de forma segura
            safe_write(final_path, data)
            result.append(final_path)

    return result


def run(start: str, end: str, serie: str, out_dir: Path, pattern: str, keep_zip: bool) -> None:
    ensure_dir(out_dir)
    periods = month_range(start, end)
    print(f"Procesando {len(periods)} periodo(s): {start} → {end}, serie {serie}")
    session = make_session()

    for p in periods:
        url = build_url(p, serie)
        zip_name = f"SF-{serie}-{p.abbrev}{p.year}.ZIP"
        zip_path = out_dir / zip_name
        try:
            downloaded = download_zip(session, url, zip_path)
            print(f"[{p}] {'Descargado' if downloaded else 'Ya existe'}: {zip_name}")
        except Exception as e:
            print(f"[{p}] ERROR descarga: {e}")
            continue

        try:
            outputs = extract_and_rename(zip_path, out_dir, pattern, p)
            if outputs:
                names = ", ".join(x.name for x in outputs)
                print(f"[{p}] Extraído y renombrado → {names}")
        except zipfile.BadZipFile:
            print(f"[{p}] ERROR: ZIP corrupto o inesperado: {zip_name}")
            continue
        except Exception as e:
            print(f"[{p}] ERROR extracción: {e}")
            continue

        if not keep_zip:
            try:
                zip_path.unlink(missing_ok=True)
                print(f"[{p}] ZIP eliminado.")
            except Exception as e:
                print(f"[{p}] Aviso: no se pudo borrar {zip_name}: {e}")


def parse_args(argv: List[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Descarga, extrae y renombra Excel de la SBS por rango de fechas.")
    ap.add_argument("--start", type=str, required=True, help="Inicio YYYY-MM (ej: 2023-01)")
    ap.add_argument("--end", type=str, required=True, help="Fin YYYY-MM (ej: 2025-06)")
    ap.add_argument("--serie", type=str, default="2101", help="Código de serie (ej: 2101)")
    ap.add_argument("--out", type=str, required=True, help="Carpeta destino para Excel (y ZIPs si keep-zip)")
    ap.add_argument(
        "--pattern",
        type=str,
        default="SF-{abbrev}{year}.xls",
        help="Patrón de nombre final. Tokens: {year}, {month_num:02d}, {month_name}, {abbrev}",
    )
    ap.add_argument("--keep-zip", action="store_true", help="Conservar ZIPs tras extraer")
    return ap.parse_args(argv)

# Inputs para correr el codigo 
def main() -> None:
    START = "2018-01"   # fecha inicio
    END   = "2025-08"   # fecha fin
    SERIE = "2101"      # serie por defecto de SF
    OUT   = Path(r"C:\Users\Raisa Sullca\OneDrive - Crece Finanzas Estratégicas\Documentos\Codigos\Web Scraping\SBS\morosidad") #definir carpeta de base de datos

    run(
        start=START,
        end=END,
        serie=SERIE,
        out_dir=OUT,
        pattern="SF-{abbrev}{year}.xls",
        keep_zip=False,
    )


if __name__ == "__main__":
    main()

