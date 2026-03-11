"""
loader.py — Load and clean any tourism CSV file.

Works with both clean files and very messy real-world files.
Handles: encoding issues, typos, invalid values, bad months/years,
         region aliases, European decimal format (comma), NA strings.

Expected columns (flexible naming — see COLUMN_ALIASES):
    region, annee, mois, visiteurs, hebergement, depense_moyenne
"""

from __future__ import annotations

import re
import pandas as pd
from pathlib import Path

# Maps canonical name → list of accepted aliases (case-insensitive)
COLUMN_ALIASES: dict[str, list[str]] = {
    "region":          ["region", "régions", "zone", "area"],
    "annee":           ["annee", "année", "year", "an"],
    "mois":            ["mois", "month", "mois_num"],
    "visiteurs":       ["visiteurs", "visitors", "nb_visiteurs", "tourists"],
    "hebergement":     ["hebergement", "hébergement", "accommodation", "type_hebergement"],
    "depense_moyenne": ["depense_moyenne", "dépense_moyenne", "avg_spend", "spending"],
}

# French month name → number
MONTH_NAME_MAP: dict[str, int] = {
    "janvier": 1, "février": 2, "fevrier": 2, "mars": 3,
    "avril": 4, "mai": 5, "juin": 6, "juillet": 7,
    "août": 8, "aout": 8, "septembre": 9, "octobre": 10,
    "novembre": 11, "décembre": 12, "decembre": 12,
    "jan": 1, "fév": 2, "fev": 2, "mar": 3, "avr": 4,
    "jui": 6, "jul": 7, "aoû": 8, "sep": 9, "oct": 10,
    "nov": 11, "déc": 12, "dec": 12,
}

# Normalized region names (handles encoding artifacts + aliases)
REGION_FIXES: dict[str, str] = {
    # aliases
    "paca":                             "PROVENCE-ALPES-CÔTE D'AZUR",
    "ile de france":                    "ÎLE-DE-FRANCE",
    "idf":                              "ÎLE-DE-FRANCE",
    # encoding artifacts for Île-de-France  (× = bad 0xc3 → Î)
    "×le-de-france":                    "ÎLE-DE-FRANCE",
    "\xc3\xaele-de-france":             "ÎLE-DE-FRANCE",
    # other common misspellings
    "auvergne-rhone-alpes":             "AUVERGNE-RHÔNE-ALPES",
    "bourgogne-franche-comte":          "BOURGOGNE-FRANCHE-COMTÉ",
    "provence-alpes-cote d'azur":       "PROVENCE-ALPES-CÔTE D'AZUR",
    "provence-alpes-c\x93te d'azur":    "PROVENCE-ALPES-CÔTE D'AZUR",
    "auvergne-rh\x93ne-alpes":          "AUVERGNE-RHÔNE-ALPES",
    "bourgogne-franche-comt\x82":       "BOURGOGNE-FRANCHE-COMTÉ",
}

# Known mis-spellings / abbreviations in the hebergement column
HEBERGEMENT_FIXES: dict[str, str] = {
    "hote":         "Hôtel",
    "hôte":         "Hôtel",
    "hotel":        "Hôtel",
    "hôtel":        "Hôtel",
    "h\x93tel":     "Hôtel",   # latin-1 artifact
    "hotl":         "Hôtel",
    "hot":          "Hôtel",
    "camping":      "Camping",
    "gite":         "Gîte",
    "gîte":         "Gîte",
    "auberge":      "Auberge",
    "airbnb":       "Airbnb",
    "appartement":  "Appartement",
    "location":     "Location",
}


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _detect_encoding(filepath: Path) -> str:
    """Try UTF-8 first, fall back to latin-1."""
    for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        try:
            with open(filepath, encoding=enc) as f:
                f.read(4096)
            return enc
        except (UnicodeDecodeError, ValueError):
            continue
    return "latin-1"


def _detect_separator(filepath: Path, encoding: str) -> str:
    """Sniff the column separator from the first line of the file."""
    with open(filepath, encoding=encoding, errors="replace") as f:
        first_line = f.readline()
    for sep in (";", ",", "\t", "|"):
        if sep in first_line:
            return sep
    return ","


def _rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename aliased column names to canonical names."""
    lower_cols = {c.lower().strip(): c for c in df.columns}
    rename_map: dict[str, str] = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in lower_cols:
                original = lower_cols[alias]
                if original != canonical:
                    rename_map[original] = canonical
                break
    return df.rename(columns=rename_map)


def _fix_region(val: str) -> str:
    """Normalize a region name: fix encoding artifacts, aliases, upper-case."""
    if pd.isna(val) or str(val).strip() in ("", "nan"):
        return ""
    v = str(val).strip()
    # Try direct alias lookup (case-insensitive)
    key = v.lower()
    if key in REGION_FIXES:
        return REGION_FIXES[key]
    # Remove accents induced by bad encoding and retry
    clean = v.upper()
    # Replace known single-byte artifacts
    clean = clean.replace("\x93", "Ô").replace("\x82", "É").replace("\x8e", "Â")
    return clean


def _fix_hebergement(val: str) -> str:
    """Normalize accommodation type."""
    if pd.isna(val) or str(val).strip() in ("", "nan", "na"):
        return ""
    v = str(val).strip().lower()
    return HEBERGEMENT_FIXES.get(v, v.capitalize())


def _parse_month(val) -> float:
    """Convert a month value to int 1-12, NaN if invalid."""
    if pd.isna(val):
        return float("nan")
    s = str(val).strip().lower()
    if s in MONTH_NAME_MAP:
        return float(MONTH_NAME_MAP[s])
    try:
        m = float(s)
        if 1 <= m <= 12:
            return m
        return float("nan")   # 0 or 13+ → invalid
    except ValueError:
        return float("nan")


def _parse_year(val) -> float:
    """Convert a year value to int, NaN if clearly invalid."""
    if pd.isna(val):
        return float("nan")
    s = str(val).strip()
    # Remove any non-digit character (handles '202A', '2O23', etc.)
    digits_only = re.sub(r"[^\d]", "", s)
    if not digits_only:
        return float("nan")
    try:
        y = int(digits_only)
        if 1900 <= y <= 2100:
            return float(y)
        return float("nan")   # year 18, 24, 20, 1820 → invalid
    except ValueError:
        return float("nan")


def _parse_visitors(val) -> float:
    """Convert visitors value to positive int, NaN if invalid."""
    if pd.isna(val):
        return float("nan")
    s = str(val).strip().lower()
    if s in ("beaucoup", "many", "much", "lots", "na", ""):
        return float("nan")
    try:
        v = float(s)
        if v > 0:
            return v
        return float("nan")   # 0 or negative → invalid
    except ValueError:
        return float("nan")


def _parse_spending(val) -> float:
    """Parse European decimal format and reject invalid values."""
    if pd.isna(val):
        return float("nan")
    s = str(val).strip()
    if s.lower() in ("na", "n/a", "", "nan"):
        return float("nan")
    # European comma → dot
    s = s.replace(",", ".")
    try:
        v = float(s)
        if v > 0:
            return v
        return float("nan")   # negative or zero → invalid
    except ValueError:
        return float("nan")


# ─────────────────────────────────────────────────────────────────────────────
# Public function
# ─────────────────────────────────────────────────────────────────────────────

def load_data(filepath: str | Path, verbose: bool = True) -> pd.DataFrame:
    """
    Load a tourism CSV file and return a clean DataFrame.

    Handles UTF-8 and latin-1 encodings, auto-detects separator,
    fixes region aliases, accommodation typos, invalid month/year/
    visitor values, European decimal format, and NA strings.

    Parameters
    ----------
    filepath : str or Path
    verbose  : bool — print cleaning summary (default True)

    Returns
    -------
    pd.DataFrame  Clean, normalised DataFrame with canonical column names.

    Raises
    ------
    FileNotFoundError  If the file does not exist.
    ValueError         If required columns are missing.
    """
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    encoding = _detect_encoding(filepath)
    sep = _detect_separator(filepath, encoding)
    df = pd.read_csv(
        filepath, sep=sep, encoding=encoding,
        dtype=str,           # read everything as string first
        on_bad_lines="warn",
        na_values=["", "NA", "N/A", "na", "n/a", "NaN"],
        keep_default_na=False,
    )
    initial_rows = len(df)
    df = _rename_columns(df)

    # Validate required columns
    required = {"region", "annee", "mois", "visiteurs"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Missing required columns: {missing}. "
            f"Available columns: {list(df.columns)}"
        )

    # ── Region ──
    df["region"] = df["region"].apply(_fix_region)
    df = df[df["region"] != ""]   # drop blank region

    # ── Year ──
    df["annee"] = df["annee"].apply(_parse_year)

    # ── Month ──
    df["mois"] = df["mois"].apply(_parse_month)

    # ── Visitors ──
    df["visiteurs"] = df["visiteurs"].apply(_parse_visitors)

    # ── Accommodation ──
    if "hebergement" in df.columns:
        df["hebergement"] = df["hebergement"].apply(_fix_hebergement)
        df["hebergement"] = df["hebergement"].replace("", pd.NA)

    # ── Spending ──
    if "depense_moyenne" in df.columns:
        df["depense_moyenne"] = df["depense_moyenne"].apply(_parse_spending)

    # ── Drop rows where key values are still invalid ──
    df = df.dropna(subset=["annee", "mois", "visiteurs"])
    df["annee"]     = df["annee"].astype(int)
    df["mois"]      = df["mois"].astype(int)
    df["visiteurs"] = df["visiteurs"].astype(int)

    dropped = initial_rows - len(df)
    if verbose:
        print(f"[loader] {filepath.name}: {initial_rows} lignes lues → "
              f"{len(df)} lignes valides ({dropped} supprimées)")

    return df.reset_index(drop=True)
