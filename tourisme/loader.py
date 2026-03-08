"""
loader.py — Load and clean any tourism CSV file.

Expected columns (flexible naming — see COLUMN_ALIASES):
    region, annee, mois, visiteurs, hebergement, depense_moyenne
"""

from __future__ import annotations

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

# Known mis-spellings / abbreviations in the hebergement column
HEBERGEMENT_FIXES: dict[str, str] = {
    "hote": "Hotel",
    "hôte": "Hôtel",
    "hotel": "Hotel",
    "hôtel": "Hôtel",
    "camping": "Camping",
    "gite": "Gîte",
    "gîte": "Gîte",
    "auberge": "Auberge",
    "airbnb": "Airbnb",
    "appartement": "Appartement",
}


def _detect_separator(filepath: str | Path) -> str:
    """Sniff the column separator from the first line of the file."""
    with open(filepath, encoding="utf-8", errors="replace") as f:
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


def load_data(filepath: str | Path) -> pd.DataFrame:
    """
    Load a tourism CSV file and return a clean DataFrame.

    Parameters
    ----------
    filepath : str or Path
        Path to the CSV file.

    Returns
    -------
    pd.DataFrame
        Cleaned and normalised DataFrame with canonical column names.

    Raises
    ------
    FileNotFoundError
        If the file does not exist.
    ValueError
        If required columns are missing after alias resolution.
    """
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    sep = _detect_separator(filepath)
    df = pd.read_csv(filepath, sep=sep, encoding="utf-8", on_bad_lines="warn")
    df = _rename_columns(df)

    # Validate required columns
    required = {"region", "annee", "mois", "visiteurs"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Missing required columns: {missing}. "
            f"Available columns: {list(df.columns)}"
        )

    # --- Clean region ---
    df["region"] = df["region"].astype(str).str.strip().str.upper()

    # --- Clean hebergement (if present) ---
    if "hebergement" in df.columns:
        df["hebergement"] = (
            df["hebergement"]
            .astype(str)
            .str.strip()
            .str.lower()
            .map(lambda v: HEBERGEMENT_FIXES.get(v, v.capitalize()))
        )

    # --- Coerce numeric columns ---
    for col in ["annee", "mois", "visiteurs"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    if "depense_moyenne" in df.columns:
        df["depense_moyenne"] = pd.to_numeric(df["depense_moyenne"], errors="coerce")

    # Drop rows where key numeric columns are null
    before = len(df)
    df = df.dropna(subset=["region", "annee", "mois", "visiteurs"])
    dropped = before - len(df)
    if dropped:
        print(f"[loader] Dropped {dropped} row(s) with missing key values.")

    df["annee"] = df["annee"].astype(int)
    df["mois"] = df["mois"].astype(int)
    df["visiteurs"] = df["visiteurs"].astype(int)

    return df.reset_index(drop=True)
