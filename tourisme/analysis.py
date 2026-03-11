"""
analysis.py — Statistical analysis helpers for tourism data.
"""

from __future__ import annotations

import pandas as pd

MONTH_NAMES = {
    1: "Jan", 2: "Fév", 3: "Mar", 4: "Avr",
    5: "Mai", 6: "Jun", 7: "Jul", 8: "Aoû",
    9: "Sep", 10: "Oct", 11: "Nov", 12: "Déc",
}


class TourismeAnalyser:
    """
    Wraps a cleaned tourism DataFrame and exposes reusable analysis methods.

    Parameters
    ----------
    df : pd.DataFrame
        Output of ``tourisme.loader.load_data()``.
    """

    def __init__(self, df: pd.DataFrame) -> None:
        self.df = df.copy()

    # ------------------------------------------------------------------
    # Overview
    # ------------------------------------------------------------------

    def overview(self) -> dict:
        """Return high-level dataset statistics as a plain dict."""
        df = self.df
        result: dict = {
            "rows": len(df),
            "columns": list(df.columns),
            "regions": sorted(df["region"].unique().tolist()),
            "years": sorted(df["annee"].unique().tolist()),
            "total_visitors": int(df["visiteurs"].sum()),
        }
        if "hebergement" in df.columns:
            result["accommodation_types"] = sorted(
                df["hebergement"].dropna().unique().tolist()
            )
        if "depense_moyenne" in df.columns:
            result["avg_spending_overall"] = round(float(df["depense_moyenne"].mean()), 2)
        return result

    # ------------------------------------------------------------------
    # Aggregations
    # ------------------------------------------------------------------

    def visitors_by_region(self) -> pd.DataFrame:
        """Total visitors per region, sorted descending."""
        return (
            self.df.groupby("region")["visiteurs"]
            .sum()
            .sort_values(ascending=False)
            .reset_index()
            .rename(columns={"visiteurs": "total_visiteurs"})
        )

    def visitors_by_year(self) -> pd.DataFrame:
        """Total visitors per year."""
        return (
            self.df.groupby("annee")["visiteurs"]
            .sum()
            .reset_index()
            .rename(columns={"visiteurs": "total_visiteurs"})
        )

    def visitors_by_month(self) -> pd.DataFrame:
        """Average visitors per calendar month across all regions/years."""
        return (
            self.df.groupby("mois")["visiteurs"]
            .mean()
            .reset_index()
            .rename(columns={"visiteurs": "moy_visiteurs"})
            .assign(mois_nom=lambda d: d["mois"].map(MONTH_NAMES))
        )

    def monthly_trend(self, region: str | None = None) -> pd.DataFrame:
        """
        Visitors over time (year-month) for one region or all regions combined.
        Returns DataFrame with columns: annee, mois, mois_nom, visiteurs.
        """
        df = self.df if region is None else self.df[self.df["region"] == region.upper()]
        grouped = (
            df.groupby(["annee", "mois"])["visiteurs"]
            .sum()
            .reset_index()
            .sort_values(["annee", "mois"])
        )
        grouped["mois_nom"] = grouped["mois"].map(MONTH_NAMES)
        grouped["label"] = grouped["annee"].astype(str) + "-" + grouped["mois"].astype(str).str.zfill(2)
        return grouped

    def spending_by_region(self) -> pd.DataFrame | None:
        """Average spending per region. Returns None if column is absent."""
        if "depense_moyenne" not in self.df.columns:
            return None
        return (
            self.df.groupby("region")["depense_moyenne"]
            .mean()
            .round(2)
            .sort_values(ascending=False)
            .reset_index()
            .rename(columns={"depense_moyenne": "depense_moy"})
        )

    def accommodation_distribution(self) -> pd.DataFrame | None:
        """Visitor count per accommodation type. Returns None if absent."""
        if "hebergement" not in self.df.columns:
            return None
        return (
            self.df.groupby("hebergement")["visiteurs"]
            .sum()
            .reset_index()
            .rename(columns={"visiteurs": "total_visiteurs"})
            .sort_values("total_visiteurs", ascending=False)
        )

    def top_months(self, n: int = 3) -> pd.DataFrame:
        """Return the top-n months by average visitor count."""
        return self.visitors_by_month().nlargest(n, "moy_visiteurs")

    # ------------------------------------------------------------------
    # Text report
    # ------------------------------------------------------------------

    def print_report(self) -> None:
        """Print a formatted text summary to stdout."""
        ov = self.overview()

        print("=" * 60)
        print("  RAPPORT D'ANALYSE TOURISTIQUE")
        print("=" * 60)
        print(f"  Lignes        : {ov['rows']}")
        print(f"  Régions       : {', '.join(ov['regions'])}")
        print(f"  Années        : {', '.join(str(y) for y in ov['years'])}")
        print(f"  Total visiteurs: {ov['total_visitors']:,}")
        if "avg_spending_overall" in ov:
            print(f"  Dépense moy.  : {ov['avg_spending_overall']} €")
        if "accommodation_types" in ov:
            print(f"  Hébergements  : {', '.join(ov['accommodation_types'])}")

        print("\n--- Visiteurs par région ---")
        print(self.visitors_by_region().to_string(index=False))

        print("\n--- Visiteurs par année ---")
        print(self.visitors_by_year().to_string(index=False))

        print("\n--- Top 3 mois (moyenne visiteurs) ---")
        print(self.top_months(3).to_string(index=False))

        sp = self.spending_by_region()
        if sp is not None:
            print("\n--- Dépense moyenne par région ---")
            print(sp.to_string(index=False))

        acc = self.accommodation_distribution()
        if acc is not None:
            print("\n--- Distribution des hébergements ---")
            print(acc.to_string(index=False))

        print("=" * 60)
