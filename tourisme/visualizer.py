"""
visualizer.py — Generate and save statistical charts for tourism data.

All charts are saved as PNG files to the specified output directory.
"""

from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # non-interactive backend — safe for CLI / servers

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns

from .analysis import TourismeAnalyser

# Apply a clean global style
sns.set_theme(style="whitegrid", palette="muted", font_scale=1.1)


class TourismeVisualizer:
    """
    Generate PNG charts from a :class:`TourismeAnalyser` instance.

    Parameters
    ----------
    analyser : TourismeAnalyser
    output_dir : str or Path
        Directory where PNG files will be written (created if absent).
    """

    def __init__(self, analyser: TourismeAnalyser, output_dir: str | Path = "output") -> None:
        self.analyser = analyser
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _save(self, fig: plt.Figure, name: str) -> Path:
        dest = self.output_dir / name
        fig.savefig(dest, bbox_inches="tight", dpi=150)
        plt.close(fig)
        print(f"  [saved] {dest}")
        return dest

    @staticmethod
    def _fmt_thousands(x: float, _pos=None) -> str:
        if x >= 1_000_000:
            return f"{x/1_000_000:.1f}M"
        if x >= 1_000:
            return f"{x/1_000:.0f}k"
        return str(int(x))

    # ------------------------------------------------------------------
    # Charts
    # ------------------------------------------------------------------

    def plot_visitors_by_region(self) -> Path:
        """Horizontal bar chart: total visitors per region."""
        data = self.analyser.visitors_by_region()

        fig, ax = plt.subplots(figsize=(10, max(4, len(data) * 0.6)))
        bars = ax.barh(data["region"], data["total_visiteurs"], color=sns.color_palette("muted"))
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(self._fmt_thousands))
        ax.set_xlabel("Total visiteurs")
        ax.set_title("Total des visiteurs par région", fontweight="bold")
        ax.invert_yaxis()

        # Annotate bars
        for bar in bars:
            w = bar.get_width()
            ax.text(
                w * 1.01, bar.get_y() + bar.get_height() / 2,
                self._fmt_thousands(w),
                va="center", ha="left", fontsize=9,
            )
        fig.tight_layout()
        return self._save(fig, "01_visiteurs_par_region.png")

    def plot_visitors_by_year(self) -> Path:
        """Bar chart: total visitors per year."""
        data = self.analyser.visitors_by_year()

        fig, ax = plt.subplots(figsize=(8, 5))
        colors = sns.color_palette("muted", len(data))
        bars = ax.bar(data["annee"].astype(str), data["total_visiteurs"], color=colors)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(self._fmt_thousands))
        ax.set_xlabel("Année")
        ax.set_ylabel("Total visiteurs")
        ax.set_title("Évolution annuelle des visiteurs", fontweight="bold")

        for bar in bars:
            h = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width() / 2, h * 1.005,
                self._fmt_thousands(h),
                ha="center", va="bottom", fontsize=9,
            )
        fig.tight_layout()
        return self._save(fig, "02_visiteurs_par_annee.png")

    def plot_monthly_seasonality(self) -> Path:
        """Line chart: average visitors across calendar months."""
        data = self.analyser.visitors_by_month()

        fig, ax = plt.subplots(figsize=(11, 5))
        ax.plot(
            data["mois_nom"], data["moy_visiteurs"],
            marker="o", linewidth=2.2, color="#2a7ae2",
        )
        ax.fill_between(data["mois_nom"], data["moy_visiteurs"], alpha=0.15, color="#2a7ae2")
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(self._fmt_thousands))
        ax.set_xlabel("Mois")
        ax.set_ylabel("Visiteurs (moy.)")
        ax.set_title("Saisonnalité mensuelle (moyenne toutes régions)", fontweight="bold")
        ax.tick_params(axis="x", rotation=30)
        fig.tight_layout()
        return self._save(fig, "03_saisonnalite_mensuelle.png")

    def plot_spending_by_region(self) -> Path | None:
        """Horizontal bar chart: average spending per region."""
        data = self.analyser.spending_by_region()
        if data is None:
            print("  [skip] Column 'depense_moyenne' not found — skipping spending chart.")
            return None

        fig, ax = plt.subplots(figsize=(10, max(4, len(data) * 0.6)))
        bars = ax.barh(data["region"], data["depense_moy"], color=sns.color_palette("coolwarm_r", len(data)))
        ax.set_xlabel("Dépense moyenne (€)")
        ax.set_title("Dépense moyenne par région", fontweight="bold")
        ax.invert_yaxis()

        for bar in bars:
            w = bar.get_width()
            ax.text(
                w * 1.005, bar.get_y() + bar.get_height() / 2,
                f"{w:.1f} €",
                va="center", ha="left", fontsize=9,
            )
        fig.tight_layout()
        return self._save(fig, "04_depense_par_region.png")

    def plot_accommodation_distribution(self) -> Path | None:
        """Pie chart: share of visitors per accommodation type."""
        data = self.analyser.accommodation_distribution()
        if data is None:
            print("  [skip] Column 'hebergement' not found — skipping accommodation chart.")
            return None

        fig, ax = plt.subplots(figsize=(7, 7))
        wedges, texts, autotexts = ax.pie(
            data["total_visiteurs"],
            labels=data["hebergement"],
            autopct="%1.1f%%",
            startangle=140,
            colors=sns.color_palette("pastel"),
        )
        for t in autotexts:
            t.set_fontsize(10)
        ax.set_title("Répartition des visiteurs par type d'hébergement", fontweight="bold")
        fig.tight_layout()
        return self._save(fig, "05_repartition_hebergement.png")

    def plot_heatmap_region_month(self) -> Path:
        """Heatmap: visitors by region × month."""
        pivot = (
            self.analyser.df
            .groupby(["region", "mois"])["visiteurs"]
            .sum()
            .unstack(fill_value=0)
        )
        pivot.columns = [f"Mois {m:02d}" for m in pivot.columns]

        fig, ax = plt.subplots(figsize=(14, max(5, len(pivot) * 0.7)))
        sns.heatmap(
            pivot / 1000,
            annot=True, fmt=".0f",
            linewidths=0.4, cmap="YlOrRd", ax=ax,
            cbar_kws={"label": "Visiteurs (milliers)"},
        )
        ax.set_title("Visiteurs par région et par mois (milliers)", fontweight="bold")
        ax.set_ylabel("Région")
        ax.set_xlabel("Mois")
        fig.tight_layout()
        return self._save(fig, "06_heatmap_region_mois.png")

    # ------------------------------------------------------------------
    # Generate all charts in one call
    # ------------------------------------------------------------------

    def generate_all(self) -> list[Path]:
        """
        Produce every available chart and return the list of saved paths.
        """
        print(f"\nGenerating charts → {self.output_dir}/\n")
        saved: list[Path] = []
        for method in [
            self.plot_visitors_by_region,
            self.plot_visitors_by_year,
            self.plot_monthly_seasonality,
            self.plot_spending_by_region,
            self.plot_accommodation_distribution,
            self.plot_heatmap_region_month,
        ]:
            result = method()
            if result is not None:
                saved.append(result)
        print(f"\nDone — {len(saved)} chart(s) saved to '{self.output_dir}/'")
        return saved
