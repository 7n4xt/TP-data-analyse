"""
main.py — CLI entry point for the tourism analysis tool.

Usage
-----
  python main.py analyse   --file data/tourisme_brut.csv
  python main.py visualize --file data/tourisme_brut.csv [--output output/]
  python main.py report    --file data/tourisme_brut.csv [--output output/]
  python main.py clean     --file data/donnees_tourisme_france_exercice.csv
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Allow running directly from the project root without installing the package
sys.path.insert(0, str(Path(__file__).parent))

from tourisme.loader import load_data
from tourisme.analysis import TourismeAnalyser
from tourisme.visualizer import TourismeVisualizer


# ---------------------------------------------------------------------------
# Sub-command handlers
# ---------------------------------------------------------------------------

def cmd_clean(args: argparse.Namespace) -> None:
    """Load and clean a file, show what was fixed, optionally save cleaned CSV."""
    df = load_data(args.file, verbose=True)
    print(f"\nAperçu des données nettoyées ({len(df)} lignes) :")
    print(df.to_string())
    if args.save:
        out = Path(args.save)
        df.to_csv(out, index=False, sep=";", encoding="utf-8")
        print(f"\n[clean] Fichier nettoyé sauvegardé : {out}")


def cmd_analyse(args: argparse.Namespace) -> None:
    """Load data and print the text analysis report."""
    df = load_data(args.file)
    analyser = TourismeAnalyser(df)
    analyser.print_report()


def cmd_visualize(args: argparse.Namespace) -> None:
    """Load data and generate all PNG charts."""
    df = load_data(args.file)
    analyser = TourismeAnalyser(df)
    visualizer = TourismeVisualizer(analyser, output_dir=args.output)
    visualizer.generate_all()


def cmd_report(args: argparse.Namespace) -> None:
    """Run both the text analysis and the chart generation."""
    df = load_data(args.file)
    analyser = TourismeAnalyser(df)
    analyser.print_report()
    visualizer = TourismeVisualizer(analyser, output_dir=args.output)
    visualizer.generate_all()


# ---------------------------------------------------------------------------
# CLI definition
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="tourisme",
        description="Outil d'analyse et de visualisation de données touristiques.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemples :
  python main.py clean     --file data/donnees_tourisme_france_exercice.csv
  python main.py clean     --file data/donnees_tourisme_france_exercice.csv --save data/clean.csv
  python main.py analyse   --file data/tourisme_brut.csv
  python main.py visualize --file data/tourisme_brut.csv --output output/
  python main.py report    --file data/tourisme_brut.csv --output output/
        """,
    )

    sub = parser.add_subparsers(dest="command", metavar="COMMANDE")
    sub.required = True

    # --- clean ---
    p_clean = sub.add_parser(
        "clean",
        help="Nettoyer un fichier CSV et afficher le résultat (+ export optionnel).",
    )
    p_clean.add_argument("--file", "-f", required=True, metavar="CSV",
                         help="Fichier CSV source (peut être très sale).")
    p_clean.add_argument("--save", "-s", default=None, metavar="OUTPUT_CSV",
                         help="Chemin pour sauvegarder le CSV nettoyé (optionnel).")
    p_clean.set_defaults(func=cmd_clean)

    # --- analyse ---
    p_analyse = sub.add_parser(
        "analyse",
        help="Afficher le rapport statistique textuel.",
    )
    p_analyse.add_argument(
        "--file", "-f",
        required=True,
        metavar="CSV",
        help="Chemin vers le fichier CSV de données touristiques.",
    )
    p_analyse.set_defaults(func=cmd_analyse)

    # --- visualize ---
    p_viz = sub.add_parser(
        "visualize",
        help="Générer les graphiques statistiques (PNG).",
    )
    p_viz.add_argument("--file", "-f", required=True, metavar="CSV",
                       help="Fichier CSV source.")
    p_viz.add_argument("--output", "-o", default="output", metavar="DIR",
                       help="Dossier de destination pour les PNG (défaut: output/).")
    p_viz.set_defaults(func=cmd_visualize)

    # --- report ---
    p_report = sub.add_parser(
        "report",
        help="Rapport texte ET graphiques en une seule commande.",
    )
    p_report.add_argument("--file", "-f", required=True, metavar="CSV",
                          help="Fichier CSV source.")
    p_report.add_argument("--output", "-o", default="output", metavar="DIR",
                          help="Dossier de destination pour les PNG (défaut: output/).")
    p_report.set_defaults(func=cmd_report)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
