from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from constants.constant import (
    DEFAULT_COMPARE_OUTPUT,
    DEFAULT_VALIDATE_OUTPUT,
)
from export.cmd import cmd_validate, cmd_compare

def main() -> int:
    start_time = time.perf_counter()
    parser = argparse.ArgumentParser(
        description="Comparaison de fichiers magasins Nickel pour la synchronisation FTP."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    compare_parser = subparsers.add_parser(
        "compare",
        help="Compare le fichier du mois précédent avec le nouveau fichier Nickel",
    )
    compare_parser.add_argument(
        "previous",
        help="Fichier Nickel du mois précédent (déjà synchronisé)",
    )
    compare_parser.add_argument("incoming", help="Nouveau fichier Nickel reçu")
    compare_parser.add_argument(
        "-o",
        "--output",
        default=str(DEFAULT_COMPARE_OUTPUT),
        help="Fichier Excel pour Filezilla (défaut: output/resultat/fichier_synchronisation.xlsx)",
    )
    compare_parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="Masquer l'indicateur de progression",
    )
    compare_parser.set_defaults(func=cmd_compare)

    validate_parser = subparsers.add_parser(
        "validate", help="Contrôle la qualité des données d'un fichier (usage CS)"
    )
    validate_parser.add_argument("fichier", help="Fichier Nickel à contrôler")
    validate_parser.add_argument(
        "-o",
        "--output",
        default=str(DEFAULT_VALIDATE_OUTPUT),
        help="Rapport Excel dans output/log (défaut: output/log/problemes_qualite.xlsx)",
    )
    validate_parser.set_defaults(func=cmd_validate)

    args = parser.parse_args()
    result = args.func(args)

    elapsed = time.perf_counter() - start_time
    print(f"\nTemps d'exécution : {elapsed:.2f} secondes")
    return result


if __name__ == "__main__":
    sys.exit(main())
