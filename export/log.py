
from __future__ import annotations
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
import pandas as pd

from constants.constant import (
    ACTION_UPDATE,
    ACTION_CLOSE,
    ACTION_CREATE,
    OUTPUT_LOG,
    ACTION_NO_CHANGE,
)
from comparison.compare import ComparisonResult
from .excel_export import *

def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def timestamped_log_path(prefix: str, stamp: str | None = None) -> Path:
    return OUTPUT_LOG / f"{prefix}_{stamp or timestamp()}.log"


def timestamped_analysis_path(stamp: str | None = None) -> Path:
    return OUTPUT_LOG / f"analyse_{stamp or timestamp()}.xlsx"


def write_log(path: Path, lines: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")

def _build_analysis_summary(
    result: ComparisonResult,
    previous: Path,
    incoming: Path,
    resultat: Path,
    dup_count: int,
) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"indicateur": "date", "valeur": datetime.now().isoformat(timespec="seconds")},
            {"indicateur": "fichier_mois_precedent", "valeur": str(previous.resolve())},
            {"indicateur": "fichier_nouveau", "valeur": str(incoming.resolve())},
            {"indicateur": "fichier_resultat", "valeur": str(resultat.resolve())},
            {"indicateur": "total_reference", "valeur": result.reference_total},
            {"indicateur": "total_entrant", "valeur": result.incoming_total},
            {"indicateur": ACTION_NO_CHANGE, "valeur": result.unchanged_count},
            {"indicateur": ACTION_UPDATE, "valeur": len(result.modified)},
            {"indicateur": ACTION_CREATE, "valeur": len(result.added)},
            {"indicateur": ACTION_CLOSE, "valeur": len(result.removed)},
            {"indicateur": "lignes_filezilla", "valeur": len(result.modified) + len(result.added) + len(result.removed)},
            {"indicateur": "doublons_entrant", "valeur": dup_count},
        ]
    )


def write_analysis_workbook(
    path: Path,
    result: ComparisonResult,
    previous: Path,
    incoming: Path,
    resultat: Path,
    dup_count: int,
    inc_duplicates: pd.DataFrame,
    quality: pd.DataFrame,
) -> None:
    write_workbook(
        path,
        {
            "resume": _build_analysis_summary(result, previous, incoming, resultat, dup_count),
            "no_change": with_action(result.unchanged, ACTION_NO_CHANGE),
            "update": with_action(result.modified, ACTION_UPDATE),
            "create": with_action(result.added, ACTION_CREATE),
            "close_permanently": with_action(result.removed, ACTION_CLOSE),
            "champs_modifies": result.changed_fields,
            "doublons": inc_duplicates,
            "qualite": quality,
        },
    )
