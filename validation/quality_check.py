from __future__ import annotations
from dataclasses import dataclass, field
import pandas as pd

from constants.constant import (
    STORE_ID_COLUMN,
    SCHEDULE_COLUMNS,
)
from comparison.progress import *
from comparison.utils import normalize_cell


def _issue(store_id: str, type_probleme: str, detail: str, gravite: str) -> dict[str, str]:
    return {
        STORE_ID_COLUMN: store_id,
        "type_probleme": type_probleme,
        "detail": detail,
        "gravite": gravite,
    }

def _to_float(value: object) -> float | None:
    text = normalize_cell(value)
    if not text:
        return None
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return None

def deduplicate_stores(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Garde la dernière occurrence par code magasin. Retourne (unique, doublons)."""
    df = df.copy()
    df[STORE_ID_COLUMN] = df[STORE_ID_COLUMN].map(normalize_cell)
    duplicated_mask = df[STORE_ID_COLUMN].duplicated(keep=False)
    duplicates = df[duplicated_mask].sort_values(STORE_ID_COLUMN)
    unique = df.drop_duplicates(subset=[STORE_ID_COLUMN], keep="last")
    return unique, duplicates

def find_data_quality_issues(df: pd.DataFrame, *, show_progress: bool = True) -> pd.DataFrame:
    """Rapport CS : erreurs bloquantes et avertissements informatifs."""
    issues: list[dict[str, str]] = []

    duplicated = df[df.duplicated(subset=[STORE_ID_COLUMN], keep=False)]
    for store_id in duplicated[STORE_ID_COLUMN].unique():
        issues.append(
            _issue(
                store_id,
                "doublon",
                "Code magasin présent plusieurs fois dans le fichier",
                "erreur",
            )
        )

    progress = ProgressReporter("Controle qualite", len(df), enabled=show_progress)
    for index, (_, row) in enumerate(df.iterrows(), start=1):
        store_id = normalize_cell(row[STORE_ID_COLUMN])
        for field, label in {
            "Nom_etablissement": "nom manquant",
            "Ligne_Adresse_1": "adresse manquante",
            "Localite": "ville manquante",
            "pos_zipcode": "code postal manquant",
            "Latitude": "latitude manquante",
            "Longitude": "longitude manquante",
        }.items():
            if not normalize_cell(row[field]):
                issues.append(_issue(store_id, "champ_manquant", label, "erreur"))

        lat = _to_float(row["Latitude"])
        lon = _to_float(row["Longitude"])
        if lat is not None and not -90 <= lat <= 90:
            issues.append(
                _issue(store_id, "coordonnee_invalide", f"latitude hors plage : {lat}", "erreur")
            )
        if lon is not None and not -180 <= lon <= 180:
            issues.append(
                _issue(store_id, "coordonnee_invalide", f"longitude hors plage : {lon}", "erreur")
            )

        # Horaire vide : avertissement CS seulement (non bloquant, pas de règle explicite dans les sources prioritaires)
        if not any(normalize_cell(row[col]) for col in SCHEDULE_COLUMNS):
            issues.append(
                _issue(
                    store_id,
                    "horaire_manquant",
                    "aucun horaire renseigné (à vérifier avec le client si besoin)",
                    "avertissement",
                )
            )

        website = normalize_cell(row["Siteweb"])
        if website and not website.startswith(("http://", "https://")):
            issues.append(
                _issue(store_id, "site_web", "URL sans schéma http/https", "avertissement")
            )
        progress.update(index)
    progress.finish()

    columns = [STORE_ID_COLUMN, "type_probleme", "detail", "gravite"]
    if not issues:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(issues).drop_duplicates()
