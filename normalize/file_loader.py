from __future__ import annotations
from pathlib import Path
import pandas as pd

from constants.constant import (
    STORE_ID_COLUMN,
    FILE_COLUMNS,
    OPTIONAL_COLUMNS,
    STATUS_OPEN,
    CLIENT_COLUMN_RENAME,
)

class LoadError(Exception):
    pass

def _is_client_format(columns: list[str]) -> bool:
    if STORE_ID_COLUMN not in columns:
        return False
    if "Ligne_adresse1" in columns or "Ligne_Adresse_1" in columns:
        return True
    return "Nom_etablissement" in columns and ("Pays" in columns or "Tag Pays" in columns)

def _find_data_sheet(path: Path, sheet_names: list[str]) -> str:
    for name in sheet_names:
        if "all_in" in name.lower():
            return name

    for name in sheet_names:
        header = pd.read_excel(path, sheet_name=name, nrows=0)
        columns = [str(col).strip() for col in header.columns]
        if _is_client_format(columns):
            row_count = len(pd.read_excel(path, sheet_name=name, usecols=[0]))
            if row_count > 0:
                return name

    raise LoadError(
        "Aucune feuille de magasins trouvée (all_in ou format Nickel Gmaps). "
        f"Feuilles disponibles : {', '.join(sheet_names)}"
    )


def _normalize_raw_dataframe(df: pd.DataFrame, statut: str = "") -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    if "Statut" in df.columns and "status" not in df.columns:
        df = df.rename(columns={"Statut": "status"})

    if _is_client_format(list(df.columns)):
        df = df.rename(columns=CLIENT_COLUMN_RENAME)
        if not statut:
            statut = STATUS_OPEN

    if "status" not in df.columns:
        df["status"] = statut

    missing = [col for col in FILE_COLUMNS if col not in df.columns]
    if missing:
        raise LoadError(f"Colonnes manquantes : {', '.join(missing)}")

    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[[*FILE_COLUMNS, *OPTIONAL_COLUMNS]].copy()
    ids = df[STORE_ID_COLUMN].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df[STORE_ID_COLUMN] = ids
    return df

def load_store_file(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise LoadError(f"Fichier introuvable : {path}")

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        workbook = pd.ExcelFile(path)
        first_sheet = pd.read_excel(path, sheet_name=workbook.sheet_names[0], nrows=0)
        columns = [str(col).strip() for col in first_sheet.columns]

        if _is_client_format(columns) or any(
            _is_client_format([str(c).strip() for c in pd.read_excel(path, sheet_name=n, nrows=0).columns])
            for n in workbook.sheet_names
        ):
            sheet_name = _find_data_sheet(path, workbook.sheet_names)
            df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
            return _normalize_raw_dataframe(df, statut=STATUS_OPEN)

        if set(FILE_COLUMNS).issubset(columns):
            df = pd.read_excel(path, dtype=str)
            return _normalize_raw_dataframe(df)

        raise LoadError(
            "Format de fichier non reconnu. Attendu : export Nickel Gmaps (feuille all_in)."
        )

    if suffix == ".csv":
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
        return _normalize_raw_dataframe(df)

    raise LoadError(f"Format non supporté : {suffix}. Utilisez .xlsx, .xls ou .csv")
