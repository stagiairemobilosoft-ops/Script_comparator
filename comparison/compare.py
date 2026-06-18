from __future__ import annotations
from dataclasses import dataclass, field
import pandas as pd

from constants.constant import (
    STORE_ID_COLUMN,
    SYNC_FIELDS,
    STATUS_OPEN,
    STATUS_CLOSED,
)
from .utils import normalize_cell
from .progress import *

@dataclass
class ComparisonResult:
    modified: pd.DataFrame = field(default_factory=pd.DataFrame)
    added: pd.DataFrame = field(default_factory=pd.DataFrame)
    removed: pd.DataFrame = field(default_factory=pd.DataFrame)
    unchanged: pd.DataFrame = field(default_factory=pd.DataFrame)
    changed_fields: pd.DataFrame = field(default_factory=pd.DataFrame)
    reference_total: int = 0
    incoming_total: int = 0

    @property
    def unchanged_count(self) -> int:
        return len(self.unchanged)

def _index_by_store(df: pd.DataFrame) -> pd.DataFrame:
    indexed = df.copy()
    indexed[STORE_ID_COLUMN] = indexed[STORE_ID_COLUMN].map(normalize_cell)
    return indexed.set_index(STORE_ID_COLUMN, drop=False)

def _rows_to_dataframe(rows: list[pd.Series], columns: pd.Index) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows).reset_index(drop=True)

def compare_store_files(
    reference: pd.DataFrame,
    incoming: pd.DataFrame,
    *,
    show_progress: bool = True,
) -> ComparisonResult:
    ref = _index_by_store(reference)
    inc = _index_by_store(incoming)

    ref_ids = set(ref.index)
    inc_ids = set(inc.index)
    added_ids = inc_ids - ref_ids
    removed_ids = ref_ids - inc_ids
    common_ids = ref_ids & inc_ids

    changed_rows: list[pd.Series] = []
    unchanged_rows: list[pd.Series] = []
    field_changes: list[dict[str, str]] = []

    common_list = sorted(common_ids)
    progress = ProgressReporter("Comparaison", len(common_list), enabled=show_progress)
    for index, store_id in enumerate(common_list, start=1):
        ref_row = ref.loc[store_id]
        inc_row = inc.loc[store_id]
        changed = [
            name
            for name in SYNC_FIELDS
            if normalize_cell(ref_row[name]) != normalize_cell(inc_row[name])
        ]
        if changed:
            changed_rows.append(inc.loc[store_id])
            for field_name in changed:
                field_changes.append(
                    {
                        STORE_ID_COLUMN: store_id,
                        "champ": field_name,
                        "ancienne_valeur": normalize_cell(ref_row[field_name]),
                        "nouvelle_valeur": normalize_cell(inc_row[field_name]),
                    }
                )
        else:
            unchanged_rows.append(inc.loc[store_id])
        progress.update(index)
    progress.finish(
        f"{len(changed_rows)} modification(s), {len(unchanged_rows)} inchange(s)"
    )

    added = incoming[incoming[STORE_ID_COLUMN].isin(added_ids)].copy()
    added["status"] = STATUS_OPEN

    removed = reference[reference[STORE_ID_COLUMN].isin(removed_ids)].copy()
    removed["status"] = STATUS_CLOSED

    return ComparisonResult(
        modified=_rows_to_dataframe(changed_rows, incoming.columns),
        added=added,
        removed=removed,
        unchanged=_rows_to_dataframe(unchanged_rows, incoming.columns),
        changed_fields=pd.DataFrame(field_changes),
        reference_total=len(ref_ids),
        incoming_total=len(inc_ids),
    )
