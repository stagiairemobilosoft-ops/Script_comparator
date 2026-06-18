from __future__ import annotations
import sys
import pandas as pd

def print_stage(step: int, total: int, message: str, *, enabled: bool = True) -> None:
    if enabled:
        print(f"[{step}/{total}] {message}", file=sys.stderr, flush=True)


def print_step_done(message: str, *, enabled: bool = True) -> None:
    if enabled:
        print(f"      {message}", file=sys.stderr, flush=True)

def normalize_cell(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none"}:
        return ""
    if text.endswith(".0") and text.replace(".", "", 1).isdigit():
        return text[:-2]
    return text
