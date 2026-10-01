"""Local CSV/XLSX/XLS tabular file parsing shared by module importers."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}


def _extension(filename: str) -> str:
    extension = Path(filename).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ValueError("Choose a .csv, .xlsx, or .xls file.")
    return extension


def list_sheets(filename: str, content: bytes) -> list[str]:
    extension = _extension(filename)
    if extension == ".csv":
        return []
    with pd.ExcelFile(BytesIO(content)) as workbook:
        return list(workbook.sheet_names)


def read_table(filename: str, content: bytes, sheet_name: str | int | None = None) -> pd.DataFrame:
    extension = _extension(filename)
    if extension == ".csv":
        if sheet_name not in (None, ""):
            raise ValueError("CSV files do not contain worksheets.")
        return pd.read_csv(BytesIO(content))
    return pd.read_excel(BytesIO(content), sheet_name=sheet_name or 0)