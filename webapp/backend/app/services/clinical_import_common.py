"""Pure helpers shared by Clinical module file imports (no database access)."""

from __future__ import annotations

import re

import pandas as pd

from . import tabular

MAX_UPLOAD_BYTES = 25 * 1024 * 1024


class ClinicalImportError(RuntimeError):
    def __init__(self, code: str, message: str, issues: list[str] | None = None):
        self.code = code
        self.issues = issues or []
        super().__init__(message)


def is_blank(value: object) -> bool:
    if value is None or (isinstance(value, str) and not value.strip()):
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def cell_text(value: object) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_frame(filename: str, content: bytes, sheet_name: str | None) -> tuple[pd.DataFrame, list[str], str | None]:
    if len(content) > MAX_UPLOAD_BYTES:
        raise ClinicalImportError("FILE_TOO_LARGE", "Clinical import files must be 25 MB or smaller.")
    try:
        sheets = tabular.list_sheets(filename, content)
        selected = (sheet_name or sheets[0]) if sheets else None
        if selected is not None and selected not in sheets:
            raise ClinicalImportError("UNKNOWN_SHEET", "Choose a worksheet that exists in the selected file.")
        frame = tabular.read_table(filename, content, selected)
    except ClinicalImportError:
        raise
    except Exception as error:
        raise ClinicalImportError("FILE_READ_FAILED", f"Could not read the selected file: {error}") from error
    frame = frame.copy()
    frame.columns = [str(column) for column in frame.columns]
    return frame, sheets, selected


def header_key(column: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", column.casefold()).strip()
