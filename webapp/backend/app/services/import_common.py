"""Pure helpers shared by module file imports (Clinical and Lesson); no database access."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from . import tabular

if TYPE_CHECKING:
    import pandas as pd

MAX_UPLOAD_BYTES = 25 * 1024 * 1024


class ImportFileError(RuntimeError):
    def __init__(self, code: str, message: str, issues: list[str] | None = None):
        self.code = code
        self.issues = issues or []
        super().__init__(message)


def is_blank(value: object) -> bool:
    if value is None or (isinstance(value, str) and not value.strip()):
        return True
    import pandas as pd

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


TAG_HEADER_ALIASES = ("tags", "tag", "tag names", "tag list")


def split_tag_cell(value: object) -> list[str]:
    """Split a tag cell on semicolons, trimming whitespace and dropping empty values."""
    return [part.strip() for part in cell_text(value).split(";") if part.strip()]


def suggest_tags_column(columns: list[str], taken: set[str]) -> str | None:
    """Suggest an unmapped column whose header contains the word "tag"; exact "Tags" headers win."""
    free = [column for column in columns if column not in taken]
    exact = next((c for c in free if header_key(c) in TAG_HEADER_ALIASES), None)
    return exact or next((c for c in free if "tag" in c.casefold()), None)


def read_frame(filename: str, content: bytes, sheet_name: str | None) -> tuple[pd.DataFrame, list[str], str | None]:
    if len(content) > MAX_UPLOAD_BYTES:
        raise ImportFileError("FILE_TOO_LARGE", "Import files must be 25 MB or smaller.")
    try:
        sheets = tabular.list_sheets(filename, content)
        selected = (sheet_name or sheets[0]) if sheets else None
        if selected is not None and selected not in sheets:
            raise ImportFileError("UNKNOWN_SHEET", "Choose a worksheet that exists in the selected file.")
        frame = tabular.read_table(filename, content, selected)
    except ImportFileError:
        raise
    except Exception as error:
        raise ImportFileError("FILE_READ_FAILED", f"Could not read the selected file: {error}") from error
    frame = frame.copy()
    frame.columns = [str(column) for column in frame.columns]
    return frame, sheets, selected


def header_key(column: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", column.casefold()).strip()
