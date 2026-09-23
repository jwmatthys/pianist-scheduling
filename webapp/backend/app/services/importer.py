"""CSV/XLSX import: preview + column-mapping commit.

Uploaded files are cached in-memory (keyed by an upload token) between the
preview step and the commit step so the user doesn't have to re-upload after
choosing column mappings. This is fine for a single-process local MVP; a
multi-worker deployment would swap this for a short-lived DB/blob cache.
"""

from __future__ import annotations

import io
import uuid

import pandas as pd

from ..models import DAYS_ORDER

# Fields the app understands; the user maps their spreadsheet's columns to these.
TARGET_FIELDS = [
    "teacher",
    "student",
    "day",
    "start_time",
    "end_time",
    "room",
    "instrument",
    "required_pianist_name",
    "need_pianist",
]

DAY_ALIASES = {
    "M": "Monday", "MON": "Monday", "MONDAY": "Monday",
    "T": "Tuesday", "TU": "Tuesday", "TUE": "Tuesday", "TUESDAY": "Tuesday",
    "W": "Wednesday", "WED": "Wednesday", "WEDNESDAY": "Wednesday",
    "R": "Thursday", "TH": "Thursday", "THU": "Thursday", "THURSDAY": "Thursday",
    "F": "Friday", "FRI": "Friday", "FRIDAY": "Friday",
    "S": "Saturday", "SAT": "Saturday", "SATURDAY": "Saturday",
    "U": "Sunday", "SUN": "Sunday", "SUNDAY": "Sunday",
}

_UPLOAD_CACHE: dict[str, pd.DataFrame] = {}


def _read_any(filename: str, content: bytes) -> pd.DataFrame:
    if filename.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(content))
    return pd.read_csv(io.BytesIO(content))


def stage_upload(filename: str, content: bytes) -> tuple[str, list[str], list[dict]]:
    df = _read_any(filename, content)
    df = df.dropna(how="all")
    token = uuid.uuid4().hex
    _UPLOAD_CACHE[token] = df
    preview_rows = df.head(20).fillna("").astype(str).to_dict(orient="records")
    return token, [str(c) for c in df.columns], preview_rows


def normalize_day(value) -> str:
    if pd.isna(value):
        return ""
    key = str(value).strip().upper()
    return DAY_ALIASES.get(key, str(value).strip().title())


def parse_time_to_minutes(value) -> int | None:
    if pd.isna(value) or value == "":
        return None
    if hasattr(value, "hour"):
        return value.hour * 60 + value.minute
    parsed = pd.to_datetime(str(value), errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.hour * 60 + parsed.minute


def commit_upload(token: str, mapping: dict[str, str | None]) -> tuple[list[dict], list[str]]:
    """Applies a field->column mapping to the staged upload.

    Returns (list of lesson dicts ready for Lesson creation, list of warnings).
    """
    df = _UPLOAD_CACHE.get(token)
    if df is None:
        raise KeyError("Upload token not found or expired; please re-upload the file.")

    warnings: list[str] = []
    lessons: list[dict] = []

    for idx, row in df.iterrows():
        def get(field):
            col = mapping.get(field)
            if not col or col not in df.columns:
                return None
            return row[col]

        day_raw = get("day")
        start_raw = get("start_time")
        if day_raw is None or pd.isna(day_raw) or start_raw is None or pd.isna(start_raw):
            warnings.append(f"Row {idx + 2}: missing day/start time -- skipped")
            continue

        day = normalize_day(day_raw)
        if day not in DAYS_ORDER:
            warnings.append(f"Row {idx + 2}: unrecognized day '{day_raw}' -- skipped")
            continue

        start_min = parse_time_to_minutes(start_raw)
        if start_min is None:
            warnings.append(f"Row {idx + 2}: unparseable start time '{start_raw}' -- skipped")
            continue

        end_raw = get("end_time")
        end_min = parse_time_to_minutes(end_raw)
        if end_min is None:
            end_min = start_min + 50  # default lesson length

        need_raw = get("need_pianist")
        need_pianist = True
        if need_raw is not None and not pd.isna(need_raw):
            need_pianist = str(need_raw).strip().upper() in ("1", "1.0", "TRUE", "YES", "Y")

        def text(field):
            v = get(field)
            if v is None or pd.isna(v):
                return ""
            return str(v).strip()

        lessons.append({
            "teacher": text("teacher"),
            "student": text("student"),
            "day": day,
            "start_minute": start_min,
            "end_minute": end_min,
            "room": text("room"),
            "instrument": text("instrument"),
            "required_pianist_name": text("required_pianist_name"),
            "need_pianist": need_pianist,
        })

    return lessons, warnings
