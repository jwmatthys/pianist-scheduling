"""Local CSV/XLSX/XLS parsing and validation for Clinical Sites import.

This module is pure parsing: it never touches the database. The router combines it
with ClinicalRepository.replace_opportunities for the commit phase.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

import pandas as pd
from pydantic import BaseModel

from ..clinical_models import normalize_tag_name, tag_identity_key
from . import tabular
from .availability import normalize_day
from .time_parsing import parse_time_of_day

MAX_UPLOAD_BYTES = 25 * 1024 * 1024

FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "site_name": ("site name", "site", "clinical site", "placement site", "name"),
    "site_address": ("site address", "address", "location"),
    "contact_name": ("contact name", "contact", "site contact"),
    "contact_email": ("contact email", "email", "e mail"),
    "contact_phone": ("contact phone", "phone", "phone number"),
    "day_of_week": ("day of week", "day", "weekday"),
    "start_time": ("start time", "start", "begin time", "begin"),
    "end_time": ("end time", "end", "finish time", "finish"),
    "session_length": ("session length", "session length minutes", "session minutes", "length", "duration"),
    "travel_time": ("travel time", "travel time minutes", "travel minutes", "travel"),
    "capacity": ("capacity", "students", "slots", "seats"),
}
FIELD_LABELS = {
    "site_name": "Site Name",
    "site_address": "Site Address",
    "contact_name": "Contact Name",
    "contact_email": "Contact Email",
    "contact_phone": "Contact Phone",
    "day_of_week": "Day of Week",
    "start_time": "Start Time",
    "end_time": "End Time",
    "session_length": "Session Length",
    "travel_time": "Travel Time",
    "capacity": "Capacity",
}
REQUIRED_FIELDS = ("site_name", "day_of_week", "start_time", "end_time", "session_length")


class ClinicalSitesImportError(RuntimeError):
    def __init__(self, code: str, message: str, issues: list[str] | None = None):
        self.code = code
        self.issues = issues or []
        super().__init__(message)


class ClinicalSitesImportInspection(BaseModel):
    sheets: list[str]
    selected_sheet: str | None
    columns: list[str]
    sample_rows: list[dict[str, str]]
    suggested_mapping: dict[str, str | None]


class ClinicalSitesImportMapping(BaseModel):
    site_name: str | None = None
    site_address: str | None = None
    contact_name: str | None = None
    contact_email: str | None = None
    contact_phone: str | None = None
    day_of_week: str | None = None
    start_time: str | None = None
    end_time: str | None = None
    session_length: str | None = None
    travel_time: str | None = None
    capacity: str | None = None


class ClinicalSitesImportPreviewRow(BaseModel):
    row_number: int
    site_name: str
    day_of_week: str | None
    start_minute: int | None
    end_minute: int | None
    session_length_minutes: int | None
    travel_minutes: int | None
    capacity: int | None
    tags: list[str]
    errors: list[str]


class ClinicalSitesImportPreview(BaseModel):
    selected_sheet: str | None
    tag_columns: list[str]
    total_rows: int
    valid_rows: int
    invalid_rows: int
    rows: list[ClinicalSitesImportPreviewRow]
    new_tags: list[str]
    reused_tags: list[str]
    existing_opportunity_count: int
    existing_placement_count: int
    can_apply: bool


@dataclass
class ParsedSite:
    row_number: int
    site_name: str = ""
    site_address: str = ""
    contact_name: str = ""
    contact_email: str = ""
    contact_phone: str = ""
    day_of_week: str | None = None
    start_minute: int | None = None
    end_minute: int | None = None
    session_length_minutes: int | None = None
    travel_minutes: int | None = None
    capacity: int | None = None
    tags: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


@dataclass
class ParsedSites:
    selected_sheet: str | None
    tag_columns: list[str]
    rows: list[ParsedSite]


def _is_blank(value: object) -> bool:
    if value is None or (isinstance(value, str) and not value.strip()):
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _text(value: object) -> str:
    if _is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _frame(filename: str, content: bytes, sheet_name: str | None) -> tuple[pd.DataFrame, list[str], str | None]:
    if len(content) > MAX_UPLOAD_BYTES:
        raise ClinicalSitesImportError("FILE_TOO_LARGE", "Clinical Sites files must be 25 MB or smaller.")
    try:
        sheets = tabular.list_sheets(filename, content)
        selected = (sheet_name or sheets[0]) if sheets else None
        if selected is not None and selected not in sheets:
            raise ClinicalSitesImportError("UNKNOWN_SHEET", "Choose a worksheet that exists in the selected file.")
        frame = tabular.read_table(filename, content, selected)
    except ClinicalSitesImportError:
        raise
    except Exception as error:
        raise ClinicalSitesImportError("FILE_READ_FAILED", f"Could not read the selected file: {error}") from error
    frame = frame.copy()
    frame.columns = [str(column) for column in frame.columns]
    return frame, sheets, selected


def _header_key(column: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", column.casefold()).strip()


def inspect_file(filename: str, content: bytes, sheet_name: str | None) -> ClinicalSitesImportInspection:
    frame, sheets, selected = _frame(filename, content, sheet_name)
    columns = list(frame.columns)
    keys = {column: _header_key(column) for column in columns}
    suggested: dict[str, str | None] = {}
    taken: set[str] = set()
    for field_name, aliases in FIELD_ALIASES.items():
        match = next(
            (column for alias in aliases for column in columns if column not in taken and keys[column] == alias),
            None,
        )
        suggested[field_name] = match
        if match:
            taken.add(match)
    return ClinicalSitesImportInspection(
        sheets=sheets,
        selected_sheet=selected,
        columns=columns,
        sample_rows=frame.head(10).fillna("").astype(str).to_dict(orient="records"),
        suggested_mapping=suggested,
    )


_MINUTES = re.compile(r"^(\d+)(?:\s*(?:m|min|mins|minute|minutes))?$", re.IGNORECASE)


def _parse_minutes(value: object) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        if not math.isfinite(number) or number != int(number):
            return None
        return int(number)
    text = str(value).strip()
    if re.fullmatch(r"-?\d+(?:\.0+)?", text):
        return int(float(text))
    match = _MINUTES.match(text)
    return int(match.group(1)) if match else None


def validate_mapping(mapping: ClinicalSitesImportMapping, columns: list[str]) -> None:
    issues: list[str] = []
    used: dict[str, str] = {}
    for field_name in FIELD_ALIASES:
        column = getattr(mapping, field_name)
        if not column:
            if field_name in REQUIRED_FIELDS:
                issues.append(f"Map the {FIELD_LABELS[field_name]} column.")
            continue
        if column not in columns:
            issues.append(f"The column mapped to {FIELD_LABELS[field_name]} is not in the selected worksheet.")
        elif column in used:
            issues.append(f"{used[column]} and {FIELD_LABELS[field_name]} cannot use the same column.")
        else:
            used[column] = FIELD_LABELS[field_name]
    if issues:
        raise ClinicalSitesImportError("INVALID_MAPPING", issues[0], issues)


def parse_sites(
    filename: str, content: bytes, sheet_name: str | None, mapping: ClinicalSitesImportMapping
) -> ParsedSites:
    frame, _, selected = _frame(filename, content, sheet_name)
    columns = list(frame.columns)
    validate_mapping(mapping, columns)
    mapped = {name: getattr(mapping, name) for name in FIELD_ALIASES if getattr(mapping, name)}
    mapped_columns = set(mapped.values())
    tag_columns = [
        column
        for column in columns
        if column not in mapped_columns and not all(_is_blank(value) for value in frame[column])
    ]

    rows: list[ParsedSite] = []
    for index, (_, record) in enumerate(frame.iterrows()):
        if all(_is_blank(value) for value in record):
            continue
        site = ParsedSite(row_number=index + 2)
        cells = {name: record[column] for name, column in mapped.items()}

        site.site_name = _text(cells.get("site_name"))
        site.site_address = _text(cells.get("site_address"))
        site.contact_name = _text(cells.get("contact_name"))
        site.contact_email = _text(cells.get("contact_email"))
        site.contact_phone = _text(cells.get("contact_phone"))
        if not site.site_name:
            site.errors.append("Site Name is required.")

        day_cell = cells.get("day_of_week")
        if _is_blank(day_cell):
            site.errors.append("Day of Week is required.")
        else:
            site.day_of_week = normalize_day(day_cell)
            if site.day_of_week is None:
                site.errors.append(f"Day of Week '{_text(day_cell)}' is not a recognized weekday.")

        for name, attribute, label in (
            ("start_time", "start_minute", "Start Time"),
            ("end_time", "end_minute", "End Time"),
        ):
            cell = cells.get(name)
            if _is_blank(cell):
                site.errors.append(f"{label} is required.")
                continue
            minute = parse_time_of_day(cell)
            if minute is None:
                site.errors.append(f"{label} '{_text(cell)}' is not a recognized time.")
            else:
                setattr(site, attribute, minute)
        if site.start_minute is not None and site.end_minute is not None and site.end_minute <= site.start_minute:
            site.errors.append("End Time must be after Start Time.")

        length_cell = cells.get("session_length")
        if _is_blank(length_cell):
            site.errors.append("Session Length is required.")
        else:
            site.session_length_minutes = _parse_minutes(length_cell)
            if site.session_length_minutes is None:
                site.errors.append(f"Session Length '{_text(length_cell)}' is not a whole number of minutes.")
            elif site.session_length_minutes <= 0:
                site.errors.append("Session Length must be greater than 0.")

        travel_cell = cells.get("travel_time")
        if _is_blank(travel_cell):
            site.travel_minutes = 0
        else:
            site.travel_minutes = _parse_minutes(travel_cell)
            if site.travel_minutes is None or site.travel_minutes < 0:
                site.errors.append(f"Travel Time '{_text(travel_cell)}' is not a whole number of minutes.")
                site.travel_minutes = None

        capacity_cell = cells.get("capacity")
        if _is_blank(capacity_cell):
            site.capacity = 1
        else:
            site.capacity = _parse_minutes(capacity_cell)
            if site.capacity is None:
                site.errors.append(f"Capacity '{_text(capacity_cell)}' is not a whole number.")
            elif site.capacity <= 0:
                site.errors.append("Capacity must be greater than 0.")

        seen: set[str] = set()
        for column in tag_columns:
            name = normalize_tag_name(_text(record[column]))
            key = tag_identity_key(name)
            if key and key not in seen:
                seen.add(key)
                site.tags.append(name)
        rows.append(site)

    if not rows:
        raise ClinicalSitesImportError("NO_ROWS", "The selected worksheet has no Placement Opportunity rows.")
    return ParsedSites(selected_sheet=selected, tag_columns=tag_columns, rows=rows)


def build_preview(
    parsed: ParsedSites,
    existing_tag_keys: set[str],
    existing_opportunity_count: int,
    existing_placement_count: int,
) -> ClinicalSitesImportPreview:
    new_tags: dict[str, str] = {}
    reused: dict[str, str] = {}
    for site in parsed.rows:
        for name in site.tags:
            key = tag_identity_key(name)
            (reused if key in existing_tag_keys else new_tags).setdefault(key, name)
    invalid = sum(1 for site in parsed.rows if site.errors)
    return ClinicalSitesImportPreview(
        selected_sheet=parsed.selected_sheet,
        tag_columns=parsed.tag_columns,
        total_rows=len(parsed.rows),
        valid_rows=len(parsed.rows) - invalid,
        invalid_rows=invalid,
        rows=[
            ClinicalSitesImportPreviewRow(
                row_number=site.row_number,
                site_name=site.site_name,
                day_of_week=site.day_of_week,
                start_minute=site.start_minute,
                end_minute=site.end_minute,
                session_length_minutes=site.session_length_minutes,
                travel_minutes=site.travel_minutes,
                capacity=site.capacity,
                tags=site.tags,
                errors=site.errors,
            )
            for site in parsed.rows
        ],
        new_tags=sorted(new_tags.values()),
        reused_tags=sorted(reused.values()),
        existing_opportunity_count=existing_opportunity_count,
        existing_placement_count=existing_placement_count,
        can_apply=invalid == 0,
    )
