"""Local CSV/XLSX/XLS import that replaces all Jury Panels and clears Panel assignments."""

from __future__ import annotations

import math
import re
from datetime import date, datetime

import pandas as pd
from pydantic import BaseModel, ValidationError
from sqlalchemy.orm import Session

from .. import module_models
from ..jury_schemas import JuryPanelFields
from . import tabular
from .availability import parse_time_minutes
from .module_lifecycle import active_session_uuid, bump_jury_revision

MAX_UPLOAD_BYTES = 25 * 1024 * 1024
MAX_ERRORS_REPORTED = 25

FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "schedule_date": ("schedule date", "scheduling date", "jury date", "date"),
    "panel_name": ("panel name", "panel", "name"),
    "room": ("room", "location"),
    "earliest_start": ("earliest start", "earliest start time", "earliest"),
    "preferred_start": ("preferred start", "preferred start time", "preferred"),
    "jury_length": ("jury length", "jury length minutes", "length", "duration"),
    "break_needed": ("break needed", "break"),
    "break_every": ("break every x juries", "break every", "break frequency"),
    "break_length": ("break length", "break duration"),
    "meal_break": ("meal break needed", "meal break", "meal"),
    "meal_start": ("meal start", "meal start time"),
    "meal_end": ("meal end", "meal end time"),
}
REQUIRED_FIELDS = ("schedule_date", "panel_name", "earliest_start", "jury_length")
FIELD_LABELS = {
    "schedule_date": "Schedule date",
    "panel_name": "Panel name",
    "earliest_start": "Earliest start",
    "jury_length": "Jury length",
}


class JuryPanelImportError(RuntimeError):
    def __init__(self, code: str, message: str, issues: list[str] | None = None):
        self.code = code
        self.issues = issues or []
        super().__init__(message)


class JuryPanelImportInspection(BaseModel):
    sheets: list[str]
    selected_sheet: str | None
    columns: list[str]
    sample_rows: list[dict[str, str]]
    suggested_mapping: dict[str, str | None]


class JuryPanelImportMapping(BaseModel):
    schedule_date: str | None = None
    panel_name: str | None = None
    room: str | None = None
    earliest_start: str | None = None
    preferred_start: str | None = None
    jury_length: str | None = None
    break_needed: str | None = None
    break_every: str | None = None
    break_length: str | None = None
    meal_break: str | None = None
    meal_start: str | None = None
    meal_end: str | None = None


class JuryPanelImportResult(BaseModel):
    panels_removed: int
    assignments_cleared: int
    panels_created: int


def _is_blank(value: object) -> bool:
    if value is None or (isinstance(value, str) and not value.strip()):
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _frame(filename: str, content: bytes, sheet_name: str | None) -> tuple[pd.DataFrame, list[str], str | None]:
    if len(content) > MAX_UPLOAD_BYTES:
        raise JuryPanelImportError("FILE_TOO_LARGE", "Jury Panel files must be 25 MB or smaller.")
    try:
        sheets = tabular.list_sheets(filename, content)
        selected = (sheet_name or sheets[0]) if sheets else None
        if selected is not None and selected not in sheets:
            raise JuryPanelImportError("UNKNOWN_SHEET", "Choose a worksheet that exists in the selected file.")
        frame = tabular.read_table(filename, content, selected)
    except JuryPanelImportError:
        raise
    except Exception as error:
        raise JuryPanelImportError("FILE_READ_FAILED", f"Could not read the selected file: {error}") from error
    frame = frame.copy()
    frame.columns = [str(column) for column in frame.columns]
    return frame, sheets, selected


def _suggest(columns: list[str], aliases: tuple[str, ...]) -> str | None:
    normalized = {column: re.sub(r"[^a-z0-9]+", " ", column.casefold()).strip() for column in columns}
    for alias in aliases:
        for column, header in normalized.items():
            if header == alias:
                return column
    return None


def inspect_file(filename: str, content: bytes, sheet_name: str | None) -> JuryPanelImportInspection:
    frame, sheets, selected = _frame(filename, content, sheet_name)
    columns = list(frame.columns)
    suggested: dict[str, str | None] = {}
    taken: set[str] = set()
    for field, aliases in FIELD_ALIASES.items():
        match = _suggest([column for column in columns if column not in taken], aliases)
        suggested[field] = match
        if match:
            taken.add(match)
    return JuryPanelImportInspection(
        sheets=sheets,
        selected_sheet=selected,
        columns=columns,
        sample_rows=frame.head(10).fillna("").astype(str).to_dict(orient="records"),
        suggested_mapping=suggested,
    )


def _parse_date(value: object) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for pattern in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%B %d, %Y", "%b %d, %Y"):
        try:
            return datetime.strptime(text.split(" 00:00:00")[0], pattern).date()
        except ValueError:
            continue
    raise ValueError(f"'{text}' is not a recognized date")


def _parse_int(value: object) -> int:
    if isinstance(value, bool):
        raise ValueError(f"'{value}' is not a whole number")
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        try:
            number = float(str(value).strip())
        except ValueError as error:
            raise ValueError(f"'{value}' is not a whole number") from error
    if not math.isfinite(number) or number != int(number):
        raise ValueError(f"'{value}' is not a whole number")
    return int(number)


_TRUE = {"true", "yes", "y", "1", "x"}
_FALSE = {"false", "no", "n", "0"}


def _parse_bool(value: object) -> bool:
    if _is_blank(value):
        return False
    if isinstance(value, bool):
        return value
    text = str(value).strip().casefold()
    if text in _TRUE or text in {"1.0"}:
        return True
    if text in _FALSE or text in {"0.0"}:
        return False
    raise ValueError(f"'{value}' is not Yes/No")


def _parse_time(value: object, *, end_of_day: bool = False) -> int:
    minute = parse_time_minutes(value, allow_end_of_day=end_of_day)
    if minute is None:
        raise ValueError(f"'{value}' is not a valid time")
    return minute


def _validate_mapping(mapping: JuryPanelImportMapping, columns: list[str]) -> None:
    issues = []
    for field in REQUIRED_FIELDS:
        if not getattr(mapping, field):
            issues.append(f"Map the {FIELD_LABELS[field]} column.")
    for field in FIELD_ALIASES:
        column = getattr(mapping, field)
        if column and column not in columns:
            issues.append(f"The column mapped to {field.replace('_', ' ')} is not in the selected worksheet.")
    if issues:
        raise JuryPanelImportError("INVALID_MAPPING", issues[0], issues)


def _parse_rows(frame: pd.DataFrame, mapping: JuryPanelImportMapping) -> list[JuryPanelFields]:
    panels: list[JuryPanelFields] = []
    issues: list[str] = []
    seen_names: dict[str, int] = {}
    mapped = {field: getattr(mapping, field) for field in FIELD_ALIASES if getattr(mapping, field)}

    for index, (_, row) in enumerate(frame.iterrows()):
        row_number = index + 2
        cells = {field: row.get(column) for field, column in mapped.items()}
        if all(_is_blank(value) for value in cells.values()):
            continue

        def cell(field: str):
            value = cells.get(field)
            return None if _is_blank(value) else value

        try:
            for field in REQUIRED_FIELDS:
                if cell(field) is None:
                    raise ValueError(f"{FIELD_LABELS[field]} is missing")
            break_needed = _parse_bool(cells.get("break_needed"))
            meal_break = _parse_bool(cells.get("meal_break"))
            earliest = _parse_time(cell("earliest_start"))
            fields = JuryPanelFields(
                panel_name=" ".join(str(cell("panel_name")).split()),
                room=" ".join(str(cell("room") or "").split()),
                jury_date=_parse_date(cell("schedule_date")),
                earliest_start_minute=earliest,
                preferred_start_minute=_parse_time(cell("preferred_start")) if cell("preferred_start") is not None else None,
                jury_length_minutes=_parse_int(cell("jury_length")),
                break_needed=break_needed,
                break_every_x_juries=_parse_int(cell("break_every")) if break_needed and cell("break_every") is not None else None,
                break_length_minutes=_parse_int(cell("break_length")) if break_needed and cell("break_length") is not None else None,
                meal_break=meal_break,
                meal_start_minute=_parse_time(cell("meal_start")) if meal_break and cell("meal_start") is not None else None,
                meal_end_minute=_parse_time(cell("meal_end"), end_of_day=True) if meal_break and cell("meal_end") is not None else None,
            )
        except ValidationError as error:
            issues.append(f"Row {row_number}: {'; '.join(item['msg'].removeprefix('Value error, ') for item in error.errors())}")
            continue
        except ValueError as error:
            issues.append(f"Row {row_number}: {error}.")
            continue

        key = fields.panel_name.casefold()
        if key in seen_names:
            issues.append(f"Row {row_number}: Panel name '{fields.panel_name}' already appears in row {seen_names[key]}.")
            continue
        seen_names[key] = row_number
        panels.append(fields)

    if issues:
        shown = issues[:MAX_ERRORS_REPORTED]
        if len(issues) > len(shown):
            shown.append(f"{len(issues) - len(shown)} more problem(s) not shown.")
        raise JuryPanelImportError("INVALID_ROWS", "The file contains rows that cannot be imported.", shown)
    if not panels:
        raise JuryPanelImportError("NO_PANELS", "The selected worksheet does not contain any Jury Panel rows.")
    return panels


def apply_import(
    db: Session,
    filename: str,
    content: bytes,
    sheet_name: str | None,
    mapping: JuryPanelImportMapping,
) -> JuryPanelImportResult:
    frame, _, _ = _frame(filename, content, sheet_name)
    _validate_mapping(mapping, list(frame.columns))
    panels = _parse_rows(frame, mapping)

    session_uuid = active_session_uuid(db)
    entries = db.query(module_models.JuryLessonEntry).filter(
        module_models.JuryLessonEntry.session_uuid == session_uuid,
        module_models.JuryLessonEntry.panel_uuid.is_not(None),
    ).all()
    for entry in entries:
        entry.panel_uuid = None
    db.flush()

    existing = db.query(module_models.JuryPanel).filter(
        module_models.JuryPanel.session_uuid == session_uuid
    ).all()
    for panel in existing:
        panel_date = db.get(module_models.JuryPanelDate, panel.panel_uuid)
        if panel_date is not None:
            db.delete(panel_date)
        db.delete(panel)
    db.flush()

    for fields in panels:
        values = fields.model_dump()
        jury_date = values.pop("jury_date")
        row = module_models.JuryPanel(session_uuid=session_uuid, **values)
        db.add(row)
        db.flush()
        db.add(module_models.JuryPanelDate(
            panel_uuid=row.panel_uuid,
            session_uuid=session_uuid,
            jury_date=jury_date,
        ))
    bump_jury_revision(db)
    db.flush()
    return JuryPanelImportResult(
        panels_removed=len(existing),
        assignments_cleared=len(entries),
        panels_created=len(panels),
    )
