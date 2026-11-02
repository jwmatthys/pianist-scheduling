"""Local CSV/XLSX/XLS parsing and validation for the Tutor Scheduling Tutor Roster import.

Pure parsing: it never touches the database. The router combines it with
TutorRepository.replace_tutors for the commit phase (Phase B re-parses the file;
preview data is never trusted).

Weekday cells and spreadsheet reading reuse the pure, module-neutral helpers that the Clinical
imports (``weekly_availability_cells`` and ``import_common``).
Mapped columns become tutor fields and weekday availability. Tags come only from the single
column explicitly mapped to Tags (semicolon-separated); unmapped columns never become tags.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from pydantic import BaseModel, Field

from ..tutor_models import (
    DEFAULT_TUTOR_CAPACITY,
    DEFAULT_LESSON_LENGTH_MINUTES,
    tutor_tag_identity_key,
    normalize_tutor_tag_name,
)
from .availability import WEEKDAYS, AvailabilityWindow
from .weekly_availability_cells import (
    mapped_day_columns,
    parse_week,
    suggest_day_columns,
)
from .import_common import (
    MAX_UPLOAD_BYTES,
    ImportFileError,
    cell_text as _text,
    header_key as _header_key,
    is_blank as _is_blank,
    read_frame as _frame,
    split_tag_cell,
    suggest_tags_column,
)
from .time_parsing import format_time_of_day

__all__ = ["MAX_UPLOAD_BYTES", "TutorTutorImportError"]

TutorTutorImportError = ImportFileError

FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "tutor_name": ("tutor name", "tutor", "teacher name", "teacher", "name"),
    "tutor_email": ("tutor email", "teacher email", "email", "e mail"),
    "room": ("room", "studio", "room number"),
    "lesson_length": ("lesson length", "lesson length minutes", "tutor minutes", "length", "duration"),
    "capacity": ("capacity", "students", "max students", "slots"),
}
FIELD_LABELS = {
    "tutor_name": "Tutor Name",
    "tutor_email": "Tutor Email",
    "room": "Room",
    "lesson_length": "Lesson Length",
    "capacity": "Capacity",
}
REQUIRED_FIELDS = ("tutor_name",)
_DAY_ABBREVIATIONS = {day: day[:3] for day in WEEKDAYS}


class TutorTutorImportInspection(BaseModel):
    sheets: list[str]
    selected_sheet: str | None
    columns: list[str]
    sample_rows: list[dict[str, str]]
    suggested_mapping: "TutorTutorImportMapping"


class TutorTutorImportMapping(BaseModel):
    tutor_name: str | None = None
    tutor_email: str | None = None
    room: str | None = None
    lesson_length: str | None = None
    capacity: str | None = None
    days: dict[str, str | None] = Field(default_factory=dict)
    tags: str | None = None


TutorTutorImportInspection.model_rebuild()


class TutorTutorImportPreviewRow(BaseModel):
    row_number: int
    tutor_name: str
    tutor_email: str
    room: str
    lesson_length_minutes: int | None
    capacity: int | None
    tags: list[str]
    availability_window_count: int
    availability: list[str]
    warnings: list[str]
    errors: list[str]


class TutorTutorImportPreview(BaseModel):
    selected_sheet: str | None
    tag_columns: list[str]
    total_rows: int
    valid_rows: int
    invalid_rows: int
    rows: list[TutorTutorImportPreviewRow]
    new_tags: list[str]
    reused_tags: list[str]
    existing_tutor_count: int
    existing_placement_count: int
    can_apply: bool


@dataclass
class ParsedTutor:
    row_number: int
    tutor_name: str = ""
    tutor_email: str = ""
    room: str = ""
    lesson_length_minutes: int | None = None
    capacity: int | None = None
    tags: list[str] = field(default_factory=list)
    windows: list[AvailabilityWindow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


@dataclass
class ParsedTutors:
    selected_sheet: str | None
    tag_columns: list[str]
    rows: list[ParsedTutor]


def inspect_file(filename: str, content: bytes, sheet_name: str | None) -> TutorTutorImportInspection:
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
    return TutorTutorImportInspection(
        sheets=sheets,
        selected_sheet=selected,
        columns=columns,
        sample_rows=frame.head(10).fillna("").astype(str).to_dict(orient="records"),
        suggested_mapping=TutorTutorImportMapping(**suggested, days=suggest_day_columns(columns, taken), tags=suggest_tags_column(columns, taken | {c for c in suggested.values() if c})),
    )


def _parse_whole_number(value: object) -> int | None:
    if isinstance(value, bool):
        return None
    text = _text(value)
    if re.fullmatch(r"-?\d+(?:\.0+)?", text):
        return int(float(text))
    match = re.fullmatch(r"(\d+)\s*(?:m|min|mins|minute|minutes)", text, re.IGNORECASE)
    return int(match.group(1)) if match else None


def validate_mapping(mapping: TutorTutorImportMapping, columns: list[str]) -> None:
    issues: list[str] = []
    used: dict[str, str] = {}

    def claim(column: str | None, label: str, required: bool = False) -> None:
        if not column:
            if required:
                issues.append(f"Map the {label} column.")
            return
        if column not in columns:
            issues.append(f"The column mapped to {label} is not in the selected worksheet.")
        elif column in used:
            issues.append(f"{used[column]} and {label} cannot use the same column.")
        else:
            used[column] = label

    for field_name in FIELD_ALIASES:
        claim(getattr(mapping, field_name), FIELD_LABELS[field_name], field_name in REQUIRED_FIELDS)
    if any(day not in WEEKDAYS for day in mapping.days):
        issues.append("Availability columns must be mapped to weekdays.")
    for day in WEEKDAYS:
        claim(mapping.days.get(day), day)
    claim(mapping.tags, "Tags")
    if issues:
        raise TutorTutorImportError("INVALID_MAPPING", issues[0], issues)


def parse_tutors(
    filename: str, content: bytes, sheet_name: str | None, mapping: TutorTutorImportMapping
) -> ParsedTutors:
    frame, _, selected = _frame(filename, content, sheet_name)
    columns = list(frame.columns)
    validate_mapping(mapping, columns)
    mapped = {name: getattr(mapping, name) for name in FIELD_ALIASES if getattr(mapping, name)}
    day_columns = mapped_day_columns(mapping.days)
    tag_columns = [mapping.tags] if mapping.tags else []

    rows: list[ParsedTutor] = []
    seen_names: set[str] = set()
    for index, (_, record) in enumerate(frame.iterrows()):
        if all(_is_blank(value) for value in record):
            continue
        row = ParsedTutor(row_number=index + 2)
        cells = {name: record[column] for name, column in mapped.items()}

        row.tutor_name = " ".join(_text(cells.get("tutor_name")).split())
        row.tutor_email = _text(cells.get("tutor_email"))
        row.room = _text(cells.get("room"))
        if not row.tutor_name:
            row.errors.append("Tutor Name is required.")
        elif row.tutor_name.casefold() in seen_names:
            row.warnings.append(f"Tutor '{row.tutor_name}' appears more than once in the file.")
        seen_names.add(row.tutor_name.casefold())

        for name, attribute, label, default in (
            ("lesson_length", "lesson_length_minutes", "Lesson Length", DEFAULT_LESSON_LENGTH_MINUTES),
            ("capacity", "capacity", "Capacity", DEFAULT_TUTOR_CAPACITY),
        ):
            cell = cells.get(name)
            if _is_blank(cell):
                setattr(row, attribute, default)
                continue
            number = _parse_whole_number(cell)
            if number is None:
                row.errors.append(f"{label} '{_text(cell)}' is not a whole number.")
            elif number <= 0:
                row.errors.append(f"{label} must be greater than 0.")
            else:
                setattr(row, attribute, number)

        week = parse_week(
            {column: record[column] for column in day_columns.values()}, day_columns, row.row_number
        )
        row.windows = week.windows
        row.warnings.extend(week.warnings)
        row.errors.extend(week.errors)

        seen: set[str] = set()
        for column in tag_columns:
            for part in split_tag_cell(record[column]):
                name = normalize_tutor_tag_name(part)
                key = tutor_tag_identity_key(name)
                if key and key not in seen:
                    seen.add(key)
                    row.tags.append(name)
        rows.append(row)

    if not rows:
        raise TutorTutorImportError("NO_ROWS", "The selected worksheet has no tutor rows.")
    return ParsedTutors(selected_sheet=selected, tag_columns=tag_columns, rows=rows)


def _describe_window(window: AvailabilityWindow) -> str:
    prefix = "" if window.status == "Available" else f"{window.status} "
    return (
        f"{prefix}{_DAY_ABBREVIATIONS[window.day]} {format_time_of_day(window.start_minute)}-"
        f"{format_time_of_day(window.end_minute)}"
    )


def build_preview(
    parsed: ParsedTutors,
    existing_tag_keys: set[str],
    existing_tutor_count: int,
    existing_placement_count: int,
) -> TutorTutorImportPreview:
    new_tags: dict[str, str] = {}
    reused: dict[str, str] = {}
    for row in parsed.rows:
        for name in row.tags:
            key = tutor_tag_identity_key(name)
            (reused if key in existing_tag_keys else new_tags).setdefault(key, name)
    invalid = sum(1 for row in parsed.rows if row.errors)
    return TutorTutorImportPreview(
        selected_sheet=parsed.selected_sheet,
        tag_columns=parsed.tag_columns,
        total_rows=len(parsed.rows),
        valid_rows=len(parsed.rows) - invalid,
        invalid_rows=invalid,
        rows=[
            TutorTutorImportPreviewRow(
                row_number=row.row_number,
                tutor_name=row.tutor_name,
                tutor_email=row.tutor_email,
                room=row.room,
                lesson_length_minutes=row.lesson_length_minutes,
                capacity=row.capacity,
                tags=row.tags,
                availability_window_count=len(row.windows),
                availability=[_describe_window(window) for window in row.windows],
                warnings=row.warnings,
                errors=row.errors,
            )
            for row in parsed.rows
        ],
        new_tags=sorted(new_tags.values()),
        reused_tags=sorted(reused.values()),
        existing_tutor_count=existing_tutor_count,
        existing_placement_count=existing_placement_count,
        can_apply=invalid == 0,
    )
