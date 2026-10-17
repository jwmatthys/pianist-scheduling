"""Local CSV/XLSX/XLS parsing and validation for the Lesson Scheduling Student Roster import.

Pure parsing: it never touches the database. The router combines it with
LessonRepository.replace_students for the commit phase (Phase B re-parses the file; preview
data is never trusted).

One import carries the roster and weekly availability together; there is no separate
availability import. Spreadsheet reading and weekday-cell parsing reuse the pure,
module-neutral helpers that the Clinical imports already proved. Unlike the Clinical
roster import (which merges by Student ID), this import is a full replacement.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pydantic import BaseModel, Field

from .availability import WEEKDAYS, AvailabilityWindow
from .clinical_availability_importer import (
    mapped_day_columns,
    parse_week,
    suggest_day_columns,
)
from .clinical_import_common import (
    MAX_UPLOAD_BYTES,
    ClinicalImportError,
    cell_text as _text,
    header_key as _header_key,
    is_blank as _is_blank,
    read_frame as _frame,
)
from .time_parsing import format_time_of_day

__all__ = ["MAX_UPLOAD_BYTES", "LessonStudentImportError"]

LessonStudentImportError = ClinicalImportError

FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "student_id": ("student id", "studentid", "id", "student number", "student", "emplid"),
    "name": ("name", "student name", "full name"),
    "email": ("email", "student email", "e mail", "email address"),
}
FIELD_LABELS = {"student_id": "Student ID", "name": "Name", "email": "Email"}
REQUIRED_FIELDS = ("student_id", "name")
_DAY_ABBREVIATIONS = {day: day[:3] for day in WEEKDAYS}


class LessonStudentImportMapping(BaseModel):
    student_id: str | None = None
    name: str | None = None
    email: str | None = None
    days: dict[str, str | None] = Field(default_factory=dict)


class LessonStudentImportInspection(BaseModel):
    sheets: list[str]
    selected_sheet: str | None
    columns: list[str]
    sample_rows: list[dict[str, str]]
    suggested_mapping: LessonStudentImportMapping


class LessonStudentImportPreviewRow(BaseModel):
    row_number: int
    student_id: str
    name: str
    email: str
    availability_window_count: int
    availability: list[str]
    warnings: list[str]
    errors: list[str]


class LessonStudentImportPreview(BaseModel):
    selected_sheet: str | None
    total_rows: int
    valid_rows: int
    invalid_rows: int
    availability_mapped: bool
    rows: list[LessonStudentImportPreviewRow]
    existing_student_count: int
    existing_placement_count: int
    can_apply: bool


@dataclass
class ParsedStudent:
    row_number: int
    student_id: str = ""
    name: str = ""
    email: str = ""
    windows: list[AvailabilityWindow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


@dataclass
class ParsedStudents:
    selected_sheet: str | None
    availability_mapped: bool
    rows: list[ParsedStudent]


def inspect_file(filename: str, content: bytes, sheet_name: str | None) -> LessonStudentImportInspection:
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
    return LessonStudentImportInspection(
        sheets=sheets,
        selected_sheet=selected,
        columns=columns,
        sample_rows=frame.head(10).fillna("").astype(str).to_dict(orient="records"),
        suggested_mapping=LessonStudentImportMapping(**suggested, days=suggest_day_columns(columns, taken)),
    )


def validate_mapping(mapping: LessonStudentImportMapping, columns: list[str]) -> None:
    issues: list[str] = []
    used: dict[str, str] = {}

    def claim(column: str | None, label: str, required: bool = False) -> None:
        if not column:
            if required:
                issues.append(f"Map the {label} column.")
        elif column not in columns:
            issues.append(f"The column mapped to {label} is not in the selected worksheet.")
        elif column in used:
            issues.append(f"{used[column]} and {label} cannot use the same column.")
        else:
            used[column] = label

    for field_name in FIELD_ALIASES:
        claim(getattr(mapping, field_name), FIELD_LABELS[field_name], field_name in REQUIRED_FIELDS)
    if any(day not in WEEKDAYS for day in mapping.days):
        issues.append("Availability columns must be mapped to Monday through Sunday.")
    for day in WEEKDAYS:
        claim(mapping.days.get(day), day)
    if issues:
        raise LessonStudentImportError("INVALID_MAPPING", issues[0], issues)


def parse_students(
    filename: str, content: bytes, sheet_name: str | None, mapping: LessonStudentImportMapping
) -> ParsedStudents:
    frame, _, selected = _frame(filename, content, sheet_name)
    validate_mapping(mapping, list(frame.columns))
    day_columns = mapped_day_columns(mapping.days)

    rows: list[ParsedStudent] = []
    for index, record in enumerate(frame.to_dict(orient="records")):
        if all(_is_blank(value) for value in record.values()):
            continue
        row = ParsedStudent(row_number=index + 2)
        row.student_id = _text(record.get(mapping.student_id))
        row.name = " ".join(_text(record.get(mapping.name)).split())
        row.email = _text(record.get(mapping.email)) if mapping.email else ""
        if not row.student_id:
            row.errors.append("Student ID is required.")
        if not row.name:
            row.errors.append("Name is required.")
        if day_columns:
            week = parse_week(record, day_columns, row.row_number)
            row.windows = week.windows
            row.warnings.extend(week.warnings)
            row.errors.extend(week.errors)
        rows.append(row)

    if not rows:
        raise LessonStudentImportError("NO_ROWS", "The selected worksheet has no student rows.")

    by_key: dict[str, list[ParsedStudent]] = {}
    for row in rows:
        if row.student_id:
            by_key.setdefault(row.student_id.casefold(), []).append(row)
    for group in by_key.values():
        if len(group) > 1:
            for row in group:
                others = ", ".join(str(other.row_number) for other in group if other is not row)
                row.errors.append(f"Student ID is duplicated in this file (also on row {others}).")
    return ParsedStudents(selected_sheet=selected, availability_mapped=bool(day_columns), rows=rows)


def _describe_window(window: AvailabilityWindow) -> str:
    prefix = "" if window.status == "Available" else f"{window.status} "
    return (
        f"{prefix}{_DAY_ABBREVIATIONS[window.day]} {format_time_of_day(window.start_minute)}-"
        f"{format_time_of_day(window.end_minute)}"
    )


def build_preview(
    parsed: ParsedStudents, existing_student_count: int, existing_placement_count: int
) -> LessonStudentImportPreview:
    invalid = sum(1 for row in parsed.rows if row.errors)
    return LessonStudentImportPreview(
        selected_sheet=parsed.selected_sheet,
        total_rows=len(parsed.rows),
        valid_rows=len(parsed.rows) - invalid,
        invalid_rows=invalid,
        availability_mapped=parsed.availability_mapped,
        rows=[
            LessonStudentImportPreviewRow(
                row_number=row.row_number,
                student_id=row.student_id,
                name=row.name,
                email=row.email,
                availability_window_count=len(row.windows),
                availability=[_describe_window(window) for window in row.windows],
                warnings=row.warnings,
                errors=row.errors,
            )
            for row in parsed.rows
        ],
        existing_student_count=existing_student_count,
        existing_placement_count=existing_placement_count,
        can_apply=invalid == 0,
    )
