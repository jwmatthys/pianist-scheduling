"""Local CSV/XLSX/XLS parsing and validation for the Tutor Scheduling Student Roster import.

Pure parsing: it never touches the database. The router combines it with
TutorRepository.replace_students for the commit phase (Phase B re-parses the file; preview
data is never trusted).

One import carries the roster and weekly availability together; there is no separate
availability import. Spreadsheet reading and weekday-cell parsing reuse the pure,
module-neutral helpers shared with the Clinical imports. Unlike the Clinical
roster import (which merges by Student ID), this import is a full replacement.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pydantic import BaseModel, Field

from ..tutor_domain import TutorTagState
from ..tutor_models import tutor_tag_identity_key, normalize_tutor_tag_name

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
)
from .time_parsing import format_time_of_day

__all__ = ["MAX_UPLOAD_BYTES", "TutorStudentImportError"]

TutorStudentImportError = ImportFileError

FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "student_id": ("student id", "studentid", "id", "student number", "student", "emplid"),
    "name": ("name", "student name", "full name"),
    "email": ("email", "student email", "e mail", "email address"),
    "required_tags": ("required tags", "required tag", "required"),
    "preferred_tags": ("preferred tags", "preferred tag", "preferred"),
    "restricted_tags": ("restricted tags", "restricted tag", "restricted"),
}
FIELD_LABELS = {
    "student_id": "Student ID",
    "name": "Name",
    "email": "Email",
    "required_tags": "Required Tags",
    "preferred_tags": "Preferred Tags",
    "restricted_tags": "Restricted Tags",
}
# Tag column -> student/tag state stored for the student.
TAG_FIELD_STATES = {
    "required_tags": TutorTagState.REQUIRED,
    "preferred_tags": TutorTagState.PREFERRED,
    "restricted_tags": TutorTagState.RESTRICTED,
}
REQUIRED_FIELDS = ("student_id", "name")
_DAY_ABBREVIATIONS = {day: day[:3] for day in WEEKDAYS}


class TutorStudentImportMapping(BaseModel):
    student_id: str | None = None
    name: str | None = None
    email: str | None = None
    required_tags: str | None = None
    preferred_tags: str | None = None
    restricted_tags: str | None = None
    days: dict[str, str | None] = Field(default_factory=dict)


class TutorStudentImportInspection(BaseModel):
    sheets: list[str]
    selected_sheet: str | None
    columns: list[str]
    sample_rows: list[dict[str, str]]
    suggested_mapping: TutorStudentImportMapping


class TutorStudentImportPreviewRow(BaseModel):
    row_number: int
    student_id: str
    name: str
    email: str
    required_tags: list[str]
    preferred_tags: list[str]
    restricted_tags: list[str]
    availability_window_count: int
    availability: list[str]
    warnings: list[str]
    errors: list[str]


class TutorStudentImportPreview(BaseModel):
    selected_sheet: str | None
    total_rows: int
    valid_rows: int
    invalid_rows: int
    availability_mapped: bool
    tag_columns: list[str]
    rows: list[TutorStudentImportPreviewRow]
    new_tags: list[str]
    reused_tags: list[str]
    existing_student_count: int
    existing_placement_count: int
    can_apply: bool


@dataclass
class ParsedStudent:
    row_number: int
    student_id: str = ""
    name: str = ""
    email: str = ""
    tags: list[tuple[str, str]] = field(default_factory=list)  # (tag name, state value)
    windows: list[AvailabilityWindow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


@dataclass
class ParsedStudents:
    selected_sheet: str | None
    availability_mapped: bool
    tag_columns: list[str]
    rows: list[ParsedStudent]


def inspect_file(filename: str, content: bytes, sheet_name: str | None) -> TutorStudentImportInspection:
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
    return TutorStudentImportInspection(
        sheets=sheets,
        selected_sheet=selected,
        columns=columns,
        sample_rows=frame.head(10).fillna("").astype(str).to_dict(orient="records"),
        suggested_mapping=TutorStudentImportMapping(**suggested, days=suggest_day_columns(columns, taken)),
    )


def validate_mapping(mapping: TutorStudentImportMapping, columns: list[str]) -> None:
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
        raise TutorStudentImportError("INVALID_MAPPING", issues[0], issues)


def _parse_tags(row: ParsedStudent, record: dict, tag_columns: dict[str, str]) -> None:
    seen: dict[str, tuple[str, str]] = {}
    for field_name, column in tag_columns.items():
        state = TAG_FIELD_STATES[field_name].value
        for part in split_tag_cell(record.get(column)):
            name = normalize_tutor_tag_name(part)
            key = tutor_tag_identity_key(name)
            if not key:
                continue
            if key not in seen:
                seen[key] = (name, state)
            elif seen[key][1] != state:
                row.errors.append(f"Tag '{name}' is listed in more than one of Required Tags, Preferred Tags and Restricted Tags.")
    row.tags = list(seen.values())


def parse_students(
    filename: str, content: bytes, sheet_name: str | None, mapping: TutorStudentImportMapping
) -> ParsedStudents:
    frame, _, selected = _frame(filename, content, sheet_name)
    validate_mapping(mapping, list(frame.columns))
    day_columns = mapped_day_columns(mapping.days)
    tag_columns = {field_name: getattr(mapping, field_name) for field_name in TAG_FIELD_STATES if getattr(mapping, field_name)}

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
        _parse_tags(row, record, tag_columns)
        rows.append(row)

    if not rows:
        raise TutorStudentImportError("NO_ROWS", "The selected worksheet has no student rows.")

    by_key: dict[str, list[ParsedStudent]] = {}
    for row in rows:
        if row.student_id:
            by_key.setdefault(row.student_id.casefold(), []).append(row)
    for group in by_key.values():
        if len(group) > 1:
            for row in group:
                others = ", ".join(str(other.row_number) for other in group if other is not row)
                row.errors.append(f"Student ID is duplicated in this file (also on row {others}).")
    return ParsedStudents(
        selected_sheet=selected,
        availability_mapped=bool(day_columns),
        tag_columns=list(tag_columns.values()),
        rows=rows,
    )


def _describe_window(window: AvailabilityWindow) -> str:
    prefix = "" if window.status == "Available" else f"{window.status} "
    return (
        f"{prefix}{_DAY_ABBREVIATIONS[window.day]} {format_time_of_day(window.start_minute)}-"
        f"{format_time_of_day(window.end_minute)}"
    )


def _names(row: ParsedStudent, state: TutorTagState) -> list[str]:
    return [name for name, value in row.tags if value == state.value]


def build_preview(
    parsed: ParsedStudents,
    existing_tag_keys: set[str],
    existing_student_count: int,
    existing_placement_count: int,
) -> TutorStudentImportPreview:
    new_tags: dict[str, str] = {}
    reused: dict[str, str] = {}
    for row in parsed.rows:
        for name, _ in row.tags:
            key = tutor_tag_identity_key(name)
            (reused if key in existing_tag_keys else new_tags).setdefault(key, name)
    invalid = sum(1 for row in parsed.rows if row.errors)
    return TutorStudentImportPreview(
        selected_sheet=parsed.selected_sheet,
        total_rows=len(parsed.rows),
        valid_rows=len(parsed.rows) - invalid,
        invalid_rows=invalid,
        availability_mapped=parsed.availability_mapped,
        tag_columns=parsed.tag_columns,
        rows=[
            TutorStudentImportPreviewRow(
                row_number=row.row_number,
                student_id=row.student_id,
                name=row.name,
                email=row.email,
                required_tags=_names(row, TAG_FIELD_STATES["required_tags"]),
                preferred_tags=_names(row, TAG_FIELD_STATES["preferred_tags"]),
                restricted_tags=_names(row, TAG_FIELD_STATES["restricted_tags"]),
                availability_window_count=len(row.windows),
                availability=[_describe_window(window) for window in row.windows],
                warnings=row.warnings,
                errors=row.errors,
            )
            for row in parsed.rows
        ],
        new_tags=sorted(new_tags.values()),
        reused_tags=sorted(reused.values()),
        existing_student_count=existing_student_count,
        existing_placement_count=existing_placement_count,
        can_apply=invalid == 0,
    )
