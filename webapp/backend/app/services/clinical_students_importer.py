"""Local CSV/XLSX/XLS parsing and validation for the unified Clinical Student Roster import.

Pure parsing: it never touches the database. The caller supplies a snapshot of the
current roster so rows can be classified as create/update; the router combines the result
with ClinicalRepository.apply_roster_import for the commit phase.

One import handles identity and availability together. Students are matched by Student ID
(case-insensitive): unknown IDs are created, known IDs are updated, and students absent
from the file are left untouched. Optional weekday columns carry recurring Availability
Windows (cell parsing lives in ``clinical_availability_importer``). When any weekday column
is mapped, each listed student's availability is replaced (blank cells mean no
availability); when none is mapped, availability is left unchanged.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Mapping

from pydantic import BaseModel, Field

from .availability import WEEKDAYS, AvailabilityWindow
from .clinical_availability_importer import (
    describe_change,
    mapped_day_columns,
    parse_week,
    suggest_day_columns,
)
from .clinical_import_common import (
    ClinicalImportError,
    cell_text,
    header_key,
    is_blank,
    read_frame,
)

FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "student_id": ("student id", "studentid", "id", "student number", "student", "emplid"),
    "name": ("name", "student name", "full name"),
    "email": ("email", "student email", "e mail", "email address"),
}
FIELD_LABELS = {"student_id": "Student ID", "name": "Name", "email": "Email"}
REQUIRED_FIELDS = ("student_id",)


class ClinicalStudentsImportMapping(BaseModel):
    student_id: str | None = None
    name: str | None = None
    email: str | None = None
    days: dict[str, str | None] = Field(default_factory=dict)


class ClinicalStudentsImportInspection(BaseModel):
    sheets: list[str]
    selected_sheet: str | None
    columns: list[str]
    sample_rows: list[dict[str, str]]
    suggested_mapping: ClinicalStudentsImportMapping


class ClinicalStudentsImportPreviewRow(BaseModel):
    row_number: int
    student_id: str
    name: str
    email: str
    action: Literal["create", "update"] | None
    student_change: str
    availability_change: str
    warnings: list[str]
    errors: list[str]


class ClinicalStudentsImportPreview(BaseModel):
    selected_sheet: str | None
    total_rows: int
    valid_rows: int
    invalid_rows: int
    students_to_create: int
    students_to_update: int
    availability_mapped: bool
    rows: list[ClinicalStudentsImportPreviewRow]
    can_apply: bool


@dataclass(frozen=True)
class ExistingStudent:
    clinical_student_id: str
    name: str
    email: str
    window_count: int


@dataclass
class ParsedStudent:
    row_number: int
    student_id: str = ""
    name: str = ""
    email: str = ""
    clinical_student_id: str | None = None
    # None means availability is not part of this import (no weekday column mapped).
    windows: list[AvailabilityWindow] | None = None
    action: Literal["create", "update"] | None = None
    student_change: str = ""
    availability_change: str = "No availability change"
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


@dataclass
class ParsedStudents:
    selected_sheet: str | None
    availability_mapped: bool
    rows: list[ParsedStudent]


def inspect_file(filename: str, content: bytes, sheet_name: str | None) -> ClinicalStudentsImportInspection:
    frame, sheets, selected = read_frame(filename, content, sheet_name)
    columns = list(frame.columns)
    keys = {column: header_key(column) for column in columns}
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
    return ClinicalStudentsImportInspection(
        sheets=sheets,
        selected_sheet=selected,
        columns=columns,
        sample_rows=frame.head(10).fillna("").astype(str).to_dict(orient="records"),
        suggested_mapping=ClinicalStudentsImportMapping(**suggested, days=suggest_day_columns(columns, taken)),
    )


def validate_mapping(mapping: ClinicalStudentsImportMapping, columns: list[str]) -> None:
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
        raise ClinicalImportError("INVALID_MAPPING", issues[0], issues)


def _student_change(existing: ExistingStudent | None, name_changed: bool, email_changed: bool) -> str:
    if existing is None:
        return "Create student"
    if name_changed and email_changed:
        return "Update name and email"
    if name_changed:
        return "Update name"
    if email_changed:
        return "Update email"
    return "No student change"


def parse_students(
    filename: str,
    content: bytes,
    sheet_name: str | None,
    mapping: ClinicalStudentsImportMapping,
    roster: Mapping[str, ExistingStudent] | None = None,
) -> ParsedStudents:
    """Parse the file against the roster snapshot (keyed by case-folded Student ID)."""
    roster = roster or {}
    frame, _, selected = read_frame(filename, content, sheet_name)
    validate_mapping(mapping, list(frame.columns))
    day_columns = mapped_day_columns(mapping.days)

    rows: list[ParsedStudent] = []
    for index, record in enumerate(frame.to_dict(orient="records")):
        if all(is_blank(value) for value in record.values()):
            continue
        row = ParsedStudent(row_number=index + 2)
        rows.append(row)
        row.student_id = cell_text(record.get(mapping.student_id))
        file_name = cell_text(record.get(mapping.name)) if mapping.name else ""
        file_email = cell_text(record.get(mapping.email)) if mapping.email else None
        if not row.student_id:
            row.errors.append("Student ID is required.")
        existing = roster.get(row.student_id.casefold()) if row.student_id else None
        if row.student_id:
            row.action = "update" if existing else "create"
            row.clinical_student_id = existing.clinical_student_id if existing else None

        if existing is None:
            row.name = file_name
            row.email = file_email or ""
            if row.student_id and not file_name:
                row.errors.append(
                    "Name is required."
                    if mapping.name
                    else "Name is required to create a new student; map the Name column."
                )
            row.student_change = "Create student" if row.action else ""
        else:
            row.name = file_name or existing.name
            row.email = existing.email if file_email is None else file_email
            row.student_change = _student_change(existing, row.name != existing.name, row.email != existing.email)

        if day_columns:
            week = parse_week(record, day_columns, row.row_number)
            row.errors.extend(week.errors)
            row.warnings.extend(week.warnings)
            if not week.errors:
                row.windows = week.windows
                row.availability_change = describe_change(existing.window_count if existing else 0, len(week.windows))

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


def build_preview(parsed: ParsedStudents) -> ClinicalStudentsImportPreview:
    invalid = sum(1 for row in parsed.rows if row.errors)
    return ClinicalStudentsImportPreview(
        selected_sheet=parsed.selected_sheet,
        total_rows=len(parsed.rows),
        valid_rows=len(parsed.rows) - invalid,
        invalid_rows=invalid,
        students_to_create=sum(1 for row in parsed.rows if row.action == "create" and not row.errors),
        students_to_update=sum(1 for row in parsed.rows if row.action == "update" and not row.errors),
        availability_mapped=parsed.availability_mapped,
        rows=[
            ClinicalStudentsImportPreviewRow(
                row_number=row.row_number,
                student_id=row.student_id,
                name=row.name,
                email=row.email,
                action=row.action,
                student_change=row.student_change,
                availability_change=row.availability_change,
                warnings=row.warnings,
                errors=row.errors,
            )
            for row in parsed.rows
        ],
        can_apply=bool(parsed.rows) and invalid == 0,
    )
