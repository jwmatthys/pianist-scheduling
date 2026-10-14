"""Local CSV/XLSX/XLS parsing and validation for Clinical Student Roster import.

Pure parsing: it never touches the database. The router combines it with
ClinicalRepository.replace_students for the commit phase. Availability columns are
not part of this import; unmapped columns are ignored.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pydantic import BaseModel

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
REQUIRED_FIELDS = ("student_id", "name")


class ClinicalStudentsImportInspection(BaseModel):
    sheets: list[str]
    selected_sheet: str | None
    columns: list[str]
    sample_rows: list[dict[str, str]]
    suggested_mapping: dict[str, str | None]


class ClinicalStudentsImportMapping(BaseModel):
    student_id: str | None = None
    name: str | None = None
    email: str | None = None


class ClinicalStudentsImportPreviewRow(BaseModel):
    row_number: int
    student_id: str
    name: str
    email: str
    errors: list[str]


class ClinicalStudentsImportPreview(BaseModel):
    selected_sheet: str | None
    total_rows: int
    valid_rows: int
    invalid_rows: int
    rows: list[ClinicalStudentsImportPreviewRow]
    existing_student_count: int
    existing_placement_count: int
    can_apply: bool


@dataclass
class ParsedStudent:
    row_number: int
    student_id: str = ""
    name: str = ""
    email: str = ""
    errors: list[str] = field(default_factory=list)


@dataclass
class ParsedStudents:
    selected_sheet: str | None
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
        suggested_mapping=suggested,
    )


def validate_mapping(mapping: ClinicalStudentsImportMapping, columns: list[str]) -> None:
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
        raise ClinicalImportError("INVALID_MAPPING", issues[0], issues)


def parse_students(
    filename: str, content: bytes, sheet_name: str | None, mapping: ClinicalStudentsImportMapping
) -> ParsedStudents:
    frame, _, selected = read_frame(filename, content, sheet_name)
    validate_mapping(mapping, list(frame.columns))
    mapped = {name: getattr(mapping, name) for name in FIELD_ALIASES if getattr(mapping, name)}

    rows: list[ParsedStudent] = []
    for index, record in enumerate(frame.to_dict(orient="records")):
        if all(is_blank(value) for value in record.values()):
            continue
        row = ParsedStudent(row_number=index + 2)
        for field_name, column in mapped.items():
            setattr(row, field_name, cell_text(record.get(column)))
        if not row.student_id:
            row.errors.append("Student ID is required.")
        if not row.name:
            row.errors.append("Name is required.")
        rows.append(row)

    by_key: dict[str, list[ParsedStudent]] = {}
    for row in rows:
        if row.student_id:
            by_key.setdefault(row.student_id.casefold(), []).append(row)
    for group in by_key.values():
        if len(group) > 1:
            for row in group:
                others = ", ".join(str(other.row_number) for other in group if other is not row)
                row.errors.append(f"Student ID is duplicated in this file (also on row {others}).")
    return ParsedStudents(selected_sheet=selected, rows=rows)


def build_preview(
    parsed: ParsedStudents, existing_student_count: int, existing_placement_count: int
) -> ClinicalStudentsImportPreview:
    invalid = sum(1 for row in parsed.rows if row.errors)
    return ClinicalStudentsImportPreview(
        selected_sheet=parsed.selected_sheet,
        total_rows=len(parsed.rows),
        valid_rows=len(parsed.rows) - invalid,
        invalid_rows=invalid,
        rows=[
            ClinicalStudentsImportPreviewRow(
                row_number=row.row_number,
                student_id=row.student_id,
                name=row.name,
                email=row.email,
                errors=row.errors,
            )
            for row in parsed.rows
        ],
        existing_student_count=existing_student_count,
        existing_placement_count=existing_placement_count,
        can_apply=bool(parsed.rows) and invalid == 0,
    )
