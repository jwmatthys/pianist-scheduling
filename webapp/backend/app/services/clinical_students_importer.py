"""Local CSV/XLSX/XLS parsing and validation for the unified Clinical Student Roster import.

Pure parsing: it never touches the database. The caller supplies a snapshot of the
current roster so rows can be classified as create/update; the router combines the result
with ClinicalRepository.apply_roster_import for the commit phase.

The file is the whole roster, one row per student, and behaves like the Lesson and Tutor
student imports: identity, weekly availability and tag states are imported together and
every listed student's availability and tag states are replaced (blank or unmapped
columns mean none). Students are matched by Student ID (case-insensitive): known IDs keep
their internal identity and placement, unknown IDs are created, and students absent from
the file are removed together with their placements. Weekday-cell parsing lives in
``weekly_availability_cells``. Tags missing from the Clinical registry are created on apply.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Mapping

from pydantic import BaseModel, Field

from ..clinical_models import TagState, normalize_tag_name, tag_identity_key
from .availability import WEEKDAYS, AvailabilityWindow
from .weekly_availability_cells import (
    describe_change,
    mapped_day_columns,
    parse_week,
    suggest_day_columns,
)
from .import_common import (
    ImportFileError,
    cell_text,
    header_key,
    is_blank,
    read_frame,
    split_tag_cell,
)

FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "student_id": ("student id", "studentid", "id", "student number", "student", "emplid"),
    "name": ("name", "student name", "full name"),
    "email": ("email", "student email", "e mail", "email address"),
}
TAG_FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "required_tags": ("required tags", "required tag", "required"),
    "preferred_tags": ("preferred tags", "preferred tag", "preferred"),
    "restricted_tags": ("restricted tags", "restricted tag", "restricted"),
}
FIELD_ALIASES.update(TAG_FIELD_ALIASES)
FIELD_LABELS = {
    "student_id": "Student ID",
    "name": "Name",
    "email": "Email",
    "required_tags": "Required Tags",
    "preferred_tags": "Preferred Tags",
    "restricted_tags": "Restricted Tags",
}
TAG_FIELD_STATES = {
    "required_tags": TagState.REQUIRED,
    "preferred_tags": TagState.PREFERRED,
    "restricted_tags": TagState.RESTRICTED,
}
REQUIRED_FIELDS = ("student_id", "name")


class ClinicalStudentsImportMapping(BaseModel):
    student_id: str | None = None
    name: str | None = None
    email: str | None = None
    required_tags: str | None = None
    preferred_tags: str | None = None
    restricted_tags: str | None = None
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
    required_tags: list[str]
    preferred_tags: list[str]
    restricted_tags: list[str]
    warnings: list[str]
    errors: list[str]


class ClinicalStudentsImportPreview(BaseModel):
    selected_sheet: str | None
    total_rows: int
    valid_rows: int
    invalid_rows: int
    students_to_create: int
    students_to_update: int
    students_to_remove: int
    placements_to_remove: int
    availability_mapped: bool
    tag_columns: list[str]
    new_tags: list[str]
    reused_tags: list[str]
    rows: list[ClinicalStudentsImportPreviewRow]
    can_apply: bool


@dataclass(frozen=True)
class ExistingStudent:
    clinical_student_id: str
    name: str
    email: str
    window_count: int
    placement_count: int = 0


@dataclass
class ParsedStudent:
    row_number: int
    student_id: str = ""
    name: str = ""
    email: str = ""
    clinical_student_id: str | None = None
    # None only while the row has errors.
    windows: list[AvailabilityWindow] | None = None
    action: Literal["create", "update"] | None = None
    student_change: str = ""
    availability_change: str = "No availability change"
    # (tag name, state) pairs; only meaningful when a tag column is mapped.
    tags: list[tuple[str, str]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


@dataclass
class ParsedStudents:
    selected_sheet: str | None
    availability_mapped: bool
    rows: list[ParsedStudent]
    tag_columns: list[str] = field(default_factory=list)


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
        raise ImportFileError("INVALID_MAPPING", issues[0], issues)


def _parse_tags(row: ParsedStudent, record: dict, tag_columns: dict[str, str]) -> None:
    seen: dict[str, tuple[str, str]] = {}
    for field_name, column in tag_columns.items():
        state = TAG_FIELD_STATES[field_name].value
        for part in split_tag_cell(record.get(column)):
            name = normalize_tag_name(part)
            key = tag_identity_key(name)
            if not key:
                continue
            if key not in seen:
                seen[key] = (name, state)
            elif seen[key][1] != state:
                row.errors.append(
                    f"Tag '{name}' is listed in more than one of Required Tags, Preferred Tags and Restricted Tags."
                )
    row.tags = list(seen.values())


def _names(row: ParsedStudent, state: TagState) -> list[str]:
    return [name for name, value in row.tags if value == state.value]


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
    tag_columns = {
        field_name: getattr(mapping, field_name) for field_name in TAG_FIELD_STATES if getattr(mapping, field_name)
    }

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
                row.errors.append("Name is required.")
            row.student_change = "Create student" if row.action else ""
        else:
            row.name = file_name
            row.email = file_email or ""
            if not file_name:
                row.errors.append("Name is required.")
            row.student_change = _student_change(existing, row.name != existing.name, row.email != existing.email)

        _parse_tags(row, record, tag_columns)

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
    return ParsedStudents(
        selected_sheet=selected,
        availability_mapped=bool(day_columns),
        rows=rows,
        tag_columns=list(tag_columns.values()),
    )


def build_preview(
    parsed: ParsedStudents,
    existing_tag_keys: set[str] | None = None,
    roster: Mapping[str, ExistingStudent] | None = None,
) -> ClinicalStudentsImportPreview:
    existing_tag_keys = existing_tag_keys or set()
    listed = {row.student_id.casefold() for row in parsed.rows if row.student_id}
    removed = [student for key, student in (roster or {}).items() if key not in listed]
    new_tags: dict[str, str] = {}
    reused: dict[str, str] = {}
    for row in parsed.rows:
        for name, _ in row.tags:
            key = tag_identity_key(name)
            (reused if key in existing_tag_keys else new_tags).setdefault(key, name)
    invalid = sum(1 for row in parsed.rows if row.errors)
    return ClinicalStudentsImportPreview(
        selected_sheet=parsed.selected_sheet,
        total_rows=len(parsed.rows),
        valid_rows=len(parsed.rows) - invalid,
        invalid_rows=invalid,
        students_to_create=sum(1 for row in parsed.rows if row.action == "create" and not row.errors),
        students_to_update=sum(1 for row in parsed.rows if row.action == "update" and not row.errors),
        students_to_remove=len(removed),
        placements_to_remove=sum(student.placement_count for student in removed),
        availability_mapped=parsed.availability_mapped,
        tag_columns=parsed.tag_columns,
        new_tags=sorted(new_tags.values()),
        reused_tags=sorted(reused.values()),
        rows=[
            ClinicalStudentsImportPreviewRow(
                row_number=row.row_number,
                student_id=row.student_id,
                name=row.name,
                email=row.email,
                action=row.action,
                student_change=row.student_change,
                availability_change=row.availability_change,
                required_tags=_names(row, TagState.REQUIRED),
                preferred_tags=_names(row, TagState.PREFERRED),
                restricted_tags=_names(row, TagState.RESTRICTED),
                warnings=row.warnings,
                errors=row.errors,
            )
            for row in parsed.rows
        ],
        can_apply=bool(parsed.rows) and invalid == 0,
    )
