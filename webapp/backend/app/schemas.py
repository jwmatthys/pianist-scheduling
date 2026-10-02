"""Pydantic request/response schemas."""

from datetime import date, datetime
from typing import Literal
from uuid import UUID

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class SessionMetadataIn(BaseModel):
    institution_name: str = Field(min_length=1, max_length=200)
    program_name: str = Field(min_length=1, max_length=200)
    term_label: str = Field(min_length=1, max_length=100)
    year: int | None = Field(default=None, ge=1000, le=9999)
    start_date: date | None = None
    end_date: date | None = None

    @field_validator("institution_name", "program_name", "term_label")
    @classmethod
    def trim_required_labels(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("This field cannot be blank.")
        return value

    @model_validator(mode="after")
    def validate_dates(self):
        if self.start_date and self.end_date and self.end_date < self.start_date:
            raise ValueError("End date must be on or after start date.")
        return self


class SessionMetadataOut(SessionMetadataIn):
    session_uuid: UUID
    created_at: datetime
    modified_at: datetime


class SessionCreateRequest(SessionMetadataIn):
    confirmed: bool = False


class SessionArchiveTerm(BaseModel):
    model_config = ConfigDict(extra="forbid")

    label: str = Field(min_length=1, max_length=100)
    year: int | None = Field(default=None, ge=1000, le=9999)
    startDate: date | None = None
    endDate: date | None = None

    @model_validator(mode="after")
    def validate_dates(self):
        if self.startDate and self.endDate and self.endDate < self.startDate:
            raise ValueError("Archive term end date precedes its start date.")
        return self


class SessionArchivePayload(BaseModel):
    model_config = ConfigDict(extra="forbid")

    path: Literal["data/session.sqlite3"]
    sha256: str = Field(pattern=r"^[0-9a-f]{64}$")


class SessionArchiveManifest(BaseModel):
    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    format: Literal["music-program-scheduler-session"]
    format_version: Literal[1] = Field(alias="formatVersion")
    session_id: UUID = Field(alias="sessionId")
    institution: str = Field(min_length=1, max_length=200)
    program: str = Field(min_length=1, max_length=200)
    term: SessionArchiveTerm
    application_version: str = Field(alias="applicationVersion", max_length=40)
    database_schema_version: int = Field(alias="databaseSchemaVersion", ge=0)
    exported_at: datetime = Field(alias="exportedAt")
    payload: SessionArchivePayload


class PianistCreate(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str
    email: str = ""
    max_hours_per_week: float | None = None


class PianistUpdate(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str | None = None
    email: str | None = None
    max_hours_per_week: float | None = None


class PianistOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    name: str
    email: str
    max_hours_per_week: float | None = None
    availability_complete: bool = False


class AvailabilitySlotIn(BaseModel):
    day: str
    slot_start_minute: int
    status: str  # Available | Tentative | Unavailable


class AvailabilityBulkIn(BaseModel):
    slots: list[AvailabilitySlotIn]


class AvailabilitySlotOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    day: str
    slot_start_minute: int
    status: str


class AvailabilityWindowColumns(BaseModel):
    start_column: str | None = None
    end_column: str | None = None


class AvailabilityImportMapping(BaseModel):
    model_config = ConfigDict(extra="forbid")

    layout: Literal["normalized", "wide"]
    person_name_column: str | None = None
    email_column: str | None = None
    day_column: str | None = None
    start_column: str | None = None
    end_column: str | None = None
    status_column: str | None = None
    wide_status: Literal["Available", "Tentative", "Unavailable"] = "Available"
    wide_windows: dict[str, list[AvailabilityWindowColumns]] = Field(default_factory=dict)


class AvailabilityImportInspection(BaseModel):
    upload_token: str
    sheets: list[str]
    selected_sheet: str | None = None
    columns: list[str]
    sample_rows: list[dict[str, str]]
    suggested_normalized: dict[str, str | None]
    suggested_wide_windows: dict[str, list[AvailabilityWindowColumns]]


class AvailabilityImportIssueOut(BaseModel):
    severity: Literal["error", "warning"]
    code: str
    message: str
    row_number: int | None = None


class AvailabilityImportWindowOut(BaseModel):
    pianist_name: str
    day: str
    start_minute: int
    end_minute: int
    status: Literal["Available", "Tentative", "Unavailable"]


class AvailabilityImportPianistOut(BaseModel):
    action: Literal["new", "invalid"]
    pianist_name: str
    email: str = ""
    max_hours_per_week: float | None = None
    days: list[str]
    row_numbers: list[int] = Field(default_factory=list)


class AvailabilityImportPreviewRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    upload_token: str
    sheet_name: str | None = None
    mapping: AvailabilityImportMapping


class AvailabilityImportPreviewOut(BaseModel):
    preview_token: str
    sheet_name: str | None = None
    rows_processed: int
    existing_pianist_count: int
    existing_assignment_count: int
    incoming_pianist_count: int
    valid_window_count: int
    existing_slots_in_scope: int
    pianists: list[AvailabilityImportPianistOut]
    windows: list[AvailabilityImportWindowOut]
    errors: list[AvailabilityImportIssueOut]
    warnings: list[AvailabilityImportIssueOut]
    can_apply: bool


class AvailabilityImportApplyRequest(BaseModel):
    preview_token: str
    confirmed: bool = False


class AvailabilityImportApplyResult(BaseModel):
    pianists_removed: int
    pianists_created: int
    assignments_cleared: int
    slots_replaced: int
    slots_created: int
    days_replaced: int
    jury_availability_windows_removed: int


class LessonCreate(BaseModel):
    teacher: str = ""
    teacher_email: str = ""
    student: str = ""
    student_id: str = ""
    day: str
    start_minute: int
    end_minute: int
    room: str = ""
    instrument: str = ""
    required_pianist_name: str = ""
    need_pianist: bool = True
    jury_required: bool = False


class LessonUpdate(BaseModel):
    teacher: str | None = None
    teacher_email: str | None = None
    student: str | None = None
    student_id: str | None = None
    day: str | None = None
    start_minute: int | None = None
    end_minute: int | None = None
    room: str | None = None
    instrument: str | None = None
    required_pianist_name: str | None = None
    need_pianist: bool | None = None
    jury_required: bool | None = None
    assigned_pianist_id: int | None = None
    clear_assigned_pianist: bool = False


class LessonOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    teacher: str
    teacher_email: str
    student: str
    student_id: str
    day: str
    start_minute: int
    end_minute: int
    room: str
    instrument: str
    required_pianist_name: str
    need_pianist: bool
    jury_required: bool = False
    assigned_pianist_id: int | None = None
    fit_quality: str
    notes: str
    hours: float
    manually_edited: bool


class ImportPreview(BaseModel):
    columns: list[str]
    rows: list[dict]
    upload_token: str


class ImportCommit(BaseModel):
    upload_token: str
    mapping: dict[str, str | None]
    save_profile_name: str | None = None


class ImportCommitResult(BaseModel):
    created: int
    skipped: int
    warnings: list[str]


class RunAssignmentResult(BaseModel):
    lessons: list[LessonOut]
    hours_by_pianist: dict[str, float]
    conflicts: list[str]
    unassigned_count: int


class ValidationResult(BaseModel):
    lessons: list[LessonOut]
    hours_by_pianist: dict[str, float]
    conflicts: list[str]
    unassigned_count: int
