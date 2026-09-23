"""Pydantic request/response schemas."""

from pydantic import BaseModel, ConfigDict


class PianistCreate(BaseModel):
    name: str
    email: str = ""
    max_hours_per_week: float | None = None


class PianistUpdate(BaseModel):
    name: str | None = None
    email: str | None = None
    max_hours_per_week: float | None = None


class PianistOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    name: str
    email: str
    max_hours_per_week: float | None = None


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


class LessonCreate(BaseModel):
    teacher: str = ""
    student: str = ""
    day: str
    start_minute: int
    end_minute: int
    room: str = ""
    instrument: str = ""
    required_pianist_name: str = ""
    need_pianist: bool = True


class LessonUpdate(BaseModel):
    teacher: str | None = None
    student: str | None = None
    day: str | None = None
    start_minute: int | None = None
    end_minute: int | None = None
    room: str | None = None
    instrument: str | None = None
    required_pianist_name: str | None = None
    need_pianist: bool | None = None
    assigned_pianist_id: int | None = None
    clear_assigned_pianist: bool = False


class LessonOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    teacher: str
    student: str
    day: str
    start_minute: int
    end_minute: int
    room: str
    instrument: str
    required_pianist_name: str
    need_pianist: bool
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
