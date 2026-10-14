"""Pydantic request/response schemas for the Clinical Placement module."""

from datetime import datetime

from pydantic import BaseModel, ConfigDict, Field


class ClinicalTagOut(BaseModel):
    id: str
    name: str
    opportunity_count: int
    student_count: int
    is_unused: bool


class ClinicalTagIn(BaseModel):
    name: str = Field(max_length=200)


class ClinicalTagDeleteOut(BaseModel):
    affected_student_count: int


class ClinicalTagRef(BaseModel):
    id: str
    name: str

    model_config = ConfigDict(from_attributes=True)


class PlacementOpportunityFields(BaseModel):
    site_name: str = Field(max_length=200)
    site_address: str = Field(default="", max_length=400)
    contact_name: str = Field(default="", max_length=200)
    contact_email: str = Field(default="", max_length=200)
    contact_phone: str = Field(default="", max_length=100)
    day_of_week: str = Field(max_length=20)
    start_minute: int
    end_minute: int
    session_length_minutes: int
    travel_minutes: int = 0
    capacity: int = 1


class PlacementOpportunityIn(PlacementOpportunityFields):
    tag_ids: list[str] = Field(default_factory=list)


class PlacementOpportunityUpdate(PlacementOpportunityFields):
    # None leaves associations unchanged; tags are normally persisted via the tags endpoint.
    tag_ids: list[str] | None = None


class PlacementOpportunityTagsIn(BaseModel):
    tag_ids: list[str]


class PlacementOpportunityOut(PlacementOpportunityFields):
    id: str
    tags: list[ClinicalTagRef]
    placement_count: int
    created_at: datetime
    updated_at: datetime


class PlacementOpportunityDeleteOut(BaseModel):
    removed_placement_count: int


class ClinicalModuleStateOut(BaseModel):
    schedule_revision: int
    source_data_revision: int
    is_schedule_stale: bool


class ClinicalStudentFields(BaseModel):
    student_id: str = Field(max_length=100)
    name: str = Field(max_length=200)
    email: str = Field(default="", max_length=200)


class ClinicalStudentOut(ClinicalStudentFields):
    id: str
    availability_window_count: int
    # Only non-neutral states are present; Neutral is the absence of an entry.
    tag_states: dict[str, str]
    created_at: datetime
    updated_at: datetime


class ClinicalStudentTagStatesIn(BaseModel):
    tag_states: dict[str, str]


class ClinicalAvailabilityWindowIO(BaseModel):
    day: str
    start_minute: int
    end_minute: int
    status: str


class ClinicalStudentAvailabilityIO(BaseModel):
    windows: list[ClinicalAvailabilityWindowIO] = Field(max_length=500)


class ClinicalReadinessSubjectOut(BaseModel):
    kind: str
    id: str
    label: str


class ClinicalReadinessWarningOut(BaseModel):
    code: str
    severity: str
    category: str
    subject: ClinicalReadinessSubjectOut | None = None


class ClinicalReadinessOut(BaseModel):
    source_data_revision: int
    is_ready: bool
    student_count: int
    opportunity_count: int
    feasible_candidate_count: int
    total_candidate_count: int
    students_with_candidates: int
    warnings: list[ClinicalReadinessWarningOut]


class ClinicalPlacementOut(BaseModel):
    student_id: str
    kind: str
    opportunity_id: str | None
    start_minute: int | None


class ClinicalUnplacedOut(BaseModel):
    student_id: str
    student_name: str
    reason: str


class ClinicalScheduleOut(BaseModel):
    schedule_revision: int
    source_data_revision: int
    is_schedule_stale: bool
    has_placements: bool
    student_count: int
    placed_count: int
    unplaced_count: int
    used_capacity: int
    total_capacity: int
    utilization: float
    preferred_tag_matches: int
    tentative_minutes: int
    unverified_placement_count: int
    unplaced: list[ClinicalUnplacedOut]
    placements: list[ClinicalPlacementOut]
    entries: list["ClinicalScheduleEntryOut"]
    verification: list["ClinicalPlacementVerificationOut"]


class ClinicalPlacementWarningOut(BaseModel):
    code: str
    message: str


class ClinicalScheduleEntryOut(BaseModel):
    student_id: str
    student_name: str
    student_number: str
    kind: str | None
    opportunity_id: str | None
    site_name: str | None
    opportunity_label: str | None
    day: str | None
    start_minute: int | None
    end_minute: int | None
    tags: list[str]
    preferred_tag_match_count: int
    tentative_minutes: int | None
    verified: bool | None
    warnings: list[ClinicalPlacementWarningOut]
    unplaced_reason: str | None


class ClinicalPlacementVerificationOut(BaseModel):
    student_id: str
    student_name: str
    opportunity_id: str | None
    site_name: str | None
    day: str | None
    start_minute: int | None
    warnings: list[ClinicalPlacementWarningOut]


class ClinicalManualPlacementIn(BaseModel):
    student_id: str
    opportunity_id: str
    start_minute: int = Field(ge=0, lt=1440)


class ClinicalPlacementWarningsOut(BaseModel):
    warnings: list[ClinicalPlacementWarningOut]


ClinicalScheduleOut.model_rebuild()
