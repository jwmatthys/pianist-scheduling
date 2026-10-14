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
