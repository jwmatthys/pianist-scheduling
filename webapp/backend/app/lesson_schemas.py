"""Pydantic request/response schemas for the Lesson Scheduling module."""

from datetime import datetime
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


class LessonModuleStateOut(BaseModel):
    schedule_revision: int
    source_data_revision: int
    is_schedule_stale: bool


class LessonReadinessSubjectOut(BaseModel):
    kind: str
    id: str
    label: str


class LessonReadinessWarningOut(BaseModel):
    code: str
    severity: str
    category: str
    subjects: list[LessonReadinessSubjectOut]


class LessonReadinessOut(BaseModel):
    source_data_revision: int
    is_ready: bool
    student_count: int
    instructor_count: int
    feasible_candidate_count: int
    total_candidate_count: int
    students_with_candidates: int
    instructors_with_schedulable_lessons: int
    warnings: list[LessonReadinessWarningOut]


class LessonScheduleOut(BaseModel):
    schedule_revision: int
    source_data_revision: int
    is_schedule_stale: bool
    has_schedule: bool
    students_scheduled: int
    students_unscheduled: int
    instructor_utilization: float
    preferred_tag_matches: int
    tentative_minutes: int
    gap_cost_minutes: int
    automatic_count: int
    manual_count: int
    manually_unplaced_count: int
    verified_count: int
    unverified_count: int
    entries: list["LessonScheduleEntryOut"]
    verification: list["LessonPlacementVerificationOut"]


class LessonPlacementWarningOut(BaseModel):
    code: str
    message: str


class LessonScheduleEntryOut(BaseModel):
    student_id: str
    student_name: str
    placement_type: str | None
    placement_state: Literal["unassigned", "assigned", "manuallyUnplaced"]
    instructor_id: str | None
    instructor_name: str | None
    day: str | None
    start_minute: int | None
    end_minute: int | None
    preferred_tag_match_count: int
    tentative_minutes: int
    verified: bool | None
    warnings: list[LessonPlacementWarningOut]
    unplaced_reason: str | None


class LessonPlacementVerificationOut(BaseModel):
    student_id: str
    student_name: str
    instructor_id: str | None
    instructor_name: str | None
    day: str | None
    start_minute: int | None
    warnings: list[LessonPlacementWarningOut]


class LessonManualPlacementIn(BaseModel):
    student_id: str
    instructor_id: str
    day: str
    start_minute: int = Field(ge=0, lt=1440)


class LessonPlacementWarningsOut(BaseModel):
    warnings: list[LessonPlacementWarningOut]


class LessonTagOut(BaseModel):
    id: str
    name: str
    instructor_count: int
    student_count: int
    is_unused: bool


class LessonTagIn(BaseModel):
    name: str = Field(max_length=200)


class LessonTagDeleteOut(BaseModel):
    affected_instructor_count: int
    affected_student_count: int


class LessonTagRef(BaseModel):
    id: str
    name: str

    model_config = ConfigDict(from_attributes=True)


class LessonInstructorFields(BaseModel):
    name: str = Field(max_length=200)
    email: str = Field(default="", max_length=200)
    room: str = Field(default="", max_length=200)
    lesson_length_minutes: int = 60
    capacity: int = 1


class LessonInstructorIn(LessonInstructorFields):
    # None leaves associations unchanged on update; empty list clears them.
    tag_ids: list[str] | None = Field(default=None, max_length=200)


class LessonInstructorTagsIn(BaseModel):
    tag_ids: list[str] = Field(max_length=200)


class LessonInstructorOut(LessonInstructorFields):
    id: str
    tags: list[LessonTagRef]
    availability_window_count: int
    created_at: datetime
    updated_at: datetime


class LessonInstructorDeleteOut(BaseModel):
    removed_placement_count: int
    placements_preserved: int


class LessonAvailabilityWindowIO(BaseModel):
    day: str
    start_minute: int
    end_minute: int
    status: str


class LessonInstructorAvailabilityIO(BaseModel):
    windows: list[LessonAvailabilityWindowIO] = Field(max_length=500)


class LessonStudentFields(BaseModel):
    student_id: str = Field(max_length=100)
    name: str = Field(max_length=200)
    email: str = Field(default="", max_length=200)


class LessonStudentIn(LessonStudentFields):
    # None leaves tag states unchanged on update; an object replaces the whole set
    # (Neutral entries are not stored).
    tag_states: dict[str, str] | None = None


class LessonStudentTagStatesIn(BaseModel):
    tag_states: dict[str, str]


class LessonStudentOut(LessonStudentFields):
    id: str
    availability_window_count: int
    # Only non-neutral states are present; Neutral is the absence of an entry.
    tag_states: dict[str, str]
    created_at: datetime
    updated_at: datetime


class LessonStudentAvailabilityIO(BaseModel):
    windows: list[LessonAvailabilityWindowIO] = Field(max_length=500)


LessonScheduleOut.model_rebuild()
