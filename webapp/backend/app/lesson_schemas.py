"""Pydantic request/response schemas for the Lesson Scheduling module."""

from datetime import datetime

from pydantic import BaseModel, ConfigDict, Field


class LessonModuleStateOut(BaseModel):
    schedule_revision: int
    source_data_revision: int
    is_schedule_stale: bool


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
