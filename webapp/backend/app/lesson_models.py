"""Lesson Scheduling module-owned persistence models.

These tables hold only Lesson Scheduling data. They have no foreign keys to, and
no dependency on, Accompanist, Jury, or Clinical tables or shared person
identities; matching names, IDs, emails, or tags across modules create no
relationship. Internal identifiers are application-generated UUID strings and
are never shown to users.
"""

from datetime import datetime
from enum import Enum
from uuid import uuid4

from sqlalchemy import (
    Boolean,
    CheckConstraint,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    UniqueConstraint,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .module_models import ModuleBase

LESSON_MODULE_ID = "lesson-scheduling"
DEFAULT_LESSON_LENGTH_MINUTES = 60
DEFAULT_INSTRUCTOR_CAPACITY = 1


class LessonTagState(str, Enum):
    NEUTRAL = "neutral"
    REQUIRED = "required"
    PREFERRED = "preferred"
    RESTRICTED = "restricted"


class LessonPlacementType(str, Enum):
    AUTOMATIC = "automatic"
    MANUAL = "manual"
    MANUALLY_UNPLACED = "manuallyUnplaced"


class LessonPlacementState(str, Enum):
    """Coordinator-facing state of a student: no placement record is UNASSIGNED."""

    UNASSIGNED = "unassigned"
    ASSIGNED = "assigned"
    MANUALLY_UNPLACED = "manuallyUnplaced"

    @classmethod
    def from_placement_type(cls, placement_type: str | None) -> "LessonPlacementState":
        if placement_type is None:
            return cls.UNASSIGNED
        if placement_type == LessonPlacementType.MANUALLY_UNPLACED.value:
            return cls.MANUALLY_UNPLACED
        return cls.ASSIGNED


class LessonUnplacedReason(str, Enum):
    NO_FEASIBLE_CANDIDATE = "NO_FEASIBLE_CANDIDATE"
    CAPACITY_LIMIT = "CAPACITY_LIMIT"
    INSTRUCTOR_OVERLAP = "INSTRUCTOR_OVERLAP"
    MANUALLY_UNPLACED = "MANUALLY_UNPLACED"


def normalize_lesson_tag_name(value: str | None) -> str:
    """Trim surrounding whitespace and uppercase; empty input yields an empty string."""
    return (value or "").strip().upper()


def lesson_tag_identity_key(value: str | None) -> str:
    """Case-insensitive identity used for Lesson Tag uniqueness."""
    return normalize_lesson_tag_name(value).casefold()


def _new_id() -> str:
    return str(uuid4())


def _in_list(column: str, values: type[Enum]) -> str:
    return f"{column} IN ({', '.join(repr(member.value) for member in values)})"


_WINDOW_CHECK = "start_minute >= 0 AND end_minute <= 1440 AND end_minute > start_minute"


class LessonTag(ModuleBase):
    __tablename__ = "lesson_tags"
    __table_args__ = (
        UniqueConstraint("normalized_name", name="uq_lesson_tags_normalized_name"),
        CheckConstraint("length(name) > 0", name="ck_lesson_tags_name"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    name: Mapped[str] = mapped_column(String(200), nullable=False)
    normalized_name: Mapped[str] = mapped_column(String(200), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class LessonInstructor(ModuleBase):
    __tablename__ = "lesson_instructors"
    __table_args__ = (
        CheckConstraint("length(name) > 0", name="ck_lesson_instructors_name"),
        CheckConstraint("lesson_length_minutes > 0", name="ck_lesson_instructors_length"),
        CheckConstraint("capacity >= 1", name="ck_lesson_instructors_capacity"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    name: Mapped[str] = mapped_column(String(200), nullable=False)
    email: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    room: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    lesson_length_minutes: Mapped[int] = mapped_column(
        Integer, nullable=False, default=DEFAULT_LESSON_LENGTH_MINUTES,
        server_default=str(DEFAULT_LESSON_LENGTH_MINUTES),
    )
    capacity: Mapped[int] = mapped_column(
        Integer, nullable=False, default=DEFAULT_INSTRUCTOR_CAPACITY,
        server_default=str(DEFAULT_INSTRUCTOR_CAPACITY),
    )
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)

    tags: Mapped[list[LessonTag]] = relationship(
        secondary="lesson_instructor_tags", order_by="LessonTag.normalized_name", viewonly=True
    )


class LessonInstructorTag(ModuleBase):
    __tablename__ = "lesson_instructor_tags"

    instructor_id: Mapped[str] = mapped_column(
        ForeignKey("lesson_instructors.id", ondelete="CASCADE"), primary_key=True
    )
    tag_id: Mapped[str] = mapped_column(
        ForeignKey("lesson_tags.id", ondelete="CASCADE"), primary_key=True, index=True
    )


class LessonInstructorAvailabilityWindow(ModuleBase):
    """One recurring weekly Availability Window; absent windows are not Unavailable."""

    __tablename__ = "lesson_instructor_availability_windows"
    __table_args__ = (
        CheckConstraint(_WINDOW_CHECK, name="ck_lesson_instructor_availability_window"),
        CheckConstraint(
            "status IN ('Available', 'Tentative', 'Unavailable')",
            name="ck_lesson_instructor_availability_status",
        ),
        CheckConstraint("source IN ('manual', 'import')", name="ck_lesson_instructor_availability_source"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    instructor_id: Mapped[str] = mapped_column(
        ForeignKey("lesson_instructors.id", ondelete="CASCADE"), nullable=False, index=True
    )
    day: Mapped[str] = mapped_column(String(20), nullable=False)
    start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    end_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    status: Mapped[str] = mapped_column(String(20), nullable=False)
    source: Mapped[str] = mapped_column(String(20), nullable=False, default="manual", server_default="manual")
    source_row: Mapped[int | None] = mapped_column(Integer, nullable=True)


class LessonStudent(ModuleBase):
    __tablename__ = "lesson_students"
    __table_args__ = (
        UniqueConstraint("student_id", name="uq_lesson_students_student_id"),
        CheckConstraint("length(student_id) > 0", name="ck_lesson_students_student_id"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    student_id: Mapped[str] = mapped_column(String(100), nullable=False)
    name: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    email: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class LessonStudentAvailabilityWindow(ModuleBase):
    """One recurring weekly Availability Window; absent windows are not Unavailable."""

    __tablename__ = "lesson_student_availability_windows"
    __table_args__ = (
        CheckConstraint(_WINDOW_CHECK, name="ck_lesson_student_availability_window"),
        CheckConstraint(
            "status IN ('Available', 'Tentative', 'Unavailable')",
            name="ck_lesson_student_availability_status",
        ),
        CheckConstraint("source IN ('manual', 'import')", name="ck_lesson_student_availability_source"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    lesson_student_id: Mapped[str] = mapped_column(
        ForeignKey("lesson_students.id", ondelete="CASCADE"), nullable=False, index=True
    )
    day: Mapped[str] = mapped_column(String(20), nullable=False)
    start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    end_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    status: Mapped[str] = mapped_column(String(20), nullable=False)
    source: Mapped[str] = mapped_column(String(20), nullable=False, default="manual", server_default="manual")
    source_row: Mapped[int | None] = mapped_column(Integer, nullable=True)


class LessonStudentTagState(ModuleBase):
    __tablename__ = "lesson_student_tag_states"
    __table_args__ = (
        CheckConstraint(_in_list("state", LessonTagState), name="ck_lesson_tag_state"),
    )

    student_id: Mapped[str] = mapped_column(
        ForeignKey("lesson_students.id", ondelete="CASCADE"), primary_key=True
    )
    tag_id: Mapped[str] = mapped_column(
        ForeignKey("lesson_tags.id", ondelete="CASCADE"), primary_key=True, index=True
    )
    state: Mapped[str] = mapped_column(String(20), nullable=False, default=LessonTagState.NEUTRAL.value)


class LessonPlacement(ModuleBase):
    """A student's single Lesson placement decision.

    A manuallyUnplaced decision has no instructor or start time.
    """

    __tablename__ = "lesson_placements"
    __table_args__ = (
        UniqueConstraint("lesson_student_id", name="uq_lesson_placements_student"),
        CheckConstraint(_in_list("placement_type", LessonPlacementType), name="ck_lesson_placement_type"),
        CheckConstraint(
            "(placement_type = 'manuallyUnplaced' "
            "AND lesson_instructor_id IS NULL AND start_minute IS NULL AND end_minute IS NULL) "
            "OR (placement_type <> 'manuallyUnplaced' "
            "AND lesson_instructor_id IS NOT NULL AND start_minute IS NOT NULL AND end_minute IS NOT NULL "
            "AND start_minute >= 0 AND start_minute < end_minute AND end_minute <= 1440)",
            name="ck_lesson_placement_shape",
        ),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    lesson_student_id: Mapped[str] = mapped_column(String(36), nullable=False)
    lesson_instructor_id: Mapped[str | None] = mapped_column(
        String(36), nullable=True, index=True
    )
    day: Mapped[str] = mapped_column(String(20), nullable=False, default="Monday", server_default="Monday")
    start_minute: Mapped[int | None] = mapped_column(Integer, nullable=True)
    end_minute: Mapped[int | None] = mapped_column(Integer, nullable=True)
    placement_type: Mapped[str] = mapped_column(String(20), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class LessonModuleState(ModuleBase):
    __tablename__ = "lesson_module_states"

    session_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    schedule_revision: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    source_data_revision: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    is_schedule_stale: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")
    modified_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class LessonScheduleResult(ModuleBase):
    """Persisted summary and source provenance for the current automatic schedule."""

    __tablename__ = "lesson_schedule_results"

    session_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    schedule_revision: Mapped[int] = mapped_column(Integer, nullable=False)
    source_data_revision: Mapped[int] = mapped_column(Integer, nullable=False)
    students_scheduled: Mapped[int] = mapped_column(Integer, nullable=False)
    students_unscheduled: Mapped[int] = mapped_column(Integer, nullable=False)
    instructor_utilization: Mapped[float] = mapped_column(Float, nullable=False)
    preferred_tag_matches: Mapped[int] = mapped_column(Integer, nullable=False)
    tentative_minutes: Mapped[int] = mapped_column(Integer, nullable=False)
    gap_cost_minutes: Mapped[int] = mapped_column(Integer, nullable=False)
    modified_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class LessonSchedulePlacement(ModuleBase):
    """Current assigned-placement snapshot, retained when roster identities change."""

    __tablename__ = "lesson_schedule_placements"

    student_id: Mapped[str] = mapped_column(String(36), primary_key=True)
    instructor_id: Mapped[str] = mapped_column(String(36), nullable=False, index=True)
    placement_type: Mapped[str] = mapped_column(
        String(20), nullable=False, default=LessonPlacementType.AUTOMATIC.value,
        server_default=LessonPlacementType.AUTOMATIC.value,
    )
    student_name: Mapped[str] = mapped_column(String(200), nullable=False, default="", server_default="")
    instructor_name: Mapped[str] = mapped_column(String(200), nullable=False, default="", server_default="")
    day: Mapped[str] = mapped_column(String(20), nullable=False)
    start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    end_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    preferred_tag_match_count: Mapped[int] = mapped_column(Integer, nullable=False)
    tentative_minutes: Mapped[int] = mapped_column(Integer, nullable=False)


class LessonUnplacedStudent(ModuleBase):
    """Unplaced student decision retained with the current schedule."""

    __tablename__ = "lesson_unplaced_students"

    student_id: Mapped[str] = mapped_column(String(36), primary_key=True)
    student_name: Mapped[str] = mapped_column(String(200), nullable=False, default="", server_default="")
    reason: Mapped[str] = mapped_column(String(40), nullable=False)
    placement_type: Mapped[str] = mapped_column(
        String(20), nullable=False, default=LessonPlacementType.AUTOMATIC.value,
        server_default=LessonPlacementType.AUTOMATIC.value,
    )
