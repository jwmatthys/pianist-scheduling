"""Tutor Scheduling module-owned persistence models.

These tables hold only Tutor Scheduling data. They have no foreign keys to, and
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

from .tutor_domain import (  # noqa: F401  (re-exported for persistence callers)
    TutorPlacementState,
    TutorPlacementType,
    TutorTagState,
    TutorUnplacedReason,
)
from .module_models import ModuleBase

TUTOR_MODULE_ID = "tutor-scheduling"
DEFAULT_LESSON_LENGTH_MINUTES = 60
DEFAULT_TUTOR_CAPACITY = 1


def normalize_tutor_tag_name(value: str | None) -> str:
    """Trim surrounding whitespace and uppercase; empty input yields an empty string."""
    return (value or "").strip().upper()


def tutor_tag_identity_key(value: str | None) -> str:
    """Case-insensitive identity used for Tutor Tag uniqueness."""
    return normalize_tutor_tag_name(value).casefold()


def _new_id() -> str:
    return str(uuid4())


def _in_list(column: str, values: type[Enum]) -> str:
    return f"{column} IN ({', '.join(repr(member.value) for member in values)})"


_WINDOW_CHECK = "start_minute >= 0 AND end_minute <= 1440 AND end_minute > start_minute"


class TutorTag(ModuleBase):
    __tablename__ = "tutor_tags"
    __table_args__ = (
        UniqueConstraint("normalized_name", name="uq_tutor_tags_normalized_name"),
        CheckConstraint("length(name) > 0", name="ck_tutor_tags_name"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    name: Mapped[str] = mapped_column(String(200), nullable=False)
    normalized_name: Mapped[str] = mapped_column(String(200), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class TutorTutor(ModuleBase):
    __tablename__ = "tutor_tutors"
    __table_args__ = (
        CheckConstraint("length(name) > 0", name="ck_tutor_tutors_name"),
        CheckConstraint("lesson_length_minutes > 0", name="ck_tutor_tutors_length"),
        CheckConstraint("capacity >= 1", name="ck_tutor_tutors_capacity"),
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
        Integer, nullable=False, default=DEFAULT_TUTOR_CAPACITY,
        server_default=str(DEFAULT_TUTOR_CAPACITY),
    )
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)

    tags: Mapped[list[TutorTag]] = relationship(
        secondary="tutor_tutor_tags", order_by="TutorTag.normalized_name", viewonly=True
    )


class TutorTutorTag(ModuleBase):
    __tablename__ = "tutor_tutor_tags"

    tutor_id: Mapped[str] = mapped_column(
        ForeignKey("tutor_tutors.id", ondelete="CASCADE"), primary_key=True
    )
    tag_id: Mapped[str] = mapped_column(
        ForeignKey("tutor_tags.id", ondelete="CASCADE"), primary_key=True, index=True
    )


class TutorTutorAvailabilityWindow(ModuleBase):
    """One recurring weekly Availability Window; absent windows are not Unavailable."""

    __tablename__ = "tutor_tutor_availability_windows"
    __table_args__ = (
        CheckConstraint(_WINDOW_CHECK, name="ck_tutor_tutor_availability_window"),
        CheckConstraint(
            "status IN ('Available', 'Tentative', 'Unavailable')",
            name="ck_tutor_tutor_availability_status",
        ),
        CheckConstraint("source IN ('manual', 'import')", name="ck_tutor_tutor_availability_source"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    tutor_id: Mapped[str] = mapped_column(
        ForeignKey("tutor_tutors.id", ondelete="CASCADE"), nullable=False, index=True
    )
    day: Mapped[str] = mapped_column(String(20), nullable=False)
    start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    end_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    status: Mapped[str] = mapped_column(String(20), nullable=False)
    source: Mapped[str] = mapped_column(String(20), nullable=False, default="manual", server_default="manual")
    source_row: Mapped[int | None] = mapped_column(Integer, nullable=True)


class TutorStudent(ModuleBase):
    __tablename__ = "tutor_students"
    __table_args__ = (
        UniqueConstraint("student_id", name="uq_tutor_students_student_id"),
        CheckConstraint("length(student_id) > 0", name="ck_tutor_students_student_id"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    student_id: Mapped[str] = mapped_column(String(100), nullable=False)
    name: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    email: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class TutorStudentAvailabilityWindow(ModuleBase):
    """One recurring weekly Availability Window; absent windows are not Unavailable."""

    __tablename__ = "tutor_student_availability_windows"
    __table_args__ = (
        CheckConstraint(_WINDOW_CHECK, name="ck_tutor_student_availability_window"),
        CheckConstraint(
            "status IN ('Available', 'Tentative', 'Unavailable')",
            name="ck_tutor_student_availability_status",
        ),
        CheckConstraint("source IN ('manual', 'import')", name="ck_tutor_student_availability_source"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    tutor_student_id: Mapped[str] = mapped_column(
        ForeignKey("tutor_students.id", ondelete="CASCADE"), nullable=False, index=True
    )
    day: Mapped[str] = mapped_column(String(20), nullable=False)
    start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    end_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    status: Mapped[str] = mapped_column(String(20), nullable=False)
    source: Mapped[str] = mapped_column(String(20), nullable=False, default="manual", server_default="manual")
    source_row: Mapped[int | None] = mapped_column(Integer, nullable=True)


class TutorStudentTagState(ModuleBase):
    __tablename__ = "tutor_student_tag_states"
    __table_args__ = (
        CheckConstraint(_in_list("state", TutorTagState), name="ck_tutor_tag_state"),
    )

    student_id: Mapped[str] = mapped_column(
        ForeignKey("tutor_students.id", ondelete="CASCADE"), primary_key=True
    )
    tag_id: Mapped[str] = mapped_column(
        ForeignKey("tutor_tags.id", ondelete="CASCADE"), primary_key=True, index=True
    )
    state: Mapped[str] = mapped_column(String(20), nullable=False, default=TutorTagState.NEUTRAL.value)


class TutorPlacement(ModuleBase):
    """A student's single Tutor placement decision.

    A manuallyUnplaced decision has no tutor or start time.
    """

    __tablename__ = "tutor_placements"
    __table_args__ = (
        UniqueConstraint("tutor_student_id", name="uq_tutor_placements_student"),
        CheckConstraint(_in_list("placement_type", TutorPlacementType), name="ck_tutor_placement_type"),
        CheckConstraint(
            "(placement_type = 'manuallyUnplaced' "
            "AND tutor_tutor_id IS NULL AND start_minute IS NULL AND end_minute IS NULL) "
            "OR (placement_type <> 'manuallyUnplaced' "
            "AND tutor_tutor_id IS NOT NULL AND start_minute IS NOT NULL AND end_minute IS NOT NULL "
            "AND start_minute >= 0 AND start_minute < end_minute AND end_minute <= 1440)",
            name="ck_tutor_placement_shape",
        ),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    tutor_student_id: Mapped[str] = mapped_column(String(36), nullable=False)
    tutor_tutor_id: Mapped[str | None] = mapped_column(
        String(36), nullable=True, index=True
    )
    day: Mapped[str] = mapped_column(String(20), nullable=False, default="Monday", server_default="Monday")
    start_minute: Mapped[int | None] = mapped_column(Integer, nullable=True)
    end_minute: Mapped[int | None] = mapped_column(Integer, nullable=True)
    placement_type: Mapped[str] = mapped_column(String(20), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class TutorModuleState(ModuleBase):
    __tablename__ = "tutor_module_states"

    session_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    schedule_revision: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    source_data_revision: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    is_schedule_stale: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")
    modified_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class TutorScheduleResult(ModuleBase):
    """Persisted summary and source provenance for the current automatic schedule."""

    __tablename__ = "tutor_schedule_results"

    session_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    schedule_revision: Mapped[int] = mapped_column(Integer, nullable=False)
    source_data_revision: Mapped[int] = mapped_column(Integer, nullable=False)
    students_scheduled: Mapped[int] = mapped_column(Integer, nullable=False)
    students_unscheduled: Mapped[int] = mapped_column(Integer, nullable=False)
    tutor_utilization: Mapped[float] = mapped_column(Float, nullable=False)
    preferred_tag_matches: Mapped[int] = mapped_column(Integer, nullable=False)
    tentative_minutes: Mapped[int] = mapped_column(Integer, nullable=False)
    gap_cost_minutes: Mapped[int] = mapped_column(Integer, nullable=False)
    modified_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class TutorSchedulePlacement(ModuleBase):
    """Current assigned-placement snapshot, retained when roster identities change."""

    __tablename__ = "tutor_schedule_placements"

    student_id: Mapped[str] = mapped_column(String(36), primary_key=True)
    tutor_id: Mapped[str] = mapped_column(String(36), nullable=False, index=True)
    placement_type: Mapped[str] = mapped_column(
        String(20), nullable=False, default=TutorPlacementType.AUTOMATIC.value,
        server_default=TutorPlacementType.AUTOMATIC.value,
    )
    student_name: Mapped[str] = mapped_column(String(200), nullable=False, default="", server_default="")
    tutor_name: Mapped[str] = mapped_column(String(200), nullable=False, default="", server_default="")
    day: Mapped[str] = mapped_column(String(20), nullable=False)
    start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    end_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    preferred_tag_match_count: Mapped[int] = mapped_column(Integer, nullable=False)
    tentative_minutes: Mapped[int] = mapped_column(Integer, nullable=False)


class TutorUnplacedStudent(ModuleBase):
    """Unplaced student decision retained with the current schedule."""

    __tablename__ = "tutor_unplaced_students"

    student_id: Mapped[str] = mapped_column(String(36), primary_key=True)
    student_name: Mapped[str] = mapped_column(String(200), nullable=False, default="", server_default="")
    reason: Mapped[str] = mapped_column(String(40), nullable=False)
    placement_type: Mapped[str] = mapped_column(
        String(20), nullable=False, default=TutorPlacementType.AUTOMATIC.value,
        server_default=TutorPlacementType.AUTOMATIC.value,
    )
