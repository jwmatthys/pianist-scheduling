"""Clinical Placement module-owned persistence models.

These tables hold only Clinical data. They have no foreign keys to, and no
dependency on, Accompanist or Jury tables or shared person identities.
Internal identifiers are application-generated UUID strings and are never
shown to users.
"""

from datetime import datetime
from enum import Enum
from uuid import uuid4

from sqlalchemy import (
    Boolean,
    CheckConstraint,
    DateTime,
    ForeignKey,
    Integer,
    String,
    UniqueConstraint,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .module_models import ModuleBase

CLINICAL_MODULE_ID = "clinical-placements"


class TagState(str, Enum):
    NEUTRAL = "neutral"
    REQUIRED = "required"
    PREFERRED = "preferred"
    RESTRICTED = "restricted"


class PlacementType(str, Enum):
    AUTOMATIC = "automatic"
    MANUAL = "manual"
    MANUALLY_UNPLACED = "manuallyUnplaced"


def normalize_tag_name(value: str | None) -> str:
    """Trim surrounding whitespace and uppercase; empty input yields an empty string."""
    return (value or "").strip().upper()


def tag_identity_key(value: str | None) -> str:
    """Case-insensitive identity used for Clinical Tag uniqueness."""
    return normalize_tag_name(value).casefold()


def _new_id() -> str:
    return str(uuid4())


def _in_list(column: str, values: type[Enum]) -> str:
    return f"{column} IN ({', '.join(repr(member.value) for member in values)})"


class ClinicalTag(ModuleBase):
    __tablename__ = "clinical_tags"
    __table_args__ = (
        UniqueConstraint("normalized_name", name="uq_clinical_tags_normalized_name"),
        CheckConstraint("length(name) > 0", name="ck_clinical_tags_name"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    name: Mapped[str] = mapped_column(String(200), nullable=False)
    normalized_name: Mapped[str] = mapped_column(String(200), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class PlacementOpportunity(ModuleBase):
    __tablename__ = "clinical_placement_opportunities"
    __table_args__ = (
        CheckConstraint(
            "start_minute >= 0 AND end_minute <= 1440 AND end_minute > start_minute",
            name="ck_clinical_opportunity_window",
        ),
        CheckConstraint("session_length_minutes > 0", name="ck_clinical_opportunity_session_length"),
        CheckConstraint("travel_minutes >= 0", name="ck_clinical_opportunity_travel"),
        CheckConstraint("capacity >= 1", name="ck_clinical_opportunity_capacity"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    site_name: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    site_address: Mapped[str] = mapped_column(String(400), nullable=False, default="")
    contact_name: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    contact_email: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    contact_phone: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    day_of_week: Mapped[str] = mapped_column(String(20), nullable=False)
    start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    end_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    session_length_minutes: Mapped[int] = mapped_column(Integer, nullable=False)
    travel_minutes: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    capacity: Mapped[int] = mapped_column(Integer, nullable=False, default=1, server_default="1")
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)

    tags: Mapped[list[ClinicalTag]] = relationship(
        secondary="clinical_opportunity_tags", order_by="ClinicalTag.normalized_name", viewonly=True
    )


class PlacementOpportunityTag(ModuleBase):
    __tablename__ = "clinical_opportunity_tags"

    opportunity_id: Mapped[str] = mapped_column(
        ForeignKey("clinical_placement_opportunities.id", ondelete="CASCADE"), primary_key=True
    )
    tag_id: Mapped[str] = mapped_column(
        ForeignKey("clinical_tags.id", ondelete="CASCADE"), primary_key=True, index=True
    )


class ClinicalStudent(ModuleBase):
    __tablename__ = "clinical_students"
    __table_args__ = (
        UniqueConstraint("student_id", name="uq_clinical_students_student_id"),
        CheckConstraint("length(student_id) > 0", name="ck_clinical_students_student_id"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    student_id: Mapped[str] = mapped_column(String(100), nullable=False)
    name: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    email: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)

    availability: Mapped[list["ClinicalStudentAvailabilityWindow"]] = relationship(
        order_by="(ClinicalStudentAvailabilityWindow.day, ClinicalStudentAvailabilityWindow.start_minute)",
        viewonly=True,
    )


class ClinicalStudentAvailabilityWindow(ModuleBase):
    """One recurring weekly Availability Window; absent windows are not Unavailable."""

    __tablename__ = "clinical_student_availability_windows"
    __table_args__ = (
        CheckConstraint(
            "start_minute >= 0 AND end_minute <= 1440 AND end_minute > start_minute",
            name="ck_clinical_availability_window",
        ),
        CheckConstraint(
            "status IN ('Available', 'Tentative', 'Unavailable')",
            name="ck_clinical_availability_status",
        ),
        CheckConstraint("source IN ('manual', 'import')", name="ck_clinical_availability_source"),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    clinical_student_id: Mapped[str] = mapped_column(
        ForeignKey("clinical_students.id", ondelete="CASCADE"), nullable=False, index=True
    )
    day: Mapped[str] = mapped_column(String(20), nullable=False)
    start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    end_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    status: Mapped[str] = mapped_column(String(20), nullable=False)
    source: Mapped[str] = mapped_column(String(20), nullable=False, default="manual", server_default="manual")
    source_row: Mapped[int | None] = mapped_column(Integer, nullable=True)


class ClinicalStudentTagState(ModuleBase):
    __tablename__ = "clinical_student_tag_states"
    __table_args__ = (
        CheckConstraint(_in_list("state", TagState), name="ck_clinical_tag_state"),
    )

    student_id: Mapped[str] = mapped_column(
        ForeignKey("clinical_students.id", ondelete="CASCADE"), primary_key=True
    )
    tag_id: Mapped[str] = mapped_column(
        ForeignKey("clinical_tags.id", ondelete="CASCADE"), primary_key=True, index=True
    )
    state: Mapped[str] = mapped_column(String(20), nullable=False, default=TagState.NEUTRAL.value)


class ClinicalPlacement(ModuleBase):
    """A student's single Clinical placement decision.

    A manuallyUnplaced decision has no opportunity or start time.
    """

    __tablename__ = "clinical_placements"
    __table_args__ = (
        UniqueConstraint("clinical_student_id", name="uq_clinical_placements_student"),
        CheckConstraint(_in_list("placement_type", PlacementType), name="ck_clinical_placement_type"),
        CheckConstraint(
            "(placement_type = 'manuallyUnplaced' "
            "AND placement_opportunity_id IS NULL AND start_minute IS NULL) "
            "OR (placement_type <> 'manuallyUnplaced' "
            "AND placement_opportunity_id IS NOT NULL AND start_minute IS NOT NULL "
            "AND start_minute >= 0 AND start_minute < 1440)",
            name="ck_clinical_placement_shape",
        ),
    )

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=_new_id)
    clinical_student_id: Mapped[str] = mapped_column(
        ForeignKey("clinical_students.id", ondelete="CASCADE"), nullable=False
    )
    placement_opportunity_id: Mapped[str | None] = mapped_column(
        ForeignKey("clinical_placement_opportunities.id", ondelete="CASCADE"), nullable=True, index=True
    )
    start_minute: Mapped[int | None] = mapped_column(Integer, nullable=True)
    placement_type: Mapped[str] = mapped_column(String(20), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class ClinicalModuleState(ModuleBase):
    __tablename__ = "clinical_module_states"

    session_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    schedule_revision: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    source_data_revision: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    is_schedule_stale: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")
    modified_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
