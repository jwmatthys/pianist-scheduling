"""Shared identity, module result, and Jury persistence models."""

import json
import re
from datetime import date, datetime
from uuid import uuid4

from sqlalchemy import (
    Boolean,
    CheckConstraint,
    Date,
    DateTime,
    ForeignKey,
    ForeignKeyConstraint,
    Integer,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column


class ModuleBase(DeclarativeBase):
    pass


def normalize_student_id(value: str | None) -> str:
    return (value or "").strip()


def normalize_student_name(value: str | None) -> str:
    return re.sub(r"\s+", " ", (value or "").strip()).casefold()


def normalized_names_json(names: set[str]) -> str:
    return json.dumps(sorted(names), ensure_ascii=True)


class PersonIdentity(ModuleBase):
    __tablename__ = "person_identities"

    person_uuid: Mapped[str] = mapped_column(String(36), primary_key=True, default=lambda: str(uuid4()))
    display_name: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class AccompanistStudentProfile(ModuleBase):
    __tablename__ = "accompanist_student_profiles"
    __table_args__ = (
        UniqueConstraint("session_uuid", "normalized_student_id", name="uq_accomp_student_id"),
    )

    person_uuid: Mapped[str] = mapped_column(ForeignKey("person_identities.person_uuid"), primary_key=True)
    session_uuid: Mapped[str] = mapped_column(String(36), nullable=False)
    student_id: Mapped[str | None] = mapped_column(String(100), nullable=True)
    normalized_student_id: Mapped[str | None] = mapped_column(String(100), nullable=True)
    seen_names_json: Mapped[str] = mapped_column(Text, nullable=False, default="[]")
    identity_conflict: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")


class AccompanistLessonIdentity(ModuleBase):
    __tablename__ = "accompanist_lesson_identities"
    __table_args__ = (UniqueConstraint("lesson_uuid", name="uq_accomp_lesson_uuid"),)

    lesson_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    lesson_uuid: Mapped[str] = mapped_column(String(36), nullable=False, default=lambda: str(uuid4()))
    student_person_uuid: Mapped[str] = mapped_column(
        ForeignKey("accompanist_student_profiles.person_uuid"), nullable=False
    )


class AccompanistLessonJuryRequirement(ModuleBase):
    __tablename__ = "accompanist_lesson_jury_requirements"

    lesson_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    jury_required: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")


class AccompanistPianistIdentity(ModuleBase):
    __tablename__ = "accompanist_pianist_identities"

    pianist_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    person_uuid: Mapped[str] = mapped_column(
        ForeignKey("person_identities.person_uuid"), nullable=False, unique=True
    )


class ModuleRevision(ModuleBase):
    __tablename__ = "module_revisions"

    session_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    module_id: Mapped[str] = mapped_column(String(80), primary_key=True)
    source_revision: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    current_result_uuid: Mapped[str | None] = mapped_column(String(36), nullable=True)
    modified_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class ModuleResult(ModuleBase):
    __tablename__ = "module_results"
    __table_args__ = (
        UniqueConstraint("session_uuid", "module_id", "result_version", name="uq_module_result_version"),
    )

    result_uuid: Mapped[str] = mapped_column(String(36), primary_key=True, default=lambda: str(uuid4()))
    session_uuid: Mapped[str] = mapped_column(String(36), nullable=False, index=True)
    module_id: Mapped[str] = mapped_column(String(80), nullable=False, index=True)
    contract_id: Mapped[str] = mapped_column(String(120), nullable=False)
    contract_version: Mapped[int] = mapped_column(Integer, nullable=False)
    payload_schema_version: Mapped[int] = mapped_column(Integer, nullable=False)
    result_version: Mapped[int] = mapped_column(Integer, nullable=False)
    source_revision: Mapped[int] = mapped_column(Integer, nullable=False)
    state: Mapped[str] = mapped_column(String(20), nullable=False, index=True)
    payload_json: Mapped[str] = mapped_column(Text, nullable=False)
    payload_sha256: Mapped[str] = mapped_column(String(64), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    finalized_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)


class ModuleResultDependency(ModuleBase):
    __tablename__ = "module_result_dependencies"
    __table_args__ = (
        ForeignKeyConstraint(["dependent_result_uuid"], ["module_results.result_uuid"]),
        ForeignKeyConstraint(["source_result_uuid"], ["module_results.result_uuid"]),
    )

    dependent_result_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    dependency_key: Mapped[str] = mapped_column(String(120), primary_key=True)
    source_result_uuid: Mapped[str] = mapped_column(String(36), nullable=False)
    source_session_uuid: Mapped[str] = mapped_column(String(36), nullable=False)
    source_contract_id: Mapped[str] = mapped_column(String(120), nullable=False)
    source_contract_version: Mapped[int] = mapped_column(Integer, nullable=False)
    source_revision: Mapped[int] = mapped_column(Integer, nullable=False)


class JuryConfiguration(ModuleBase):
    __tablename__ = "jury_configurations"

    session_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    jury_date: Mapped[date | None] = mapped_column(Date, nullable=True)  # Legacy v7 source used by v8 Panel-date backfill.
    input_revision: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    roster_source_result_uuid: Mapped[str | None] = mapped_column(String(36), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    modified_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class JuryPanel(ModuleBase):
    __tablename__ = "jury_panels"
    __table_args__ = (
        UniqueConstraint("session_uuid", "panel_name", name="uq_jury_panel_name"),
    )

    panel_uuid: Mapped[str] = mapped_column(String(36), primary_key=True, default=lambda: str(uuid4()))
    session_uuid: Mapped[str] = mapped_column(String(36), nullable=False, index=True)
    panel_name: Mapped[str] = mapped_column(String(200), nullable=False)
    room: Mapped[str] = mapped_column(String(200), nullable=False, default="")
    earliest_start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    preferred_start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    jury_length_minutes: Mapped[int] = mapped_column(Integer, nullable=False)
    break_needed: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")
    break_every_x_juries: Mapped[int | None] = mapped_column(Integer, nullable=True)
    break_length_minutes: Mapped[int | None] = mapped_column(Integer, nullable=True)
    meal_break: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")
    meal_start_minute: Mapped[int | None] = mapped_column(Integer, nullable=True)
    meal_end_minute: Mapped[int | None] = mapped_column(Integer, nullable=True)


class JuryPanelDate(ModuleBase):
    __tablename__ = "jury_panel_dates"

    panel_uuid: Mapped[str] = mapped_column(
        ForeignKey("jury_panels.panel_uuid", ondelete="CASCADE"), primary_key=True
    )
    session_uuid: Mapped[str] = mapped_column(String(36), nullable=False, index=True)
    jury_date: Mapped[date] = mapped_column(Date, nullable=False)


class JuryLessonEntry(ModuleBase):
    __tablename__ = "jury_lesson_entries"
    __table_args__ = (
        ForeignKeyConstraint(["student_person_uuid"], ["person_identities.person_uuid"]),
        ForeignKeyConstraint(["panel_uuid"], ["jury_panels.panel_uuid"]),
    )

    session_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    source_lesson_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    student_person_uuid: Mapped[str] = mapped_column(String(36), nullable=False)
    jury_required: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")
    panel_uuid: Mapped[str | None] = mapped_column(String(36), nullable=True)


class JuryPianistAvailabilityDeclaration(ModuleBase):
    __tablename__ = "jury_pianist_availability_declarations"

    session_uuid: Mapped[str] = mapped_column(String(36), primary_key=True)
    pianist_person_uuid: Mapped[str] = mapped_column(
        ForeignKey("person_identities.person_uuid"), primary_key=True
    )
    jury_date: Mapped[date] = mapped_column(Date, primary_key=True)
    is_complete: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")
    declared_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    modified_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)


class JuryPianistAvailableWindow(ModuleBase):
    __tablename__ = "jury_pianist_available_windows"
    __table_args__ = (
        ForeignKeyConstraint(
            ["session_uuid", "pianist_person_uuid", "jury_date"],
            [
                "jury_pianist_availability_declarations.session_uuid",
                "jury_pianist_availability_declarations.pianist_person_uuid",
                "jury_pianist_availability_declarations.jury_date",
            ],
            ondelete="CASCADE",
        ),
        CheckConstraint(
            "start_minute >= 0 AND end_minute <= 1440 AND end_minute > start_minute",
            name="ck_jury_available_window",
        ),
    )

    window_uuid: Mapped[str] = mapped_column(String(36), primary_key=True, default=lambda: str(uuid4()))
    session_uuid: Mapped[str] = mapped_column(String(36), nullable=False)
    pianist_person_uuid: Mapped[str] = mapped_column(String(36), nullable=False)
    jury_date: Mapped[date] = mapped_column(Date, nullable=False)
    start_minute: Mapped[int] = mapped_column(Integer, nullable=False)
    end_minute: Mapped[int] = mapped_column(Integer, nullable=False)