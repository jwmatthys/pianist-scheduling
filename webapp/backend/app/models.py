"""SQLAlchemy ORM models for the pianist scheduling webapp.

Every top-level table carries an ``organization_id`` even though the MVP
only ever has one organization. This keeps the schema ready for real
multi-tenant auth later without any migration surgery.
"""

from datetime import datetime

from sqlalchemy import (
    Boolean,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .database import Base

DAYS_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


class Organization(Base):
    __tablename__ = "organizations"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(200), default="Default Organization")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    pianists: Mapped[list["Pianist"]] = relationship(back_populates="organization")
    lessons: Mapped[list["Lesson"]] = relationship(back_populates="organization")
    import_profiles: Mapped[list["ImportProfile"]] = relationship(back_populates="organization")


class Pianist(Base):
    __tablename__ = "pianists"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    organization_id: Mapped[int] = mapped_column(ForeignKey("organizations.id"), default=1, index=True)
    name: Mapped[str] = mapped_column(String(200))
    email: Mapped[str] = mapped_column(String(200), default="")
    max_hours_per_week: Mapped[float | None] = mapped_column(Float, nullable=True)

    organization: Mapped["Organization"] = relationship(back_populates="pianists")
    availability: Mapped[list["AvailabilitySlot"]] = relationship(
        back_populates="pianist", cascade="all, delete-orphan"
    )
    lessons: Mapped[list["Lesson"]] = relationship(back_populates="assigned_pianist")


class AvailabilitySlot(Base):
    """One 30-minute availability block for a pianist on a given weekday."""

    __tablename__ = "availability_slots"
    __table_args__ = (
        UniqueConstraint("pianist_id", "day", "slot_start_minute", name="uq_slot"),
    )

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    pianist_id: Mapped[int] = mapped_column(ForeignKey("pianists.id"), index=True)
    day: Mapped[str] = mapped_column(String(20))  # one of DAYS_ORDER
    slot_start_minute: Mapped[int] = mapped_column(Integer)  # minutes since midnight
    status: Mapped[str] = mapped_column(String(20))  # Available | Tentative | Unavailable

    pianist: Mapped["Pianist"] = relationship(back_populates="availability")


class Lesson(Base):
    __tablename__ = "lessons"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    organization_id: Mapped[int] = mapped_column(ForeignKey("organizations.id"), default=1, index=True)

    teacher: Mapped[str] = mapped_column(String(200), default="")
    student: Mapped[str] = mapped_column(String(200), default="")
    day: Mapped[str] = mapped_column(String(20))  # one of DAYS_ORDER
    start_minute: Mapped[int] = mapped_column(Integer)
    end_minute: Mapped[int] = mapped_column(Integer)
    room: Mapped[str] = mapped_column(String(200), default="")
    instrument: Mapped[str] = mapped_column(String(200), default="")
    required_pianist_name: Mapped[str] = mapped_column(String(200), default="")
    need_pianist: Mapped[bool] = mapped_column(Boolean, default=True)

    assigned_pianist_id: Mapped[int | None] = mapped_column(ForeignKey("pianists.id"), nullable=True)
    fit_quality: Mapped[str] = mapped_column(String(20), default="")
    notes: Mapped[str] = mapped_column(Text, default="")
    hours: Mapped[float] = mapped_column(Float, default=0.0)
    manually_edited: Mapped[bool] = mapped_column(Boolean, default=False)

    organization: Mapped["Organization"] = relationship(back_populates="lessons")
    assigned_pianist: Mapped["Pianist | None"] = relationship(back_populates="lessons")


class ImportProfile(Base):
    """A saved column-mapping preset so recurring imports skip re-mapping."""

    __tablename__ = "import_profiles"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    organization_id: Mapped[int] = mapped_column(ForeignKey("organizations.id"), default=1, index=True)
    name: Mapped[str] = mapped_column(String(200))
    mapping_json: Mapped[str] = mapped_column(Text)

    organization: Mapped["Organization"] = relationship(back_populates="import_profiles")
