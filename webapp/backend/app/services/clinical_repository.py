"""Clinical Placement repository: module-owned persistence operations.

Foreign-key cascades are not enforced by the active SQLite connection, so every
dependent-row removal is performed explicitly here. The caller owns the
transaction (operations flush; they do not commit).
"""

from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Sequence

from sqlalchemy import func
from sqlalchemy.orm import Session

from .. import models
from ..clinical_models import (
    ClinicalModuleState,
    ClinicalPlacement,
    ClinicalStudent,
    ClinicalStudentAvailabilityWindow,
    ClinicalStudentTagState,
    ClinicalTag,
    PlacementOpportunity,
    PlacementOpportunityTag,
    PlacementType,
    TagState,
    normalize_tag_name,
    tag_identity_key,
)
from .availability import AvailabilityWindow, normalize_day


class ClinicalValidationError(ValueError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


@dataclass(frozen=True)
class PlacementOpportunityInput:
    day_of_week: str
    start_minute: int
    end_minute: int
    session_length_minutes: int
    site_name: str = ""
    site_address: str = ""
    contact_name: str = ""
    contact_email: str = ""
    contact_phone: str = ""
    travel_minutes: int = 0
    capacity: int = 1


@dataclass(frozen=True)
class ClinicalTagSummary:
    tag: ClinicalTag
    opportunity_count: int
    student_count: int

    @property
    def is_unused(self) -> bool:
        """Unused means no current Placement Opportunity references the tag."""
        return self.opportunity_count == 0


def _now() -> datetime:
    return datetime.utcnow()


@dataclass
class ClinicalReplaceResult:
    opportunities_removed: int
    placements_removed: int
    opportunities_created: int
    tags_created: int


class ClinicalRepository:
    def __init__(self, db: Session):
        self.db = db

    # --- module state -------------------------------------------------

    def _session_uuid(self) -> str:
        rows = self.db.query(models.SchedulingSession.session_uuid).all()
        if len(rows) != 1:
            raise ValueError("The active database must contain exactly one Scheduling Session.")
        return rows[0][0]

    def get_state(self) -> ClinicalModuleState:
        session_uuid = self._session_uuid()
        state = self.db.get(ClinicalModuleState, session_uuid)
        if state is None:
            state = ClinicalModuleState(
                session_uuid=session_uuid,
                schedule_revision=0,
                source_data_revision=0,
                is_schedule_stale=False,
                modified_at=_now(),
            )
            self.db.add(state)
            self.db.flush()
        return state

    def record_source_change(self) -> ClinicalModuleState:
        """Bump the source revision; flag an existing schedule stale without altering it."""
        state = self.get_state()
        state.source_data_revision += 1
        if self.db.query(ClinicalPlacement.id).first() is not None:
            state.is_schedule_stale = True
        state.modified_at = _now()
        self.db.flush()
        return state

    # --- tags ---------------------------------------------------------

    def list_tags(self) -> list[ClinicalTag]:
        return self.db.query(ClinicalTag).order_by(ClinicalTag.normalized_name).all()

    def find_tag(self, name: str) -> ClinicalTag | None:
        key = tag_identity_key(name)
        if not key:
            return None
        return self.db.query(ClinicalTag).filter(ClinicalTag.normalized_name == key).one_or_none()

    def _clean_tag_name(self, name: str) -> str:
        cleaned = normalize_tag_name(name)
        if not cleaned:
            raise ClinicalValidationError("INVALID_TAG", "A tag name is required.")
        return cleaned

    def _insert_tag(self, name: str) -> ClinicalTag:
        cleaned = self._clean_tag_name(name)
        if self.find_tag(cleaned) is not None:
            raise ClinicalValidationError("DUPLICATE_TAG", f"The tag '{cleaned}' already exists.")
        tag = ClinicalTag(name=cleaned, normalized_name=tag_identity_key(cleaned))
        self.db.add(tag)
        self.db.flush()
        return tag

    def create_tag(self, name: str) -> ClinicalTag:
        tag = self._insert_tag(name)
        self.record_source_change()
        return tag

    def get_or_create_tag(self, name: str) -> ClinicalTag:
        return self.find_tag(self._clean_tag_name(name)) or self.create_tag(name)

    def rename_tag(self, tag_id: str, name: str) -> ClinicalTag:
        tag = self._require(ClinicalTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
        cleaned = self._clean_tag_name(name)
        existing = self.find_tag(cleaned)
        if existing is not None and existing.id != tag.id:
            raise ClinicalValidationError("DUPLICATE_TAG", f"The tag '{cleaned}' already exists.")
        tag.name = cleaned
        tag.normalized_name = tag_identity_key(cleaned)
        tag.updated_at = _now()
        self.db.flush()
        self.record_source_change()
        return tag

    def is_tag_used_by_opportunity(self, tag_id: str) -> bool:
        return (
            self.db.query(PlacementOpportunityTag)
            .filter(PlacementOpportunityTag.tag_id == tag_id)
            .first()
            is not None
        )

    def count_students_with_tag_selection(self, tag_id: str) -> int:
        """Students whose Required/Preferred/Restricted selection would be lost on delete."""
        return (
            self.db.query(func.count(func.distinct(ClinicalStudentTagState.student_id)))
            .filter(
                ClinicalStudentTagState.tag_id == tag_id,
                ClinicalStudentTagState.state != TagState.NEUTRAL.value,
            )
            .scalar()
            or 0
        )

    def delete_tag(self, tag_id: str) -> int:
        """Delete a tag and its Clinical associations; returns affected student count."""
        tag = self._require(ClinicalTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
        affected_students = self.count_students_with_tag_selection(tag.id)
        self.db.query(PlacementOpportunityTag).filter(PlacementOpportunityTag.tag_id == tag.id).delete()
        self.db.query(ClinicalStudentTagState).filter(ClinicalStudentTagState.tag_id == tag.id).delete()
        self.db.delete(tag)
        self.db.flush()
        self.record_source_change()
        return affected_students

    def tag_summaries(self) -> list["ClinicalTagSummary"]:
        """Every registered tag with current-use information for the Tag Manager."""
        opportunity_counts = dict(
            self.db.query(PlacementOpportunityTag.tag_id, func.count(PlacementOpportunityTag.opportunity_id))
            .group_by(PlacementOpportunityTag.tag_id)
            .all()
        )
        student_counts = dict(
            self.db.query(
                ClinicalStudentTagState.tag_id, func.count(func.distinct(ClinicalStudentTagState.student_id))
            )
            .filter(ClinicalStudentTagState.state != TagState.NEUTRAL.value)
            .group_by(ClinicalStudentTagState.tag_id)
            .all()
        )
        return [
            ClinicalTagSummary(
                tag=tag,
                opportunity_count=opportunity_counts.get(tag.id, 0),
                student_count=student_counts.get(tag.id, 0),
            )
            for tag in self.list_tags()
        ]

    # --- placement opportunities -------------------------------------

    def _validated_opportunity_values(self, values: PlacementOpportunityInput) -> dict:
        if not values.site_name.strip():
            raise ClinicalValidationError("INVALID_SITE_NAME", "Site name is required.")
        day = normalize_day(values.day_of_week)
        if day is None:
            raise ClinicalValidationError("INVALID_DAY", "A valid day of the week is required.")
        if not 0 <= values.start_minute < values.end_minute <= 24 * 60:
            raise ClinicalValidationError("INVALID_WINDOW", "The end time must be after the start time.")
        if values.session_length_minutes <= 0:
            raise ClinicalValidationError("INVALID_SESSION_LENGTH", "Session length must be positive.")
        if values.travel_minutes < 0:
            raise ClinicalValidationError("INVALID_TRAVEL", "Travel time cannot be negative.")
        if values.capacity < 1:
            raise ClinicalValidationError("INVALID_CAPACITY", "Capacity must be at least 1.")
        return {
            "site_name": values.site_name.strip(),
            "site_address": values.site_address.strip(),
            "contact_name": values.contact_name.strip(),
            "contact_email": values.contact_email.strip(),
            "contact_phone": values.contact_phone.strip(),
            "day_of_week": day,
            "start_minute": values.start_minute,
            "end_minute": values.end_minute,
            "session_length_minutes": values.session_length_minutes,
            "travel_minutes": values.travel_minutes,
            "capacity": values.capacity,
        }

    def _set_opportunity_tags(self, opportunity: PlacementOpportunity, tag_names: Iterable[str]) -> None:
        tags: dict[str, ClinicalTag] = {}
        for name in tag_names:
            if not normalize_tag_name(name):
                continue
            tag = self.find_tag(name) or self._insert_tag(name)
            tags[tag.id] = tag
        self.db.query(PlacementOpportunityTag).filter(
            PlacementOpportunityTag.opportunity_id == opportunity.id
        ).delete()
        for tag_id in tags:
            self.db.add(PlacementOpportunityTag(opportunity_id=opportunity.id, tag_id=tag_id))
        self.db.flush()
        self.db.expire(opportunity, ["tags"])

    def _set_opportunity_tag_ids(self, opportunity: PlacementOpportunity, tag_ids: Iterable[str]) -> None:
        unique_ids = list(dict.fromkeys(tag_ids))
        found = {
            tag.id for tag in self.db.query(ClinicalTag).filter(ClinicalTag.id.in_(unique_ids)).all()
        } if unique_ids else set()
        if found != set(unique_ids):
            raise ClinicalValidationError("TAG_NOT_FOUND", "A selected tag no longer exists.")
        self.db.query(PlacementOpportunityTag).filter(
            PlacementOpportunityTag.opportunity_id == opportunity.id
        ).delete()
        for tag_id in unique_ids:
            self.db.add(PlacementOpportunityTag(opportunity_id=opportunity.id, tag_id=tag_id))
        self.db.flush()
        self.db.expire(opportunity, ["tags"])

    def create_opportunity(
        self,
        values: PlacementOpportunityInput,
        tag_names: Iterable[str] = (),
        tag_ids: Iterable[str] | None = None,
    ) -> PlacementOpportunity:
        opportunity = PlacementOpportunity(**self._validated_opportunity_values(values))
        self.db.add(opportunity)
        self.db.flush()
        if tag_ids is not None:
            self._set_opportunity_tag_ids(opportunity, tag_ids)
        else:
            self._set_opportunity_tags(opportunity, tag_names)
        self.record_source_change()
        return opportunity

    def set_opportunity_tags(self, opportunity_id: str, tag_ids: Iterable[str]) -> PlacementOpportunity:
        """Replace one opportunity's tag associations; never creates or deletes tags."""
        opportunity = self._require(
            PlacementOpportunity, opportunity_id, "OPPORTUNITY_NOT_FOUND", "The Placement Opportunity no longer exists."
        )
        self._set_opportunity_tag_ids(opportunity, tag_ids)
        opportunity.updated_at = _now()
        self.db.flush()
        self.record_source_change()
        return opportunity

    def list_opportunities(self) -> list[PlacementOpportunity]:
        return self.db.query(PlacementOpportunity).order_by(
            PlacementOpportunity.site_name, PlacementOpportunity.created_at, PlacementOpportunity.id
        ).all()

    def update_opportunity(
        self,
        opportunity_id: str,
        values: PlacementOpportunityInput,
        tag_names: Iterable[str] | None = None,
        tag_ids: Iterable[str] | None = None,
    ) -> PlacementOpportunity:
        opportunity = self._require(
            PlacementOpportunity, opportunity_id, "OPPORTUNITY_NOT_FOUND", "The Placement Opportunity no longer exists."
        )
        for key, value in self._validated_opportunity_values(values).items():
            setattr(opportunity, key, value)
        opportunity.updated_at = _now()
        if tag_ids is not None:
            self._set_opportunity_tag_ids(opportunity, tag_ids)
        elif tag_names is not None:
            self._set_opportunity_tags(opportunity, tag_names)
        self.db.flush()
        self.record_source_change()
        return opportunity

    def replace_opportunities(
        self, sites: Iterable[tuple[PlacementOpportunityInput, Iterable[str]]]
    ) -> "ClinicalReplaceResult":
        """Full replacement of Placement Opportunities (all-or-nothing within the caller's transaction).

        Clears every placement decision, reuses tags by normalized name, creates missing tags,
        never deletes tags, and leaves students, availability, and tag states untouched.
        """
        prepared = [(self._validated_opportunity_values(values), list(names)) for values, names in sites]
        opportunities_removed = self.db.query(func.count(PlacementOpportunity.id)).scalar() or 0
        placements_removed = self.clear_all_placements()
        self.db.query(PlacementOpportunityTag).delete()
        self.db.query(PlacementOpportunity).delete()
        self.db.flush()

        tags: dict[str, ClinicalTag] = {}
        tags_created = 0
        for values, names in prepared:
            opportunity = PlacementOpportunity(**values)
            self.db.add(opportunity)
            self.db.flush()
            linked: set[str] = set()
            for name in names:
                key = tag_identity_key(name)
                if not key or key in linked:
                    continue
                tag = tags.get(key) or self.find_tag(name)
                if tag is None:
                    tag = self._insert_tag(name)
                    tags_created += 1
                tags[key] = tag
                linked.add(key)
                self.db.add(PlacementOpportunityTag(opportunity_id=opportunity.id, tag_id=tag.id))
            self.db.flush()
        self.record_source_change()
        return ClinicalReplaceResult(
            opportunities_removed=opportunities_removed,
            placements_removed=placements_removed,
            opportunities_created=len(prepared),
            tags_created=tags_created,
        )

    def count_placements_for_opportunity(self, opportunity_id: str) -> int:
        return (
            self.db.query(func.count(ClinicalPlacement.id))
            .filter(ClinicalPlacement.placement_opportunity_id == opportunity_id)
            .scalar()
            or 0
        )

    def delete_opportunity(self, opportunity_id: str) -> int:
        """Delete an opportunity and only its placements; returns removed placement count."""
        opportunity = self._require(
            PlacementOpportunity, opportunity_id, "OPPORTUNITY_NOT_FOUND", "The Placement Opportunity no longer exists."
        )
        removed = self.db.query(ClinicalPlacement).filter(
            ClinicalPlacement.placement_opportunity_id == opportunity.id
        ).delete()
        self.db.query(PlacementOpportunityTag).filter(
            PlacementOpportunityTag.opportunity_id == opportunity.id
        ).delete()
        self.db.delete(opportunity)
        self.db.flush()
        self.record_source_change()
        return removed

    # --- students -----------------------------------------------------

    def create_student(
        self,
        student_id: str,
        name: str = "",
        email: str = "",
        availability: Sequence[AvailabilityWindow] = (),
    ) -> ClinicalStudent:
        cleaned_id = (student_id or "").strip()
        if not cleaned_id:
            raise ClinicalValidationError("INVALID_STUDENT_ID", "Student ID is required.")
        duplicate = (
            self.db.query(ClinicalStudent.id)
            .filter(func.lower(ClinicalStudent.student_id) == cleaned_id.lower())
            .first()
        )
        if duplicate is not None:
            raise ClinicalValidationError("DUPLICATE_STUDENT_ID", "A student with that Student ID already exists.")
        student = ClinicalStudent(student_id=cleaned_id, name=name.strip(), email=email.strip())
        self.db.add(student)
        self.db.flush()
        self.replace_student_availability(student.id, availability, record_change=False)
        self.record_source_change()
        return student

    def delete_student(self, clinical_student_id: str) -> None:
        student = self._require(
            ClinicalStudent, clinical_student_id, "STUDENT_NOT_FOUND", "The student no longer exists."
        )
        self.db.query(ClinicalPlacement).filter(ClinicalPlacement.clinical_student_id == student.id).delete()
        self.db.query(ClinicalStudentTagState).filter(ClinicalStudentTagState.student_id == student.id).delete()
        self.db.query(ClinicalStudentAvailabilityWindow).filter(
            ClinicalStudentAvailabilityWindow.clinical_student_id == student.id
        ).delete()
        self.db.delete(student)
        self.db.flush()
        self.record_source_change()

    def replace_student_availability(
        self,
        clinical_student_id: str,
        windows: Sequence[AvailabilityWindow],
        *,
        record_change: bool = True,
    ) -> None:
        student = self._require(
            ClinicalStudent, clinical_student_id, "STUDENT_NOT_FOUND", "The student no longer exists."
        )
        for window in windows:
            if normalize_day(window.day) != window.day or not 0 <= window.start_minute < window.end_minute <= 24 * 60:
                raise ClinicalValidationError("INVALID_AVAILABILITY_WINDOW", "An Availability Window is invalid.")
        self.db.query(ClinicalStudentAvailabilityWindow).filter(
            ClinicalStudentAvailabilityWindow.clinical_student_id == clinical_student_id
        ).delete()
        for window in windows:
            self.db.add(ClinicalStudentAvailabilityWindow(
                clinical_student_id=clinical_student_id,
                day=window.day,
                start_minute=window.start_minute,
                end_minute=window.end_minute,
                status=window.status,
                source=window.source,
                source_row=window.source_row,
            ))
        self.db.flush()
        self.db.expire(student, ["availability"])
        if record_change:
            self.record_source_change()

    def student_availability(self, clinical_student_id: str) -> list[AvailabilityWindow]:
        rows = (
            self.db.query(ClinicalStudentAvailabilityWindow)
            .filter(ClinicalStudentAvailabilityWindow.clinical_student_id == clinical_student_id)
            .order_by(ClinicalStudentAvailabilityWindow.day, ClinicalStudentAvailabilityWindow.start_minute)
            .all()
        )
        return [
            AvailabilityWindow(
                day=row.day,
                start_minute=row.start_minute,
                end_minute=row.end_minute,
                status=row.status,
                source=row.source,
                source_row=row.source_row,
            )
            for row in rows
        ]

    def set_tag_state(self, clinical_student_id: str, tag_id: str, state: TagState) -> None:
        """Persist a student/tag state; Neutral is represented by the absence of a row."""
        self._require(ClinicalStudent, clinical_student_id, "STUDENT_NOT_FOUND", "The student no longer exists.")
        self._require(ClinicalTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
        state = TagState(state)
        row = self.db.get(ClinicalStudentTagState, (clinical_student_id, tag_id))
        if state is TagState.NEUTRAL:
            if row is not None:
                self.db.delete(row)
        elif row is None:
            self.db.add(ClinicalStudentTagState(student_id=clinical_student_id, tag_id=tag_id, state=state.value))
        else:
            row.state = state.value
        self.db.flush()
        self.record_source_change()

    def get_tag_state(self, clinical_student_id: str, tag_id: str) -> TagState:
        row = self.db.get(ClinicalStudentTagState, (clinical_student_id, tag_id))
        return TagState(row.state) if row is not None else TagState.NEUTRAL

    # --- placements ---------------------------------------------------

    def set_placement(
        self,
        clinical_student_id: str,
        placement_type: PlacementType,
        placement_opportunity_id: str | None = None,
        start_minute: int | None = None,
    ) -> ClinicalPlacement:
        """Create or replace the student's single placement decision."""
        self._require(ClinicalStudent, clinical_student_id, "STUDENT_NOT_FOUND", "The student no longer exists.")
        placement_type = PlacementType(placement_type)
        if placement_type is PlacementType.MANUALLY_UNPLACED:
            if placement_opportunity_id is not None or start_minute is not None:
                raise ClinicalValidationError(
                    "INVALID_PLACEMENT", "A manually unplaced decision has no opportunity or start time."
                )
        else:
            if placement_opportunity_id is None or start_minute is None:
                raise ClinicalValidationError(
                    "INVALID_PLACEMENT", "A placement requires an opportunity and a start time."
                )
            if not 0 <= start_minute < 24 * 60:
                raise ClinicalValidationError("INVALID_PLACEMENT", "The start time is not a valid time of day.")
            self._require(
                PlacementOpportunity,
                placement_opportunity_id,
                "OPPORTUNITY_NOT_FOUND",
                "The Placement Opportunity no longer exists.",
            )
        placement = (
            self.db.query(ClinicalPlacement)
            .filter(ClinicalPlacement.clinical_student_id == clinical_student_id)
            .one_or_none()
        )
        if placement is None:
            placement = ClinicalPlacement(clinical_student_id=clinical_student_id)
            self.db.add(placement)
        placement.placement_type = placement_type.value
        placement.placement_opportunity_id = placement_opportunity_id
        placement.start_minute = start_minute
        placement.updated_at = _now()
        self.db.flush()
        return placement

    def get_placement(self, clinical_student_id: str) -> ClinicalPlacement | None:
        return (
            self.db.query(ClinicalPlacement)
            .filter(ClinicalPlacement.clinical_student_id == clinical_student_id)
            .one_or_none()
        )

    def clear_all_placements(self) -> int:
        """Delete every placement decision and reset result/staleness state."""
        removed = self.db.query(ClinicalPlacement).delete()
        state = self.get_state()
        state.schedule_revision = 0
        state.is_schedule_stale = False
        state.modified_at = _now()
        self.db.flush()
        return removed

    # --- helpers ------------------------------------------------------

    def _require(self, model, primary_key: str, code: str, message: str):
        row = self.db.get(model, primary_key)
        if row is None:
            raise ClinicalValidationError(code, message)
        return row
