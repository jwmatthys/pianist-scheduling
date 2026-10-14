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
from .availability import AvailabilityWindow, normalize_day, normalize_windows
from .clinical_candidates import generate_candidates
from .clinical_optimizer import optimize
from .clinical_placement_review import review_schedule
from .clinical_readiness import analyze_readiness
from .clinical_scheduling_models import (
    OpportunityInput,
    OptimizationResult,
    PlacementKind,
    PlacementRecord,
    PlacementVerification,
    PlacementWarning,
    PreservedPlacement,
    ScheduleEntry,
    SchedulingInput,
    ScheduleSummary,
    StudentInput,
    TagPreference,
)
from .clinical_schedule_summary import summarize_schedule
from .time_parsing import format_time_of_day


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


@dataclass(frozen=True)
class RosterImportRow:
    """One validated roster-import row. ``windows`` is None when availability is not imported."""

    student_id: str
    name: str
    email: str
    clinical_student_id: str | None
    windows: Sequence[AvailabilityWindow] | None


@dataclass(frozen=True)
class ClinicalRosterImportResult:
    students_created: int
    students_updated: int
    availability_updated: int
    windows_imported: int


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

    def clear_opportunities(self) -> tuple[int, int]:
        """Delete every Placement Opportunity and the placements that reference one.

        Manually unplaced decisions have no site, so they and the tag registry are kept.
        Returns (opportunities removed, placements removed) and records one source change.
        """
        removed_placements = self.db.query(ClinicalPlacement).filter(
            ClinicalPlacement.placement_opportunity_id.isnot(None)
        ).delete(synchronize_session=False)
        self.db.query(PlacementOpportunityTag).delete(synchronize_session=False)
        removed = self.db.query(PlacementOpportunity).delete(synchronize_session=False)
        self.db.expire_all()
        if removed:
            self.record_source_change()
        return removed, removed_placements

    def clear_students(self) -> tuple[int, int]:
        """Delete every Clinical student with their tag states, availability, and placements.

        Clinical Sites and the tag registry are kept. Returns (students removed, placements
        removed) and records one source change.
        """
        removed_placements = self.db.query(ClinicalPlacement).delete(synchronize_session=False)
        self.db.query(ClinicalStudentTagState).delete(synchronize_session=False)
        self.db.query(ClinicalStudentAvailabilityWindow).delete(synchronize_session=False)
        removed = self.db.query(ClinicalStudent).delete(synchronize_session=False)
        self.db.expire_all()
        if removed:
            self.record_source_change()
        return removed, removed_placements

    # --- students -----------------------------------------------------

    def _clean_student_values(
        self, student_id: str, name: str, email: str, *, require_name: bool, exclude_id: str | None = None
    ) -> tuple[str, str, str]:
        cleaned_id = (student_id or "").strip()
        cleaned_name = (name or "").strip()
        if not cleaned_id:
            raise ClinicalValidationError("INVALID_STUDENT_ID", "Student ID is required.")
        if require_name and not cleaned_name:
            raise ClinicalValidationError("INVALID_STUDENT_NAME", "Name is required.")
        query = self.db.query(ClinicalStudent.id).filter(func.lower(ClinicalStudent.student_id) == cleaned_id.lower())
        if exclude_id is not None:
            query = query.filter(ClinicalStudent.id != exclude_id)
        if query.first() is not None:
            raise ClinicalValidationError("DUPLICATE_STUDENT_ID", "A student with that Student ID already exists.")
        return cleaned_id, cleaned_name, (email or "").strip()

    def create_student(
        self,
        student_id: str,
        name: str = "",
        email: str = "",
        availability: Sequence[AvailabilityWindow] = (),
        *,
        require_name: bool = False,
    ) -> ClinicalStudent:
        cleaned_id, cleaned_name, cleaned_email = self._clean_student_values(
            student_id, name, email, require_name=require_name
        )
        student = ClinicalStudent(student_id=cleaned_id, name=cleaned_name, email=cleaned_email)
        self.db.add(student)
        self.db.flush()
        self.replace_student_availability(student.id, availability, record_change=False)
        self.record_source_change()
        return student

    def get_student(self, clinical_student_id: str) -> ClinicalStudent:
        return self._require(
            ClinicalStudent, clinical_student_id, "STUDENT_NOT_FOUND", "The student no longer exists."
        )

    def list_students(self) -> list[ClinicalStudent]:
        return self.db.query(ClinicalStudent).order_by(func.lower(ClinicalStudent.name), ClinicalStudent.student_id).all()

    def update_student(self, clinical_student_id: str, student_id: str, name: str, email: str = "") -> ClinicalStudent:
        student = self._require(
            ClinicalStudent, clinical_student_id, "STUDENT_NOT_FOUND", "The student no longer exists."
        )
        student.student_id, student.name, student.email = self._clean_student_values(
            student_id, name, email, require_name=True, exclude_id=student.id
        )
        student.updated_at = _now()
        self.db.flush()
        self.record_source_change()
        return student

    def student_tag_states(self) -> dict[str, dict[str, TagState]]:
        """Stored non-neutral tag states keyed by student then tag (Neutral is never stored)."""
        states: dict[str, dict[str, TagState]] = {}
        for row in self.db.query(ClinicalStudentTagState).all():
            states.setdefault(row.student_id, {})[row.tag_id] = TagState(row.state)
        return states

    def student_availability_counts(self) -> dict[str, int]:
        rows = (
            self.db.query(
                ClinicalStudentAvailabilityWindow.clinical_student_id, func.count(ClinicalStudentAvailabilityWindow.id)
            )
            .group_by(ClinicalStudentAvailabilityWindow.clinical_student_id)
            .all()
        )
        return {student_id: count for student_id, count in rows}

    def set_student_tag_states(self, clinical_student_id: str, states: dict[str, TagState]) -> None:
        """Replace one student's tag states with the supplied tag-to-state map in a single revision bump."""
        self._require(ClinicalStudent, clinical_student_id, "STUDENT_NOT_FOUND", "The student no longer exists.")
        wanted: dict[str, TagState] = {}
        for tag_id, state in states.items():
            self._require(ClinicalTag, tag_id, "TAG_NOT_FOUND", "A selected tag no longer exists.")
            try:
                state = TagState(state)
            except ValueError as error:
                raise ClinicalValidationError("INVALID_TAG_STATE", "Choose a valid tag state.") from error
            if state is not TagState.NEUTRAL:
                wanted[tag_id] = state
        self.db.query(ClinicalStudentTagState).filter(
            ClinicalStudentTagState.student_id == clinical_student_id
        ).delete()
        for tag_id, state in wanted.items():
            self.db.add(ClinicalStudentTagState(student_id=clinical_student_id, tag_id=tag_id, state=state.value))
        self.db.flush()
        self.record_source_change()

    def roster_import_snapshot(self) -> dict[str, tuple[str, str, str, int]]:
        """Case-folded Student ID -> (internal id, name, email, stored window count) for import matching."""
        counts = self.student_availability_counts()
        return {
            student.student_id.casefold(): (student.id, student.name, student.email, counts.get(student.id, 0))
            for student in self.list_students()
        }

    def apply_roster_import(self, rows: Sequence[RosterImportRow]) -> ClinicalRosterImportResult:
        """Create/update students and optionally replace their availability as one source-data change.

        All-or-nothing within the caller's transaction. Students absent from ``rows``, sites,
        tags, tag states, and placements are untouched; the revision is bumped exactly once.
        """
        seen: set[str] = set()
        created = updated = availability_updated = windows_imported = 0
        for row in rows:
            key = (row.student_id or "").strip().casefold()
            if key in seen:
                raise ClinicalValidationError("DUPLICATE_STUDENT_ID", "Student IDs must be unique in the import.")
            seen.add(key)
            windows: list[AvailabilityWindow] | None = None
            if row.windows is not None:
                windows, issues = normalize_windows(list(row.windows))
                if any(issue.severity == "error" for issue in issues):
                    raise ClinicalValidationError(
                        "INVALID_AVAILABILITY_WINDOW",
                        "Availability Windows must be valid and must not overlap with a different status.",
                    )
            if row.clinical_student_id is None:
                cleaned_id, cleaned_name, cleaned_email = self._clean_student_values(
                    row.student_id, row.name, row.email, require_name=True
                )
                student = ClinicalStudent(student_id=cleaned_id, name=cleaned_name, email=cleaned_email)
                self.db.add(student)
                self.db.flush()
                created += 1
            else:
                student = self._require(
                    ClinicalStudent, row.clinical_student_id, "STUDENT_NOT_FOUND", "The student no longer exists."
                )
                # Matching is case-insensitive; the stored Student ID spelling is kept.
                _, student.name, student.email = self._clean_student_values(
                    student.student_id, row.name, row.email, require_name=True, exclude_id=student.id
                )
                student.updated_at = _now()
                self.db.flush()
                updated += 1
            if windows is not None:
                self.replace_student_availability(student.id, windows, record_change=False)
                availability_updated += 1
                windows_imported += len(windows)
        if rows:
            self.record_source_change()
        return ClinicalRosterImportResult(created, updated, availability_updated, windows_imported)

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

    def set_student_availability(
        self, clinical_student_id: str, windows: Sequence[AvailabilityWindow]
    ) -> list[AvailabilityWindow]:
        """Replace one student's recurring weekly windows as a single source-data change.

        Same-status overlaps and adjacent windows are merged; conflicting overlaps and
        invalid windows are rejected so nothing partial is stored.
        """
        self._require(ClinicalStudent, clinical_student_id, "STUDENT_NOT_FOUND", "The student no longer exists.")
        normalized, issues = normalize_windows(list(windows))
        if any(issue.severity == "error" for issue in issues):
            raise ClinicalValidationError(
                "INVALID_AVAILABILITY_WINDOW",
                "Availability Windows must be valid and must not overlap with a different status.",
            )
        self.replace_student_availability(clinical_student_id, normalized)
        return self.student_availability(clinical_student_id)

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

    def scheduling_input(self) -> SchedulingInput:
        """Immutable snapshot of scheduling inputs, stamped with the current source-data revision.

        Candidate and readiness results derived from it are valid only for that
        revision; any source-data change increments it.
        """
        tags = {tag.id: tag.name for tag in self.list_tags()}
        availability: dict[str, list[AvailabilityWindow]] = {}
        for row in self.db.query(ClinicalStudentAvailabilityWindow).order_by(
            ClinicalStudentAvailabilityWindow.day, ClinicalStudentAvailabilityWindow.start_minute
        ):
            availability.setdefault(row.clinical_student_id, []).append(AvailabilityWindow(
                day=row.day, start_minute=row.start_minute, end_minute=row.end_minute,
                status=row.status, source=row.source, source_row=row.source_row,
            ))
        states = self.student_tag_states()
        preferences = {
            TagState.REQUIRED: TagPreference.REQUIRED,
            TagState.PREFERRED: TagPreference.PREFERRED,
            TagState.RESTRICTED: TagPreference.RESTRICTED,
        }
        students = tuple(
            StudentInput(
                id=student.id,
                label=student.name or student.student_id,
                student_number=student.student_id,
                availability=tuple(availability.get(student.id, ())),
                tag_preferences={
                    tag_id: preferences[state]
                    for tag_id, state in states.get(student.id, {}).items()
                    if state in preferences
                },
            )
            for student in self.list_students()
        )
        opportunities = tuple(
            OpportunityInput(
                id=item.id,
                label=(
                    f"{item.site_name}, {item.day_of_week[:3]} "
                    f"{format_time_of_day(item.start_minute)}-{format_time_of_day(item.end_minute)}"
                ),
                day=item.day_of_week,
                start_minute=item.start_minute,
                end_minute=item.end_minute,
                session_length_minutes=item.session_length_minutes,
                travel_minutes=item.travel_minutes,
                capacity=item.capacity,
                tag_ids=frozenset(tag.id for tag in item.tags),
                site_name=item.site_name,
            )
            for item in self.list_opportunities()
        )
        return SchedulingInput(self.get_state().source_data_revision, students, opportunities, tags)

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

    # --- schedule generation -----------------------------------------------

    def set_manual_placement(
        self, clinical_student_id: str, placement_opportunity_id: str, start_minute: int
    ) -> ClinicalPlacement:
        """Assign or move a student manually, preserving all policy warnings for review."""
        if isinstance(start_minute, bool) or not isinstance(start_minute, int) or not 0 <= start_minute < 24 * 60:
            raise ClinicalValidationError("INVALID_PLACEMENT", "The start time is not a valid time of day.")
        placement = self.set_placement(
            clinical_student_id, PlacementType.MANUAL, placement_opportunity_id, start_minute
        )
        self._record_schedule_edit()
        return placement

    def set_manually_unplaced(self, clinical_student_id: str) -> ClinicalPlacement:
        placement = self.set_placement(clinical_student_id, PlacementType.MANUALLY_UNPLACED)
        self._record_schedule_edit()
        return placement

    def preview_manual_placement(
        self, clinical_student_id: str, placement_opportunity_id: str, start_minute: int
    ) -> tuple[PlacementWarning, ...]:
        if isinstance(start_minute, bool) or not isinstance(start_minute, int) or not 0 <= start_minute < 24 * 60:
            raise ClinicalValidationError("INVALID_PLACEMENT", "The start time is not a valid time of day.")
        self.get_student(clinical_student_id)
        self._require(
            PlacementOpportunity,
            placement_opportunity_id,
            "OPPORTUNITY_NOT_FOUND",
            "The Placement Opportunity no longer exists.",
        )
        records = [r for r in self.list_placement_records() if r.student_id != clinical_student_id]
        records.append(
            PlacementRecord(
                student_id=clinical_student_id,
                kind=PlacementKind.MANUAL,
                opportunity_id=placement_opportunity_id,
                start_minute=start_minute,
            )
        )
        entries, _ = review_schedule(self.scheduling_input(), records)
        return next(entry.warnings for entry in entries if entry.student_id == clinical_student_id)

    def remove_manual_override(self, clinical_student_id: str) -> bool:
        placement = self.get_placement(clinical_student_id)
        if placement is None or placement.placement_type == PlacementType.AUTOMATIC.value:
            raise ClinicalValidationError(
                "NO_MANUAL_OVERRIDE", "The student does not have a manual placement override to remove."
            )
        self.db.delete(placement)
        self._record_schedule_edit()
        self.db.flush()
        return True

    def remove_placement_decision(self, clinical_student_id: str) -> bool:
        placement = self.get_placement(clinical_student_id)
        if placement is None:
            raise ClinicalValidationError("PLACEMENT_NOT_FOUND", "The student has no placement decision to remove.")
        self.db.delete(placement)
        self._record_schedule_edit()
        self.db.flush()
        return True

    def _record_schedule_edit(self) -> None:
        state = self.get_state()
        state.schedule_revision += 1
        state.modified_at = _now()
        self.db.flush()

    def list_placement_records(self) -> list[PlacementRecord]:
        return [
            PlacementRecord(
                student_id=row.clinical_student_id,
                kind=PlacementKind(row.placement_type),
                opportunity_id=row.placement_opportunity_id,
                start_minute=row.start_minute,
            )
            for row in self.db.query(ClinicalPlacement).order_by(ClinicalPlacement.clinical_student_id)
        ]

    def schedule_review(self) -> tuple[tuple[ScheduleEntry, ...], tuple[PlacementVerification, ...]]:
        """Return schedule rows and verification warnings against current source data."""
        scheduling_input = self.scheduling_input()
        return review_schedule(scheduling_input, self.list_placement_records(), generate_candidates(scheduling_input))

    def generate_schedule(self) -> OptimizationResult:
        """Replace automatic placements with a fresh optimal solution.

        Refuses when readiness has blocking errors. Manual and Manually unplaced
        decisions are preserved and treated as occupying capacity. Increments the
        schedule revision and clears staleness; Sites and Students are untouched.
        """
        scheduling_input = self.scheduling_input()
        candidates = generate_candidates(scheduling_input)
        if not analyze_readiness(scheduling_input, candidates).is_ready:
            raise ClinicalValidationError(
                "NOT_READY", "Placements cannot be generated until the blocking readiness errors are resolved."
            )
        self.db.query(ClinicalPlacement).filter(
            ClinicalPlacement.placement_type == PlacementType.AUTOMATIC.value
        ).delete(synchronize_session=False)
        preserved = [
            PreservedPlacement(item.student_id, item.opportunity_id) for item in self.list_placement_records()
        ]
        result = optimize(scheduling_input, candidates, preserved)
        for item in result.placements:
            self.set_placement(
                item.student_id, PlacementType.AUTOMATIC, item.opportunity_id, item.selected_start_minute
            )
        state = self.get_state()
        state.schedule_revision += 1
        state.is_schedule_stale = False
        state.modified_at = _now()
        self.db.flush()
        return result

    def schedule_summary(self) -> ScheduleSummary:
        scheduling_input = self.scheduling_input()
        candidates = generate_candidates(scheduling_input)
        return summarize_schedule(
            scheduling_input, candidates, self.list_placement_records()
        )

    def schedule_snapshot(
        self,
    ) -> tuple[ScheduleSummary, tuple[ScheduleEntry, ...], tuple[PlacementVerification, ...]]:
        """Return summary, rows, and verification from one consistent input snapshot."""
        scheduling_input = self.scheduling_input()
        candidates = generate_candidates(scheduling_input)
        records = self.list_placement_records()
        return (
            summarize_schedule(scheduling_input, candidates, records),
            *review_schedule(scheduling_input, records, candidates),
        )

    # --- helpers ------------------------------------------------------

    def _require(self, model, primary_key: str, code: str, message: str):
        row = self.db.get(model, primary_key)
        if row is None:
            raise ClinicalValidationError(code, message)
        return row
