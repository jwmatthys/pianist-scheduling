"""Tutor Scheduling persistence and typed orchestration repository.

The repository owns roster, availability, placement, and schedule persistence,
revision/staleness state, and orchestration over pure candidate, readiness, and
optimizer services.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Mapping, Sequence

from sqlalchemy import case, func
from sqlalchemy.orm import Session

from .. import models
from ..tutor_models import (
    DEFAULT_TUTOR_CAPACITY,
    DEFAULT_LESSON_LENGTH_MINUTES,
    TutorTutor,
    TutorTutorAvailabilityWindow,
    TutorTutorTag,
    TutorModuleState,
    TutorPlacement,
    TutorPlacementType,
    TutorSchedulePlacement,
    TutorScheduleResult,
    TutorStudent,
    TutorStudentAvailabilityWindow,
    TutorStudentTagState,
    TutorTag,
    TutorTagState,
    TutorUnplacedReason,
    TutorUnplacedStudent,
    tutor_tag_identity_key,
    normalize_tutor_tag_name,
)
from .availability import AvailabilityWindow, WEEKDAYS, normalize_day, normalize_windows
from .tutor_candidates import generate_tutor_candidates
from .tutor_scheduling_models import (
    TutorCandidate,
    TutorCandidateSet,
    TutorTutorInput,
    TutorOptimizationResult,
    TutorPlacementVerification,
    TutorReadinessReport,
    TutorSchedule,
    TutorScheduleEntry,
    TutorScheduleStatistics,
    TutorScheduledPlacement,
    TutorSchedulingInput,
    TutorStoredAssignment,
    TutorStoredUnplaced,
    TutorStudentInput,
    TutorTagInput,
    TutorUnplacedStudentResult,
)
from .tutor_readiness import analyze_tutor_readiness
from .tutor_optimizer import (
    _gap_cost,
    _schedulable_tutor_capacity,
    _usable_coverage,
    optimize_tutor_schedule,
)
from .tutor_placement_review import review_tutor_schedule


class TutorValidationError(ValueError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


@dataclass(frozen=True)
class TutorInput:
    name: str
    email: str = ""
    room: str = ""
    lesson_length_minutes: int = DEFAULT_LESSON_LENGTH_MINUTES
    capacity: int = DEFAULT_TUTOR_CAPACITY


@dataclass(frozen=True)
class TutorTagSummary:
    tag: TutorTag
    tutor_count: int
    student_count: int

    @property
    def is_unused(self) -> bool:
        return self.tutor_count == 0 and self.student_count == 0


@dataclass(frozen=True)
class StudentInput:
    student_id: str
    name: str
    email: str = ""


@dataclass(frozen=True)
class TutorStudentReplaceResult:
    students_removed: int
    placements_removed: int
    placements_preserved: int
    students_created: int
    tags_created: int = 0


@dataclass(frozen=True)
class TutorReplaceResult:
    tutors_removed: int
    placements_removed: int
    placements_preserved: int
    tutors_created: int
    tags_created: int


def _now() -> datetime:
    return datetime.utcnow()


def _is_int(value: object) -> bool:
    return isinstance(value, int) and not isinstance(value, bool)


class TutorRepository:
    def __init__(self, db: Session):
        self.db = db

    # --- module state -------------------------------------------------

    def _session_uuid(self) -> str:
        rows = self.db.query(models.SchedulingSession.session_uuid).all()
        if len(rows) != 1:
            raise ValueError("The active database must contain exactly one Scheduling Session.")
        return rows[0][0]

    def get_state(self) -> TutorModuleState:
        session_uuid = self._session_uuid()
        state = self.db.get(TutorModuleState, session_uuid)
        if state is None:
            state = TutorModuleState(
                session_uuid=session_uuid,
                schedule_revision=0,
                source_data_revision=0,
                is_schedule_stale=False,
                modified_at=_now(),
            )
            self.db.add(state)
            self.db.flush()
        return state

    def record_source_change(self, *, schedule_exists: bool | None = None) -> TutorModuleState:
        """Bump the source revision; flag an existing schedule stale without altering it."""
        state = self.get_state()
        state.source_data_revision += 1
        if schedule_exists is None:
            schedule_exists = self._has_schedule_data()
        if schedule_exists:
            state.is_schedule_stale = True
        state.modified_at = _now()
        self.db.flush()
        return state

    def _record_schedule_edit(self) -> None:
        state = self.get_state()
        state.schedule_revision += 1
        state.modified_at = _now()
        self.db.flush()

    def _has_schedule_data(self) -> bool:
        return (
            self.db.query(TutorPlacement.id).first() is not None
            or self.db.query(TutorScheduleResult.session_uuid).first() is not None
        )

    # --- tags ---------------------------------------------------------

    def list_tags(self) -> list[TutorTag]:
        return self.db.query(TutorTag).order_by(TutorTag.normalized_name).all()

    def find_tag(self, name: str) -> TutorTag | None:
        key = tutor_tag_identity_key(name)
        if not key:
            return None
        return self.db.query(TutorTag).filter(TutorTag.normalized_name == key).one_or_none()

    def _clean_tag_name(self, name: str) -> str:
        cleaned = normalize_tutor_tag_name(name)
        if not cleaned:
            raise TutorValidationError("INVALID_TAG", "A tag name is required.")
        return cleaned

    def create_tag(self, name: str) -> TutorTag:
        cleaned = self._clean_tag_name(name)
        if self.find_tag(cleaned) is not None:
            raise TutorValidationError("DUPLICATE_TAG", f"The tag '{cleaned}' already exists.")
        tag = self._insert_tag(cleaned)
        self.record_source_change()
        return tag

    def _insert_tag(self, cleaned_name: str) -> TutorTag:
        tag = TutorTag(name=cleaned_name, normalized_name=tutor_tag_identity_key(cleaned_name))
        self.db.add(tag)
        self.db.flush()
        return tag

    def tag_summaries(self) -> list["TutorTagSummary"]:
        """Every registered tag with current-use information for the Tag Manager."""
        tutor_counts = dict(
            self.db.query(TutorTutorTag.tag_id, func.count(TutorTutorTag.tutor_id))
            .group_by(TutorTutorTag.tag_id)
            .all()
        )
        student_counts = dict(
            self.db.query(TutorStudentTagState.tag_id, func.count(func.distinct(TutorStudentTagState.student_id)))
            .filter(TutorStudentTagState.state != TutorTagState.NEUTRAL.value)
            .group_by(TutorStudentTagState.tag_id)
            .all()
        )
        return [
            TutorTagSummary(
                tag=tag,
                tutor_count=tutor_counts.get(tag.id, 0),
                student_count=student_counts.get(tag.id, 0),
            )
            for tag in self.list_tags()
        ]

    def rename_tag(self, tag_id: str, name: str) -> TutorTag:
        tag = self._require(TutorTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
        cleaned = self._clean_tag_name(name)
        existing = self.find_tag(cleaned)
        if existing is not None and existing.id != tag.id:
            raise TutorValidationError("DUPLICATE_TAG", f"The tag '{cleaned}' already exists.")
        tag.name = cleaned
        tag.normalized_name = tutor_tag_identity_key(cleaned)
        tag.updated_at = _now()
        self.db.flush()
        self.record_source_change()
        return tag

    def delete_tag(self, tag_id: str) -> None:
        """Delete a tag and its Tutor Scheduling associations."""
        tag = self._require(TutorTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
        self.db.query(TutorTutorTag).filter(TutorTutorTag.tag_id == tag.id).delete()
        self.db.query(TutorStudentTagState).filter(TutorStudentTagState.tag_id == tag.id).delete()
        self.db.delete(tag)
        self.db.flush()
        self.record_source_change()

    # --- tutors --------------------------------------------------

    def _clean_tutor_values(
        self, name: str, email: str, room: str, lesson_length_minutes: int, capacity: int
    ) -> dict:
        cleaned_name = (name or "").strip()
        if not cleaned_name:
            raise TutorValidationError("INVALID_TUTOR_NAME", "Tutor name is required.")
        if not _is_int(lesson_length_minutes) or lesson_length_minutes <= 0:
            raise TutorValidationError("INVALID_LESSON_LENGTH", "Lesson length must be a positive number of minutes.")
        if not _is_int(capacity) or capacity < 1:
            raise TutorValidationError("INVALID_CAPACITY", "Capacity must be at least 1.")
        return {
            "name": cleaned_name,
            "email": (email or "").strip(),
            "room": (room or "").strip(),
            "lesson_length_minutes": lesson_length_minutes,
            "capacity": capacity,
        }

    def create_tutor(
        self,
        name: str,
        email: str = "",
        room: str = "",
        lesson_length_minutes: int = DEFAULT_LESSON_LENGTH_MINUTES,
        capacity: int = DEFAULT_TUTOR_CAPACITY,
        availability: Sequence[AvailabilityWindow] = (),
        tag_ids: Iterable[str] = (),
    ) -> TutorTutor:
        values = self._clean_tutor_values(name, email, room, lesson_length_minutes, capacity)
        tutor = TutorTutor(**values)
        self.db.add(tutor)
        self.db.flush()
        self._replace_tutor_availability(tutor.id, availability)
        self._replace_tutor_tags(tutor, tag_ids)
        self.record_source_change()
        return tutor

    def get_tutor(self, tutor_id: str) -> TutorTutor:
        return self._require(TutorTutor, tutor_id, "TUTOR_NOT_FOUND", "The tutor no longer exists.")

    def list_tutors(self) -> list[TutorTutor]:
        return self.db.query(TutorTutor).order_by(func.lower(TutorTutor.name), TutorTutor.id).all()

    def update_tutor(
        self,
        tutor_id: str,
        name: str,
        email: str = "",
        room: str = "",
        lesson_length_minutes: int = DEFAULT_LESSON_LENGTH_MINUTES,
        capacity: int = DEFAULT_TUTOR_CAPACITY,
        tag_ids: Iterable[str] | None = None,
    ) -> TutorTutor:
        """Update fields (and tags when given) as a single source-data change."""
        tutor = self.get_tutor(tutor_id)
        values = self._clean_tutor_values(name, email, room, lesson_length_minutes, capacity)
        if tag_ids is not None:
            self._replace_tutor_tags(tutor, tag_ids)
        for field, value in values.items():
            setattr(tutor, field, value)
        tutor.updated_at = _now()
        self.db.flush()
        self.record_source_change()
        return tutor

    def set_tutor_tags(self, tutor_id: str, tag_ids: Iterable[str]) -> TutorTutor:
        tutor = self.get_tutor(tutor_id)
        self._replace_tutor_tags(tutor, tag_ids)
        self.record_source_change()
        return tutor

    def delete_tutor(self, tutor_id: str) -> int:
        """Delete a tutor without deleting schedule decisions that reference it."""
        tutor = self.get_tutor(tutor_id)
        self.db.query(TutorTutorTag).filter(TutorTutorTag.tutor_id == tutor.id).delete()
        self.db.query(TutorTutorAvailabilityWindow).filter(
            TutorTutorAvailabilityWindow.tutor_id == tutor.id
        ).delete()
        self.db.delete(tutor)
        self.db.flush()
        self.record_source_change()
        return 0

    def _replace_tutor_tags(self, tutor: TutorTutor, tag_ids: Iterable[str]) -> None:
        wanted = list(dict.fromkeys(tag_ids))
        for tag_id in wanted:
            self._require(TutorTag, tag_id, "TAG_NOT_FOUND", "A selected tag no longer exists.")
        self.db.query(TutorTutorTag).filter(TutorTutorTag.tutor_id == tutor.id).delete()
        for tag_id in wanted:
            self.db.add(TutorTutorTag(tutor_id=tutor.id, tag_id=tag_id))
        self.db.flush()
        self.db.expire(tutor, ["tags"])

    def replace_tutors(
        self, tutors: Iterable[tuple[TutorInput, Sequence[AvailabilityWindow], Iterable[str]]]
    ) -> "TutorReplaceResult":
        """Replace tutors without deleting schedule decisions."""
        prepared = [
            (
                self._clean_tutor_values(
                    values.name, values.email, values.room, values.lesson_length_minutes, values.capacity
                ),
                self._validated_windows(windows),
                list(names),
            )
            for values, windows, names in tutors
        ]
        tutors_removed = self.db.query(func.count(TutorTutor.id)).scalar() or 0
        placements_removed = 0
        placements_preserved = self.db.query(func.count(TutorPlacement.id)).scalar() or 0
        self.db.query(TutorTutorTag).delete()
        self.db.query(TutorTutorAvailabilityWindow).delete()
        self.db.query(TutorTutor).delete()
        self.db.flush()

        tags: dict[str, TutorTag] = {}
        tags_created = 0
        for values, windows, names in prepared:
            tutor = TutorTutor(**values)
            self.db.add(tutor)
            self.db.flush()
            self._replace_tutor_availability(tutor.id, windows)
            linked: set[str] = set()
            for name in names:
                key = tutor_tag_identity_key(name)
                if not key or key in linked:
                    continue
                tag = tags.get(key) or self.find_tag(name)
                if tag is None:
                    tag = self._insert_tag(normalize_tutor_tag_name(name))
                    tags_created += 1
                tags[key] = tag
                linked.add(key)
                self.db.add(TutorTutorTag(tutor_id=tutor.id, tag_id=tag.id))
            self.db.flush()
        self.record_source_change()
        return TutorReplaceResult(
            tutors_removed=tutors_removed,
            placements_removed=placements_removed,
            placements_preserved=placements_preserved,
            tutors_created=len(prepared),
            tags_created=tags_created,
        )

    def tutor_availability_counts(self) -> dict[str, int]:
        return dict(
            self.db.query(
                TutorTutorAvailabilityWindow.tutor_id, func.count(TutorTutorAvailabilityWindow.id)
            )
            .group_by(TutorTutorAvailabilityWindow.tutor_id)
            .all()
        )

    def tutor_availability_minutes(self) -> dict[str, tuple[int, int]]:
        """Weekly (Available, Tentative) minutes per tutor; Unavailable windows are not counted."""
        duration = TutorTutorAvailabilityWindow.end_minute - TutorTutorAvailabilityWindow.start_minute
        rows = (
            self.db.query(
                TutorTutorAvailabilityWindow.tutor_id,
                func.coalesce(func.sum(case((TutorTutorAvailabilityWindow.status == "Available", duration), else_=0)), 0),
                func.coalesce(func.sum(case((TutorTutorAvailabilityWindow.status == "Tentative", duration), else_=0)), 0),
            )
            .group_by(TutorTutorAvailabilityWindow.tutor_id)
            .all()
        )
        return {key: (int(available), int(tentative)) for key, available, tentative in rows}

    # --- availability -------------------------------------------------

    @staticmethod
    def _validated_windows(windows: Sequence[AvailabilityWindow]) -> list[AvailabilityWindow]:
        """Merge same-status overlaps; reject invalid or conflicting windows so nothing partial is stored."""
        for window in windows:
            if normalize_day(window.day) != window.day or not 0 <= window.start_minute < window.end_minute <= 24 * 60:
                raise TutorValidationError("INVALID_AVAILABILITY_WINDOW", "An Availability Window is invalid.")
        normalized, issues = normalize_windows(list(windows))
        if any(issue.severity == "error" for issue in issues):
            raise TutorValidationError(
                "INVALID_AVAILABILITY_WINDOW",
                "Availability Windows must be valid and must not overlap with a different status.",
            )
        return normalized

    def _replace_tutor_availability(self, tutor_id: str, windows: Sequence[AvailabilityWindow]) -> None:
        validated = self._validated_windows(windows)
        self.db.query(TutorTutorAvailabilityWindow).filter(
            TutorTutorAvailabilityWindow.tutor_id == tutor_id
        ).delete()
        for window in validated:
            self.db.add(TutorTutorAvailabilityWindow(
                tutor_id=tutor_id,
                day=window.day,
                start_minute=window.start_minute,
                end_minute=window.end_minute,
                status=window.status,
                source=window.source,
                source_row=window.source_row,
            ))
        self.db.flush()

    def set_tutor_availability(
        self, tutor_id: str, windows: Sequence[AvailabilityWindow]
    ) -> list[AvailabilityWindow]:
        """Replace one tutor's recurring weekly windows as a single source-data change."""
        self.get_tutor(tutor_id)
        self._replace_tutor_availability(tutor_id, windows)
        self.record_source_change()
        return self.tutor_availability(tutor_id)

    def tutor_availability(self, tutor_id: str) -> list[AvailabilityWindow]:
        rows = (
            self.db.query(TutorTutorAvailabilityWindow)
            .filter(TutorTutorAvailabilityWindow.tutor_id == tutor_id)
            .order_by(TutorTutorAvailabilityWindow.day, TutorTutorAvailabilityWindow.start_minute)
            .all()
        )
        return [_window(row) for row in rows]

    def _replace_student_availability(self, tutor_student_id: str, windows: Sequence[AvailabilityWindow]) -> None:
        validated = self._validated_windows(windows)
        self.db.query(TutorStudentAvailabilityWindow).filter(
            TutorStudentAvailabilityWindow.tutor_student_id == tutor_student_id
        ).delete()
        for window in validated:
            self.db.add(TutorStudentAvailabilityWindow(
                tutor_student_id=tutor_student_id,
                day=window.day,
                start_minute=window.start_minute,
                end_minute=window.end_minute,
                status=window.status,
                source=window.source,
                source_row=window.source_row,
            ))
        self.db.flush()

    def set_student_availability(
        self, tutor_student_id: str, windows: Sequence[AvailabilityWindow]
    ) -> list[AvailabilityWindow]:
        """Replace one student's recurring weekly windows as a single source-data change."""
        self.get_student(tutor_student_id)
        self._replace_student_availability(tutor_student_id, windows)
        self.record_source_change()
        return self.student_availability(tutor_student_id)

    def student_availability(self, tutor_student_id: str) -> list[AvailabilityWindow]:
        rows = (
            self.db.query(TutorStudentAvailabilityWindow)
            .filter(TutorStudentAvailabilityWindow.tutor_student_id == tutor_student_id)
            .order_by(TutorStudentAvailabilityWindow.day, TutorStudentAvailabilityWindow.start_minute)
            .all()
        )
        return [_window(row) for row in rows]

    # --- students -----------------------------------------------------

    def _clean_student_values(
        self,
        student_id: str,
        name: str,
        email: str,
        exclude_id: str | None = None,
        check_unique: bool = True,
    ) -> tuple[str, str, str]:
        cleaned_id = (student_id or "").strip()
        if not cleaned_id:
            raise TutorValidationError("INVALID_STUDENT_ID", "Student ID is required.")
        if check_unique:
            query = self.db.query(TutorStudent.id).filter(func.lower(TutorStudent.student_id) == cleaned_id.lower())
            if exclude_id is not None:
                query = query.filter(TutorStudent.id != exclude_id)
            if query.first() is not None:
                raise TutorValidationError("DUPLICATE_STUDENT_ID", "A student with that Student ID already exists.")
        cleaned_name = " ".join((name or "").split())
        if not cleaned_name:
            raise TutorValidationError("INVALID_STUDENT_NAME", "Name is required.")
        return cleaned_id, cleaned_name, (email or "").strip()

    def create_student(
        self,
        student_id: str,
        name: str = "",
        email: str = "",
        availability: Sequence[AvailabilityWindow] = (),
        tag_states: Mapping[str, TutorTagState | str] | None = None,
    ) -> TutorStudent:
        cleaned_id, cleaned_name, cleaned_email = self._clean_student_values(
            student_id, name, email
        )
        validated_states = self._validated_tag_states(tag_states or {})
        student = TutorStudent(student_id=cleaned_id, name=cleaned_name, email=cleaned_email)
        self.db.add(student)
        self.db.flush()
        self._replace_student_availability(student.id, availability)
        self._replace_student_tag_states(student.id, validated_states)
        self.record_source_change()
        return student

    def get_student(self, tutor_student_id: str) -> TutorStudent:
        return self._require(TutorStudent, tutor_student_id, "STUDENT_NOT_FOUND", "The student no longer exists.")

    def list_students(self) -> list[TutorStudent]:
        return self.db.query(TutorStudent).order_by(func.lower(TutorStudent.name), TutorStudent.student_id).all()

    def update_student(
        self,
        tutor_student_id: str,
        student_id: str,
        name: str,
        email: str = "",
        tag_states: Mapping[str, TutorTagState | str] | None = None,
    ) -> TutorStudent:
        """Update fields (and, when given, the full tag-state set) as one source-data change."""
        student = self.get_student(tutor_student_id)
        cleaned = self._clean_student_values(
            student_id, name, email, exclude_id=student.id
        )
        validated_states = self._validated_tag_states(tag_states) if tag_states is not None else None
        student.student_id, student.name, student.email = cleaned
        student.updated_at = _now()
        self.db.flush()
        if validated_states is not None:
            self._replace_student_tag_states(student.id, validated_states)
        self.record_source_change()
        return student

    def delete_student(self, tutor_student_id: str) -> None:
        student = self.get_student(tutor_student_id)
        self.db.query(TutorStudentTagState).filter(TutorStudentTagState.student_id == student.id).delete()
        self.db.query(TutorStudentAvailabilityWindow).filter(
            TutorStudentAvailabilityWindow.tutor_student_id == student.id
        ).delete()
        self.db.delete(student)
        self.db.flush()
        self.record_source_change()

    # --- student tag states -------------------------------------------

    def student_tag_states(self) -> dict[str, dict[str, TutorTagState]]:
        """Stored non-neutral tag states keyed by student then tag (Neutral is never stored)."""
        states: dict[str, dict[str, TutorTagState]] = {}
        for row in self.db.query(TutorStudentTagState).all():
            states.setdefault(row.student_id, {})[row.tag_id] = TutorTagState(row.state)
        return states

    def get_scheduling_input(self) -> TutorSchedulingInput:
        """Return a typed Tutor scheduling snapshot at the current source revision."""
        source_data_revision = self.get_state().source_data_revision
        tutors = self.list_tutors()
        students = self.list_students()

        tutor_ids = [item.id for item in tutors]
        student_ids = [item.id for item in students]
        tutor_windows: dict[str, list[AvailabilityWindow]] = {key: [] for key in tutor_ids}
        student_windows: dict[str, list[AvailabilityWindow]] = {key: [] for key in student_ids}
        tutor_tags: dict[str, set[str]] = {key: set() for key in tutor_ids}
        student_states: dict[str, list[tuple[str, TutorTagState]]] = {key: [] for key in student_ids}

        for row in (
            self.db.query(TutorTutorAvailabilityWindow)
            .order_by(TutorTutorAvailabilityWindow.day, TutorTutorAvailabilityWindow.start_minute)
            .all()
        ):
            if row.tutor_id in tutor_windows:
                tutor_windows[row.tutor_id].append(_window(row))
        for row in (
            self.db.query(TutorStudentAvailabilityWindow)
            .order_by(TutorStudentAvailabilityWindow.day, TutorStudentAvailabilityWindow.start_minute)
            .all()
        ):
            if row.tutor_student_id in student_windows:
                student_windows[row.tutor_student_id].append(_window(row))
        for row in self.db.query(TutorTutorTag).all():
            if row.tutor_id in tutor_tags:
                tutor_tags[row.tutor_id].add(row.tag_id)
        for row in self.db.query(TutorStudentTagState).all():
            if row.student_id in student_states:
                student_states[row.student_id].append((row.tag_id, TutorTagState(row.state)))

        return TutorSchedulingInput(
            source_data_revision=source_data_revision,
            students=tuple(
                TutorStudentInput(
                    id=student.id,
                    name=student.name,
                    availability=tuple(student_windows[student.id]),
                    tag_states=tuple(sorted(student_states[student.id])),
                )
                for student in students
            ),
            tutors=tuple(
                TutorTutorInput(
                    id=tutor.id,
                    name=tutor.name,
                    lesson_length_minutes=tutor.lesson_length_minutes,
                    capacity=tutor.capacity,
                    tag_ids=frozenset(tutor_tags[tutor.id]),
                    availability=tuple(tutor_windows[tutor.id]),
                )
                for tutor in tutors
            ),
            tags=tuple(
                TutorTagInput(id=tag.id, name=tag.name)
                for tag in self.list_tags()
            ),
        )

    def generate_candidate_set(self) -> TutorCandidateSet:
        """Generate Tutor candidates from the current typed snapshot."""
        return generate_tutor_candidates(self.get_scheduling_input())

    def get_readiness_report(self) -> TutorReadinessReport:
        """Orchestrate a revision-consistent snapshot, candidate set, and readiness analysis."""
        scheduling_input = self.get_scheduling_input()
        candidate_set = generate_tutor_candidates(scheduling_input)
        return analyze_tutor_readiness(scheduling_input, candidate_set)

    def _stored_assignments(
        self, scheduling_input: TutorSchedulingInput
    ) -> tuple[TutorStoredAssignment, ...]:
        students = {item.id: item for item in scheduling_input.students}
        tutors = {item.id: item for item in scheduling_input.tutors}
        previous = {
            item.student_id: item
            for item in self.db.query(TutorSchedulePlacement).all()
        }
        assignments: list[TutorStoredAssignment] = []
        for row in self.list_placements():
            if row.placement_type == TutorPlacementType.MANUALLY_UNPLACED.value:
                continue
            if row.tutor_tutor_id is None or row.start_minute is None:
                continue
            prior = previous.get(row.tutor_student_id)
            tutor = tutors.get(row.tutor_tutor_id)
            student = students.get(row.tutor_student_id)
            end_minute = row.end_minute or (
                prior.end_minute if prior is not None
                else row.start_minute + (
                    tutor.lesson_length_minutes if tutor is not None
                    else DEFAULT_LESSON_LENGTH_MINUTES
                )
            )
            preferred, tentative = self._placement_metrics(
                student, tutor, row.day, row.start_minute, end_minute
            ) if student is not None and tutor is not None else (
                prior.preferred_tag_match_count if prior is not None else 0,
                prior.tentative_minutes if prior is not None else 0,
            )
            assignments.append(TutorStoredAssignment(
                student_id=row.tutor_student_id,
                placement_type=row.placement_type,
                tutor_id=row.tutor_tutor_id,
                tutor_name=tutor.name if tutor is not None else (
                    prior.tutor_name if prior is not None else row.tutor_tutor_id
                ),
                student_name=student.name if student is not None else (
                    prior.student_name if prior is not None else row.tutor_student_id
                ),
                day=row.day,
                start_minute=row.start_minute,
                end_minute=end_minute,
                preferred_tag_match_count=preferred,
                tentative_minutes=tentative,
            ))
        return tuple(assignments)

    @staticmethod
    def _placement_metrics(student, tutor, day: str, start: int, end: int) -> tuple[int, int]:
        states = dict(student.tag_states)
        preferred = sum(
            state is TutorTagState.PREFERRED and tag_id in tutor.tag_ids
            for tag_id, state in states.items()
        )
        student_tentative = sum(
            max(0, min(end, window.end_minute) - max(start, window.start_minute))
            for window in student.availability
            if window.day == day and window.status == "Tentative"
        )
        tutor_tentative = sum(
            max(0, min(end, window.end_minute) - max(start, window.start_minute))
            for window in tutor.availability
            if window.day == day and window.status == "Tentative"
        )
        return preferred, student_tentative + tutor_tentative

    def _stored_unplaced(self) -> tuple[TutorStoredUnplaced, ...]:
        students = {item.id: item.name for item in self.list_students()}
        decisions = {
            item.tutor_student_id: item for item in self.list_placements()
            if item.placement_type == TutorPlacementType.MANUALLY_UNPLACED.value
        }
        stored = {
            item.student_id: item for item in self.db.query(TutorUnplacedStudent).all()
        }
        for student_id, row in decisions.items():
            snapshot = stored.get(student_id)
            stored[student_id] = TutorUnplacedStudent(
                student_id=student_id,
                student_name=students.get(student_id, snapshot.student_name if snapshot else student_id),
                reason=TutorUnplacedReason.MANUALLY_UNPLACED.value,
                placement_type=TutorPlacementType.MANUALLY_UNPLACED.value,
            )
        return tuple(
            TutorStoredUnplaced(
                student_id=row.student_id,
                student_name=students.get(row.student_id, row.student_name or row.student_id),
                placement_type=row.placement_type,
                reason=row.reason,
            )
            for row in stored.values()
        )

    def _schedule_review(
        self, scheduling_input: TutorSchedulingInput | None = None
    ) -> tuple[
        TutorSchedulingInput,
        tuple[TutorStoredAssignment, ...],
        tuple[TutorStoredUnplaced, ...],
        tuple[TutorScheduleEntry, ...],
        tuple[TutorPlacementVerification, ...],
    ]:
        scheduling_input = scheduling_input or self.get_scheduling_input()
        candidates = generate_tutor_candidates(scheduling_input)
        assignments = self._stored_assignments(scheduling_input)
        unplaced = self._stored_unplaced()
        entries, verification = review_tutor_schedule(
            scheduling_input, candidates, assignments, unplaced
        )
        return scheduling_input, assignments, unplaced, entries, verification

    def _schedule_statistics(
        self,
        scheduling_input: TutorSchedulingInput,
        assignments: Sequence[TutorStoredAssignment],
        entries: Sequence[TutorScheduleEntry],
    ) -> TutorScheduleStatistics:
        tutor_by_id = {item.id: item for item in scheduling_input.tutors}
        capacities = {
            item.id: _schedulable_tutor_capacity(item)
            for item in scheduling_input.tutors
        }
        candidate_rows = []
        for item in assignments:
            tutor = tutor_by_id.get(item.tutor_id)
            if tutor is None or item.day not in WEEKDAYS:
                continue
            candidate_rows.append(TutorCandidate(
                student_id=item.student_id,
                tutor_id=item.tutor_id,
                day=item.day,
                tutor_start_minute=item.start_minute,
                tutor_end_minute=item.end_minute,
                preferred_tag_match_count=item.preferred_tag_match_count,
                tentative_minutes=item.tentative_minutes,
                issues=(),
                is_feasible=True,
                tutor_capacity=tutor.capacity,
            ))
        coverage = {
            item.id: _usable_coverage(item.availability)
            for item in scheduling_input.tutors
        }
        placed_count = len(assignments)
        return TutorScheduleStatistics(
            students_scheduled=placed_count,
            students_unscheduled=max(0, len(entries) - placed_count),
            tutor_utilization=placed_count / sum(capacities.values()) if sum(capacities.values()) else 0.0,
            preferred_tag_matches=sum(item.preferred_tag_match_count for item in assignments),
            tentative_minutes=sum(item.tentative_minutes for item in assignments),
            gap_cost_minutes=_gap_cost(candidate_rows, coverage),
        )

    def _refresh_schedule_snapshots(
        self,
        scheduling_input: TutorSchedulingInput,
        *,
        automatic_unplaced: Sequence[TutorUnplacedStudentResult] | None = None,
    ) -> tuple[TutorStoredAssignment, ...]:
        old_assignments = {
            item.student_id: item for item in self.db.query(TutorSchedulePlacement).all()
        }
        old_unplaced = {
            item.student_id: item for item in self.db.query(TutorUnplacedStudent).all()
        }
        students = {item.id: item for item in scheduling_input.students}
        tutors = {item.id: item for item in scheduling_input.tutors}
        rows: list[TutorStoredAssignment] = []
        self.db.query(TutorSchedulePlacement).delete()
        for decision in self.list_placements():
            if decision.placement_type == TutorPlacementType.MANUALLY_UNPLACED.value:
                continue
            if decision.tutor_tutor_id is None or decision.start_minute is None:
                continue
            old = old_assignments.get(decision.tutor_student_id)
            student = students.get(decision.tutor_student_id)
            tutor = tutors.get(decision.tutor_tutor_id)
            end_minute = decision.end_minute or (
                old.end_minute if old else decision.start_minute + (
                    tutor.lesson_length_minutes if tutor else DEFAULT_LESSON_LENGTH_MINUTES
                )
            )
            preferred, tentative = self._placement_metrics(
                student, tutor, decision.day, decision.start_minute, end_minute
            ) if student is not None and tutor is not None else (
                old.preferred_tag_match_count if old else 0,
                old.tentative_minutes if old else 0,
            )
            item = TutorStoredAssignment(
                student_id=decision.tutor_student_id,
                placement_type=decision.placement_type,
                tutor_id=decision.tutor_tutor_id,
                tutor_name=tutor.name if tutor else (
                    old.tutor_name if old else decision.tutor_tutor_id
                ),
                student_name=student.name if student else (
                    old.student_name if old else decision.tutor_student_id
                ),
                day=decision.day,
                start_minute=decision.start_minute,
                end_minute=end_minute,
                preferred_tag_match_count=preferred,
                tentative_minutes=tentative,
            )
            rows.append(item)
            self.db.add(TutorSchedulePlacement(
                student_id=item.student_id,
                tutor_id=item.tutor_id,
                placement_type=item.placement_type,
                student_name=item.student_name,
                tutor_name=item.tutor_name,
                day=item.day,
                start_minute=item.start_minute,
                end_minute=item.end_minute,
                preferred_tag_match_count=item.preferred_tag_match_count,
                tentative_minutes=item.tentative_minutes,
            ))

        automatic_unplaced_ids: set[str] = set()
        if automatic_unplaced is not None:
            self.db.query(TutorUnplacedStudent).filter(
                TutorUnplacedStudent.placement_type == TutorPlacementType.AUTOMATIC.value
            ).delete()
            for item in automatic_unplaced:
                student = students.get(item.student_id)
                automatic_unplaced_ids.add(item.student_id)
                self.db.add(TutorUnplacedStudent(
                    student_id=item.student_id,
                    student_name=student.name if student else (
                        old_unplaced[item.student_id].student_name
                        if item.student_id in old_unplaced else item.student_id
                    ),
                    reason=item.reason.value,
                    placement_type=TutorPlacementType.AUTOMATIC.value,
                ))

        assigned_ids = {item.student_id for item in rows}
        manual_unplaced_ids = {
            item.tutor_student_id for item in self.list_placements()
            if item.placement_type == TutorPlacementType.MANUALLY_UNPLACED.value
        }
        self.db.query(TutorUnplacedStudent).filter(
            TutorUnplacedStudent.placement_type == TutorPlacementType.MANUALLY_UNPLACED.value
        ).delete()
        if manual_unplaced_ids:
            # An earlier automatic unplaced row for the same student must be replaced.
            self.db.query(TutorUnplacedStudent).filter(
                TutorUnplacedStudent.student_id.in_(manual_unplaced_ids)
            ).delete(synchronize_session=False)
        for student_id in manual_unplaced_ids:
            previous = old_unplaced.get(student_id)
            student = students.get(student_id)
            self.db.add(TutorUnplacedStudent(
                student_id=student_id,
                student_name=student.name if student else (previous.student_name if previous else student_id),
                reason=TutorUnplacedReason.MANUALLY_UNPLACED.value,
                placement_type=TutorPlacementType.MANUALLY_UNPLACED.value,
            ))
        if automatic_unplaced is not None:
            decisions = {item.tutor_student_id for item in self.list_placements()}
            for student in scheduling_input.students:
                if (
                    student.id not in assigned_ids
                    and student.id not in manual_unplaced_ids
                    and student.id not in decisions
                    and student.id not in automatic_unplaced_ids
                ):
                    self.db.add(TutorUnplacedStudent(
                        student_id=student.id,
                        student_name=student.name,
                        reason=TutorUnplacedReason.NO_FEASIBLE_CANDIDATE.value,
                        placement_type=TutorPlacementType.AUTOMATIC.value,
                    ))
        self.db.flush()
        return tuple(rows)

    def generate_schedule(self) -> TutorSchedule:
        """Replace automatic assignments while preserving manual decisions."""
        scheduling_input = self.get_scheduling_input()
        candidate_set = generate_tutor_candidates(scheduling_input)
        readiness = analyze_tutor_readiness(scheduling_input, candidate_set)
        if not readiness.is_ready:
            raise TutorValidationError(
                "NOT_READY", "Tutor Scheduling inputs are not ready for schedule generation."
            )

        decisions = self.list_placements()
        manual_ids = {
            item.tutor_student_id for item in decisions
            if item.placement_type in {
                TutorPlacementType.MANUAL.value,
                TutorPlacementType.MANUALLY_UNPLACED.value,
            }
        }
        fixed = tuple(
            item for item in self._stored_assignments(scheduling_input)
            if item.placement_type == TutorPlacementType.MANUAL.value
        )
        manually_unplaced_ids = frozenset(
            item.tutor_student_id for item in decisions
            if item.placement_type == TutorPlacementType.MANUALLY_UNPLACED.value
        )
        optimized = optimize_tutor_schedule(
            scheduling_input,
            candidate_set,
            fixed_placements=tuple(
                TutorScheduledPlacement(
                    student_id=item.student_id,
                    tutor_id=item.tutor_id,
                    day=item.day,
                    start_minute=item.start_minute,
                    end_minute=item.end_minute,
                    preferred_tag_match_count=item.preferred_tag_match_count,
                    tentative_minutes=item.tentative_minutes,
                )
                for item in fixed
            ),
            excluded_student_ids=manually_unplaced_ids,
        )
        session_uuid = self._session_uuid()
        state = self.get_state()
        self.db.query(TutorPlacement).filter(
            TutorPlacement.placement_type == TutorPlacementType.AUTOMATIC.value
        ).delete()
        for placement in optimized.placements:
            if placement.student_id in manual_ids:
                continue
            self.db.add(TutorPlacement(
                tutor_student_id=placement.student_id,
                tutor_tutor_id=placement.tutor_id,
                day=placement.day,
                start_minute=placement.start_minute,
                end_minute=placement.end_minute,
                placement_type=TutorPlacementType.AUTOMATIC.value,
                created_at=_now(),
                updated_at=_now(),
            ))
        assignment_rows = self._refresh_schedule_snapshots(
            scheduling_input, automatic_unplaced=optimized.unplaced_students
        )
        _, _, _, entries, _ = self._schedule_review(scheduling_input)
        statistics = self._schedule_statistics(scheduling_input, assignment_rows, entries)
        optimized = TutorOptimizationResult(
            source_data_revision=optimized.source_data_revision,
            placements=optimized.placements,
            unplaced_students=optimized.unplaced_students,
            statistics=statistics,
        )
        self._save_schedule_result(optimized, state.schedule_revision + 1, session_uuid)
        state.schedule_revision += 1
        state.is_schedule_stale = False
        state.modified_at = _now()
        self.db.flush()
        return self.get_schedule()

    def _save_schedule_result(
        self, result: TutorOptimizationResult, schedule_revision: int, session_uuid: str
    ) -> None:
        stats = result.statistics
        row = self.db.get(TutorScheduleResult, session_uuid)
        if row is None:
            row = TutorScheduleResult(session_uuid=session_uuid)
            self.db.add(row)
        row.schedule_revision = schedule_revision
        row.source_data_revision = result.source_data_revision
        row.students_scheduled = stats.students_scheduled
        row.students_unscheduled = stats.students_unscheduled
        row.tutor_utilization = stats.tutor_utilization
        row.preferred_tag_matches = stats.preferred_tag_matches
        row.tentative_minutes = stats.tentative_minutes
        row.gap_cost_minutes = stats.gap_cost_minutes
        row.modified_at = _now()

    def get_schedule(self) -> TutorSchedule:
        """Retrieve assignments, status, verification, and summary from current snapshots."""
        state = self.get_state()
        result = self.db.get(TutorScheduleResult, self._session_uuid())
        scheduling_input, assignments, _, entries, verification = self._schedule_review()
        placements = tuple(
            TutorScheduledPlacement(
                student_id=item.student_id,
                tutor_id=item.tutor_id,
                day=item.day,
                start_minute=item.start_minute,
                end_minute=item.end_minute,
                preferred_tag_match_count=item.preferred_tag_match_count,
                tentative_minutes=item.tentative_minutes,
            )
            for item in assignments
        )
        unplaced_students = tuple(
            TutorUnplacedStudentResult(
                student_id=item.student_id,
                reason=TutorUnplacedReason(item.reason),
            )
            for item in self._stored_unplaced()
        )
        has_schedule = (
            result is not None
            or bool(self.list_placements())
            or bool(self.db.query(TutorUnplacedStudent.student_id).first())
        )
        statistics = (
            self._schedule_statistics(scheduling_input, assignments, entries)
            if has_schedule else TutorScheduleStatistics(0, 0, 0.0, 0, 0, 0)
        )
        source_data_revision = (
            result.source_data_revision if result is not None else state.source_data_revision
        )
        verification_count = sum(item.verified is False for item in entries)
        return TutorSchedule(
            schedule_revision=result.schedule_revision if result is not None else state.schedule_revision,
            source_data_revision=source_data_revision,
            is_schedule_stale=(
                state.is_schedule_stale
                or source_data_revision != state.source_data_revision
            ),
            has_schedule=has_schedule,
            placements=placements,
            unplaced_students=unplaced_students,
            statistics=statistics,
            entries=entries,
            verification=verification,
            automatic_count=sum(
                item.placement_type == TutorPlacementType.AUTOMATIC.value and item.day is not None
                for item in entries
            ),
            manual_count=sum(
                item.placement_type == TutorPlacementType.MANUAL.value and item.day is not None
                for item in entries
            ),
            manually_unplaced_count=sum(
                item.placement_type == TutorPlacementType.MANUALLY_UNPLACED.value for item in entries
            ),
            verified_count=sum(item.verified is True for item in entries),
            unverified_count=verification_count,
        )

    def clear_schedule(self) -> TutorSchedule:
        """Clear automatic assignments while retaining manual decisions and their state."""
        session_uuid = self._session_uuid()
        self.db.query(TutorPlacement).filter(
            TutorPlacement.placement_type == TutorPlacementType.AUTOMATIC.value
        ).delete()
        self.db.query(TutorUnplacedStudent).filter(
            TutorUnplacedStudent.placement_type == TutorPlacementType.AUTOMATIC.value
        ).delete()
        state = self.get_state()
        state.schedule_revision += 1
        state.modified_at = _now()
        assignments = self._refresh_schedule_snapshots(self.get_scheduling_input())
        if assignments or self.db.query(TutorUnplacedStudent.student_id).first() is not None:
            scheduling_input, _, _, entries, _ = self._schedule_review()
            stats = self._schedule_statistics(scheduling_input, assignments, entries)
            result = TutorOptimizationResult(
                source_data_revision=(
                    self.db.get(TutorScheduleResult, session_uuid).source_data_revision
                    if self.db.get(TutorScheduleResult, session_uuid) is not None
                    else state.source_data_revision
                ),
                placements=(),
                unplaced_students=(),
                statistics=stats,
            )
            self._save_schedule_result(result, state.schedule_revision, session_uuid)
        else:
            self.db.query(TutorScheduleResult).filter(
                TutorScheduleResult.session_uuid == session_uuid
            ).delete()
            state.is_schedule_stale = False
        self.db.flush()
        return self.get_schedule()

    def set_tag_state(self, tutor_student_id: str, tag_id: str, state: TutorTagState) -> None:
        """Persist a student/tag state; Neutral is represented by the absence of a row."""
        self.get_student(tutor_student_id)
        self._require(TutorTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
        try:
            state = TutorTagState(state)
        except ValueError as error:
            raise TutorValidationError("INVALID_TAG_STATE", "Choose a valid tag state.") from error
        row = self.db.get(TutorStudentTagState, (tutor_student_id, tag_id))
        if state is TutorTagState.NEUTRAL:
            if row is not None:
                self.db.delete(row)
        elif row is None:
            self.db.add(TutorStudentTagState(student_id=tutor_student_id, tag_id=tag_id, state=state.value))
        else:
            row.state = state.value
        self.db.flush()
        self.record_source_change()

    def _validated_tag_states(self, states: Mapping[str, TutorTagState | str]) -> dict[str, TutorTagState]:
        validated: dict[str, TutorTagState] = {}
        for tag_id, raw in states.items():
            self._require(TutorTag, tag_id, "TAG_NOT_FOUND", "The tag no longer exists.")
            try:
                validated[tag_id] = TutorTagState(raw)
            except ValueError as error:
                raise TutorValidationError("INVALID_TAG_STATE", "Choose a valid tag state.") from error
        return validated

    def _replace_student_tag_states(self, tutor_student_id: str, states: Mapping[str, TutorTagState]) -> None:
        """Store only non-neutral states; Neutral is the absence of a row."""
        self.db.query(TutorStudentTagState).filter(TutorStudentTagState.student_id == tutor_student_id).delete()
        for tag_id, state in states.items():
            if state is not TutorTagState.NEUTRAL:
                self.db.add(TutorStudentTagState(student_id=tutor_student_id, tag_id=tag_id, state=state.value))
        self.db.flush()

    def set_student_tag_states(
        self, tutor_student_id: str, states: Mapping[str, TutorTagState | str]
    ) -> TutorStudent:
        """Replace a student's whole tag-state set as a single source-data change."""
        student = self.get_student(tutor_student_id)
        self._replace_student_tag_states(student.id, self._validated_tag_states(states))
        self.record_source_change()
        return student

    def student_availability_counts(self) -> dict[str, int]:
        return dict(
            self.db.query(
                TutorStudentAvailabilityWindow.tutor_student_id, func.count(TutorStudentAvailabilityWindow.id)
            )
            .group_by(TutorStudentAvailabilityWindow.tutor_student_id)
            .all()
        )

    def replace_students(
        self,
        students: Iterable[tuple[StudentInput, Sequence[AvailabilityWindow], Sequence[tuple[str, TutorTagState | str]]]],
    ) -> TutorStudentReplaceResult:
        """Full replacement of students as one source-data change (caller owns the transaction).

        Existing placement decisions are retained and the resulting schedule becomes stale.
        Student availability and tag state are replaced and imported tags missing from the registry are created.
        """
        prepared: list[tuple[tuple[str, str, str], list[AvailabilityWindow], list[tuple[str, TutorTagState]]]] = []
        seen: set[str] = set()
        for values, windows, tag_pairs in students:
            # Existing students are about to be replaced, so only uniqueness within the batch matters.
            cleaned = self._clean_student_values(
                values.student_id, values.name, values.email, check_unique=False
            )
            if cleaned[0].casefold() in seen:
                raise TutorValidationError("DUPLICATE_STUDENT_ID", "A student with that Student ID already exists.")
            seen.add(cleaned[0].casefold())
            states = [(name, TutorTagState(state)) for name, state in tag_pairs]
            prepared.append((cleaned, self._validated_windows(windows), states))
        students_removed = self.db.query(func.count(TutorStudent.id)).scalar() or 0
        placements_removed = 0
        placements_preserved = self.db.query(func.count(TutorPlacement.id)).scalar() or 0
        self.db.query(TutorStudentTagState).delete()
        self.db.query(TutorStudentAvailabilityWindow).delete()
        self.db.query(TutorStudent).delete()
        self.db.flush()
        tags: dict[str, TutorTag] = {}
        tags_created = 0
        for (student_id, name, email), windows, states in prepared:
            student = TutorStudent(student_id=student_id, name=name, email=email)
            self.db.add(student)
            self.db.flush()
            self._replace_student_availability(student.id, windows)
            resolved: dict[str, TutorTagState] = {}
            for tag_name, state in states:
                key = tutor_tag_identity_key(tag_name)
                if not key:
                    continue
                tag = tags.get(key) or self.find_tag(tag_name)
                if tag is None:
                    tag = self._insert_tag(normalize_tutor_tag_name(tag_name))
                    tags_created += 1
                tags[key] = tag
                resolved[tag.id] = state
            self._replace_student_tag_states(student.id, resolved)
        self.record_source_change()
        return TutorStudentReplaceResult(
            students_removed=students_removed,
            placements_removed=placements_removed,
            placements_preserved=placements_preserved,
            students_created=len(prepared),
            tags_created=tags_created,
        )

    def get_tag_state(self, tutor_student_id: str, tag_id: str) -> TutorTagState:
        row = self.db.get(TutorStudentTagState, (tutor_student_id, tag_id))
        return TutorTagState(row.state) if row is not None else TutorTagState.NEUTRAL

    # --- placements ---------------------------------------------------

    def set_placement(
        self,
        tutor_student_id: str,
        placement_type: TutorPlacementType,
        tutor_tutor_id: str | None = None,
        start_minute: int | None = None,
        day: str = "Monday",
    ) -> TutorPlacement:
        """Create or replace the student's single placement decision."""
        self.get_student(tutor_student_id)
        try:
            placement_type = TutorPlacementType(placement_type)
        except ValueError as error:
            raise TutorValidationError("INVALID_PLACEMENT", "Choose a valid placement type.") from error
        if placement_type is TutorPlacementType.MANUALLY_UNPLACED:
            if tutor_tutor_id is not None or start_minute is not None:
                raise TutorValidationError(
                    "INVALID_PLACEMENT", "A manually unplaced decision has no tutor or start time."
                )
            end_minute = None
            day = "Monday"
        else:
            if tutor_tutor_id is None or start_minute is None:
                raise TutorValidationError("INVALID_PLACEMENT", "A placement requires a tutor and a start time.")
            if not _is_int(start_minute) or not 0 <= start_minute < 24 * 60:
                raise TutorValidationError("INVALID_PLACEMENT", "The start time is not a valid time of day.")
            normalized_day = normalize_day(day)
            if normalized_day is None:
                raise TutorValidationError("INVALID_PLACEMENT", "Choose a valid day of the week.")
            day = normalized_day
            tutor = self.get_tutor(tutor_tutor_id)
            end_minute = start_minute + tutor.lesson_length_minutes
            if end_minute > 24 * 60:
                raise TutorValidationError(
                    "INVALID_PLACEMENT", "The session must end before the end of the day."
                )
        placement = self.get_placement(tutor_student_id)
        if placement is None:
            placement = TutorPlacement(tutor_student_id=tutor_student_id)
            self.db.add(placement)
        placement.placement_type = placement_type.value
        placement.tutor_tutor_id = tutor_tutor_id
        placement.day = day
        placement.start_minute = start_minute
        placement.end_minute = end_minute
        placement.updated_at = _now()
        self.db.flush()
        self._record_schedule_edit()
        self._persist_schedule_after_edit()
        return placement

    def set_manual_placement(
        self, tutor_student_id: str, tutor_id: str, day: str, start_minute: int
    ) -> TutorPlacement:
        return self.set_placement(
            tutor_student_id,
            TutorPlacementType.MANUAL,
            tutor_id,
            start_minute,
            day,
        )

    def set_manually_unplaced(self, tutor_student_id: str) -> TutorPlacement:
        return self.set_placement(tutor_student_id, TutorPlacementType.MANUALLY_UNPLACED)

    def get_placement(self, tutor_student_id: str) -> TutorPlacement | None:
        return (
            self.db.query(TutorPlacement)
            .filter(TutorPlacement.tutor_student_id == tutor_student_id)
            .one_or_none()
        )

    def list_placements(self) -> list[TutorPlacement]:
        return self.db.query(TutorPlacement).order_by(TutorPlacement.created_at, TutorPlacement.id).all()

    def remove_placement(self, tutor_student_id: str) -> bool:
        placement = self.get_placement(tutor_student_id)
        if placement is None:
            raise TutorValidationError("PLACEMENT_NOT_FOUND", "The student has no placement decision to remove.")
        self.db.delete(placement)
        self.db.query(TutorUnplacedStudent).filter(
            TutorUnplacedStudent.student_id == tutor_student_id,
            TutorUnplacedStudent.placement_type == placement.placement_type,
        ).delete()
        self.db.flush()
        self._record_schedule_edit()
        self._persist_schedule_after_edit()
        return True

    def remove_manual_override(self, tutor_student_id: str) -> bool:
        placement = self.get_placement(tutor_student_id)
        if placement is None or placement.placement_type == TutorPlacementType.AUTOMATIC.value:
            raise TutorValidationError(
                "NO_MANUAL_OVERRIDE", "The student does not have a manual placement override to remove."
            )
        return self.remove_placement(tutor_student_id)

    def _persist_schedule_after_edit(self) -> None:
        scheduling_input = self.get_scheduling_input()
        existing_result = self.db.get(TutorScheduleResult, self._session_uuid())
        assignments = self._refresh_schedule_snapshots(scheduling_input)
        _, _, _, entries, _ = self._schedule_review(scheduling_input)
        if not self.list_placements() and not self.db.query(TutorUnplacedStudent.student_id).first():
            if existing_result is not None:
                self.db.delete(existing_result)
            return
        statistics = self._schedule_statistics(scheduling_input, assignments, entries)
        result = TutorOptimizationResult(
            source_data_revision=(
                existing_result.source_data_revision
                if existing_result is not None else scheduling_input.source_data_revision
            ),
            placements=tuple(
                TutorScheduledPlacement(
                    student_id=item.student_id,
                    tutor_id=item.tutor_id,
                    day=item.day,
                    start_minute=item.start_minute,
                    end_minute=item.end_minute,
                    preferred_tag_match_count=item.preferred_tag_match_count,
                    tentative_minutes=item.tentative_minutes,
                )
                for item in assignments
            ),
            unplaced_students=(),
            statistics=statistics,
        )
        state = self.get_state()
        self._save_schedule_result(result, state.schedule_revision, self._session_uuid())
        state.is_schedule_stale = (
            state.is_schedule_stale
            or result.source_data_revision != state.source_data_revision
        )

    def clear_all_placements(self) -> int:
        """Delete every placement decision and reset result/staleness state."""
        removed = self.db.query(TutorPlacement).delete()
        self.db.query(TutorSchedulePlacement).delete()
        self.db.query(TutorUnplacedStudent).delete()
        self.db.query(TutorScheduleResult).delete()
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
            raise TutorValidationError(code, message)
        return row


def _window(row) -> AvailabilityWindow:
    return AvailabilityWindow(
        day=row.day,
        start_minute=row.start_minute,
        end_minute=row.end_minute,
        status=row.status,
        source=row.source,
        source_row=row.source_row,
    )
