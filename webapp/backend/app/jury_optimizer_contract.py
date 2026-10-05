"""Pure-domain DTOs and behavioral contract for a future Jury optimizer."""

from dataclasses import dataclass
from datetime import date
from enum import StrEnum
from typing import Protocol
from uuid import UUID


MINUTES_PER_DAY = 24 * 60
ACCOMPANIST_RESULT_CONTRACT = "accompanist.assignment-result"


@dataclass(frozen=True)
class AccompanistResultReference:
    session_uuid: UUID
    result_uuid: UUID
    contract_id: str
    contract_version: int
    source_revision: int
    state: str

    def __post_init__(self) -> None:
        if self.contract_id != ACCOMPANIST_RESULT_CONTRACT or self.contract_version != 2:
            raise ValueError("The Jury optimizer accepts only Accompanist result contract v2.")
        if self.state != "finalized":
            raise ValueError("The Jury optimizer accepts only finalized Accompanist results.")
        if self.source_revision < 0:
            raise ValueError("Accompanist source revision cannot be negative.")


@dataclass(frozen=True)
class ReadinessApproval:
    session_uuid: UUID
    source_result_uuid: UUID
    source_revision: int
    jury_input_revision: int
    approved: bool

    def __post_init__(self) -> None:
        if not self.approved:
            raise ValueError("Jury optimizer inputs must have readiness approval.")
        if self.source_revision < 0 or self.jury_input_revision < 0:
            raise ValueError("Readiness revisions cannot be negative.")


@dataclass(frozen=True)
class FinalizedPianistFact:
    person_uuid: UUID
    display_name: str


@dataclass(frozen=True)
class FinalizedAccompanistAssignmentFact:
    source_lesson_uuid: UUID
    student_person_uuid: UUID
    student_display_name: str
    instrument: str
    teacher: str
    pianist_required: bool
    jury_required: bool
    assigned_pianist: FinalizedPianistFact | None


@dataclass(frozen=True)
class ScheduleEntry:
    panel_uuid: UUID
    jury_date: date
    start_minute: int
    end_minute: int

    def __post_init__(self) -> None:
        if not 0 <= self.start_minute < self.end_minute <= MINUTES_PER_DAY:
            raise ValueError("Schedule entry must fit within the Jury day.")


@dataclass(frozen=True)
class ScheduledLesson(ScheduleEntry):
    source_lesson_uuid: UUID
    student_person_uuid: UUID
    pianist_person_uuid: UUID | None


@dataclass(frozen=True)
class BreakEntry(ScheduleEntry):
    after_jury_count: int

    def __post_init__(self) -> None:
        super().__post_init__()
        if self.after_jury_count <= 0:
            raise ValueError("Periodic Break must follow at least one Jury.")


@dataclass(frozen=True)
class MealBreakEntry(ScheduleEntry):
    pass


@dataclass(frozen=True)
class JuryLessonEntry:
    source_lesson_uuid: UUID
    student_person_uuid: UUID
    panel_uuid: UUID | None


@dataclass(frozen=True)
class JuryPanelDate:
    panel_uuid: UUID
    jury_date: date


@dataclass(frozen=True)
class MealBreakConfiguration:
    start_minute: int
    end_minute: int

    def __post_init__(self) -> None:
        if not 0 <= self.start_minute < self.end_minute <= MINUTES_PER_DAY:
            raise ValueError("Meal Break must be a valid interval within the Jury day.")


@dataclass(frozen=True)
class PeriodicBreakConfiguration:
    every_x_juries: int
    length_minutes: int

    def __post_init__(self) -> None:
        if self.every_x_juries <= 0 or self.length_minutes <= 0:
            raise ValueError("Periodic Break count and length must be positive.")


@dataclass(frozen=True)
class JuryPanelTimingConfiguration:
    panel_uuid: UUID
    earliest_start_minute: int
    preferred_start_minute: int
    jury_length_minutes: int
    meal_break: MealBreakConfiguration | None = None
    periodic_break: PeriodicBreakConfiguration | None = None

    def __post_init__(self) -> None:
        if not 0 <= self.earliest_start_minute < MINUTES_PER_DAY:
            raise ValueError("Earliest Start must be within the Jury day.")
        if not self.earliest_start_minute <= self.preferred_start_minute < MINUTES_PER_DAY:
            raise ValueError("Preferred Start must be on or after Earliest Start and within the Jury day.")
        if not 0 < self.jury_length_minutes <= MINUTES_PER_DAY:
            raise ValueError("Jury Length must be a positive number of minutes no longer than one day.")
        if self.earliest_start_minute + self.jury_length_minutes > MINUTES_PER_DAY:
            raise ValueError("Jury Length must fit within the Jury day from Earliest Start.")


@dataclass(frozen=True)
class JuryAvailabilityWindow:
    pianist_person_uuid: UUID
    jury_date: date
    start_minute: int
    end_minute: int

    def __post_init__(self) -> None:
        if not 0 <= self.start_minute < self.end_minute <= MINUTES_PER_DAY:
            raise ValueError("Jury Availability Window must be a valid interval within the Jury day.")


@dataclass(frozen=True)
class JuryOptimizerInput:
    accompanist_result: AccompanistResultReference
    readiness: ReadinessApproval
    accompanist_assignments: tuple[FinalizedAccompanistAssignmentFact, ...]
    jury_lesson_entries: tuple[JuryLessonEntry, ...]
    panel_dates: tuple[JuryPanelDate, ...]
    panel_timings: tuple[JuryPanelTimingConfiguration, ...]
    pianist_availability_windows: tuple[JuryAvailabilityWindow, ...]

    def __post_init__(self) -> None:
        result = self.accompanist_result
        approval = self.readiness
        if (
            approval.session_uuid != result.session_uuid
            or approval.source_result_uuid != result.result_uuid
            or approval.source_revision != result.source_revision
        ):
            raise ValueError("Readiness approval does not match the finalized Accompanist result.")

        facts_by_lesson = {fact.source_lesson_uuid: fact for fact in self.accompanist_assignments}
        entries_by_lesson = {entry.source_lesson_uuid: entry for entry in self.jury_lesson_entries}
        dates_by_panel = {item.panel_uuid: item.jury_date for item in self.panel_dates}
        timings_by_panel = {item.panel_uuid: item for item in self.panel_timings}
        if len(facts_by_lesson) != len(self.accompanist_assignments):
            raise ValueError("Finalized Accompanist facts must have unique Lesson UUIDs.")
        if len(entries_by_lesson) != len(self.jury_lesson_entries):
            raise ValueError("Jury lesson entries must have unique Lesson UUIDs.")
        if not entries_by_lesson.keys() <= facts_by_lesson.keys():
            raise ValueError("Jury lesson entries must reference finalized source Lessons.")
        if len(dates_by_panel) != len(self.panel_dates) or len(timings_by_panel) != len(self.panel_timings):
            raise ValueError("Panel dates and timing configurations must have unique Panel UUIDs.")
        if dates_by_panel.keys() != timings_by_panel.keys():
            raise ValueError("Every configured Jury Panel must have exactly one Jury Date.")

        for lesson_uuid, fact in facts_by_lesson.items():
            entry = entries_by_lesson.get(lesson_uuid)
            if fact.jury_required and entry is None:
                raise ValueError("Every Jury-required lesson must have a Jury lesson entry.")
            if entry is not None and entry.student_person_uuid != fact.student_person_uuid:
                raise ValueError("Jury lesson identity must match its finalized source lesson.")
            if fact.jury_required and (entry is None or entry.panel_uuid is None):
                raise ValueError("Every Jury-required lesson must have an authoritative Panel assignment.")
            if entry is not None and entry.panel_uuid is not None and entry.panel_uuid not in dates_by_panel:
                raise ValueError("Jury lesson Panel assignment must reference a configured Panel.")
            if not fact.jury_required or not fact.pianist_required:
                continue
            if fact.assigned_pianist is None:
                raise ValueError("A Jury-required lesson requiring a pianist must have a fixed assignment.")
            assert entry is not None and entry.panel_uuid is not None
            jury_date = dates_by_panel[entry.panel_uuid]
            if not any(
                window.pianist_person_uuid == fact.assigned_pianist.person_uuid
                and window.jury_date == jury_date
                for window in self.pianist_availability_windows
            ):
                raise ValueError("Readiness-approved inputs must include the fixed pianist's Jury Availability Windows.")

        availability_keys = [(window.pianist_person_uuid, window.jury_date) for window in self.pianist_availability_windows]
        for key in set(availability_keys):
            windows = sorted(
                (window for window in self.pianist_availability_windows if (window.pianist_person_uuid, window.jury_date) == key),
                key=lambda window: (window.start_minute, window.end_minute),
            )
            if any(current.start_minute < previous.end_minute for previous, current in zip(windows, windows[1:])):
                raise ValueError("Jury Availability Windows for one Pianist and date must not overlap.")


class UnscheduledReasonCode(StrEnum):
    NO_FEASIBLE_INTERVAL = "no_feasible_interval"
    FIXED_PIANIST_CONFLICT = "fixed_pianist_conflict"
    STUDENT_CONFLICT = "student_conflict"
    AVAILABILITY_TOO_SHORT = "availability_too_short"
    DAY_BOUND_EXCEEDED = "day_bound_exceeded"


class DiagnosticSeverity(StrEnum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


@dataclass(frozen=True)
class ScheduledJuryEntry:
    source_lesson_uuid: UUID
    student_person_uuid: UUID
    panel_uuid: UUID
    jury_date: date
    pianist_person_uuid: UUID | None
    start_minute: int
    end_minute: int

    def __post_init__(self) -> None:
        if not 0 <= self.start_minute < self.end_minute <= MINUTES_PER_DAY:
            raise ValueError("Scheduled Jury entry must fit within the Jury day.")


@dataclass(frozen=True)
class MealBreakEntry:
    panel_uuid: UUID
    jury_date: date
    start_minute: int
    end_minute: int

    def __post_init__(self) -> None:
        if not 0 <= self.start_minute < self.end_minute <= MINUTES_PER_DAY:
            raise ValueError("Meal-Break entry must fit within the Jury day.")


@dataclass(frozen=True)
class PeriodicBreakEntry:
    panel_uuid: UUID
    jury_date: date
    start_minute: int
    end_minute: int

    def __post_init__(self) -> None:
        if not 0 <= self.start_minute < self.end_minute <= MINUTES_PER_DAY:
            raise ValueError("Periodic-Break entry must fit within the Jury day.")


@dataclass(frozen=True)
class UnscheduledLesson:
    source_lesson_uuid: UUID
    reason_code: UnscheduledReasonCode
    explanation: str

    def __post_init__(self) -> None:
        if not self.explanation.strip():
            raise ValueError("Unscheduled Jury entries require an explicit explanation.")


@dataclass(frozen=True)
class OptimizerDiagnostic:
    code: str
    severity: DiagnosticSeverity
    message: str


@dataclass(frozen=True)
class ConflictExplanation:
    code: str
    message: str
    source_lesson_uuids: tuple[UUID, ...] = ()
    person_uuid: UUID | None = None
    panel_uuid: UUID | None = None


@dataclass(frozen=True)
class JuryScheduleResult:
    session_uuid: UUID
    source_result_uuid: UUID
    source_contract_id: str
    source_contract_version: int
    source_revision: int
    jury_input_revision: int
    required_lesson_uuids: tuple[UUID, ...]
    schedule_entries: tuple[ScheduleEntry, ...]
    unscheduled_lessons: tuple[UnscheduledLesson, ...]
    diagnostics: tuple[OptimizerDiagnostic, ...] = ()
    conflict_explanations: tuple[ConflictExplanation, ...] = ()

    def __post_init__(self) -> None:
        required = set(self.required_lesson_uuids)
        if len(required) != len(self.required_lesson_uuids):
            raise ValueError("Required Lesson UUIDs must be unique.")
        if any(not isinstance(entry, (ScheduledLesson, BreakEntry, MealBreakEntry)) for entry in self.schedule_entries):
            raise TypeError("Schedule entries must be Jury, Periodic-Break, or Meal-Break DTOs.")
        scheduled = [entry.source_lesson_uuid for entry in self.schedule_entries if isinstance(entry, ScheduledLesson)]
        unscheduled = [entry.source_lesson_uuid for entry in self.unscheduled_lessons]
        if len(scheduled) != len(set(scheduled)) or len(unscheduled) != len(set(unscheduled)):
            raise ValueError("Each Jury-required lesson may have only one outcome.")
        if set(scheduled) & set(unscheduled):
            raise ValueError("A Jury-required lesson cannot be both scheduled and unscheduled.")
        if set(scheduled) | set(unscheduled) != required:
            raise ValueError("Every Jury-required lesson must be scheduled or explicitly unscheduled.")
        if self.source_revision < 0 or self.jury_input_revision < 0:
            raise ValueError("Result revisions cannot be negative.")
        if self.source_contract_id != ACCOMPANIST_RESULT_CONTRACT or self.source_contract_version != 2:
            raise ValueError("Jury result provenance must reference Accompanist result contract v2.")
        events_by_panel: dict[tuple[UUID, date], list[ScheduleEntry]] = {}
        for entry in self.schedule_entries:
            events_by_panel.setdefault((entry.panel_uuid, entry.jury_date), []).append(entry)
        for events in events_by_panel.values():
            events.sort(key=lambda entry: (entry.start_minute, entry.end_minute))
            if any(current.start_minute < previous.end_minute for previous, current in zip(events, events[1:])):
                raise ValueError("Schedule entries for one Panel and date must not overlap.")

    @property
    def scheduled_lessons(self) -> tuple[ScheduledLesson, ...]:
        return tuple(entry for entry in self.schedule_entries if isinstance(entry, ScheduledLesson))

    @property
    def meal_break_entries(self) -> tuple[MealBreakEntry, ...]:
        return tuple(entry for entry in self.schedule_entries if isinstance(entry, MealBreakEntry))

    @property
    def periodic_break_entries(self) -> tuple[BreakEntry, ...]:
        return tuple(entry for entry in self.schedule_entries if isinstance(entry, BreakEntry))


# Names retained for the Phase 1 DTOs while callers move to the Phase 2 result vocabulary.
ScheduledJuryEntry = ScheduledLesson
UnscheduledJuryEntry = UnscheduledLesson
PeriodicBreakEntry = BreakEntry
JuryOptimizerResult = JuryScheduleResult


class JuryOptimizer(Protocol):
    """Pure-domain scheduling operation over readiness-approved DTOs."""

    def optimize(self, inputs: JuryOptimizerInput) -> JuryScheduleResult:
        ...