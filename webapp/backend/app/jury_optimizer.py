"""Deterministic, persistence-independent Jury scheduling core."""

from dataclasses import dataclass
from datetime import date
from uuid import UUID

from .jury_optimizer_contract import (
    MINUTES_PER_DAY,
    BreakEntry,
    ConflictExplanation,
    DiagnosticSeverity,
    FinalizedAccompanistAssignmentFact,
    JuryAvailabilityWindow,
    JuryOptimizerInput,
    JuryPanelTimingConfiguration,
    JuryScheduleResult,
    MealBreakEntry,
    OptimizerDiagnostic,
    ScheduleEntry,
    ScheduledLesson,
    UnscheduledLesson,
    UnscheduledReasonCode,
)


@dataclass
class _PanelState:
    panel_uuid: UUID
    jury_date: date
    timing: JuryPanelTimingConfiguration
    lessons: list[FinalizedAccompanistAssignmentFact]
    cursor_minute: int = 0
    juries_since_break: int = 0


@dataclass(frozen=True)
class _Placement:
    start_minute: int
    end_minute: int
    events_before: tuple[ScheduleEntry, ...]
    juries_since_break: int


@dataclass(frozen=True)
class _PlacementFailure:
    reason: UnscheduledReasonCode
    blocking_lesson_uuids: tuple[UUID, ...] = ()


class DeterministicJuryOptimizer:
    """Schedule each Panel on a finite minute horizon with fixed Pianist bookings."""

    def optimize(self, inputs: JuryOptimizerInput) -> JuryScheduleResult:
        facts_by_lesson = {fact.source_lesson_uuid: fact for fact in inputs.accompanist_assignments}
        entries_by_lesson = {entry.source_lesson_uuid: entry for entry in inputs.jury_lesson_entries}
        dates_by_panel = {item.panel_uuid: item.jury_date for item in inputs.panel_dates}
        timings_by_panel = {item.panel_uuid: item for item in inputs.panel_timings}
        windows_by_person_date: dict[tuple[UUID, date], list[JuryAvailabilityWindow]] = {}
        for window in inputs.pianist_availability_windows:
            windows_by_person_date.setdefault((window.pianist_person_uuid, window.jury_date), []).append(window)
        for windows in windows_by_person_date.values():
            windows.sort(key=lambda item: (item.start_minute, item.end_minute))

        lessons_by_panel: dict[UUID, list[FinalizedAccompanistAssignmentFact]] = {
            panel_uuid: [] for panel_uuid in timings_by_panel
        }
        for fact in inputs.accompanist_assignments:
            if not fact.jury_required:
                continue
            lesson_entry = entries_by_lesson[fact.source_lesson_uuid]
            lessons_by_panel[lesson_entry.panel_uuid].append(fact)

        panel_states: list[_PanelState] = []
        for panel_uuid in sorted(timings_by_panel, key=lambda item: item.hex):
            timing = timings_by_panel[panel_uuid]
            jury_date = dates_by_panel[panel_uuid]
            panel_lessons = lessons_by_panel[panel_uuid]
            panel_lessons.sort(key=lambda fact: fact.source_lesson_uuid.hex)
            panel_states.append(_PanelState(
                panel_uuid=panel_uuid,
                jury_date=jury_date,
                timing=timing,
                lessons=panel_lessons,
                cursor_minute=timing.earliest_start_minute,
            ))

        schedule_entries: list[ScheduleEntry] = []
        unscheduled: list[UnscheduledLesson] = []
        diagnostics: list[OptimizerDiagnostic] = []
        conflicts: list[ConflictExplanation] = []
        bookings: dict[tuple[UUID, date], list[ScheduledLesson]] = {}
        student_bookings: dict[tuple[UUID, date], list[ScheduledLesson]] = {}

        while any(state.lessons for state in panel_states):
            candidates = []
            made_progress = False
            for state in panel_states:
                for fact in tuple(state.lessons):
                    placement = self._find_placement(
                        fact,
                        state,
                        windows_by_person_date,
                        bookings,
                        student_bookings,
                    )
                    if isinstance(placement, _PlacementFailure):
                        unscheduled.append(self._unscheduled(fact, placement.reason))
                        diagnostics.append(OptimizerDiagnostic(
                            code=placement.reason.value,
                            severity=DiagnosticSeverity.WARNING,
                            message=self._reason_message(placement.reason),
                        ))
                        if placement.reason == UnscheduledReasonCode.FIXED_PIANIST_CONFLICT:
                            conflicts.append(ConflictExplanation(
                                code=placement.reason.value,
                                message=self._reason_message(placement.reason),
                                source_lesson_uuids=(fact.source_lesson_uuid, *placement.blocking_lesson_uuids),
                                person_uuid=fact.assigned_pianist.person_uuid if fact.assigned_pianist else None,
                                panel_uuid=state.panel_uuid,
                            ))
                        if placement.reason == UnscheduledReasonCode.STUDENT_CONFLICT:
                            conflicts.append(ConflictExplanation(
                                code=placement.reason.value,
                                message=self._reason_message(placement.reason),
                                source_lesson_uuids=(fact.source_lesson_uuid, *placement.blocking_lesson_uuids),
                                person_uuid=fact.student_person_uuid,
                                panel_uuid=state.panel_uuid,
                            ))
                        state.lessons.remove(fact)
                        made_progress = True
                        continue
                    candidates.append((
                        placement.start_minute,
                        placement.end_minute,
                        self._lesson_flexibility(fact, state, windows_by_person_date),
                        state.panel_uuid.hex,
                        fact.source_lesson_uuid.hex,
                        state,
                        fact,
                        placement,
                    ))

            if not candidates:
                if not made_progress:
                    raise RuntimeError("Jury optimizer failed to advance its finite lesson cursor.")
                continue

            _, _, _, _, _, state, fact, placement = min(candidates, key=lambda item: item[:5])
            schedule_entries.extend(placement.events_before)
            scheduled = ScheduledLesson(
                panel_uuid=state.panel_uuid,
                jury_date=state.jury_date,
                start_minute=placement.start_minute,
                end_minute=placement.end_minute,
                source_lesson_uuid=fact.source_lesson_uuid,
                student_person_uuid=fact.student_person_uuid,
                pianist_person_uuid=fact.assigned_pianist.person_uuid if fact.assigned_pianist else None,
            )
            schedule_entries.append(scheduled)
            student_bookings.setdefault((scheduled.student_person_uuid, scheduled.jury_date), []).append(scheduled)
            if scheduled.pianist_person_uuid is not None:
                bookings.setdefault((scheduled.pianist_person_uuid, scheduled.jury_date), []).append(scheduled)
            state.cursor_minute = placement.end_minute
            state.juries_since_break = placement.juries_since_break + 1
            state.lessons.remove(fact)

        schedule_entries.sort(key=self._event_order_key)
        unscheduled.sort(key=lambda item: item.source_lesson_uuid.hex)
        required_lesson_uuids = tuple(sorted(
            (fact.source_lesson_uuid for fact in inputs.accompanist_assignments if fact.jury_required),
            key=lambda item: item.hex,
        ))
        return JuryScheduleResult(
            session_uuid=inputs.accompanist_result.session_uuid,
            source_result_uuid=inputs.accompanist_result.result_uuid,
            source_contract_id=inputs.accompanist_result.contract_id,
            source_contract_version=inputs.accompanist_result.contract_version,
            source_revision=inputs.accompanist_result.source_revision,
            jury_input_revision=inputs.readiness.jury_input_revision,
            required_lesson_uuids=required_lesson_uuids,
            schedule_entries=tuple(schedule_entries),
            unscheduled_lessons=tuple(unscheduled),
            diagnostics=tuple(diagnostics),
            conflict_explanations=tuple(conflicts),
        )

    @staticmethod
    def _lesson_flexibility(
        fact: FinalizedAccompanistAssignmentFact,
        state: _PanelState,
        windows_by_person_date: dict[tuple[UUID, date], list[JuryAvailabilityWindow]],
    ) -> int:
        pianist = fact.assigned_pianist
        if pianist is None:
            return MINUTES_PER_DAY + 1
        return sum(
            max(0, window.end_minute - max(window.start_minute, state.timing.earliest_start_minute))
            for window in windows_by_person_date.get((pianist.person_uuid, state.jury_date), ())
        )

    def _find_placement(
        self,
        fact: FinalizedAccompanistAssignmentFact,
        state: _PanelState,
        windows_by_person_date: dict[tuple[UUID, date], list[JuryAvailabilityWindow]],
        bookings: dict[tuple[UUID, date], list[ScheduledLesson]],
        student_bookings: dict[tuple[UUID, date], list[ScheduledLesson]],
    ) -> _Placement | _PlacementFailure:
        duration = state.timing.jury_length_minutes
        pianist_uuid = fact.assigned_pianist.person_uuid if fact.assigned_pianist else None
        availability = windows_by_person_date.get((pianist_uuid, state.jury_date), ()) if pianist_uuid else ()
        pianist_bookings = bookings.get((pianist_uuid, state.jury_date), ()) if pianist_uuid else ()
        own_bookings = student_bookings.get((fact.student_person_uuid, state.jury_date), ())
        candidate = state.cursor_minute
        juries_since_break = state.juries_since_break
        events: list[ScheduleEntry] = []
        meal_added = False
        had_pianist_conflict = False
        had_student_conflict = False
        blocking_lessons: set[UUID] = set()
        blocking_student_lessons: set[UUID] = set()
        crossed_day_bound = False

        while candidate + duration <= MINUTES_PER_DAY:
            meal = state.timing.meal_break
            if meal is not None and not meal_added and state.cursor_minute < meal.end_minute and candidate >= meal.end_minute:
                events.append(self._meal_entry(state, meal.start_minute, meal.end_minute))
                meal_added = True
                juries_since_break = 0

            if meal is not None and candidate < meal.end_minute and candidate + duration > meal.start_minute:
                events.append(self._meal_entry(state, meal.start_minute, meal.end_minute))
                meal_added = True
                juries_since_break = 0
                candidate = meal.end_minute
                continue

            periodic = state.timing.periodic_break
            if periodic is not None and juries_since_break >= periodic.every_x_juries:
                break_end = candidate + periodic.length_minutes
                if meal is not None and candidate < meal.end_minute and break_end > meal.start_minute:
                    events.append(self._meal_entry(state, meal.start_minute, meal.end_minute))
                    meal_added = True
                    juries_since_break = 0
                    candidate = meal.end_minute
                    continue
                events.append(BreakEntry(
                    panel_uuid=state.panel_uuid,
                    jury_date=state.jury_date,
                    start_minute=candidate,
                    end_minute=break_end,
                    after_jury_count=periodic.every_x_juries,
                ))
                candidate = break_end
                juries_since_break = 0
                continue

            if pianist_uuid is not None:
                containing_window = next((
                    window for window in availability
                    if window.start_minute <= candidate and candidate + duration <= window.end_minute
                ), None)
                if containing_window is None:
                    next_window = next((window for window in availability if window.start_minute > candidate), None)
                    if next_window is None:
                        crossed_day_bound = candidate + duration > MINUTES_PER_DAY
                        break
                    candidate = next_window.start_minute
                    continue

                overlapping = [
                    booking for booking in pianist_bookings
                    if candidate < booking.end_minute and booking.start_minute < candidate + duration
                ]
                if overlapping:
                    had_pianist_conflict = True
                    blocking_lessons.update(booking.source_lesson_uuid for booking in overlapping)
                    candidate = max(booking.end_minute for booking in overlapping)
                    continue

            student_overlap = [
                booking for booking in own_bookings
                if candidate < booking.end_minute and booking.start_minute < candidate + duration
            ]
            if student_overlap:
                had_student_conflict = True
                blocking_student_lessons.update(booking.source_lesson_uuid for booking in student_overlap)
                candidate = max(booking.end_minute for booking in student_overlap)
                continue

            return _Placement(
                start_minute=candidate,
                end_minute=candidate + duration,
                events_before=tuple(events),
                juries_since_break=juries_since_break,
            )

        if candidate + duration > MINUTES_PER_DAY:
            crossed_day_bound = True
        if had_pianist_conflict:
            reason = UnscheduledReasonCode.FIXED_PIANIST_CONFLICT
        elif had_student_conflict:
            return _PlacementFailure(
                UnscheduledReasonCode.STUDENT_CONFLICT,
                tuple(sorted(blocking_student_lessons, key=lambda item: item.hex)),
            )
        elif pianist_uuid is not None and not any(
            window.end_minute - max(window.start_minute, state.timing.earliest_start_minute) >= duration
            for window in availability
        ):
            reason = UnscheduledReasonCode.AVAILABILITY_TOO_SHORT
        elif crossed_day_bound:
            reason = UnscheduledReasonCode.DAY_BOUND_EXCEEDED
        else:
            reason = UnscheduledReasonCode.NO_FEASIBLE_INTERVAL
        return _PlacementFailure(reason, tuple(sorted(blocking_lessons, key=lambda item: item.hex)))

    @staticmethod
    def _meal_entry(state: _PanelState, start_minute: int, end_minute: int) -> MealBreakEntry:
        return MealBreakEntry(
            panel_uuid=state.panel_uuid,
            jury_date=state.jury_date,
            start_minute=start_minute,
            end_minute=end_minute,
        )

    @staticmethod
    def _unscheduled(
        fact: FinalizedAccompanistAssignmentFact,
        reason: UnscheduledReasonCode,
    ) -> UnscheduledLesson:
        return UnscheduledLesson(
            source_lesson_uuid=fact.source_lesson_uuid,
            reason_code=reason,
            explanation=DeterministicJuryOptimizer._reason_message(reason),
        )

    @staticmethod
    def _reason_message(reason: UnscheduledReasonCode) -> str:
        return {
            UnscheduledReasonCode.NO_FEASIBLE_INTERVAL: "No interval satisfies this Panel's timing and break constraints.",
            UnscheduledReasonCode.FIXED_PIANIST_CONFLICT: "The assigned Pianist is occupied during every remaining legal interval on this Panel date.",
            UnscheduledReasonCode.STUDENT_CONFLICT: "The Student is already scheduled for another Jury during every remaining legal interval on this Panel date.",
            UnscheduledReasonCode.AVAILABILITY_TOO_SHORT: "The assigned Pianist's Jury Availability Windows do not contain an interval long enough for this Jury.",
            UnscheduledReasonCode.DAY_BOUND_EXCEEDED: "No legal interval remains before the end of this Jury day.",
        }[reason]

    @staticmethod
    def _event_order_key(entry: ScheduleEntry) -> tuple[date, int, str, int, str]:
        entry_type_order = 1 if isinstance(entry, ScheduledLesson) else 0
        lesson_key = entry.source_lesson_uuid.hex if isinstance(entry, ScheduledLesson) else ""
        return entry.jury_date, entry.start_minute, entry.panel_uuid.hex, entry_type_order, lesson_key