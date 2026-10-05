from dataclasses import replace
import unittest

from app.jury_optimizer_contract import (
    BreakEntry,
    JuryScheduleResult,
    MealBreakEntry,
    ScheduledLesson,
    UnscheduledLesson,
    UnscheduledReasonCode,
)
from app.jury_optimizer import DeterministicJuryOptimizer
from tests.fixtures.jury_optimizer_fixtures import make_fixture, make_stale_accompanist_fixture


class JuryOptimizerDomainContractTests(unittest.TestCase):
    def test_synthetic_scenario_matrix_covers_required_inputs(self):
        scenarios = (
            "single-panel",
            "multiple-panels",
            "multiple-dates",
            "shared-pianist-conflicts",
            "shared-pianist-overbooked",
            "unavailable-pianist",
            "impossible-schedule",
            "meal-breaks",
            "meal-reset",
            "periodic-breaks",
            "day-bound",
            "fill-before-late-window",
            "assigned-optional",
            "preferred-start",
            "duplicate-display-names",
        )
        for scenario in scenarios:
            with self.subTest(scenario=scenario):
                self.assertTrue(make_fixture(scenario).inputs.readiness.approved)

    def test_duplicate_display_names_keep_distinct_uuid_identity(self):
        fixture = make_fixture("duplicate-display-names")
        facts = fixture.inputs.accompanist_assignments

        self.assertEqual(facts[0].student_display_name, facts[1].student_display_name)
        self.assertNotEqual(facts[0].student_person_uuid, facts[1].student_person_uuid)
        self.assertNotEqual(facts[0].source_lesson_uuid, facts[1].source_lesson_uuid)

    def test_stale_accompanist_dependency_is_rejected_before_optimizer_input(self):
        with self.assertRaisesRegex(ValueError, "does not match"):
            make_stale_accompanist_fixture()

    def test_sixty_minute_duration_and_break_configurations_are_valid_contract_inputs(self):
        fixture = make_fixture("periodic-breaks")
        timing = fixture.inputs.panel_timings[0]

        self.assertEqual(timing.jury_length_minutes, 60)
        self.assertEqual(timing.periodic_break.every_x_juries, 2)
        self.assertEqual(timing.periodic_break.length_minutes, 10)
        self.assertEqual(make_fixture("meal-breaks").inputs.panel_timings[0].meal_break.end_minute, 750)

    def test_result_requires_exactly_one_disposition_for_every_required_lesson(self):
        fixture = make_fixture("single-panel")
        lesson_uuid = fixture.lesson_uuids[0]
        fact = fixture.inputs.accompanist_assignments[0]
        panel_uuid = fixture.panel_uuids[0]
        panel_date = fixture.inputs.panel_dates[0].jury_date
        scheduled = ScheduledLesson(
            panel_uuid=panel_uuid,
            jury_date=panel_date,
            start_minute=540,
            end_minute=600,
            source_lesson_uuid=lesson_uuid,
            student_person_uuid=fact.student_person_uuid,
            pianist_person_uuid=None,
        )
        result_fields = dict(
            session_uuid=fixture.inputs.accompanist_result.session_uuid,
            source_result_uuid=fixture.inputs.accompanist_result.result_uuid,
            source_contract_id=fixture.inputs.accompanist_result.contract_id,
            source_contract_version=fixture.inputs.accompanist_result.contract_version,
            source_revision=fixture.inputs.accompanist_result.source_revision,
            jury_input_revision=fixture.inputs.readiness.jury_input_revision,
            required_lesson_uuids=(lesson_uuid,),
        )

        JuryScheduleResult(**result_fields, schedule_entries=(scheduled,), unscheduled_lessons=())
        with self.assertRaisesRegex(ValueError, "Every Jury-required lesson"):
            JuryScheduleResult(**result_fields, schedule_entries=(), unscheduled_lessons=())
        with self.assertRaisesRegex(ValueError, "cannot be both"):
            JuryScheduleResult(
                **result_fields,
                schedule_entries=(scheduled,),
                unscheduled_lessons=(UnscheduledLesson(
                    lesson_uuid,
                    UnscheduledReasonCode.NO_FEASIBLE_INTERVAL,
                    "No legal interval remains.",
                ),),
            )


class JuryOptimizerBehaviorContractTests(unittest.TestCase):
    def setUp(self):
        self.optimizer = DeterministicJuryOptimizer()

    def optimize(self, inputs):
        return self.optimizer.optimize(inputs)

    def test_impossible_schedules_return_explicit_unscheduled_reasons(self):
        fixture = make_fixture("impossible-schedule")
        result = self.optimize(fixture.inputs)
        self.assertEqual({entry.source_lesson_uuid for entry in result.unscheduled_lessons}, set(fixture.lesson_uuids))
        self.assertTrue(all(entry.explanation for entry in result.unscheduled_lessons))
        self.assertTrue(all(entry.reason_code for entry in result.unscheduled_lessons))

    def test_shared_pianist_conflicts_are_never_overlapped(self):
        fixture = make_fixture("shared-pianist-conflicts")
        result = self.optimize(fixture.inputs)
        self.assertFalse(_has_shared_pianist_overlap(result.scheduled_lessons))
        self.assertEqual(len(result.scheduled_lessons), len(fixture.lesson_uuids))

    def test_pianist_overcapacity_returns_explicit_conflict_without_substitution(self):
        fixture = make_fixture("shared-pianist-overbooked")
        result = self.optimize(fixture.inputs)

        self.assertEqual(len(result.scheduled_lessons), 1)
        self.assertEqual(len(result.unscheduled_lessons), 1)
        self.assertEqual(result.unscheduled_lessons[0].reason_code, UnscheduledReasonCode.FIXED_PIANIST_CONFLICT)
        self.assertEqual(len(result.conflict_explanations), 1)
        self.assertEqual(result.conflict_explanations[0].person_uuid, fixture.pianist_uuids[0])

    def test_same_student_is_never_scheduled_in_two_places_at_once(self):
        fixture = make_fixture("multiple-panels")
        student_uuid = fixture.inputs.accompanist_assignments[0].student_person_uuid
        inputs = replace(
            fixture.inputs,
            accompanist_assignments=tuple(
                replace(fact, student_person_uuid=student_uuid, pianist_required=False, assigned_pianist=None)
                for fact in fixture.inputs.accompanist_assignments
            ),
            jury_lesson_entries=tuple(
                replace(entry, student_person_uuid=student_uuid) for entry in fixture.inputs.jury_lesson_entries
            ),
        )
        result = self.optimize(inputs)

        self.assertEqual(len(result.scheduled_lessons), 2)
        first, second = sorted(result.scheduled_lessons, key=lambda entry: entry.start_minute)
        self.assertFalse(_overlaps(first.start_minute, first.end_minute, second.start_minute, second.end_minute))

    def test_panel_assignment_and_date_are_authoritative(self):
        fixture = make_fixture("multiple-panels")
        result = self.optimize(fixture.inputs)
        expected = {entry.source_lesson_uuid: entry.panel_uuid for entry in fixture.inputs.jury_lesson_entries}
        expected_dates = {item.panel_uuid: item.jury_date for item in fixture.inputs.panel_dates}
        self.assertTrue(all(entry.panel_uuid == expected[entry.source_lesson_uuid] for entry in result.scheduled_lessons))
        self.assertTrue(all(entry.jury_date == expected_dates[entry.panel_uuid] for entry in result.scheduled_lessons))

    def test_finalized_pianist_assignments_are_never_substituted(self):
        fixture = make_fixture("shared-pianist-conflicts")
        result = self.optimize(fixture.inputs)
        expected = {
            fact.source_lesson_uuid: fact.assigned_pianist.person_uuid
            for fact in fixture.inputs.accompanist_assignments
        }
        self.assertTrue(all(
            entry.pianist_person_uuid == expected[entry.source_lesson_uuid]
            for entry in result.scheduled_lessons
        ))

    def test_availability_windows_are_hard_constraints(self):
        fixture = make_fixture("unavailable-pianist")
        result = self.optimize(fixture.inputs)
        self.assertFalse(_outside_availability(result.scheduled_lessons, fixture.inputs.pianist_availability_windows))

    def test_identical_inputs_produce_identical_results(self):
        fixture = make_fixture("preferred-start")
        self.assertEqual(self.optimize(fixture.inputs), self.optimize(fixture.inputs))

    def test_result_is_independent_of_dto_collection_order(self):
        fixture = make_fixture("multiple-panels")
        reversed_inputs = replace(
            fixture.inputs,
            accompanist_assignments=tuple(reversed(fixture.inputs.accompanist_assignments)),
            jury_lesson_entries=tuple(reversed(fixture.inputs.jury_lesson_entries)),
            panel_dates=tuple(reversed(fixture.inputs.panel_dates)),
            panel_timings=tuple(reversed(fixture.inputs.panel_timings)),
            pianist_availability_windows=tuple(reversed(fixture.inputs.pianist_availability_windows)),
        )

        self.assertEqual(self.optimize(fixture.inputs), self.optimize(reversed_inputs))

    def test_duration_day_bounds_and_configured_breaks_are_respected(self):
        for scenario in ("meal-breaks", "periodic-breaks"):
            with self.subTest(scenario=scenario):
                fixture = make_fixture(scenario)
                result = self.optimize(fixture.inputs)
                timing_by_panel = {item.panel_uuid: item for item in fixture.inputs.panel_timings}
                for entry in result.scheduled_lessons:
                    self.assertEqual(
                        entry.end_minute - entry.start_minute,
                        timing_by_panel[entry.panel_uuid].jury_length_minutes,
                    )
                    self.assertLessEqual(entry.end_minute, 24 * 60)
                breaks = result.meal_break_entries + result.periodic_break_entries
                self.assertTrue(all(
                    not _overlaps(entry.start_minute, entry.end_minute, break_entry.start_minute, break_entry.end_minute)
                    for entry in result.scheduled_lessons
                    for break_entry in breaks
                    if entry.panel_uuid == break_entry.panel_uuid and entry.jury_date == break_entry.jury_date
                ))

    def test_every_required_lesson_receives_one_complete_outcome(self):
        fixture = make_fixture("single-panel")
        result = self.optimize(fixture.inputs)
        outcomes = [entry.source_lesson_uuid for entry in result.scheduled_lessons]
        outcomes.extend(entry.source_lesson_uuid for entry in result.unscheduled_lessons)
        self.assertCountEqual(outcomes, fixture.lesson_uuids)
        self.assertEqual(len(outcomes), len(set(outcomes)))
        self.assertEqual(result.session_uuid, fixture.inputs.accompanist_result.session_uuid)
        self.assertEqual(result.source_result_uuid, fixture.inputs.accompanist_result.result_uuid)
        self.assertEqual(result.source_revision, fixture.inputs.accompanist_result.source_revision)
        self.assertEqual(result.source_contract_version, 2)
        self.assertEqual(result.jury_input_revision, fixture.inputs.readiness.jury_input_revision)

    def test_multiple_dates_allow_same_pianist_at_same_time(self):
        fixture = make_fixture("multiple-dates")
        result = self.optimize(fixture.inputs)

        self.assertEqual(len(result.scheduled_lessons), 2)
        self.assertEqual({entry.jury_date for entry in result.scheduled_lessons}, {item.jury_date for item in fixture.inputs.panel_dates})
        self.assertEqual({entry.start_minute for entry in result.scheduled_lessons}, {540})

    def test_unavailable_fixed_pianist_is_not_replaced(self):
        fixture = make_fixture("unavailable-pianist")
        result = self.optimize(fixture.inputs)

        self.assertEqual(len(result.scheduled_lessons), 0)
        self.assertEqual(len(result.unscheduled_lessons), 1)
        self.assertEqual(result.unscheduled_lessons[0].reason_code, UnscheduledReasonCode.AVAILABILITY_TOO_SHORT)

    def test_day_bound_is_enforced(self):
        fixture = make_fixture("day-bound")
        result = self.optimize(fixture.inputs)

        self.assertEqual(len(result.scheduled_lessons), 1)
        self.assertEqual(result.scheduled_lessons[0].end_minute, 1440)
        self.assertEqual(result.unscheduled_lessons[0].reason_code, UnscheduledReasonCode.DAY_BOUND_EXCEEDED)

    def test_periodic_breaks_follow_configured_jury_count_and_length(self):
        fixture = make_fixture("periodic-breaks")
        result = self.optimize(fixture.inputs)

        self.assertEqual(len(result.periodic_break_entries), 2)
        self.assertEqual([(entry.start_minute, entry.end_minute) for entry in result.periodic_break_entries], [(660, 670), (790, 800)])
        self.assertTrue(all(isinstance(entry, BreakEntry) for entry in result.periodic_break_entries))

    def test_meal_break_resets_periodic_break_counter(self):
        fixture = make_fixture("meal-reset")
        result = self.optimize(fixture.inputs)

        self.assertEqual([(entry.start_minute, entry.end_minute) for entry in result.meal_break_entries], [(660, 690)])
        self.assertEqual([(entry.start_minute, entry.end_minute) for entry in result.periodic_break_entries], [(750, 760)])
        self.assertTrue(all(isinstance(entry, MealBreakEntry) for entry in result.meal_break_entries))

    def test_single_panel_schedules_earliest_and_multi_panel_demand_is_serialized_by_pianist(self):
        single = self.optimize(make_fixture("single-panel").inputs)
        multiple = self.optimize(make_fixture("multiple-panels").inputs)

        self.assertEqual(single.scheduled_lessons[0].start_minute, 540)
        self.assertEqual(len(multiple.scheduled_lessons), 2)
        self.assertFalse(_has_shared_pianist_overlap(multiple.scheduled_lessons))

    def test_flexible_lesson_fills_early_time_before_a_late_fixed_pianist_window(self):
        fixture = make_fixture("fill-before-late-window")
        result = self.optimize(fixture.inputs)

        self.assertEqual(len(result.scheduled_lessons), 2)
        panel_zero_lessons = [entry for entry in result.scheduled_lessons if entry.panel_uuid == fixture.panel_uuids[0]]
        self.assertEqual(sorted(entry.start_minute for entry in panel_zero_lessons), [540, 600])

    def test_assigned_pianist_is_retained_even_when_pianist_is_not_required(self):
        fixture = make_fixture("assigned-optional")
        result = self.optimize(fixture.inputs)

        self.assertFalse(fixture.inputs.accompanist_assignments[0].pianist_required)
        self.assertEqual(result.scheduled_lessons[0].pianist_person_uuid, fixture.pianist_uuids[0])


def _has_shared_pianist_overlap(entries):
    for index, left in enumerate(entries):
        if left.pianist_person_uuid is None:
            continue
        for right in entries[index + 1:]:
            if left.pianist_person_uuid != right.pianist_person_uuid or left.jury_date != right.jury_date:
                continue
            if left.start_minute < right.end_minute and right.start_minute < left.end_minute:
                return True
    return False


def _outside_availability(entries, windows):
    for entry in entries:
        if entry.pianist_person_uuid is None:
            continue
        if not any(
            window.pianist_person_uuid == entry.pianist_person_uuid
            and window.jury_date == entry.jury_date
            and window.start_minute <= entry.start_minute
            and entry.end_minute <= window.end_minute
            for window in windows
        ):
            return True
    return False


def _overlaps(first_start, first_end, second_start, second_end):
    return first_start < second_end and second_start < first_end