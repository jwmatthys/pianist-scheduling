from datetime import time
import unittest

import pandas as pd

import generate_pianist_schedule as cli
from app.services import scheduling as desktop


def minutes_to_time(minutes):
    return time(minutes // 60, minutes % 60)


def lesson(student, start, end, required="", location="Room A", need_pianist=True):
    return {
        "Lesson Teacher Name": "Teacher Example",
        "Student Name": student,
        "Lesson Day": "Monday",
        "Lesson Start Time": minutes_to_time(start),
        "Lesson End Time": minutes_to_time(end),
        "Lesson Location": location,
        "Instrument": "Violin",
        "Need Pianist": need_pianist,
        "Required Pianist": required,
    }


def cli_assign(lessons, pianists):
    frame = pd.DataFrame(lessons)
    return cli.assign_lessons(frame, pianists)


def available_slots(start, end, status="Available"):
    return {
        "Monday": {minute: status for minute in range(start, end, 30)}
    }


class FitCharacterizationTests(unittest.TestCase):
    def test_fit_tiers_match_for_available_tentative_near_and_unavailable(self):
        cases = [
            (540, 600, {540: "Available", 570: "Available"}, cli.FIT_FULL),
            (540, 600, {540: "Available", 570: "Tentative"}, cli.FIT_PARTIAL),
            (555, 570, {570: "Available"}, cli.FIT_NEAR),
            (540, 600, {540: "Unavailable", 570: "Unavailable"}, cli.FIT_NONE),
        ]
        for start, end, slots, expected in cases:
            with self.subTest(start=start, end=end, slots=slots):
                availability = {"Monday": slots}
                cli_score, _ = cli.get_fit(
                    "Monday", minutes_to_time(start), minutes_to_time(end), availability
                )
                desktop_score, _ = desktop.get_fit("Monday", start, end, availability)
                self.assertEqual(cli_score, expected)
                self.assertEqual(desktop_score, expected)

    def test_required_name_substring_matching_agrees(self):
        pianist_availability = available_slots(540, 600)
        cli_results, _ = cli_assign(
            [lesson("Student One", 540, 600, required="Ari")],
            [("Ari Lee", 4, pianist_availability)],
        )
        desktop_lessons = [
            desktop.EngineLesson(1, "Monday", 540, 600, required_pianist_name="Ari")
        ]
        desktop_pianists = [desktop.EnginePianist(1, "Ari Lee", 4, pianist_availability)]
        desktop.assign_lessons(desktop_lessons, desktop_pianists)

        self.assertEqual(cli_results[0]["Assigned Accompanist"], "Ari Lee")
        self.assertEqual(desktop_lessons[0].assigned_pianist_id, 1)

    def test_required_pianist_is_assigned_unavailable_and_over_cap_warning_agrees(self):
        unavailable = available_slots(540, 600, "Unavailable")
        cli_results, _ = cli_assign(
            [lesson("Student One", 540, 600, required="Ari Lee")],
            [("Ari Lee", 0.5, unavailable)],
        )
        desktop_lessons = [
            desktop.EngineLesson(1, "Monday", 540, 600, required_pianist_name="Ari Lee")
        ]
        desktop_pianists = [desktop.EnginePianist(1, "Ari Lee", 0.5, unavailable)]
        desktop.assign_lessons(desktop_lessons, desktop_pianists)

        self.assertEqual(cli_results[0]["Assigned Accompanist"], "Ari Lee")
        self.assertIn("UNAVAILABLE", cli_results[0]["Notes"])
        self.assertIn("OVER CAP", cli_results[0]["Notes"])
        self.assertEqual(desktop_lessons[0].assigned_pianist_id, 1)
        self.assertIn("UNAVAILABLE", desktop_lessons[0].notes)
        self.assertIn("OVER CAP", desktop_lessons[0].notes)

    def test_over_cap_warning_uses_current_union_hours_not_existing_notes(self):
        lessons = [
            desktop.EngineLesson(
                1, "Monday", 540, 600, assigned_pianist_id=1, notes="OVER CAP: stale"
            ),
            desktop.EngineLesson(
                2, "Monday", 570, 630, assigned_pianist_id=1, notes="OVER CAP: stale"
            ),
        ]
        pianist = desktop.EnginePianist(1, "Ari Lee", 2)

        desktop.recompute_hours_and_conflicts(lessons, [pianist])
        self.assertTrue(all("OVER CAP" not in lesson.notes for lesson in lessons))

        pianist.max_hours = 1.25
        desktop.recompute_hours_and_conflicts(lessons, [pianist])
        self.assertTrue(all("OVER CAP: Ari Lee at 1.5h / 1.25h cap" in lesson.notes
                            for lesson in lessons))

        lessons[0].assigned_pianist_id = None
        desktop.recompute_hours_and_conflicts(lessons, [pianist])
        self.assertNotIn("OVER CAP", lessons[0].notes)

    def test_ordinary_assignment_prefers_under_cap_candidate_in_both_engines(self):
        full_availability = available_slots(540, 600)
        partial_availability = {"Monday": {540: "Available", 570: "Tentative"}}
        cli_results, _ = cli_assign(
            [lesson("Student One", 540, 600)],
            [("Over Cap Full Fit", 0.5, full_availability),
             ("Within Cap Partial Fit", 2, partial_availability)],
        )
        desktop_lessons = [desktop.EngineLesson(1, "Monday", 540, 600)]
        desktop_pianists = [
            desktop.EnginePianist(1, "Over Cap Full Fit", 0.5, full_availability),
            desktop.EnginePianist(2, "Within Cap Partial Fit", 2, partial_availability),
        ]
        desktop.assign_lessons(desktop_lessons, desktop_pianists)

        self.assertEqual(cli_results[0]["Assigned Accompanist"], "Within Cap Partial Fit")
        self.assertEqual(desktop_lessons[0].assigned_pianist_id, 2)

    def test_full_fit_beats_partial_fit_when_both_candidates_are_under_cap(self):
        full_availability = available_slots(540, 600)
        partial_availability = {"Monday": {540: "Available", 570: "Tentative"}}
        cli_results, _ = cli_assign(
            [lesson("Student One", 540, 600)],
            [("Full Fit", 4, full_availability), ("Partial Fit", 4, partial_availability)],
        )
        desktop_lessons = [desktop.EngineLesson(1, "Monday", 540, 600)]
        desktop_pianists = [
            desktop.EnginePianist(1, "Full Fit", 4, full_availability),
            desktop.EnginePianist(2, "Partial Fit", 4, partial_availability),
        ]
        desktop.assign_lessons(desktop_lessons, desktop_pianists)

        self.assertEqual(cli_results[0]["Assigned Accompanist"], "Full Fit")
        self.assertEqual(desktop_lessons[0].assigned_pianist_id, 1)


class AssignmentDifferenceTests(unittest.TestCase):
    def test_both_engines_preserve_compatible_overlap_as_allowed_union_time(self):
        availability = available_slots(540, 630)
        lessons = [lesson("Student One", 540, 600), lesson("Student Two", 570, 630)]
        cli_results, cli_hours = cli_assign(
            lessons, [("Ari Lee", 4, availability)]
        )
        desktop_lessons = [
            desktop.EngineLesson(1, "Monday", 540, 600),
            desktop.EngineLesson(2, "Monday", 570, 630),
        ]
        desktop_pianists = [desktop.EnginePianist(1, "Ari Lee", 4, availability)]
        desktop_hours, desktop_conflicts = desktop.assign_lessons(
            desktop_lessons, desktop_pianists
        )

        self.assertEqual(cli_results[1]["Assigned Accompanist"], "Ari Lee")
        self.assertEqual(cli_results[1]["Fit Quality"], "Overlap")
        self.assertEqual(cli_hours["Ari Lee"], 1.5)
        self.assertEqual(desktop_lessons[1].assigned_pianist_id, 1)
        self.assertEqual(desktop_lessons[0].fit_quality, "Overlap")
        self.assertEqual(desktop_lessons[1].fit_quality, "Overlap")
        self.assertEqual(desktop_hours["Ari Lee"], 1.5)
        self.assertEqual(desktop_conflicts, [])

    def test_overlap_beyond_existing_30_minute_limit_is_unassigned(self):
        availability = available_slots(540, 630)
        lessons = [lesson("Student One", 540, 600), lesson("Student Two", 569, 629)]
        cli_results, _ = cli_assign(lessons, [("Ari Lee", 4, availability)])
        desktop_lessons = [
            desktop.EngineLesson(1, "Monday", 540, 600),
            desktop.EngineLesson(2, "Monday", 569, 629),
        ]
        desktop.assign_lessons(
            desktop_lessons, [desktop.EnginePianist(1, "Ari Lee", 4, availability)]
        )

        self.assertEqual(cli_results[1]["Assigned Accompanist"], "UNASSIGNED")
        self.assertIsNone(desktop_lessons[1].assigned_pianist_id)
        self.assertIn("permitted overlap", cli_results[1]["Notes"])
        self.assertIn("permitted overlap", desktop_lessons[1].notes)

    def test_union_hours_for_manually_overlapping_desktop_assignments(self):
        lessons = [
            desktop.EngineLesson(1, "Monday", 540, 600, assigned_pianist_id=1),
            desktop.EngineLesson(2, "Monday", 570, 630, assigned_pianist_id=1),
            desktop.EngineLesson(3, "Tuesday", 540, 570, assigned_pianist_id=1),
        ]
        pianists = [desktop.EnginePianist(1, "Ari Lee", 4)]

        hours, conflicts = desktop.recompute_hours_and_conflicts(lessons, pianists)

        self.assertEqual(hours["Ari Lee"], 2.0)
        self.assertEqual([item.hours for item in lessons], [1.0, 0.5, 0.5])
        self.assertEqual(len(conflicts), 1)

    def test_incompatible_ordinary_overlap_is_unassigned_by_both_engines(self):
        availability = available_slots(540, 600)
        lessons = [lesson("Student One", 540, 600), lesson("Student Two", 540, 600)]
        cli_results, _ = cli_assign(lessons, [("Ari Lee", 4, availability)])
        desktop_lessons = [
            desktop.EngineLesson(1, "Monday", 540, 600),
            desktop.EngineLesson(2, "Monday", 540, 600),
        ]
        desktop.assign_lessons(
            desktop_lessons, [desktop.EnginePianist(1, "Ari Lee", 4, availability)]
        )

        self.assertEqual(cli_results[1]["Assigned Accompanist"], "UNASSIGNED")
        self.assertIn("already assigned", cli_results[1]["Notes"])
        self.assertIsNone(desktop_lessons[1].assigned_pianist_id)
        self.assertIn("already assigned", desktop_lessons[1].notes)

    def test_required_pianist_overlap_is_kept_and_flagged_by_both_engines(self):
        availability = available_slots(540, 600)
        lessons = [
            lesson("Student One", 540, 600, required="Ari Lee"),
            lesson("Student Two", 540, 600, required="Ari Lee"),
        ]
        cli_results, _ = cli_assign(lessons, [("Ari Lee", 4, availability)])
        desktop_lessons = [
            desktop.EngineLesson(1, "Monday", 540, 600, required_pianist_name="Ari Lee"),
            desktop.EngineLesson(2, "Monday", 540, 600, required_pianist_name="Ari Lee"),
        ]
        _, desktop_conflicts = desktop.assign_lessons(
            desktop_lessons, [desktop.EnginePianist(1, "Ari Lee", 4, availability)]
        )

        self.assertEqual(
            [row["Assigned Accompanist"] for row in cli_results], ["Ari Lee", "Ari Lee"]
        )
        self.assertTrue(all("CONFLICT" in row["Notes"] for row in cli_results))
        self.assertEqual([item.assigned_pianist_id for item in desktop_lessons], [1, 1])
        self.assertTrue(all("CONFLICT" in item.notes for item in desktop_lessons))
        self.assertEqual(len(desktop_conflicts), 1)

    def test_unavailable_ordinary_lesson_is_unassigned_with_reason(self):
        unavailable = available_slots(540, 600, "Unavailable")
        cli_results, _ = cli_assign(
            [lesson("Student One", 540, 600)], [("Ari Lee", 4, unavailable)]
        )
        desktop_lessons = [desktop.EngineLesson(1, "Monday", 540, 600)]
        desktop.assign_lessons(
            desktop_lessons, [desktop.EnginePianist(1, "Ari Lee", 4, unavailable)]
        )

        self.assertEqual(cli_results[0]["Assigned Accompanist"], "UNASSIGNED")
        self.assertEqual(cli_results[0]["Fit Quality"], "None")
        self.assertIn("sufficient availability", cli_results[0]["Notes"])
        self.assertIsNone(desktop_lessons[0].assigned_pianist_id)
        self.assertIn("sufficient availability", desktop_lessons[0].notes)

    def test_zero_pianists_returns_unassigned_with_reason_in_both_engines(self):
        cli_results, _ = cli_assign([lesson("Student One", 540, 600)], [])
        self.assertEqual(cli_results[0]["Assigned Accompanist"], "UNASSIGNED")
        self.assertIn("No pianists", cli_results[0]["Notes"])

        desktop_lessons = [desktop.EngineLesson(1, "Monday", 540, 600)]
        hours, conflicts = desktop.assign_lessons(desktop_lessons, [])
        self.assertIsNone(desktop_lessons[0].assigned_pianist_id)
        self.assertIn("No pianists", desktop_lessons[0].notes)
        self.assertEqual(hours, {})
        self.assertEqual(conflicts, [])

    def test_manual_lock_is_preserved_and_blocks_competing_desktop_assignment(self):
        availability = available_slots(540, 600)
        lessons = [
            desktop.EngineLesson(1, "Monday", 540, 600, assigned_pianist_id=1),
            desktop.EngineLesson(2, "Monday", 540, 600),
        ]
        pianists = [
            desktop.EnginePianist(1, "Locked Pianist", 4, availability),
            desktop.EnginePianist(2, "Available Pianist", 4, availability),
        ]

        desktop.assign_lessons(lessons, pianists, locked_ids={1})

        self.assertEqual(lessons[0].assigned_pianist_id, 1)
        self.assertEqual(lessons[1].assigned_pianist_id, 2)


class SchedulePenaltyCharacterizationTests(unittest.TestCase):
    def test_contiguous_blocks_and_gaps_match_between_engines(self):
        contiguous = [("Monday", 540, 600), ("Monday", 600, 660)]
        separated = [("Monday", 540, 600), ("Monday", 630, 660)]
        two_days = [("Monday", 540, 600), ("Tuesday", 540, 600)]

        for assignments, expected in [
            (contiguous, (1, 1, 0)),
            (separated, (1, 2, 30)),
            (two_days, (2, 2, 0)),
        ]:
            with self.subTest(assignments=assignments):
                self.assertEqual(cli.schedule_penalty(assignments), expected)
                self.assertEqual(desktop.schedule_penalty(assignments), expected)

    def test_overlapping_interval_gap_uses_merged_window_end_in_both_engines(self):
        assignments = [
            ("Monday", 540, 600),
            ("Monday", 555, 570),
            ("Monday", 615, 660),
        ]

        self.assertEqual(cli.schedule_penalty(assignments), (1, 2, 15))
        self.assertEqual(desktop.schedule_penalty(assignments), (1, 2, 15))

    def test_lesson_location_does_not_change_cli_assignments(self):
        availability = available_slots(540, 630)
        pianists = [("Ari Lee", 4, availability), ("Bea Lin", 4, availability)]
        same_room = [lesson("Student One", 540, 600, location="Room A")]
        different_room = [lesson("Student One", 540, 600, location="Off-campus Site Z")]

        same_result, _ = cli_assign(same_room, pianists)
        different_result, _ = cli_assign(different_room, pianists)

        self.assertEqual(
            same_result[0]["Assigned Accompanist"],
            different_result[0]["Assigned Accompanist"],
        )
        self.assertFalse(hasattr(desktop.EngineLesson(1, "Monday", 540, 600), "location"))


if __name__ == "__main__":
    unittest.main()
