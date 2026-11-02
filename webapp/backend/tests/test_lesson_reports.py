"""Synthetic tests for the data Lesson reports consume and the saved-report contract they use.

Lesson report content is built client-side from the schedule API (see lessonReport.test.ts);
these tests pin the API fields those reports rely on and the module-scoped report persistence.
"""

import unittest

from app import module_models, schemas
from app.routers import lesson_scheduling as lesson, reports
from tests.test_lesson_foundation import LessonTestCase, window


def saved_payload(stamp: str, markdown: str = "# Lesson Schedule") -> schemas.SavedReportIn:
    return schemas.SavedReportIn(
        markdown=markdown,
        page_break_headings=["Instructor Placement Roster", "Unplaced Students"],
        table=schemas.SavedReportTable(
            name="Lesson Schedule",
            columns=["Student", "Instructor"],
            rows=[["Synthetic Student", "Synthetic Instructor"]],
        ),
        source_stamp=stamp,
    )


class LessonReportSourceTests(LessonTestCase):
    def setUp(self):
        super().setUp()
        self.repo.create_instructor(
            "Synthetic Instructor", room="Studio 1", lesson_length_minutes=60,
            availability=[window(start=540, end=600)],
        )
        for number in range(2):
            self.repo.create_student(
                f"S-{number}", f"Synthetic Student {number}", availability=[window(start=540, end=600)]
            )

    def test_schedule_api_exposes_placed_and_unplaced_report_fields(self):
        schedule = lesson.generate_schedule(self.db)

        placed = [entry for entry in schedule.entries if entry.day is not None]
        unplaced = [entry for entry in schedule.entries if entry.day is None]
        self.assertEqual(len(placed), 1)
        self.assertEqual(placed[0].placement_type, "automatic")
        self.assertEqual(placed[0].instructor_name, "Synthetic Instructor")
        self.assertIsNotNone(placed[0].day)
        self.assertIsNotNone(placed[0].start_minute)
        self.assertIsNotNone(placed[0].end_minute)
        self.assertEqual(len(unplaced), 1)
        self.assertEqual(unplaced[0].unplaced_reason, "CAPACITY_LIMIT")

    def test_automatically_unplaced_entry_is_unassigned_and_not_counted_as_automatic(self):
        schedule = lesson.generate_schedule(self.db)

        unplaced = next(entry for entry in schedule.entries if entry.day is None)
        self.assertEqual(unplaced.placement_state, "unassigned")
        self.assertEqual(schedule.automatic_count, 1)
        self.assertEqual(schedule.manual_count, 0)

    def test_manually_unplaced_entries_are_reported_as_such(self):
        schedule = lesson.generate_schedule(self.db)
        unplaced_id = next(entry.student_id for entry in schedule.entries if entry.day is None)

        refreshed = lesson.manually_unplace_student(unplaced_id, self.db)

        entry = next(item for item in refreshed.entries if item.student_id == unplaced_id)
        self.assertEqual(entry.placement_type, "manuallyUnplaced")
        self.assertEqual(entry.unplaced_reason, "MANUALLY_UNPLACED")

    def test_stale_state_changes_the_report_source_stamp_inputs(self):
        generated = lesson.generate_schedule(self.db)
        stamp = (generated.schedule_revision, generated.source_data_revision)
        self.assertFalse(generated.is_schedule_stale)

        self.repo.create_student("S-9", "Late Student", availability=[window()])

        stale = lesson.get_schedule(self.db)
        self.assertTrue(stale.is_schedule_stale)
        self.assertEqual((stale.schedule_revision, stale.source_data_revision), stamp)
        self.assertGreater(self.repo.get_state().source_data_revision, stale.source_data_revision)
        self.assertEqual(len(stale.entries), 3)

    def test_verification_warnings_are_available_to_the_warnings_section(self):
        generated = lesson.generate_schedule(self.db)
        placed = next(entry for entry in generated.entries if entry.day is not None)

        self.repo.set_student_availability(placed.student_id, [])
        stale = lesson.get_schedule(self.db)

        warned = [item for item in stale.verification if item.student_id == placed.student_id]
        self.assertEqual(len(warned), 1)
        self.assertTrue(warned[0].warnings)
        self.assertTrue(all(warning.message for warning in warned[0].warnings))

    def test_saved_lesson_report_round_trips_and_is_scoped_to_the_lesson_module(self):
        reports.put_saved_report(saved_payload("2:3"), "lesson", self.db)
        reports.put_saved_report(saved_payload("1:1", "# Clinical"), "clinical", self.db)
        reports.put_saved_report(saved_payload("4:5", "# Lesson Schedule v2"), "lesson", self.db)

        loaded = reports.get_saved_report("lesson", self.db)

        self.assertEqual(loaded.markdown, "# Lesson Schedule v2")
        self.assertEqual(loaded.source_stamp, "4:5")
        self.assertEqual(loaded.page_break_headings, ["Instructor Placement Roster", "Unplaced Students"])
        self.assertEqual(reports.get_saved_report("clinical", self.db).markdown, "# Clinical")
        self.assertEqual(self.db.query(module_models.SavedReport).count(), 2)


if __name__ == "__main__":
    unittest.main()
