import unittest

from app.services import scheduling
from app.services.scheduling import EngineLesson, EnginePianist


def _pianist():
    slots = {minute: "Available" for minute in range(540, 660, 30)}
    return EnginePianist(id=1, name="Synthetic Pianist", max_hours=None, avail={"Monday": slots})


def _lesson(lesson_id, start, end, fit="Full"):
    return EngineLesson(
        id=lesson_id, day="Monday", start_min=start, end_min=end,
        assigned_pianist_id=1, fit_quality=fit, student_name=f"Student {lesson_id}",
    )


class DoubleBookedFitTests(unittest.TestCase):
    def test_double_booking_is_labelled_overlap_and_recovers(self):
        pianist = _pianist()
        first, second = _lesson(1, 540, 600), _lesson(2, 560, 620)
        _, conflicts = scheduling.recompute_hours_and_conflicts([first, second], [pianist])
        self.assertEqual(len(conflicts), 1)
        self.assertEqual(first.fit_quality, "Overlap (>30 min)")
        self.assertEqual(second.fit_quality, "Overlap (>30 min)")

        second.start_min, second.end_min = 600, 660
        _, conflicts = scheduling.recompute_hours_and_conflicts([first, second], [pianist])
        self.assertEqual(conflicts, [])
        self.assertEqual((first.fit_quality, second.fit_quality), ("Full", "Full"))

    def test_label_reports_half_lesson_threshold(self):
        self.assertEqual(scheduling.overlap_fit_label(20, 50), "Overlap (<25 min)")
        self.assertEqual(scheduling.overlap_fit_label(25, 50), "Overlap (<25 min)")
        self.assertEqual(scheduling.overlap_fit_label(26, 50), "Overlap (>25 min)")

    def test_small_overlap_is_warning_sized_and_manual_label_restores(self):
        pianist = _pianist()
        first, second = _lesson(1, 540, 600, "Manual"), _lesson(2, 590, 650)
        first.manual = True
        scheduling.recompute_hours_and_conflicts([first, second], [pianist])
        self.assertEqual(first.fit_quality, "Overlap (<30 min)")
        second.start_min, second.end_min = 600, 660
        scheduling.recompute_hours_and_conflicts([first, second], [pianist])
        self.assertEqual((first.fit_quality, second.fit_quality), ("Manual", "Full"))


if __name__ == "__main__":
    unittest.main()
