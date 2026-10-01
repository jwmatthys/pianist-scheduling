import importlib.util
from io import BytesIO
from pathlib import Path
from datetime import time
import unittest
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook


SCRIPT_PATH = Path(__file__).resolve().parents[3] / "generate_jury_schedule.py"
SPEC = importlib.util.spec_from_file_location("jury_scheduler_under_test", SCRIPT_PATH)
jury = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(jury)


def make_panel(area, *, start=540, length=10, room="Room A", hourly_break=False, lunch_break=False):
    return {
        "Area": area,
        "Jury Length": length,
        "Start Time": time(start // 60, start % 60),
        "Hourly Break": hourly_break,
        "Lunch Break": lunch_break,
        "Location": room,
    }


def student(name, pianist="", needs_pianist=False, instrument="Voice"):
    return (name, instrument, needs_pianist, pianist)


def student_slots(schedule, area):
    return [slot for slot in schedule[area]["slots"] if slot["type"] == "student"]


def assignment_frame(rows):
    return pd.DataFrame(rows, columns=["Student", "Accompanist"])


class JurySchedulerCharacterizationTests(unittest.TestCase):
    def test_student_roster_filters_jury_flag_and_exposes_unassigned_required_pianist_behavior(self):
        lessons = pd.DataFrame([
            {"Student Name": "Jane Example", "Instrument": "Voice", "Need Pianist": 1, "Jury": 1, "Area": "Voice"},
            {"Student Name": "Alex Sample", "Instrument": "Cello", "Need Pianist": 0, "Jury": 1, "Area": "Strings"},
            {"Student Name": "Unassigned Student", "Instrument": "Violin", "Need Pianist": 1, "Jury": 1, "Area": "Strings"},
            {"Student Name": "Jury Exempt", "Instrument": "Piano", "Need Pianist": 1, "Jury": 0, "Area": "Piano"},
        ])
        assignments = assignment_frame([
            ("Jane Example", "Susan Roberts"),
            ("Unassigned Student", "UNASSIGNED"),
        ])

        def read_excel(_workbook, *, sheet_name, **_kwargs):
            if sheet_name == "Lessons":
                return lessons
            return assignments

        with patch.object(jury.pd, "read_excel", side_effect=read_excel):
            roster = jury.load_students(object(), object())

        self.assertEqual(roster["Voice"], [("Jane Example", "Voice", True, "Susan Roberts")])
        self.assertEqual(roster["Strings"], [
            ("Alex Sample", "Cello", False, ""),
            ("Unassigned Student", "Violin", False, ""),
        ])
        self.assertNotIn("Piano", roster)

    def test_pianist_availability_is_binary_and_tentative_is_unavailable(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Pianist - Susan Roberts"
        sheet.append(["Name:", "Susan Roberts"])
        sheet.append([None, None])
        sheet.append(["Time", "Jury Day Availability"])
        sheet.append([time(9, 0), "Available"])
        sheet.append([time(9, 30), "Tentative"])
        sheet.append([time(10, 0), None])
        contents = BytesIO()
        workbook.save(contents)

        with pd.ExcelFile(BytesIO(contents.getvalue())) as pianists_workbook:
            unavailable = jury.load_pianist_unavailability(pianists_workbook)

        self.assertEqual(unavailable["Susan Roberts"], set(range(570, 630)))
        self.assertNotIn(540, unavailable["Susan Roberts"])

    def test_assigned_pianist_is_fixed_even_when_another_pianist_is_free(self):
        schedule = jury.build_schedule(
            {"Voice": [student("Jane Example", "Susan Roberts", True)]},
            pd.DataFrame([make_panel("Voice")]),
            {"Susan Roberts": set(range(540, 550)), "Michael Lee": set()},
        )

        slots = student_slots(schedule, "Voice")
        self.assertEqual([(slot["time"], slot["pianist"]) for slot in slots], [(550, "Susan Roberts")])

    def test_global_pianist_bookings_prevent_cross_panel_overlap_across_rooms(self):
        schedule = jury.build_schedule(
            {
                "Strings": [student("Jane Example", "Susan Roberts", True)],
                "Voice": [student("Mark Sample", "Susan Roberts", True)],
            },
            pd.DataFrame([
                make_panel("Strings", room="Room 204"),
                make_panel("Voice", room="Room 105"),
            ]),
            {},
        )

        bookings = [student_slots(schedule, area)[0] for area in ("Strings", "Voice")]
        self.assertEqual([booking["time"] for booking in bookings], [540, 550])
        self.assertFalse(
            bookings[0]["time"] < bookings[1]["time"] + bookings[1]["duration"]
            and bookings[1]["time"] < bookings[0]["time"] + bookings[0]["duration"]
        )

    def test_same_room_panels_are_serialized_but_room_is_not_global_resource(self):
        same_room = jury.build_schedule(
            {"Strings": [student("A"), student("B")], "Voice": [student("C")]},
            pd.DataFrame([
                make_panel("Strings", room="Shared Room", length=10),
                make_panel("Voice", room="Shared Room", length=20),
            ]),
            {},
        )
        self.assertEqual(same_room["Voice"]["actual_start"], 560)
        self.assertEqual(same_room["Voice"]["slot_min"], 20)

        separate_rooms = jury.build_schedule(
            {"Strings": [student("A")], "Voice": [student("B")]},
            pd.DataFrame([
                make_panel("Strings", room="Room 204"),
                make_panel("Voice", room="Room 105"),
            ]),
            {},
        )
        self.assertEqual(separate_rooms["Strings"]["actual_start"], 540)
        self.assertEqual(separate_rooms["Voice"]["actual_start"], 540)

    def test_pianist_groups_are_compacted_and_availability_changes_group_order(self):
        compact = jury.build_schedule(
            {"Voice": [
                student("A1", "Susan Roberts", True),
                student("B1", "Michael Lee", True),
                student("A2", "Susan Roberts", True),
                student("B2", "Michael Lee", True),
            ]},
            pd.DataFrame([make_panel("Voice")]),
            {},
        )
        self.assertEqual([slot["student"] for slot in student_slots(compact, "Voice")], ["A1", "A2", "B1", "B2"])

        availability_order = jury.build_schedule(
            {"Voice": [
                student("Late Pianist First In Input", "Susan Roberts", True),
                student("Ready Pianist", "Michael Lee", True),
            ]},
            pd.DataFrame([make_panel("Voice")]),
            {"Susan Roberts": set(range(540, 550))},
        )
        self.assertEqual(
            [slot["student"] for slot in student_slots(availability_order, "Voice")],
            ["Ready Pianist", "Late Pianist First In Input"],
        )

    def test_temporarily_blocked_student_is_deferred_and_revisited(self):
        booked = {"Susan Roberts": set(range(550, 560))}
        schedule = jury.schedule_area(
            [
                student("Susan 1", "Susan Roberts", True),
                student("Susan 2", "Susan Roberts", True),
                student("Michael", "Michael Lee", True),
            ],
            540,
            10,
            False,
            False,
            {"Michael Lee": set(range(540, 550))},
            booked,
            {name: set(minutes) for name, minutes in booked.items()},
        )[0]

        self.assertEqual(
            [(slot["student"], slot["time"]) for slot in schedule],
            [("Susan 1", 540), ("Michael", 550), ("Susan 2", 560)],
        )

    def test_panel_start_delay_shifts_to_reduce_large_early_gap(self):
        schedule = jury.build_schedule(
            {"Voice": [student("No Piano 1"), student("No Piano 2"), student("Susan Student", "Susan", True)]},
            pd.DataFrame([make_panel("Voice", start=540)]),
            {"Susan": set(range(540, 630))},
        )

        self.assertEqual(schedule["Voice"]["earliest"], 540)
        self.assertEqual(schedule["Voice"]["actual_start"], 580)
        self.assertEqual(student_slots(schedule, "Voice")[-1]["time"], 630)

    def test_hourly_break_cycle_is_slot_derived_and_skips_break_for_last_student(self):
        schedule = jury.build_schedule(
            {"Voice": [student(f"Student {index}") for index in range(7)]},
            pd.DataFrame([make_panel("Voice", hourly_break=True)]),
            {},
        )

        breaks = [slot for slot in schedule["Voice"]["slots"] if slot["type"] == "break"]
        self.assertEqual([(item["time"], item["label"]) for item in breaks], [(590, "10-Minute Break")])
        self.assertEqual(sum(slot["type"] == "student" for slot in schedule["Voice"]["slots"]), 7)

    def test_lunch_is_fixed_at_noon_plus_current_slot_end_and_resets_break_counter(self):
        late_lunch = jury.build_schedule(
            {"Voice": [student("Before lunch"), student("After lunch")]},
            pd.DataFrame([make_panel("Voice", start=710, length=20, lunch_break=True)]),
            {},
        )
        lunch = next(slot for slot in late_lunch["Voice"]["slots"] if slot["type"] == "lunch")
        self.assertEqual(lunch["time"], 730)
        self.assertEqual(student_slots(late_lunch, "Voice")[1]["time"], 760)

        reset_schedule = jury.build_schedule(
            {"Voice": [student(f"Meal student {index}") for index in range(10)]},
            pd.DataFrame([make_panel("Voice", start=690, hourly_break=True, lunch_break=True)]),
            {},
        )
        entries = reset_schedule["Voice"]["slots"]
        lunch_time = next(slot["time"] for slot in entries if slot["type"] == "lunch")
        periodic_break_times = [slot["time"] for slot in entries if slot["type"] == "break"]
        self.assertEqual(lunch_time, 720)
        self.assertEqual(periodic_break_times, [800])

    def test_sixty_minute_jury_with_hourly_break_raises_zero_cycle_error(self):
        with self.assertRaises(ZeroDivisionError):
            jury.build_schedule(
                {"Voice": [student("One"), student("Two")]},
                pd.DataFrame([make_panel("Voice", length=60, hourly_break=True)]),
                {},
            )

    def test_fully_unavailable_required_pianist_is_scheduled_after_midnight(self):
        schedule = jury.build_schedule(
            {"Voice": [student("Unscheduled in horizon", "Susan Roberts", True)]},
            pd.DataFrame([make_panel("Voice")]),
            {"Susan Roberts": set(range(24 * 60))},
        )

        self.assertEqual(student_slots(schedule, "Voice")[0]["time"], 24 * 60)
        self.assertEqual(jury.fmt(24 * 60), "12:00 PM")

    def test_schedule_is_deterministic_for_identical_inputs(self):
        students = {"Voice": [student("A", "Susan Roberts", True), student("B"), student("C", "Susan Roberts", True)]}
        jury_info = pd.DataFrame([make_panel("Voice", hourly_break=True)])
        first = jury.build_schedule(students, jury_info, {})
        second = jury.build_schedule(students, jury_info, {})

        self.assertEqual(first, second)


if __name__ == "__main__":
    unittest.main()