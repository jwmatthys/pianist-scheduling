import unittest
from datetime import time

from app.services.time_parsing import format_time_of_day, parse_time_of_day


class TimeParsingTests(unittest.TestCase):
    def test_parses_common_formats(self):
        cases = {"8:05": 485, "08:05": 485, "13:00:00": 780, "1:30 PM": 810, "1:30pm": 810,
                 "8 AM": 480, "12 PM": 720, "12:00 AM": 0, " 9:00 a.m. ": 540}
        for text, minutes in cases.items():
            self.assertEqual(parse_time_of_day(text), minutes, text)

    def test_parses_excel_values(self):
        self.assertEqual(parse_time_of_day(time(14, 15)), 855)
        self.assertEqual(parse_time_of_day(0.5), 720)

    def test_rejects_invalid(self):
        for bad in ("", "noon", "25:00", "13 PM", "9:75", None, True):
            self.assertIsNone(parse_time_of_day(bad), bad)

    def test_formats_twelve_hour(self):
        self.assertEqual(format_time_of_day(0), "12:00 AM")
        self.assertEqual(format_time_of_day(785), "1:05 PM")
        self.assertEqual(format_time_of_day(720), "12:00 PM")


if __name__ == "__main__":
    unittest.main()
