"""
Sample Workbook Generator

Writes three sample input workbooks matching the schema documented in the
README, so the scheduling scripts can be exercised without real data:

  - lesson_information.xlsx    ("Lessons" sheet)
  - pianist_availability.xlsx  ("Pianist - <Name>" sheets)
  - jury_information.xlsx      ("Jury Information" sheet)
"""

import random
from openpyxl import Workbook
from datetime import time

DAYS_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
INSTRUMENTS = ["Piano", "Violin", "Cello", "Flute", "Voice", "Clarinet", "Trumpet"]
ROOMS = ["Room 101", "Room 102", "Room 103", "Room 104"]
AREAS = ["Area A", "Area B", "Area C"]
TEACHERS = [f"Teacher {i}" for i in range(1, 6)]
FIRST_NAMES = ["Alice", "Ben", "Cara", "David", "Ella", "Finn", "Grace", "Henry",
               "Iris", "Jack", "Kara", "Liam", "Mia", "Noah", "Olive", "Paul",
               "Quinn", "Rosa", "Sam", "Tara"]
LAST_NAMES = ["Chen", "Torres", "Nguyen", "Osei", "Ramirez", "Walsh", "Kim",
              "Ford", "Patel", "Singh", "Diaz", "Cohen", "Wright", "Ito"]

PIANIST_NAMES = ["Grace Kim", "Henry Ford", "Ivy Sato", "Jonas Weber", "Mei Lin"]
PIANIST_MAX_HOURS = [12, 10, 8, 15, 10]

STUDENT_COUNT = 50
RANDOM_SEED = 42


def make_lessons(rng):
    lessons = []
    for i in range(STUDENT_COUNT):
        student = f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)}"
        instrument = rng.choice(INSTRUMENTS)
        day = rng.choice(DAYS_ORDER)
        start_hour = rng.randint(8, 16)
        start_minute = rng.choice([0, 30])
        start = time(start_hour, start_minute)
        duration = rng.choice([30, 45, 60])
        end_minutes = start_hour * 60 + start_minute + duration
        end = time(end_minutes // 60, end_minutes % 60)
        room = rng.choice(ROOMS)
        area = rng.choice(AREAS)
        teacher = rng.choice(TEACHERS)
        lessons.append((teacher, student, day, start, end, room, instrument, area))
    return lessons


def write_lessons(path, rng):
    wb = Workbook()
    ws = wb.active
    ws.title = "Lessons"
    headers = ["Lesson Teacher Name", "Student Name", "Lesson Day",
               "Lesson Start Time", "Lesson End Time", "Lesson Location",
               "Instrument", "Need Pianist", "Required Accompanist",
               "Jury", "Area"]
    ws.append(headers)
    for teacher, student, day, start, end, room, instrument, area in make_lessons(rng):
        ws.append([teacher, student, day, start, end, room, instrument,
                   1, "", 1, area])
    wb.save(path)


def make_day_availability(rng):
    """Return a list of 20 half-hour statuses (8:00-18:00) with a mix of
    Available/Tentative/Unavailable blocks."""
    statuses = ["Available"] * 20
    # Carve out 1-3 contiguous unavailable/tentative blocks per day
    for _ in range(rng.randint(1, 3)):
        block_len = rng.randint(1, 4)
        block_start = rng.randint(0, 20 - block_len)
        status = rng.choice(["Unavailable", "Tentative"])
        for j in range(block_start, block_start + block_len):
            statuses[j] = status
    return statuses


def write_pianists(path, rng):
    wb = Workbook()
    wb.remove(wb.active)
    for name, max_hours in zip(PIANIST_NAMES, PIANIST_MAX_HOURS):
        ws = wb.create_sheet(f"Pianist - {name}")
        ws.append(["Name:", name])
        ws.append(["Max weekly hours:", max_hours])
        ws.append([])
        ws.append(["Time", *DAYS_ORDER, "Jury Day Availability"])

        day_statuses = {day: make_day_availability(rng) for day in DAYS_ORDER}
        jury_statuses = make_day_availability(rng)

        for slot in range(20):  # 8:00 - 18:00 in 30-minute slots
            minutes = 8 * 60 + slot * 30
            row_time = time(minutes // 60, minutes % 60)
            row = [row_time] + [day_statuses[day][slot] for day in DAYS_ORDER] \
                + [jury_statuses[slot]]
            ws.append(row)
    wb.save(path)


def write_jury_info(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Jury Information"
    ws.append(["Area", "Start Time", "Jury Length", "Location",
               "Hourly Break", "Lunch Break"])
    for area, room in zip(AREAS, ROOMS):
        ws.append([area, time(9, 0), 15, room, 1, 1])
    wb.save(path)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate sample scheduling workbooks.")
    parser.add_argument("--seed", default=str(RANDOM_SEED),
                         help="Integer seed for reproducible output, or 'random' for a new seed each run")
    args = parser.parse_args()

    seed = random.randrange(2**32) if args.seed == "random" else int(args.seed)
    rng = random.Random(seed)
    write_lessons("lesson_information.xlsx", rng)
    write_pianists("pianist_availability.xlsx", rng)
    write_jury_info("jury_information.xlsx")
    print(f"Wrote lesson_information.xlsx, pianist_availability.xlsx, jury_information.xlsx (seed={seed})")
