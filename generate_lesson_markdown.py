"""Generate pianist and instructor schedules from a master lesson workbook."""

import argparse
from datetime import datetime, time, timedelta
from pathlib import Path
import re

import pandas as pd


DEFAULT_INPUT = "lesson_final_fa26.xlsx"
DEFAULT_SHEET = "RosterReportAppliedLessonsFall2"
DAY_ORDER = {
    "M": (0, "Monday"),
    "MON": (0, "Monday"),
    "MONDAY": (0, "Monday"),
    "T": (1, "Tuesday"),
    "TU": (1, "Tuesday"),
    "TUE": (1, "Tuesday"),
    "TUESDAY": (1, "Tuesday"),
    "W": (2, "Wednesday"),
    "WED": (2, "Wednesday"),
    "WEDNESDAY": (2, "Wednesday"),
    "R": (3, "Thursday"),
    "TH": (3, "Thursday"),
    "THU": (3, "Thursday"),
    "THURSDAY": (3, "Thursday"),
    "F": (4, "Friday"),
    "FRI": (4, "Friday"),
    "FRIDAY": (4, "Friday"),
    "S": (5, "Saturday"),
    "SAT": (5, "Saturday"),
    "SATURDAY": (5, "Saturday"),
    "U": (6, "Sunday"),
    "SUN": (6, "Sunday"),
    "SUNDAY": (6, "Sunday"),
}


def clean(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def email_address(value):
    value = clean(value)
    match = re.search(r"<([^<>\s]+@[^<>\s]+)>", value)
    return match.group(1) if match else value


def format_time(value):
    if pd.isna(value):
        return ""
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        return value.strftime("%I:%M %p").lstrip("0")
    parsed = pd.to_datetime(value, errors="coerce")
    if not pd.isna(parsed):
        return parsed.strftime("%I:%M %p").lstrip("0")
    return clean(value)


def add_minutes(value, minutes):
    if pd.isna(value):
        return ""
    if isinstance(value, datetime):
        return value + timedelta(minutes=minutes)
    if isinstance(value, time):
        base = datetime(2000, 1, 1, value.hour, value.minute, value.second)
        return base + timedelta(minutes=minutes)
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return ""
    return parsed + timedelta(minutes=minutes)


def normalize_day(value):
    key = clean(value).upper()
    return DAY_ORDER.get(key, (99, clean(value)))[1]


def load_lessons(path, sheet):
    workbook = pd.ExcelFile(path)
    if sheet not in workbook.sheet_names:
        raise ValueError(
            f"Sheet '{sheet}' not found. Available sheets: {', '.join(workbook.sheet_names)}"
        )

    data = workbook.parse(sheet)
    required = {
        "Instructor First",
        "Instructor Last",
        "Day",
        "Start",
        "End",
        "Room",
        "Student First Name",
        "Student Last Name",
        "Pianist First",
        "Pianist Last",
    }
    missing = sorted(required - set(data.columns))
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    lessons = []
    for _, row in data.iterrows():
        preferred = row["Student Preferred"] if "Student Preferred" in data.columns else ""
        student_first = clean(preferred) or clean(row["Student First Name"])
        student = " ".join(
            part for part in (student_first, clean(row["Student Last Name"])) if part
        )
        pianist = " ".join(
            part for part in (clean(row["Pianist First"]), clean(row["Pianist Last"])) if part
        )
        instructor = " ".join(
            part for part in (clean(row["Instructor First"]), clean(row["Instructor Last"])) if part
        )
        instructor_last = clean(row["Instructor Last"])
        instructor_email = clean(row.get("Instructor Email", ""))
        pianist_email = email_address(row.get("Pianist Email", ""))
        if not student or not pianist:
            continue

        day = normalize_day(row["Day"])
        day_index = DAY_ORDER.get(clean(row["Day"]).upper(), (99, day))[0]
        end_value = row["End"]
        if pd.isna(end_value):
            end_value = add_minutes(row["Start"], 50)
        lessons.append(
            {
                "student": student,
                "pianist": pianist,
                "pianist_email": pianist_email,
                "instructor": instructor or "Unknown instructor",
                "instructor_last": instructor_last or "Unknown instructor",
                "instructor_email": instructor_email,
                "day": day,
                "day_index": day_index,
                "start": format_time(row["Start"]),
                "end": format_time(end_value),
                "room": clean(row["Room"]) or "No room listed",
                "start_value": row["Start"],
            }
        )

    return lessons


def sort_key(lesson):
    value = lesson["start_value"]
    if isinstance(value, datetime):
        minutes = value.hour * 60 + value.minute
    elif isinstance(value, time):
        minutes = value.hour * 60 + value.minute
    else:
        parsed = pd.to_datetime(value, errors="coerce")
        minutes = parsed.hour * 60 + parsed.minute if not pd.isna(parsed) else 9999
    return lesson["day_index"], minutes, lesson["student"].lower()


def lesson_line(lesson, include_pianist=False):
    if include_pianist:
        return (
            f"- **{lesson['day']} {lesson['start']}–{lesson['end']}**"
            f" — {lesson['student']} — {lesson['room']} — {lesson['pianist']}"
            f" — {lesson['pianist_email']}"
        )
    return (
        f"- **{lesson['day']} {lesson['start']}–{lesson['end']}**"
        f" — {lesson['student']} — {lesson['room']} — {lesson['instructor_last']}"
        f" — {lesson['instructor_email']}"
    )


def build_markdown(lessons, source_name):
    by_pianist = {}
    by_instructor = {}
    for lesson in lessons:
        by_pianist.setdefault(lesson["pianist"], []).append(lesson)
        by_instructor.setdefault(lesson["instructor"], []).append(lesson)

    lines = [f"# Lesson Schedules", "", f"Source: `{source_name}`", ""]
    lines.append("## Pianist Schedules")
    lines.append("")
    for pianist in sorted(by_pianist, key=str.casefold):
        pianist_lessons = sorted(by_pianist[pianist], key=sort_key)
        lines.extend([f"### {pianist}", ""])
        lines.extend(lesson_line(lesson) for lesson in pianist_lessons)
        lines.append("")

    lines.extend(["## Students by Instructor", ""])
    for instructor in sorted(by_instructor, key=str.casefold):
        instructor_lessons = sorted(
            by_instructor[instructor], key=sort_key
        )
        lines.extend([f"### {instructor}", ""])
        lines.extend(lesson_line(lesson, include_pianist=True) for lesson in instructor_lessons)
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def main():
    parser = argparse.ArgumentParser(
        description="Generate pianist and instructor Markdown schedules from a master Excel workbook."
    )
    parser.add_argument(
        "--input", "-i", default=DEFAULT_INPUT,
        help=f"Master Excel workbook (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--output", "-o",
        help="Markdown output path (default: same name as input with .md extension)",
    )
    parser.add_argument(
        "--sheet", default=DEFAULT_SHEET,
        help=f"Worksheet name (default: {DEFAULT_SHEET})",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else input_path.with_suffix(".md")
    lessons = load_lessons(input_path, args.sheet)
    output_path.write_text(build_markdown(lessons, input_path.name), encoding="utf-8")
    print(f"Wrote {len(lessons)} assigned lessons to '{output_path}'")
    print(f"  Pianists: {len({lesson['pianist'] for lesson in lessons})}")
    print(f"  Instructors: {len({lesson['instructor'] for lesson in lessons})}")


if __name__ == "__main__":
    main()
