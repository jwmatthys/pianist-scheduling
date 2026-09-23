"""Markdown report generation, adapted from generate_lesson_markdown.py to
read from the webapp's Lesson/Pianist ORM rows instead of an Excel workbook.
"""

from __future__ import annotations

from ..models import DAYS_ORDER, Lesson, Pianist

DAY_INDEX = {day: i for i, day in enumerate(DAYS_ORDER)}


def _fmt_time(minutes: int) -> str:
    h, m = divmod(minutes, 60)
    period = "AM" if h < 12 else "PM"
    h12 = h % 12
    if h12 == 0:
        h12 = 12
    return f"{h12}:{m:02d} {period}"


def _lesson_line(lesson: Lesson, pianist_name: str, pianist_email: str, include_pianist: bool) -> str:
    window = f"{lesson.day} {_fmt_time(lesson.start_minute)}\u2013{_fmt_time(lesson.end_minute)}"
    if include_pianist:
        return (
            f"- **{window}** \u2014 {lesson.student} \u2014 {lesson.room or 'No room listed'}"
            f" \u2014 {pianist_name} \u2014 {pianist_email}"
        )
    return (
        f"- **{window}** \u2014 {lesson.room or 'No room listed'} \u2014 {lesson.student}"
        f" \u2014 {lesson.teacher or 'Unknown instructor'}"
    )


def _sort_key(lesson: Lesson):
    return DAY_INDEX.get(lesson.day, 99), lesson.start_minute, lesson.student.lower()


def build_markdown(lessons: list[Lesson], pianists_by_id: dict[int, Pianist], source_name: str) -> str:
    assigned = [l for l in lessons if l.assigned_pianist_id is not None]

    by_pianist: dict[str, list[Lesson]] = {}
    by_teacher: dict[str, list[Lesson]] = {}
    for lesson in assigned:
        pianist = pianists_by_id.get(lesson.assigned_pianist_id)
        pname = pianist.name if pianist else "Unknown"
        by_pianist.setdefault(pname, []).append(lesson)
        by_teacher.setdefault(lesson.teacher or "Unknown instructor", []).append(lesson)

    lines = ["# Lesson Schedules", "", f"Source: `{source_name}`", ""]

    lines.append("## Pianist Schedules")
    lines.append("")
    for pname in sorted(by_pianist, key=str.casefold):
        pianist_lessons = sorted(by_pianist[pname], key=_sort_key)
        pianist_email = next(
            (p.email for p in pianists_by_id.values() if p.name == pname), ""
        )
        lines.extend([f"### {pname}", ""])
        lines.extend(_lesson_line(l, pname, pianist_email, include_pianist=False) for l in pianist_lessons)
        lines.append("")

    lines.extend(["## Students by Instructor", ""])
    for teacher in sorted(by_teacher, key=str.casefold):
        teacher_lessons = sorted(by_teacher[teacher], key=_sort_key)
        lines.extend([f"### {teacher}", ""])
        for lesson in teacher_lessons:
            pianist = pianists_by_id.get(lesson.assigned_pianist_id)
            pname = pianist.name if pianist else "Unknown"
            pemail = pianist.email if pianist else ""
            lines.append(_lesson_line(lesson, pname, pemail, include_pianist=True))
        lines.append("")

    lines.extend(["## Students by Name", ""])
    for lesson in sorted(assigned, key=lambda l: l.student.casefold()):
        pianist = pianists_by_id.get(lesson.assigned_pianist_id)
        pname = pianist.name if pianist else "Unknown"
        pemail = pianist.email if pianist else ""
        lines.append(f"- **{lesson.student}** \u2014 {pname} \u2014 {pemail}")

    unassigned = [l for l in lessons if l.assigned_pianist_id is None and l.need_pianist]
    if unassigned:
        lines.extend(["", "## Unassigned Lessons", ""])
        for lesson in sorted(unassigned, key=_sort_key):
            lines.append(
                f"- **{lesson.day} {_fmt_time(lesson.start_minute)}\u2013{_fmt_time(lesson.end_minute)}**"
                f" \u2014 {lesson.student} \u2014 {lesson.teacher}"
            )

    return "\n".join(lines).rstrip() + "\n"
