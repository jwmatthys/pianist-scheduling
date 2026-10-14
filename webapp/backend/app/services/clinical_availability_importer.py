"""Pure weekly-availability cell parsing shared by the Clinical Student Roster import.

This is a subcomponent of ``clinical_students_importer``: it turns the weekday cells of
one spreadsheet row into normalized Availability Windows. It has no database access and
no knowledge of students or files.

Each weekday cell holds text such as ``8:00-12:00``, ``T 1:00 PM-4:00 PM`` or
``8:00-10:00; T 1:00-4:00``. Entries default to Available; a ``T``/``Tentative`` prefix
marks Tentative. Anything not listed is simply not stored (and so stays unavailable).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Mapping

from .availability import WEEKDAYS, AvailabilityStatus, AvailabilityWindow, normalize_day, normalize_windows
from .clinical_import_common import cell_text
from .time_parsing import format_time_of_day, parse_time_of_day

MAX_ENTRIES_PER_CELL = 20

_STATUS_PREFIX = re.compile(r"^(tentative|available|t)(?=[\s:\d])\s*:?\s*", re.IGNORECASE)
_RANGE_SPLIT = re.compile(r"\s*(?:-|–|—|\bto\b)\s*", re.IGNORECASE)
_BARE_HOUR = re.compile(r"^\d{1,2}$")
_HAS_MERIDIEM = re.compile(r"[ap]\.?m", re.IGNORECASE)
_EARLIEST_AM_MINUTE = 7 * 60


@dataclass
class WeekParse:
    windows: list[AvailabilityWindow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def suggest_day_columns(columns: list[str], taken: set[str]) -> dict[str, str | None]:
    """Suggest a source column for each weekday (Monday/Mon, Tuesday/Tue/Tues, ...)."""
    days: dict[str, str | None] = {}
    used = set(taken)
    for day in WEEKDAYS:
        match = next((column for column in columns if column not in used and normalize_day(column.strip()) == day), None)
        days[day] = match
        if match:
            used.add(match)
    return days


def mapped_day_columns(days: Mapping[str, str | None]) -> dict[str, str]:
    return {day: column for day in WEEKDAYS if (column := days.get(day))}


def _parse_time(text: str) -> int | None:
    text = text.strip()
    if _BARE_HOUR.match(text):
        text = f"{text}:00"
    return parse_time_of_day(text)


def parse_cell(day: str, text: str, row_number: int) -> tuple[list[AvailabilityWindow], list[str]]:
    """Parse one weekday cell into Availability Windows and error messages."""
    windows: list[AvailabilityWindow] = []
    errors: list[str] = []
    entries = [entry.strip() for entry in text.split(";") if entry.strip()]
    if len(entries) > MAX_ENTRIES_PER_CELL:
        return [], [f"{day}: too many availability entries (the limit is {MAX_ENTRIES_PER_CELL})."]
    for entry in entries:
        status: AvailabilityStatus = "Available"
        body = entry
        if match := _STATUS_PREFIX.match(entry):
            status = "Tentative" if match.group(1).casefold() in ("t", "tentative") else "Available"
            body = entry[match.end():]
        parts = _RANGE_SPLIT.split(body.strip())
        if len(parts) != 2 or not all(parts):
            errors.append(f"{day}: '{entry}' is not a valid time range such as 8:00-12:00.")
            continue
        start, end = _parse_time(parts[0]), _parse_time(parts[1])
        if start is None or end is None:
            errors.append(f"{day}: '{entry}' contains a time that could not be read.")
            continue
        if not _HAS_MERIDIEM.search(body) and start < _EARLIEST_AM_MINUTE and 0 < end < 12 * 60:
            # Unmarked times before 7:00 are read as afternoon (e.g. "T 1:00-4:00"), because
            # the product's availability grid starts at 7:00 AM.
            start, end = start + 12 * 60, end + 12 * 60
        if end <= start:
            errors.append(f"{day}: '{entry}' ends at or before it starts.")
            continue
        windows.append(AvailabilityWindow(
            day=day, start_minute=start, end_minute=end, status=status, source="import", source_row=row_number
        ))
    return windows, errors


def _describe_conflicts(windows: list[AvailabilityWindow]) -> list[str]:
    messages: list[str] = []
    for index, first in enumerate(windows):
        for second in windows[index + 1:]:
            if (
                first.day == second.day
                and first.status != second.status
                and first.start_minute < second.end_minute
                and second.start_minute < first.end_minute
            ):
                messages.append(
                    f"{first.day}: {first.status} {format_time_of_day(first.start_minute)}-"
                    f"{format_time_of_day(first.end_minute)} overlaps {second.status} "
                    f"{format_time_of_day(second.start_minute)}-{format_time_of_day(second.end_minute)}."
                )
    return messages


def parse_week(record: Mapping[str, object], day_columns: Mapping[str, str], row_number: int) -> WeekParse:
    """Parse every mapped weekday cell of one row; blank cells contribute no windows."""
    collected: list[AvailabilityWindow] = []
    result = WeekParse()
    for day, column in day_columns.items():
        text = cell_text(record.get(column))
        if not text:
            continue
        windows, errors = parse_cell(day, text, row_number)
        collected.extend(windows)
        result.errors.extend(errors)
    conflicts = _describe_conflicts(collected)
    result.errors.extend(conflicts)
    if not conflicts:
        normalized, issues = normalize_windows(collected)
        result.warnings.extend(issue.message for issue in issues if issue.severity == "warning")
        result.errors.extend(issue.message for issue in issues if issue.severity == "error")
        result.windows = normalized
    return result


def describe_change(existing: int, imported: int) -> str:
    if existing == 0 and imported == 0:
        return "No availability change"
    if existing == 0:
        return f"Add {imported} availability window{'s' if imported != 1 else ''}"
    if imported == 0:
        return f"Clear availability ({existing} window{'s' if existing != 1 else ''})"
    return f"Replace {existing} availability window{'s' if existing != 1 else ''} with {imported}"
