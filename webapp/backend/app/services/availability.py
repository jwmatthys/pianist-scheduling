"""Shared availability values and interval normalization."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import datetime, time
from typing import Literal

WEEKDAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")

AvailabilityStatus = Literal["Available", "Tentative", "Unavailable"]
IssueSeverity = Literal["error", "warning"]

_DAY_ALIASES = {
    "monday": "Monday", "mon": "Monday",
    "tuesday": "Tuesday", "tue": "Tuesday", "tues": "Tuesday",
    "wednesday": "Wednesday", "wed": "Wednesday",
    "thursday": "Thursday", "thu": "Thursday", "thur": "Thursday", "thurs": "Thursday",
    "friday": "Friday", "fri": "Friday",
    "saturday": "Saturday", "sat": "Saturday",
    "sunday": "Sunday", "sun": "Sunday",
}
_TIME_PATTERN = re.compile(r"^(\d{1,2}):(\d{2})(?:\s*([AaPp][Mm]))?$")
_STATUS_LOOKUP = {status.casefold(): status for status in ("Available", "Tentative", "Unavailable")}


@dataclass(frozen=True)
class AvailabilityWindow:
    day: str
    start_minute: int
    end_minute: int
    status: AvailabilityStatus
    source: Literal["manual", "import"] = "manual"
    source_row: int | None = None


@dataclass(frozen=True)
class AvailabilityIssue:
    severity: IssueSeverity
    code: str
    message: str
    row_number: int | None = None


def normalize_day(value: object) -> str | None:
    if not isinstance(value, str):
        return None
    key = value.strip().casefold().rstrip(".")
    return _DAY_ALIASES.get(key)


def parse_time_minutes(value: object, *, allow_end_of_day: bool = False) -> int | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, datetime):
        return value.hour * 60 + value.minute
    if isinstance(value, time):
        return value.hour * 60 + value.minute
    if isinstance(value, (int, float)):
        if not math.isfinite(float(value)) or not 0 <= float(value) < 1:
            return None
        return int(round(float(value) * 24 * 60))

    match = _TIME_PATTERN.fullmatch(str(value).strip())
    if not match:
        return None
    hour, minute = int(match.group(1)), int(match.group(2))
    meridiem = match.group(3)
    if minute > 59:
        return None
    if meridiem:
        if not 1 <= hour <= 12:
            return None
        hour %= 12
        if meridiem.casefold() == "pm":
            hour += 12
    elif hour == 24 and minute == 0 and allow_end_of_day:
        return 24 * 60
    elif hour > 23:
        return None
    return hour * 60 + minute


def normalize_status(value: object) -> AvailabilityStatus | None:
    if not isinstance(value, str):
        return None
    return _STATUS_LOOKUP.get(value.strip().casefold())  # type: ignore[return-value]


def validate_window(window: AvailabilityWindow) -> AvailabilityIssue | None:
    if window.day not in WEEKDAYS:
        return AvailabilityIssue("error", "INVALID_DAY", f"Unrecognized weekday '{window.day}'.", window.source_row)
    if not 0 <= window.start_minute < 24 * 60:
        return AvailabilityIssue("error", "INVALID_START", "Start time must be within the selected weekday.", window.source_row)
    if not 0 < window.end_minute <= 24 * 60 or window.end_minute <= window.start_minute:
        return AvailabilityIssue("error", "INVALID_RANGE", "End time must be later than start time on the same weekday.", window.source_row)
    if window.status not in _STATUS_LOOKUP.values():
        return AvailabilityIssue("error", "INVALID_STATUS", f"Unrecognized availability status '{window.status}'.", window.source_row)
    return None


def normalize_windows(
    windows: list[AvailabilityWindow],
) -> tuple[list[AvailabilityWindow], list[AvailabilityIssue]]:
    """Merge same-status overlaps/adjacency and report conflicting overlaps."""
    day_order = {day: index for index, day in enumerate(WEEKDAYS)}
    ordered = sorted(windows, key=lambda item: (day_order.get(item.day, len(WEEKDAYS)), item.start_minute, item.end_minute))
    normalized: list[AvailabilityWindow] = []
    issues: list[AvailabilityIssue] = []

    for window in ordered:
        invalid = validate_window(window)
        if invalid:
            issues.append(invalid)
            normalized.append(window)
            continue

        duplicates = [
            existing for existing in normalized
            if (existing.day, existing.start_minute, existing.end_minute, existing.status)
            == (window.day, window.start_minute, window.end_minute, window.status)
        ]
        if duplicates:
            issues.append(AvailabilityIssue(
                "warning", "DUPLICATE_WINDOW", f"Duplicate {window.day} availability window was ignored.", window.source_row
            ))
            continue

        conflicts = [
            existing for existing in normalized
            if existing.day == window.day
            and window.start_minute < existing.end_minute
            and window.end_minute > existing.start_minute
            and window.status != existing.status
        ]
        if conflicts:
            issues.append(AvailabilityIssue(
                "error", "CONFLICTING_OVERLAP",
                f"{window.day} {window.start_minute}-{window.end_minute} overlaps a different availability status.",
                window.source_row,
            ))
            normalized.append(window)
            continue

        mergeable = [
            existing for existing in normalized
            if existing.day == window.day
            and existing.status == window.status
            and window.start_minute <= existing.end_minute
            and window.end_minute >= existing.start_minute
        ]
        if mergeable:
            start_minute = min([window.start_minute, *(item.start_minute for item in mergeable)])
            end_minute = max([window.end_minute, *(item.end_minute for item in mergeable)])
            normalized = [item for item in normalized if item not in mergeable]
            normalized.append(AvailabilityWindow(
                day=window.day,
                start_minute=start_minute,
                end_minute=end_minute,
                status=window.status,
                source=window.source,
                source_row=window.source_row,
            ))
        else:
            normalized.append(window)

    return sorted(normalized, key=lambda item: (day_order.get(item.day, len(WEEKDAYS)), item.start_minute, item.end_minute)), issues