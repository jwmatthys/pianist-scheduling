"""Domain-neutral time-of-day parsing and formatting shared by imports and forms.

Times are minutes after midnight. This module builds on the shared
``availability.parse_time_minutes`` primitive and only widens the accepted text
formats; it has no module-specific knowledge.
"""

from __future__ import annotations

import re

from .availability import parse_time_minutes

_BARE_HOUR = re.compile(r"^(\d{1,2})\s*([AaPp])\.?[Mm]\.?$")
_SPACED_MERIDIEM = re.compile(r"^(\d{1,2}:\d{2})(?::00)?\s*([AaPp])\.?[Mm]\.?$")
_SECONDS = re.compile(r"^(\d{1,2}:\d{2}):00$")
_HOUR_ONLY = re.compile(r"^\d{1,2}$")
_MERIDIEM = re.compile(r"[ap]\.?m\.?", re.IGNORECASE)


def parse_time_of_day(value: object) -> int | None:
    """Parse Excel time cells, H:MM / HH:MM(:00), and 12-hour AM/PM text into minutes."""
    if isinstance(value, str):
        text = value.strip()
        if match := _BARE_HOUR.match(text):
            text = f"{match.group(1)}:00 {match.group(2)}M"
        elif match := _SPACED_MERIDIEM.match(text):
            text = f"{match.group(1)} {match.group(2)}M"
        elif match := _SECONDS.match(text):
            text = match.group(1)
        return parse_time_minutes(text)
    return parse_time_minutes(value)


def parse_time_range(start_text: str, end_text: str) -> tuple[int | None, int | None]:
    """Parse an imported time range, inferring an afternoon end when needed.

    Unmarked ranges that reverse numerically are read as a morning start and
    afternoon end. Existing unmarked early-hour ranges such as ``1:00-4:00``
    continue to be interpreted as afternoon times.
    """
    start_value, end_value = start_text.strip(), end_text.strip()
    if _HOUR_ONLY.fullmatch(start_value):
        start_value = f"{start_value}:00"
    if _HOUR_ONLY.fullmatch(end_value):
        end_value = f"{end_value}:00"

    start = parse_time_of_day(start_value)
    end = parse_time_of_day(end_value)
    has_meridiem = bool(_MERIDIEM.search(start_text) or _MERIDIEM.search(end_text))
    if start is None or end is None or has_meridiem:
        return start, end

    if end < start:
        end += 12 * 60
    elif start < 7 * 60 and 0 < end < 12 * 60:
        start += 12 * 60
        end += 12 * 60
    return start, end


def format_time_of_day(minutes: int) -> str:
    """Format minutes after midnight as a 12-hour AM/PM string, e.g. ``1:05 PM``."""
    hour, minute = divmod(minutes, 60)
    hour12 = hour % 12 or 12
    return f"{hour12}:{minute:02d} {'AM' if hour % 24 < 12 else 'PM'}"
