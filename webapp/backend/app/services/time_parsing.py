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


def format_time_of_day(minutes: int) -> str:
    """Format minutes after midnight as a 12-hour AM/PM string, e.g. ``1:05 PM``."""
    hour, minute = divmod(minutes, 60)
    hour12 = hour % 12 or 12
    return f"{hour12}:{minute:02d} {'AM' if hour % 24 < 12 else 'PM'}"
