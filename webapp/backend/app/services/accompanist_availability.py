"""Adapter from shared windows to the Accompanist module's legacy slots."""

from dataclasses import dataclass

from .availability import AvailabilityIssue, AvailabilityWindow, validate_window

SLOT_MINUTES = 30


@dataclass(frozen=True)
class AccompanistAvailabilitySlot:
    day: str
    slot_start_minute: int
    status: str


def windows_to_accompanist_slots(
    windows: list[AvailabilityWindow],
) -> tuple[list[AccompanistAvailabilitySlot], list[AvailabilityIssue]]:
    slots: list[AccompanistAvailabilitySlot] = []
    issues: list[AvailabilityIssue] = []
    for window in windows:
        invalid = validate_window(window)
        if invalid:
            issues.append(invalid)
            continue
        if window.start_minute % SLOT_MINUTES or window.end_minute % SLOT_MINUTES:
            issues.append(AvailabilityIssue(
                "error",
                "UNSUPPORTED_ACCOMPANIST_BOUNDARY",
                "Accompanist availability must start and end on a 30-minute boundary; no rounding was applied.",
                window.source_row,
            ))
            continue
        if window.status == "Unavailable":
            continue
        slots.extend(
            AccompanistAvailabilitySlot(window.day, minute, window.status)
            for minute in range(window.start_minute, window.end_minute, SLOT_MINUTES)
        )
    return slots, issues