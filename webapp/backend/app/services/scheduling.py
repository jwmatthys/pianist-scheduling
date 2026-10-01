"""Core best-fit assignment engine.

This is a direct adaptation of the fit-scoring / assignment-loop logic from
the original ``generate_pianist_schedule.py`` CLI tool, retargeted to work
against the webapp's ORM objects (``Lesson`` / ``Pianist`` / ``AvailabilitySlot``)
instead of pandas DataFrames read from Excel. The scoring tiers, tie-break
order, required-pianist handling, conflict prevention, and workload-cap logic
are preserved for the desktop application.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..models import DAYS_ORDER

NEAR_FIT_MARGIN_MINUTES = 15
OVERLAP_MAX_MINUTES = 30

FIT_FULL = 4
FIT_PARTIAL = 3
FIT_NEAR = 2
FIT_OVERLAP = 1
FIT_NONE = 0

FIT_LABELS = {
    FIT_FULL: "Full",
    FIT_PARTIAL: "Partial",
    FIT_NEAR: "Near",
    FIT_OVERLAP: "Overlap",
    FIT_NONE: "None",
}


@dataclass
class EngineLesson:
    """Lightweight in-memory representation of a Lesson row for the engine."""

    id: int
    day: str
    start_min: int
    end_min: int
    required_pianist_name: str = ""
    need_pianist: bool = True

    assigned_pianist_id: int | None = None
    fit_quality: str = ""
    notes: str = ""
    hours: float = 0.0


@dataclass
class EnginePianist:
    id: int
    name: str
    max_hours: float | None
    # avail[day][slot_start_minute] = "Available" | "Tentative" | "Unavailable"
    avail: dict = field(default_factory=dict)


def get_fit(day: str, start_min: int, end_min: int, avail: dict) -> tuple[int, bool]:
    """Returns (fit_score, has_tentative), ignoring overlap with other lessons."""
    day_avail = avail.get(day, {})
    if not day_avail:
        return FIT_NONE, False

    covered = [
        (sm, st) for sm, st in day_avail.items()
        if sm < end_min and sm + 30 > start_min
    ]
    if covered:
        statuses = [st for _, st in covered]
        if all(s == "Available" for s in statuses):
            return FIT_FULL, False
        if all(s in ("Available", "Tentative") for s in statuses):
            return FIT_PARTIAL, True

    for slot_min, status in day_avail.items():
        if status in ("Available", "Tentative"):
            diff = slot_min - start_min
            if 0 <= diff <= NEAR_FIT_MARGIN_MINUTES:
                return FIT_NEAR, (status == "Tentative")

    return FIT_NONE, False


def get_fit_for_window(day: str, start_min: int, end_min: int, avail: dict) -> int:
    day_avail = avail.get(day, {})
    covered = [
        status for slot_min, status in day_avail.items()
        if slot_min < end_min and slot_min + 30 > start_min
    ]
    if not covered:
        return FIT_NONE
    if all(status == "Available" for status in covered):
        return FIT_FULL
    if all(status in ("Available", "Tentative") for status in covered):
        return FIT_PARTIAL
    return FIT_NONE


def overlap_minutes(a_start, a_end, b_start, b_end) -> int:
    return max(0, min(a_end, b_end) - max(a_start, b_start))


def merge_intervals(windows: list[tuple[int, int]]) -> list[tuple[int, int]]:
    merged: list[tuple[int, int]] = []
    for start, end in sorted(windows):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def has_conflict(existing_assigns, day, s_min, e_min) -> bool:
    return any(
        d == day and overlap_minutes(s_min, e_min, es, ee) > 0
        for d, es, ee in existing_assigns
    )


def assigned_minutes(assignments) -> int:
    """Union-of-windows duration in minutes, so overlapping lessons don't double count."""
    by_day: dict[str, list[tuple[int, int]]] = {}
    for day, start, end in assignments:
        by_day.setdefault(day, []).append((start, end))

    total = 0
    for windows in by_day.values():
        total += sum(end - start for start, end in merge_intervals(windows))
    return total


def projected_hours(assignments, lesson) -> float:
    return assigned_minutes(assignments + [lesson]) / 60.0


def schedule_penalty(assignments):
    """(active_days, total_blocks, total_gap_minutes) -- lower is better."""
    if not assignments:
        return 0, 0, 0
    by_day: dict[str, list[tuple[int, int]]] = {}
    for day, s, e in assignments:
        by_day.setdefault(day, []).append((s, e))

    total_blocks = 0
    total_gap_min = 0
    for slots in by_day.values():
        merged = merge_intervals(slots)
        total_blocks += len(merged)
        total_gap_min += sum(
            merged[index][0] - merged[index - 1][1]
            for index in range(1, len(merged))
        )

    return len(by_day), total_blocks, total_gap_min


def resolve_required_name(raw: str, pianist_names: list[str]) -> str | None:
    if not raw:
        return None
    raw_l = raw.strip().lower()
    for name in pianist_names:
        if name == raw.strip():
            return name
    for name in pianist_names:
        if name.lower() == raw_l:
            return name
    for name in pianist_names:
        if raw_l in name.lower():
            return name
    for name in pianist_names:
        if name.lower() in raw_l:
            return name
    return None


def assign_lessons(
    lessons: list[EngineLesson],
    pianists: list[EnginePianist],
    locked_ids: set[int] | None = None,
):
    """Runs the best-fit assignment algorithm in place on ``lessons``.

    Lessons whose id is in ``locked_ids`` (e.g. manually reassigned by a user)
    keep their current ``assigned_pianist_id`` untouched, but are still seeded
    into the pianists' committed schedules so the algorithm won't double-book
    them and so overlap/tie-break logic still sees them.

    Returns (hours_by_pianist_name: dict[str, float], conflicts: list[str]).
    """
    day_order = {d: i for i, d in enumerate(DAYS_ORDER)}
    locked_ids = locked_ids or set()

    locked = [l for l in lessons if l.id in locked_ids]
    to_assign = [lesson for lesson in lessons if lesson.need_pianist and lesson.id not in locked_ids]

    def sort_key(lesson: EngineLesson):
        required_first = 0 if lesson.required_pianist_name.strip() else 1
        return (required_first, day_order.get(lesson.day, 99), lesson.start_min)

    to_assign.sort(key=sort_key)

    pianist_names = [p.name for p in pianists]
    pianist_by_name = {p.name: p for p in pianists}
    pianist_by_id = {p.id: p for p in pianists}
    current_assigns: dict[int, list[tuple[str, int, int]]] = {p.id: [] for p in pianists}
    results_so_far: list[EngineLesson] = []

    # Seed manually-locked assignments so the algorithm respects them as
    # already-committed schedule time.
    for lesson in locked:
        if lesson.assigned_pianist_id is not None and lesson.assigned_pianist_id in current_assigns:
            current_assigns[lesson.assigned_pianist_id].append(
                (lesson.day, lesson.start_min, lesson.end_min)
            )
        results_so_far.append(lesson)

    for lesson in to_assign:
        day, start, end = lesson.day, lesson.start_min, lesson.end_min
        lesson.notes = ""
        lesson.fit_quality = ""
        lesson.assigned_pianist_id = None

        required_raw = lesson.required_pianist_name.strip()
        required = resolve_required_name(required_raw, pianist_names) if required_raw else None

        flags: list[str] = []
        assigned: EnginePianist | None = None
        fit_score = FIT_NONE
        has_tentative = False

        if required_raw:
            if required is None:
                flags.append(f"\u26a0 REQUIRED PIANIST '{required_raw}' NOT FOUND \u2014 lesson unassigned")
            else:
                p = pianist_by_name[required]
                fit_score, has_tentative = get_fit(day, start, end, p.avail)
                assigned = p
                if has_conflict(current_assigns[p.id], day, start, end):
                    flags.append(f"\u26a0 REQUIRED PIANIST '{required}' has a CONFLICTING lesson at this time")
                if fit_score == FIT_NONE:
                    flags.append(f"\u26a0 REQUIRED PIANIST '{required}' is UNAVAILABLE for this time")
                elif fit_score == FIT_NEAR:
                    flags.append(f"\u2139 NEAR FIT: '{required}' available within 15 min of lesson start")
                elif fit_score == FIT_PARTIAL:
                    flags.append(f"\u2139 PARTIAL FIT: '{required}' has Tentative availability for this slot")
                if has_tentative:
                    flags.append("\u2139 TENTATIVE availability")
                mh = p.max_hours
                if mh and projected_hours(current_assigns[p.id], (day, start, end)) > mh:
                    flags.append(f"\u26a0 OVER CAP: Exceeds {mh}h weekly limit")
                if required_raw.lower() != required.lower():
                    flags.append(f"\u2139 Required pianist '{required_raw}' matched to '{required}'")
        else:
            candidates = []
            for p in pianists:
                fit, tentative = get_fit(day, start, end, p.avail)
                conflict = has_conflict(current_assigns[p.id], day, start, end)
                hours_after = projected_hours(current_assigns[p.id], (day, start, end))
                mh = p.max_hours
                over_cap = bool(mh and hours_after > mh)
                workload = (hours_after / mh) if mh else hours_after
                new_assigns = current_assigns[p.id] + [(day, start, end)]
                sched_penalty = schedule_penalty(new_assigns)
                candidates.append({
                    "pianist": p, "fit": fit, "tentative": tentative, "conflict": conflict,
                    "over_cap": over_cap, "workload": workload, "schedule_penalty": sched_penalty,
                })

            non_conflicting = [c for c in candidates if not c["conflict"]]
            best_standard_fit = max((c["fit"] for c in non_conflicting), default=FIT_NONE)

            if best_standard_fit >= FIT_NEAR:
                under_cap = [c for c in non_conflicting if c["fit"] >= FIT_NEAR and not c["over_cap"]]
                eligible = under_cap or [c for c in non_conflicting if c["fit"] >= FIT_NEAR]
                best_eligible_fit = max(c["fit"] for c in eligible)
                pool = [c for c in eligible if c["fit"] == best_eligible_fit]
                pool.sort(key=lambda c: (c["over_cap"], c["tentative"], c["workload"], c["schedule_penalty"]))
                chosen = pool[0]
                assigned = chosen["pianist"]
                fit_score = chosen["fit"]
                has_tentative = chosen["tentative"]
            else:
                overlap_candidate = None
                for previous in results_so_far:
                    if previous.id in locked_ids or previous.required_pianist_name.strip():
                        continue
                    if previous.assigned_pianist_id is None or previous.day != day:
                        continue
                    overlap = overlap_minutes(start, end, previous.start_min, previous.end_min)
                    if overlap <= 0 or overlap > OVERLAP_MAX_MINUTES:
                        continue
                    pianist = pianist_by_id.get(previous.assigned_pianist_id)
                    if pianist is None:
                        continue
                    combined_start = min(start, previous.start_min)
                    combined_end = max(end, previous.end_min)
                    combined_fit = get_fit_for_window(
                        day, combined_start, combined_end, pianist.avail
                    )
                    if combined_fit < FIT_PARTIAL:
                        continue
                    projected = projected_hours(
                        current_assigns[pianist.id], (day, start, end)
                    )
                    over_cap = bool(pianist.max_hours and projected > pianist.max_hours)
                    overlap_candidate = (pianist, previous, over_cap)
                    break

                if overlap_candidate is not None:
                    assigned, previous, over_cap = overlap_candidate
                    fit_score = FIT_OVERLAP
                    flags.append(
                        f"\u2139 OVERLAP FIT: shares accompanist with lesson #{previous.id}"
                    )
                    previous.fit_quality = FIT_LABELS[FIT_OVERLAP]
                    previous_note = (
                        f"\u2139 OVERLAP FIT: shares accompanist with lesson #{lesson.id}"
                    )
                    if previous_note not in previous.notes:
                        previous.notes = f"{previous.notes} | {previous_note}".strip(" |")
                    if over_cap:
                        flags.append("\u26a0 OVER CAP: pianist exceeds weekly limit")
                elif not candidates:
                    flags.append("No pianists are available; lesson left unassigned.")
                elif all(candidate["conflict"] for candidate in candidates):
                    flags.append(
                        "All pianists are already assigned during this lesson; "
                        "no permitted overlap fit was found. Lesson left unassigned."
                    )
                else:
                    flags.append(
                        "No pianist has sufficient availability and no permitted "
                        "overlap fit was found. Lesson left unassigned."
                    )

        if assigned is not None:
            lesson.assigned_pianist_id = assigned.id
            current_assigns[assigned.id].append((day, start, end))
            lesson.fit_quality = FIT_LABELS[fit_score]
        else:
            lesson.fit_quality = "None"
            if not flags:
                flags = ["No valid pianist candidate; lesson left unassigned."]

        lesson.notes = " | ".join(f for f in flags if f)
        results_so_far.append(lesson)

    return recompute_hours_and_conflicts(lessons, pianists)


def recompute_hours_and_conflicts(lessons: list["EngineLesson"], pianists: list["EnginePianist"]):
    """Recomputes per-pianist union-hours and detects double-bookings.

    Safe to call after manual reassignment too (does not change assignments,
    only derived hours/notes/conflicts).
    """
    max_hours_map = {p.id: p.max_hours for p in pianists}
    name_by_id = {p.id: p.name for p in pianists}

    # Marginal covered-time per lesson so a pianist's row totals sum to their
    # union-based weekly total even when lessons overlap.
    prior_windows: dict[int, list[tuple[str, int, int]]] = {p.id: [] for p in pianists}
    by_pianist_lessons: dict[int, list[EngineLesson]] = {}
    for lesson in lessons:
        if lesson.assigned_pianist_id is None:
            lesson.hours = 0.0
            lesson.notes = " | ".join(
                note for note in lesson.notes.split(" | ")
                if note and "OVER CAP" not in note
            )
            continue
        by_pianist_lessons.setdefault(lesson.assigned_pianist_id, []).append(lesson)

    for pid, plessons in by_pianist_lessons.items():
        plessons.sort(key=lambda l: (day_index(l.day), l.start_min))
        for lesson in plessons:
            window = (lesson.day, lesson.start_min, lesson.end_min)
            before = assigned_minutes(prior_windows[pid])
            prior_windows[pid].append(window)
            after = assigned_minutes(prior_windows[pid])
            lesson.hours = round((after - before) / 60.0, 2)

    hours_by_pianist_id = {
        pid: round(sum(l.hours for l in by_pianist_lessons.get(pid, [])), 2)
        for pid in max_hours_map
    }

    # Conflict detection across ALL lessons currently assigned (covers manual edits too).
    conflicts: list[str] = []
    conflict_notes: dict[int, list[str]] = {}
    for pid, plessons in by_pianist_lessons.items():
        plessons_sorted = sorted(plessons, key=lambda l: (day_index(l.day), l.start_min))
        for i in range(len(plessons_sorted)):
            for j in range(i + 1, len(plessons_sorted)):
                a, b = plessons_sorted[i], plessons_sorted[j]
                if a.day != b.day:
                    continue
                overlap = overlap_minutes(a.start_min, a.end_min, b.start_min, b.end_min)
                allowed_overlap = (
                    overlap <= OVERLAP_MAX_MINUTES
                    and a.fit_quality == FIT_LABELS[FIT_OVERLAP]
                    and b.fit_quality == FIT_LABELS[FIT_OVERLAP]
                    and not a.required_pianist_name.strip()
                    and not b.required_pianist_name.strip()
                )
                if overlap > 0 and not allowed_overlap:
                    conflicts.append(
                        f"\u26a0 {name_by_id.get(pid, pid)} is double-booked on {a.day}: "
                        f"lessons #{a.id} and #{b.id} overlap"
                    )
                    conflict_notes.setdefault(a.id, []).append(
                        f"\u26a0 CONFLICT: overlaps lesson #{b.id} on {a.day}"
                    )
                    conflict_notes.setdefault(b.id, []).append(
                        f"\u26a0 CONFLICT: overlaps lesson #{a.id} on {b.day}"
                    )

    for lesson in lessons:
        base_notes = [
            note for note in lesson.notes.split(" | ")
            if note and "CONFLICT" not in note
        ]
        base_notes.extend(conflict_notes.get(lesson.id, []))
        lesson.notes = " | ".join(base_notes)

    # Update over-cap notes.
    for pid, plessons in by_pianist_lessons.items():
        mh = max_hours_map.get(pid)
        total = hours_by_pianist_id.get(pid, 0.0)
        for lesson in plessons:
            base_notes = [n for n in lesson.notes.split(" | ") if n and "OVER CAP" not in n]
            if mh and total > mh:
                base_notes.append(f"\u26a0 OVER CAP: {name_by_id.get(pid)} at {total}h / {mh}h cap")
            lesson.notes = " | ".join(base_notes)

    return (
        {name_by_id[pid]: hours for pid, hours in hours_by_pianist_id.items()},
        conflicts,
    )


def day_index(day: str) -> int:
    return DAYS_ORDER.index(day) if day in DAYS_ORDER else 99
