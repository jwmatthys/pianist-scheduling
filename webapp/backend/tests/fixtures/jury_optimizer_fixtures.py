"""Deterministic synthetic inputs for the future Jury optimizer contract."""

from dataclasses import dataclass
from datetime import date
from uuid import UUID, uuid5, NAMESPACE_URL

from app.jury_optimizer_contract import (
    ACCOMPANIST_RESULT_CONTRACT,
    AccompanistResultReference,
    FinalizedAccompanistAssignmentFact,
    FinalizedPianistFact,
    JuryAvailabilityWindow,
    JuryLessonEntry,
    JuryOptimizerInput,
    JuryPanelDate,
    JuryPanelTimingConfiguration,
    MealBreakConfiguration,
    PeriodicBreakConfiguration,
    ReadinessApproval,
)


FIXTURE_NAMESPACE = NAMESPACE_URL
FIXTURE_DATE = date(2027, 5, 1)


def fixture_uuid(name: str) -> UUID:
    return uuid5(FIXTURE_NAMESPACE, f"jury-optimizer-phase1:{name}")


@dataclass(frozen=True)
class JuryOptimizerFixture:
    inputs: JuryOptimizerInput
    lesson_uuids: tuple[UUID, ...]
    panel_uuids: tuple[UUID, ...]
    pianist_uuids: tuple[UUID, ...]


def make_fixture(scenario: str = "single-panel") -> JuryOptimizerFixture:
    session_uuid = fixture_uuid("session")
    result_uuid = fixture_uuid("accompanist-result")
    source_revision = 7
    jury_revision = 3
    accompanist_result = AccompanistResultReference(
        session_uuid=session_uuid,
        result_uuid=result_uuid,
        contract_id=ACCOMPANIST_RESULT_CONTRACT,
        contract_version=2,
        source_revision=source_revision,
        state="finalized",
    )
    readiness = ReadinessApproval(
        session_uuid=session_uuid,
        source_result_uuid=result_uuid,
        source_revision=source_revision,
        jury_input_revision=jury_revision,
        approved=True,
    )

    if scenario in {"meal-reset", "periodic-breaks"}:
        lesson_count = 5
    elif scenario == "meal-breaks":
        lesson_count = 3
    elif scenario == "fill-before-late-window":
        lesson_count = 3
    elif scenario == "day-bound":
        lesson_count = 2
    else:
        lesson_count = 2 if scenario in {"shared-pianist-conflicts", "shared-pianist-overbooked", "impossible-schedule", "multiple-panels", "multiple-dates", "duplicate-display-names"} else 1
    pianist_count = 2 if scenario == "duplicate-display-names" else 1
    panel_count = 2 if scenario in {"multiple-panels", "multiple-dates", "shared-pianist-overbooked", "fill-before-late-window"} else 1
    lesson_uuids = tuple(fixture_uuid(f"{scenario}:lesson:{index}") for index in range(lesson_count))
    panel_uuids = tuple(fixture_uuid(f"{scenario}:panel:{index}") for index in range(panel_count))
    pianist_uuids = tuple(fixture_uuid(f"{scenario}:pianist:{index}") for index in range(pianist_count))
    panel_dates = tuple(
        JuryPanelDate(panel_uuid, date(2027, 5, index + 1) if scenario == "multiple-dates" else FIXTURE_DATE)
        for index, panel_uuid in enumerate(panel_uuids)
    )

    if scenario == "preferred-start":
        earliest, preferred = 540, 600
    elif scenario == "unavailable-pianist":
        earliest, preferred = 600, 600
    elif scenario == "meal-breaks":
        earliest, preferred = 690, 690
    elif scenario == "meal-reset":
        earliest, preferred = 600, 600
    elif scenario == "day-bound":
        earliest, preferred = 1410, 1410
    else:
        earliest, preferred = 540, 540
    jury_length = 60 if scenario in {"impossible-schedule", "meal-breaks", "periodic-breaks", "fill-before-late-window", "shared-pianist-overbooked"} else 30
    panel_timings = []
    for panel_uuid in panel_uuids:
        meal = MealBreakConfiguration(720, 750) if scenario == "meal-breaks" else None
        if scenario == "meal-reset":
            meal = MealBreakConfiguration(660, 690)
        periodic = PeriodicBreakConfiguration(2, 10) if scenario in {"periodic-breaks", "meal-reset"} else None
        panel_timings.append(JuryPanelTimingConfiguration(
            panel_uuid=panel_uuid,
            earliest_start_minute=earliest,
            preferred_start_minute=preferred,
            jury_length_minutes=jury_length,
            meal_break=meal,
            periodic_break=periodic,
        ))

    facts = []
    jury_entries = []
    availability = []
    availability_keys = set()
    for index, lesson_uuid in enumerate(lesson_uuids):
        student_uuid = fixture_uuid(f"{scenario}:student:{index}")
        assigned_pianist = None
        pianist_required = scenario not in {"single-panel", "preferred-start", "day-bound"}
        if scenario in {"fill-before-late-window", "assigned-optional"} and index == 0:
            pianist_required = False
        if pianist_required or scenario == "assigned-optional":
            pianist_uuid = pianist_uuids[index % len(pianist_uuids)]
            pianist_name = "Same Synthetic Name" if scenario == "duplicate-display-names" else f"Synthetic Pianist {index}"
            assigned_pianist = FinalizedPianistFact(pianist_uuid, pianist_name)
        student_name = "Duplicate Synthetic Student" if scenario == "duplicate-display-names" else f"Synthetic Student {index}"
        facts.append(FinalizedAccompanistAssignmentFact(
            source_lesson_uuid=lesson_uuid,
            student_person_uuid=student_uuid,
            student_display_name=student_name,
            instrument="Voice",
            teacher="Synthetic Teacher",
            pianist_required=pianist_required,
            jury_required=True,
            assigned_pianist=assigned_pianist,
        ))
        if scenario == "fill-before-late-window":
            panel_uuid = panel_uuids[0] if index < 2 else panel_uuids[1]
        else:
            panel_uuid = panel_uuids[index % len(panel_uuids)]
        jury_entries.append(JuryLessonEntry(lesson_uuid, student_uuid, panel_uuid))
        if assigned_pianist is not None:
            panel_date = next(item.jury_date for item in panel_dates if item.panel_uuid == panel_uuid)
            availability_key = (assigned_pianist.person_uuid, panel_date)
            if availability_key in availability_keys:
                continue
            availability_keys.add(availability_key)
            if scenario == "unavailable-pianist":
                window_start, window_end = 540, 570
            elif scenario == "impossible-schedule":
                window_start, window_end = 540, 570
            elif scenario == "shared-pianist-overbooked":
                window_start, window_end = 540, 600
            elif scenario == "fill-before-late-window":
                window_start, window_end = 600, 660
            elif scenario == "meal-breaks":
                window_start, window_end = 690, 960
            elif scenario == "periodic-breaks":
                window_start, window_end = 540, 1200
            elif scenario == "meal-reset":
                window_start, window_end = 600, 960
            else:
                window_start, window_end = 540, 660
            availability.append(JuryAvailabilityWindow(
                pianist_person_uuid=assigned_pianist.person_uuid,
                jury_date=panel_date,
                start_minute=window_start,
                end_minute=window_end,
            ))

    if scenario == "stale-accompanist-result":
        readiness = ReadinessApproval(
            session_uuid=session_uuid,
            source_result_uuid=result_uuid,
            source_revision=source_revision + 1,
            jury_input_revision=jury_revision,
            approved=True,
        )

    inputs = JuryOptimizerInput(
        accompanist_result=accompanist_result,
        readiness=readiness,
        accompanist_assignments=tuple(facts),
        jury_lesson_entries=tuple(jury_entries),
        panel_dates=panel_dates,
        panel_timings=tuple(panel_timings),
        pianist_availability_windows=tuple(availability),
    )
    return JuryOptimizerFixture(inputs, lesson_uuids, panel_uuids, pianist_uuids)


def make_stale_accompanist_fixture() -> JuryOptimizerFixture:
    return make_fixture("stale-accompanist-result")