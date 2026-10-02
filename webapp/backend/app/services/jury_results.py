"""Readiness-gated Jury schedule generation and module-result lifecycle."""

import hashlib
import json
from datetime import date, datetime
from uuid import UUID, uuid4

from sqlalchemy import func
from sqlalchemy.orm import Session

from .. import module_models
from ..jury_optimizer import DeterministicJuryOptimizer
from ..jury_optimizer_contract import (
    ACCOMPANIST_RESULT_CONTRACT,
    AccompanistResultReference,
    FinalizedAccompanistAssignmentFact,
    FinalizedPianistFact,
    BreakEntry,
    JuryAvailabilityWindow,
    JuryLessonEntry,
    JuryOptimizerInput,
    JuryPanelDate,
    JuryPanelTimingConfiguration,
    MealBreakConfiguration,
    MealBreakEntry,
    PeriodicBreakConfiguration,
    ReadinessApproval,
    ScheduledLesson,
)
from ..jury_schemas import (
    JuryConflictExplanationOut,
    JuryMealBreakEvent,
    JuryOptimizerDiagnosticOut,
    JuryPanelTimelineOut,
    JuryPeriodicBreakEvent,
    JuryReadinessOut,
    JuryResultPanelSnapshot,
    JuryScheduleEventOut,
    JuryScheduleHistoryItemOut,
    JuryScheduleViewOut,
    JuryScheduledLessonEvent,
    JuryScheduleWarningOut,
    JuryStoredSchedulePayload,
    JuryUnscheduledLessonOut,
    ReadinessIssue,
)
from . import accompanist_results, jury
from .module_lifecycle import (
    ACCOMPANIST_MODULE_ID,
    active_session_uuid,
)


JURY_MODULE_ID = "juries"
JURY_RESULT_CONTRACT = "jury.schedule-result"
JURY_RESULT_CONTRACT_VERSION = 1
JURY_PAYLOAD_SCHEMA_VERSION = 1
ACCOMPANIST_DEPENDENCY_KEY = ACCOMPANIST_RESULT_CONTRACT


class JuryGenerationError(ValueError):
    def __init__(
        self,
        code: str,
        message: str,
        readiness: JuryReadinessOut | None = None,
    ):
        self.code = code
        self.readiness = readiness
        super().__init__(message)


def _jury_module_revision(db: Session, session_uuid: str) -> module_models.ModuleRevision:
    revision = db.get(module_models.ModuleRevision, (session_uuid, JURY_MODULE_ID))
    if revision is None:
        configuration = db.get(module_models.JuryConfiguration, session_uuid)
        if configuration is None:
            raise JuryGenerationError("JURY_CONFIGURATION_MISSING", "Jury configuration is unavailable.")
        revision = module_models.ModuleRevision(
            session_uuid=session_uuid,
            module_id=JURY_MODULE_ID,
            source_revision=configuration.input_revision,
            current_result_uuid=None,
            modified_at=datetime.utcnow(),
        )
        db.add(revision)
        db.flush()
    return revision


def _optimizer_input(db: Session, result, readiness_report: JuryReadinessOut) -> JuryOptimizerInput:
    session_uuid = active_session_uuid(db)
    facts = tuple(
        FinalizedAccompanistAssignmentFact(
            source_lesson_uuid=entry.source_lesson_uuid,
            student_person_uuid=entry.student_person_uuid,
            student_display_name=entry.student_display_name,
            instrument=entry.instrument,
            teacher=entry.teacher,
            pianist_required=entry.pianist_required,
            jury_required=entry.jury_required,
            assigned_pianist=(
                FinalizedPianistFact(
                    person_uuid=entry.assigned_pianist.person_uuid,
                    display_name=entry.assigned_pianist.display_name,
                )
                if entry.assigned_pianist else None
            ),
        )
        for entry in result.payload.entries
    )
    fact_by_lesson = {str(fact.source_lesson_uuid): fact for fact in facts}
    required_lesson_uuids = {
        lesson_uuid for lesson_uuid, fact in fact_by_lesson.items() if fact.jury_required
    }
    jury_rows = db.query(module_models.JuryLessonEntry).filter(
        module_models.JuryLessonEntry.session_uuid == session_uuid,
        module_models.JuryLessonEntry.source_lesson_uuid.in_(required_lesson_uuids),
    ).order_by(module_models.JuryLessonEntry.source_lesson_uuid).all() if required_lesson_uuids else []
    jury_entries = tuple(
        JuryLessonEntry(
            source_lesson_uuid=UUID(row.source_lesson_uuid),
            student_person_uuid=UUID(row.student_person_uuid),
            panel_uuid=UUID(row.panel_uuid) if row.panel_uuid else None,
        )
        for row in jury_rows
    )
    used_panel_uuids = {str(entry.panel_uuid) for entry in jury_entries if entry.panel_uuid}
    panel_rows = db.query(module_models.JuryPanel).filter(
        module_models.JuryPanel.session_uuid == session_uuid,
        module_models.JuryPanel.panel_uuid.in_(used_panel_uuids),
    ).order_by(module_models.JuryPanel.panel_uuid).all() if used_panel_uuids else []
    panels_by_uuid = {row.panel_uuid: row for row in panel_rows}
    panel_date_rows = db.query(module_models.JuryPanelDate).filter(
        module_models.JuryPanelDate.session_uuid == session_uuid,
        module_models.JuryPanelDate.panel_uuid.in_(used_panel_uuids),
    ).all() if used_panel_uuids else []
    dates_by_panel = {row.panel_uuid: row.jury_date for row in panel_date_rows}

    panel_dates = []
    panel_timings = []
    for panel_uuid in sorted(used_panel_uuids):
        panel = panels_by_uuid[panel_uuid]
        jury_date = dates_by_panel[panel_uuid]
        panel_dates.append(JuryPanelDate(UUID(panel_uuid), jury_date))
        panel_timings.append(JuryPanelTimingConfiguration(
            panel_uuid=UUID(panel_uuid),
            earliest_start_minute=panel.earliest_start_minute,
            preferred_start_minute=panel.preferred_start_minute,
            jury_length_minutes=panel.jury_length_minutes,
            meal_break=(
                MealBreakConfiguration(panel.meal_start_minute, panel.meal_end_minute)
                if panel.meal_break else None
            ),
            periodic_break=(
                PeriodicBreakConfiguration(panel.break_every_x_juries, panel.break_length_minutes)
                if panel.break_needed else None
            ),
        ))

    pianist_uuids = {
        str(fact.assigned_pianist.person_uuid)
        for fact in facts
        if fact.jury_required and fact.assigned_pianist is not None
    }
    jury_dates = set(dates_by_panel.values())
    declarations = db.query(module_models.JuryPianistAvailabilityDeclaration).filter(
        module_models.JuryPianistAvailabilityDeclaration.session_uuid == session_uuid,
        module_models.JuryPianistAvailabilityDeclaration.pianist_person_uuid.in_(pianist_uuids),
        module_models.JuryPianistAvailabilityDeclaration.jury_date.in_(jury_dates),
        module_models.JuryPianistAvailabilityDeclaration.is_complete.is_(True),
    ).all() if pianist_uuids and jury_dates else []
    availability = []
    for declaration in declarations:
        windows = db.query(module_models.JuryPianistAvailableWindow).filter_by(
            session_uuid=session_uuid,
            pianist_person_uuid=declaration.pianist_person_uuid,
            jury_date=declaration.jury_date,
        ).order_by(module_models.JuryPianistAvailableWindow.start_minute).all()
        availability.extend(JuryAvailabilityWindow(
            pianist_person_uuid=UUID(declaration.pianist_person_uuid),
            jury_date=declaration.jury_date,
            start_minute=window.start_minute,
            end_minute=window.end_minute,
        ) for window in windows)

    return JuryOptimizerInput(
        accompanist_result=AccompanistResultReference(
            session_uuid=result.session_uuid,
            result_uuid=result.result_uuid,
            contract_id=result.contract_id,
            contract_version=result.contract_version,
            source_revision=result.source_revision,
            state=result.state,
        ),
        readiness=ReadinessApproval(
            session_uuid=result.session_uuid,
            source_result_uuid=readiness_report.source_result_uuid,
            source_revision=readiness_report.source_revision,
            jury_input_revision=readiness_report.jury_input_revision,
            approved=readiness_report.ready,
        ),
        accompanist_assignments=facts,
        jury_lesson_entries=jury_entries,
        panel_dates=tuple(panel_dates),
        panel_timings=tuple(panel_timings),
        pianist_availability_windows=tuple(availability),
    )


def _stored_payload(db: Session, inputs: JuryOptimizerInput, result, readiness_report: JuryReadinessOut) -> JuryStoredSchedulePayload:
    facts = {fact.source_lesson_uuid: fact for fact in inputs.accompanist_assignments}
    entries = {entry.source_lesson_uuid: entry for entry in inputs.jury_lesson_entries}
    date_by_panel = {item.panel_uuid: item.jury_date for item in inputs.panel_dates}
    panels = {
        row.panel_uuid: row
        for row in db.query(module_models.JuryPanel).filter(
            module_models.JuryPanel.panel_uuid.in_([str(item.panel_uuid) for item in inputs.panel_timings])
        ).all()
    }
    events: list[JuryScheduleEventOut] = []
    for entry in result.schedule_entries:
        if isinstance(entry, ScheduledLesson):
            fact = facts[entry.source_lesson_uuid]
            events.append(JuryScheduledLessonEvent(
                source_lesson_uuid=entry.source_lesson_uuid,
                student_person_uuid=fact.student_person_uuid,
                student_display_name=fact.student_display_name,
                instrument=fact.instrument,
                panel_uuid=entry.panel_uuid,
                jury_date=entry.jury_date,
                pianist_person_uuid=entry.pianist_person_uuid,
                pianist_display_name=fact.assigned_pianist.display_name if fact.assigned_pianist else None,
                start_minute=entry.start_minute,
                end_minute=entry.end_minute,
            ))
        elif isinstance(entry, MealBreakEntry):
            events.append(JuryMealBreakEvent(
                panel_uuid=entry.panel_uuid,
                jury_date=entry.jury_date,
                start_minute=entry.start_minute,
                end_minute=entry.end_minute,
            ))
        elif isinstance(entry, BreakEntry):
            events.append(JuryPeriodicBreakEvent(
                panel_uuid=entry.panel_uuid,
                jury_date=entry.jury_date,
                start_minute=entry.start_minute,
                end_minute=entry.end_minute,
                after_jury_count=entry.after_jury_count,
            ))
        else:
            raise JuryGenerationError("UNSUPPORTED_SCHEDULE_ENTRY", "Optimizer returned an unsupported schedule event.")

    unscheduled = []
    for entry in result.unscheduled_lessons:
        fact = facts[entry.source_lesson_uuid]
        jury_entry = entries[entry.source_lesson_uuid]
        panel_uuid = jury_entry.panel_uuid
        unscheduled.append(JuryUnscheduledLessonOut(
            source_lesson_uuid=entry.source_lesson_uuid,
            student_person_uuid=fact.student_person_uuid,
            student_display_name=fact.student_display_name,
            instrument=fact.instrument,
            panel_uuid=panel_uuid,
            jury_date=date_by_panel[panel_uuid],
            reason_code=entry.reason_code.value,
            explanation=entry.explanation,
        ))

    return JuryStoredSchedulePayload(
        session_uuid=result.session_uuid,
        source_result_uuid=result.source_result_uuid,
        source_contract_id=result.source_contract_id,
        source_contract_version=result.source_contract_version,
        source_revision=result.source_revision,
        jury_input_revision=result.jury_input_revision,
        required_lesson_uuids=list(result.required_lesson_uuids),
        events=events,
        unscheduled_lessons=unscheduled,
        diagnostics=[JuryOptimizerDiagnosticOut(
            code=item.code,
            severity=item.severity.value,
            message=item.message,
        ) for item in result.diagnostics],
        conflict_explanations=[JuryConflictExplanationOut(
            code=item.code,
            message=item.message,
            source_lesson_uuids=list(item.source_lesson_uuids),
            person_uuid=item.person_uuid,
            panel_uuid=item.panel_uuid,
        ) for item in result.conflict_explanations],
        panel_snapshots=[JuryResultPanelSnapshot(
            panel_uuid=timing.panel_uuid,
            panel_name=panels[str(timing.panel_uuid)].panel_name,
            jury_date=date_by_panel[timing.panel_uuid],
            earliest_start_minute=timing.earliest_start_minute,
            preferred_start_minute=timing.preferred_start_minute,
            jury_length_minutes=timing.jury_length_minutes,
            break_every_x_juries=timing.periodic_break.every_x_juries if timing.periodic_break else None,
            break_length_minutes=timing.periodic_break.length_minutes if timing.periodic_break else None,
            meal_start_minute=timing.meal_break.start_minute if timing.meal_break else None,
            meal_end_minute=timing.meal_break.end_minute if timing.meal_break else None,
        ) for timing in inputs.panel_timings],
        readiness_warnings=[issue for issue in readiness_report.issues if issue.severity == "warning"],
    )


def _payload_json(payload: JuryStoredSchedulePayload) -> str:
    return json.dumps(payload.model_dump(mode="json"), ensure_ascii=True, separators=(",", ":"), sort_keys=True)


def _payload_for(row: module_models.ModuleResult) -> JuryStoredSchedulePayload:
    if row.payload_schema_version != JURY_PAYLOAD_SCHEMA_VERSION:
        raise JuryGenerationError("UNSUPPORTED_JURY_RESULT", "This Jury schedule result version is not supported.")
    return JuryStoredSchedulePayload.model_validate_json(row.payload_json)


def _stale_reasons(db: Session, row: module_models.ModuleResult, payload: JuryStoredSchedulePayload) -> list[str]:
    reasons = []
    accompanist_revision = db.get(module_models.ModuleRevision, (row.session_uuid, ACCOMPANIST_MODULE_ID))
    dependency = db.get(
        module_models.ModuleResultDependency,
        (row.result_uuid, ACCOMPANIST_DEPENDENCY_KEY),
    )
    if dependency is None:
        reasons.append("accompanist_result_changed")
    else:
        if accompanist_revision is None or accompanist_revision.source_revision != dependency.source_revision:
            reasons.append("accompanist_source_revision_changed")
        if accompanist_revision is None or accompanist_revision.current_result_uuid != dependency.source_result_uuid:
            reasons.append("accompanist_result_changed")
    configuration = db.get(module_models.JuryConfiguration, row.session_uuid)
    if configuration is None or configuration.input_revision != payload.jury_input_revision:
        reasons.append("jury_inputs_changed")
    return reasons


def _view(db: Session, row: module_models.ModuleResult) -> JuryScheduleViewOut:
    payload = _payload_for(row)
    stale_reasons = _stale_reasons(db, row, payload)
    events_by_panel: dict[UUID, list[JuryScheduleEventOut]] = {}
    for event in payload.events:
        events_by_panel.setdefault(event.panel_uuid, []).append(event)
    timelines = [JuryPanelTimelineOut(
        panel_uuid=snapshot.panel_uuid,
        panel_name=snapshot.panel_name,
        jury_date=snapshot.jury_date,
        events=events_by_panel.get(snapshot.panel_uuid, []),
    ) for snapshot in payload.panel_snapshots]
    warnings = [JuryScheduleWarningOut(
        code=issue.code,
        severity="warning",
        message=issue.message,
    ) for issue in payload.readiness_warnings]
    warnings.extend(JuryScheduleWarningOut(
        code=item.code,
        severity="warning" if item.severity == "info" else item.severity,
        message=item.message,
    ) for item in payload.diagnostics if item.severity != "info")
    if stale_reasons:
        warnings.append(JuryScheduleWarningOut(
            code="JURY_RESULT_STALE",
            severity="warning",
            message="This saved Jury schedule was generated from older Accompanist or Jury inputs.",
        ))
    return JuryScheduleViewOut(
        result_uuid=UUID(row.result_uuid),
        session_uuid=UUID(row.session_uuid),
        contract_id=row.contract_id,
        contract_version=row.contract_version,
        result_version=row.result_version,
        state=row.state,
        created_at=row.created_at,
        source_result_uuid=payload.source_result_uuid,
        source_contract_id=payload.source_contract_id,
        source_contract_version=payload.source_contract_version,
        source_revision=payload.source_revision,
        jury_input_revision=payload.jury_input_revision,
        stale=bool(stale_reasons),
        stale_reasons=stale_reasons,
        panel_timelines=timelines,
        unscheduled_lessons=payload.unscheduled_lessons,
        warnings=warnings,
    )


def _history_item(db: Session, row: module_models.ModuleResult) -> JuryScheduleHistoryItemOut:
    payload = _payload_for(row)
    return JuryScheduleHistoryItemOut(
        result_uuid=UUID(row.result_uuid),
        session_uuid=UUID(row.session_uuid),
        result_version=row.result_version,
        state=row.state,
        created_at=row.created_at,
        source_result_uuid=payload.source_result_uuid,
        source_revision=payload.source_revision,
        jury_input_revision=payload.jury_input_revision,
        stale=bool(_stale_reasons(db, row, payload)),
        stale_reasons=_stale_reasons(db, row, payload),
        scheduled_count=sum(event.kind == "jury" for event in payload.events),
        unscheduled_count=len(payload.unscheduled_lessons),
    )


def generate_schedule(
    db: Session,
    *,
    expected_jury_input_revision: int | None = None,
    optimizer=None,
) -> JuryScheduleViewOut:
    session_uuid = active_session_uuid(db)
    configuration = db.get(module_models.JuryConfiguration, session_uuid)
    if configuration is None:
        raise JuryGenerationError("JURY_CONFIGURATION_MISSING", "Jury configuration is unavailable.")
    if expected_jury_input_revision is not None and configuration.input_revision != expected_jury_input_revision:
        raise JuryGenerationError("JURY_INPUT_REVISION_CHANGED", "Jury inputs changed; refresh readiness and try again.")

    current_source = accompanist_results.get_current_accompanist_result(db)
    readiness_report = jury.readiness(db)
    if current_source is None or not readiness_report.ready:
        raise JuryGenerationError(
            "JURY_NOT_READY",
            "Resolve Jury readiness blockers before generating a schedule.",
            readiness_report,
        )

    try:
        optimizer_input = _optimizer_input(db, current_source, readiness_report)
    except (KeyError, TypeError, ValueError) as error:
        raise JuryGenerationError("INVALID_OPTIMIZER_INPUT", "Validated Jury inputs could not be projected to the optimizer contract.", readiness_report) from error

    optimizer = optimizer or DeterministicJuryOptimizer()
    optimized = optimizer.optimize(optimizer_input)
    if (
        optimized.session_uuid != optimizer_input.accompanist_result.session_uuid
        or optimized.source_result_uuid != optimizer_input.accompanist_result.result_uuid
        or optimized.source_revision != optimizer_input.accompanist_result.source_revision
        or optimized.jury_input_revision != optimizer_input.readiness.jury_input_revision
    ):
        raise JuryGenerationError("OPTIMIZER_PROVENANCE_MISMATCH", "Optimizer output provenance does not match its input snapshot.")

    latest_configuration = db.get(module_models.JuryConfiguration, session_uuid)
    latest_source = accompanist_results.get_current_accompanist_result(db)
    if (
        latest_configuration is None
        or latest_configuration.input_revision != optimizer_input.readiness.jury_input_revision
        or latest_source is None
        or latest_source.result_uuid != current_source.result_uuid
        or latest_source.source_revision != current_source.source_revision
    ):
        raise JuryGenerationError("INPUT_REVISION_CHANGED", "Scheduling inputs changed during generation; no result was stored.")

    revision = _jury_module_revision(db, session_uuid)
    if revision.source_revision != configuration.input_revision:
        revision.source_revision = configuration.input_revision
    if revision.current_result_uuid:
        previous = db.get(module_models.ModuleResult, revision.current_result_uuid)
        if previous is not None and previous.state == "draft":
            previous.state = "superseded"

    payload = _stored_payload(db, optimizer_input, optimized, readiness_report)
    serialized_payload = _payload_json(payload)
    latest_version = db.query(func.max(module_models.ModuleResult.result_version)).filter(
        module_models.ModuleResult.session_uuid == session_uuid,
        module_models.ModuleResult.module_id == JURY_MODULE_ID,
    ).scalar()
    result_uuid = str(uuid4())
    now = datetime.utcnow()
    row = module_models.ModuleResult(
        result_uuid=result_uuid,
        session_uuid=session_uuid,
        module_id=JURY_MODULE_ID,
        contract_id=JURY_RESULT_CONTRACT,
        contract_version=JURY_RESULT_CONTRACT_VERSION,
        payload_schema_version=JURY_PAYLOAD_SCHEMA_VERSION,
        result_version=(latest_version or 0) + 1,
        source_revision=current_source.source_revision,
        state="draft",
        payload_json=serialized_payload,
        payload_sha256=hashlib.sha256(serialized_payload.encode("utf-8")).hexdigest(),
        created_at=now,
        finalized_at=None,
    )
    db.add(row)
    db.flush()
    db.add(module_models.ModuleResultDependency(
        dependent_result_uuid=result_uuid,
        dependency_key=ACCOMPANIST_DEPENDENCY_KEY,
        source_result_uuid=str(current_source.result_uuid),
        source_session_uuid=str(current_source.session_uuid),
        source_contract_id=current_source.contract_id,
        source_contract_version=current_source.contract_version,
        source_revision=current_source.source_revision,
    ))
    revision.current_result_uuid = result_uuid
    revision.source_revision = configuration.input_revision
    revision.modified_at = now
    db.flush()
    return _view(db, row)


def get_current_schedule(db: Session) -> JuryScheduleViewOut | None:
    session_uuid = active_session_uuid(db)
    revision = db.get(module_models.ModuleRevision, (session_uuid, JURY_MODULE_ID))
    if revision is None or revision.current_result_uuid is None:
        return None
    row = db.get(module_models.ModuleResult, revision.current_result_uuid)
    if row is None or row.module_id != JURY_MODULE_ID or row.contract_id != JURY_RESULT_CONTRACT:
        return None
    return _view(db, row)


def get_schedule_by_uuid(db: Session, result_uuid: str) -> JuryScheduleViewOut | None:
    row = db.get(module_models.ModuleResult, result_uuid)
    if row is None or row.module_id != JURY_MODULE_ID or row.contract_id != JURY_RESULT_CONTRACT:
        return None
    return _view(db, row)


def list_schedule_history(db: Session) -> list[JuryScheduleHistoryItemOut]:
    session_uuid = active_session_uuid(db)
    rows = db.query(module_models.ModuleResult).filter(
        module_models.ModuleResult.session_uuid == session_uuid,
        module_models.ModuleResult.module_id == JURY_MODULE_ID,
        module_models.ModuleResult.contract_id == JURY_RESULT_CONTRACT,
        module_models.ModuleResult.contract_version == JURY_RESULT_CONTRACT_VERSION,
    ).order_by(module_models.ModuleResult.result_version.desc()).all()
    return [_history_item(db, row) for row in rows]