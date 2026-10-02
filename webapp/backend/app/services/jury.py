"""Jury-owned persistence operations and pre-optimization readiness checks."""

from datetime import date, datetime
from uuid import UUID

from sqlalchemy.orm import Session

from .. import models, module_models
from ..jury_schemas import (
    AccompanistAssignmentPayload,
    AvailableWindowIn,
    JuryAvailabilityIn,
    JuryAvailabilityOut,
    JuryConfigurationOut,
    JuryLessonEntryOut,
    JuryPanelFields,
    JuryPanelOut,
    JuryReadinessOut,
    ReadinessIssue,
    ResultEnvelope,
)
from . import accompanist_results
from .module_lifecycle import (
    active_session_uuid,
    bump_accompanist_revision,
    bump_jury_revision,
    get_lesson_jury_required,
    module_revision,
    set_lesson_jury_required,
)


class JuryDataError(ValueError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


def get_configuration(db: Session) -> JuryConfigurationOut:
    session_uuid = active_session_uuid(db)
    row = db.get(module_models.JuryConfiguration, session_uuid)
    if row is None:
        raise JuryDataError("JURY_CONFIGURATION_MISSING", "Jury configuration is unavailable.")
    return JuryConfigurationOut.model_validate(row, from_attributes=True)


def _panel_out(db: Session, row: module_models.JuryPanel) -> JuryPanelOut:
    panel_date = db.get(module_models.JuryPanelDate, row.panel_uuid)
    return JuryPanelOut(
        panel_uuid=row.panel_uuid,
        session_uuid=row.session_uuid,
        panel_name=row.panel_name,
        room=row.room,
        jury_date=panel_date.jury_date if panel_date else None,
        earliest_start_minute=row.earliest_start_minute,
        preferred_start_minute=row.preferred_start_minute,
        jury_length_minutes=row.jury_length_minutes,
        break_needed=row.break_needed,
        break_every_x_juries=row.break_every_x_juries,
        break_length_minutes=row.break_length_minutes,
        meal_break=row.meal_break,
        meal_start_minute=row.meal_start_minute,
        meal_end_minute=row.meal_end_minute,
    )


def list_panels(db: Session) -> list[JuryPanelOut]:
    session_uuid = active_session_uuid(db)
    rows = db.query(module_models.JuryPanel).filter(
        module_models.JuryPanel.session_uuid == session_uuid
    ).order_by(module_models.JuryPanel.panel_name, module_models.JuryPanel.panel_uuid).all()
    return [_panel_out(db, row) for row in rows]


def create_panel(db: Session, fields: JuryPanelFields) -> JuryPanelOut:
    session_uuid = active_session_uuid(db)
    values = fields.model_dump()
    jury_date = values.pop("jury_date")
    row = module_models.JuryPanel(
        session_uuid=session_uuid,
        **values,
    )
    db.add(row)
    bump_jury_revision(db)
    db.flush()
    db.add(module_models.JuryPanelDate(
        panel_uuid=row.panel_uuid,
        session_uuid=session_uuid,
        jury_date=jury_date,
    ))
    db.flush()
    return _panel_out(db, row)


def update_panel(db: Session, panel_uuid: str, fields: JuryPanelFields) -> JuryPanelOut:
    session_uuid = active_session_uuid(db)
    row = db.get(module_models.JuryPanel, panel_uuid)
    if row is None or row.session_uuid != session_uuid:
        raise JuryDataError("PANEL_NOT_FOUND", "Jury Panel not found.")
    values = fields.model_dump()
    jury_date = values.pop("jury_date")
    panel_date = db.get(module_models.JuryPanelDate, panel_uuid)
    date_changed = panel_date is None or panel_date.jury_date != jury_date
    if any(getattr(row, key) != value for key, value in values.items()) or date_changed:
        for key, value in values.items():
            setattr(row, key, value)
        if panel_date is None:
            db.add(module_models.JuryPanelDate(
                panel_uuid=panel_uuid,
                session_uuid=session_uuid,
                jury_date=jury_date,
            ))
        else:
            panel_date.jury_date = jury_date
        bump_jury_revision(db)
    return _panel_out(db, row)


def delete_panel(db: Session, panel_uuid: str) -> None:
    session_uuid = active_session_uuid(db)
    row = db.get(module_models.JuryPanel, panel_uuid)
    if row is None or row.session_uuid != session_uuid:
        raise JuryDataError("PANEL_NOT_FOUND", "Jury Panel not found.")
    entries = db.query(module_models.JuryLessonEntry).filter(
        module_models.JuryLessonEntry.session_uuid == session_uuid,
        module_models.JuryLessonEntry.panel_uuid == panel_uuid,
    ).all()
    for entry in entries:
        entry.panel_uuid = None
    panel_date = db.get(module_models.JuryPanelDate, panel_uuid)
    if panel_date is not None:
        db.delete(panel_date)
    db.delete(row)
    bump_jury_revision(db)


def sync_roster_from_current_result(db: Session) -> list[JuryLessonEntryOut]:
    result = accompanist_results.get_current_accompanist_result(db)
    if result is None:
        raise JuryDataError(
            "NO_CURRENT_FINALIZED_RESULT",
            "Finalize the current Accompanist schedule before synchronizing Jury entries.",
        )
    session_uuid = active_session_uuid(db)
    configuration = db.get(module_models.JuryConfiguration, session_uuid)
    if configuration is None:
        raise JuryDataError("JURY_CONFIGURATION_MISSING", "Jury configuration is unavailable.")

    changed = False
    result_entries = result.payload.entries
    current_by_lesson_uuid = {str(item.source_lesson_uuid): item for item in result_entries}
    existing_rows = db.query(module_models.JuryLessonEntry).filter_by(
        session_uuid=session_uuid
    ).all()
    for entry in existing_rows:
        if entry.source_lesson_uuid not in current_by_lesson_uuid:
            db.delete(entry)
            changed = True
    for lesson_uuid, source in current_by_lesson_uuid.items():
        entry = db.get(module_models.JuryLessonEntry, (session_uuid, lesson_uuid))
        if entry is None:
            db.add(module_models.JuryLessonEntry(
                session_uuid=session_uuid,
                source_lesson_uuid=lesson_uuid,
                student_person_uuid=str(source.student_person_uuid),
                panel_uuid=None,
            ))
            changed = True
        elif entry.student_person_uuid != str(source.student_person_uuid):
            entry.student_person_uuid = str(source.student_person_uuid)
            changed = True
    if configuration.roster_source_result_uuid != str(result.result_uuid):
        configuration.roster_source_result_uuid = str(result.result_uuid)
        changed = True
    if changed:
        bump_jury_revision(db)
        db.flush()
    return list_entries(db, result)


def _entry_out(
    session_uuid: str,
    row: module_models.JuryLessonEntry,
    source,
    source_flags: tuple[bool, bool] | None = None,
) -> JuryLessonEntryOut:
    jury_required, pianist_required = source_flags or (source.jury_required, source.pianist_required)
    return JuryLessonEntryOut(
        session_uuid=session_uuid,
        source_lesson_uuid=source.source_lesson_uuid,
        student_person_uuid=source.student_person_uuid,
        student_display_name=source.student_display_name,
        instrument=source.instrument,
        teacher=source.teacher,
        pianist_required=pianist_required,
        assigned_pianist=source.assigned_pianist,
        jury_required=jury_required,
        panel_uuid=row.panel_uuid,
    )


def list_entries(
    db: Session,
    result: ResultEnvelope | None = None,
) -> list[JuryLessonEntryOut]:
    current_result = accompanist_results.get_current_accompanist_result(db)
    if result is None:
        result = current_result
    if (
        result is None
        or result.state != "finalized"
        or current_result is None
        or result.result_uuid != current_result.result_uuid
    ):
        return []
    session_uuid = active_session_uuid(db)
    source_by_uuid = {str(entry.source_lesson_uuid): entry for entry in result.payload.entries}
    rows = db.query(module_models.JuryLessonEntry).filter(
        module_models.JuryLessonEntry.session_uuid == session_uuid
    ).order_by(module_models.JuryLessonEntry.source_lesson_uuid).all()
    return [
        _entry_out(
            session_uuid,
            row,
            source_by_uuid[str(row.source_lesson_uuid)],
        )
        for row in rows
        if str(row.source_lesson_uuid) in source_by_uuid
    ]


def update_lesson_entry(
    db: Session,
    source_lesson_uuid: str,
    *,
    panel_uuid: str | None,
) -> JuryLessonEntryOut:
    result = accompanist_results.get_current_accompanist_result(db)
    if result is None:
        raise JuryDataError("NO_CURRENT_FINALIZED_RESULT", "No current finalized Accompanist result is available.")
    source_by_uuid = {str(item.source_lesson_uuid): item for item in result.payload.entries}
    source = source_by_uuid.get(source_lesson_uuid)
    if source is None:
        raise JuryDataError("SOURCE_LESSON_NOT_FOUND", "Lesson is not present in the current finalized Accompanist result.")
    if not source.jury_required and panel_uuid is not None:
        raise JuryDataError(
            "JURY_NOT_REQUIRED",
            "A Panel can only be assigned to a lesson marked Jury Required in the finalized Accompanist result.",
        )

    session_uuid = active_session_uuid(db)
    if panel_uuid is not None:
        panel = db.get(module_models.JuryPanel, panel_uuid)
        if panel is None or panel.session_uuid != session_uuid:
            raise JuryDataError("PANEL_NOT_FOUND", "Selected Jury Panel does not exist in this session.")

    row = db.get(module_models.JuryLessonEntry, (session_uuid, source_lesson_uuid))
    if row is None:
        row = module_models.JuryLessonEntry(
            session_uuid=session_uuid,
            source_lesson_uuid=source_lesson_uuid,
            student_person_uuid=str(source.student_person_uuid),
            panel_uuid=panel_uuid,
        )
        db.add(row)
        changed = True
    else:
        changed = (
            row.student_person_uuid != str(source.student_person_uuid)
            or row.panel_uuid != panel_uuid
        )
        row.student_person_uuid = str(source.student_person_uuid)
        row.panel_uuid = panel_uuid
    if changed:
        bump_jury_revision(db)
        db.flush()
    return _entry_out(session_uuid, row, source)


def assign_unassigned_panels_by_instrument(
    db: Session,
) -> tuple[list[JuryLessonEntryOut], int]:
    result = accompanist_results.get_current_accompanist_result(db)
    if result is None:
        raise JuryDataError("NO_CURRENT_FINALIZED_RESULT", "No current finalized Accompanist result is available.")

    session_uuid = active_session_uuid(db)
    panels = db.query(module_models.JuryPanel).filter_by(
        session_uuid=session_uuid
    ).order_by(module_models.JuryPanel.panel_uuid).all()
    panels_by_name: dict[str, list[module_models.JuryPanel]] = {}
    for panel in panels:
        panels_by_name.setdefault(panel.panel_name.casefold(), []).append(panel)

    source_by_lesson = {
        str(source.source_lesson_uuid): source
        for source in result.payload.entries
    }
    entries = db.query(module_models.JuryLessonEntry).filter_by(
        session_uuid=session_uuid
    ).order_by(module_models.JuryLessonEntry.source_lesson_uuid).all()
    assigned_count = 0
    for entry in entries:
        if entry.panel_uuid is not None:
            continue
        source = source_by_lesson.get(entry.source_lesson_uuid)
        if source is None or not source.instrument:
            continue
        matches = panels_by_name.get(source.instrument.casefold(), [])
        if len(matches) == 1:
            entry.panel_uuid = matches[0].panel_uuid
            assigned_count += 1

    if assigned_count:
        bump_jury_revision(db)
        db.flush()
    return list_entries(db, result), assigned_count


def update_lesson_jury_required(
    db: Session,
    source_lesson_uuid: str,
    value: bool,
) -> JuryLessonEntryOut:
    result = accompanist_results.get_latest_accompanist_result(db)
    if result is None:
        raise JuryDataError("NO_CURRENT_FINALIZED_RESULT", "No current finalized Accompanist result is available.")
    source = next(
        (item for item in result.payload.entries if str(item.source_lesson_uuid) == source_lesson_uuid),
        None,
    )
    if source is None:
        raise JuryDataError("SOURCE_LESSON_NOT_FOUND", "Lesson is not present in the current finalized Accompanist result.")
    identity = db.query(module_models.AccompanistLessonIdentity).filter_by(
        lesson_uuid=source_lesson_uuid
    ).one_or_none()
    lesson = db.get(models.Lesson, identity.lesson_id) if identity else None
    if lesson is None:
        raise JuryDataError("SOURCE_LESSON_NOT_FOUND", "The source Lesson no longer exists.")

    if set_lesson_jury_required(db, lesson.id, value):
        bump_accompanist_revision(db)
    session_uuid = active_session_uuid(db)
    row = db.get(module_models.JuryLessonEntry, (session_uuid, source_lesson_uuid))
    if row is None:
        row = module_models.JuryLessonEntry(
            session_uuid=session_uuid,
            source_lesson_uuid=source_lesson_uuid,
            student_person_uuid=str(source.student_person_uuid),
            panel_uuid=None,
        )
        db.add(row)
    return _entry_out(
        session_uuid,
        row,
        source,
        (value, lesson.need_pianist),
    )


def save_availability(
    db: Session,
    pianist_person_uuid: str,
    jury_date: date,
    payload: JuryAvailabilityIn,
) -> JuryAvailabilityOut:
    session_uuid = active_session_uuid(db)
    panel_for_date = db.query(module_models.JuryPanelDate.panel_uuid).filter_by(
        session_uuid=session_uuid,
        jury_date=jury_date,
    ).first()
    if panel_for_date is None:
        raise JuryDataError(
            "JURY_DATE_MISMATCH",
            "Enter Jury Availability Windows for a date used by a Jury Panel.",
        )
    pianist = db.get(module_models.PersonIdentity, pianist_person_uuid)
    if pianist is None or db.query(module_models.AccompanistPianistIdentity).filter_by(
        person_uuid=pianist_person_uuid
    ).first() is None:
        raise JuryDataError("PIANIST_NOT_FOUND", "Pianist identity is not available from Accompanist Scheduling.")

    windows = sorted(payload.windows, key=lambda item: (item.start_minute, item.end_minute))
    for previous, current in zip(windows, windows[1:]):
        if current.start_minute < previous.end_minute:
            raise JuryDataError("OVERLAPPING_AVAILABILITY", "Jury Availability Windows must not overlap.")

    declaration = db.get(
        module_models.JuryPianistAvailabilityDeclaration,
        (session_uuid, pianist_person_uuid, jury_date),
    )
    now = datetime.utcnow()
    is_complete = bool(windows)
    if declaration is None:
        declaration = module_models.JuryPianistAvailabilityDeclaration(
            session_uuid=session_uuid,
            pianist_person_uuid=pianist_person_uuid,
            jury_date=jury_date,
            is_complete=is_complete,
            declared_at=now if is_complete else None,
            modified_at=now,
        )
        db.add(declaration)
    else:
        declaration.is_complete = is_complete
        declaration.declared_at = now if is_complete else None
        declaration.modified_at = now
        db.query(module_models.JuryPianistAvailableWindow).filter_by(
            session_uuid=session_uuid,
            pianist_person_uuid=pianist_person_uuid,
            jury_date=jury_date,
        ).delete(synchronize_session=False)

    for window in windows:
        db.add(module_models.JuryPianistAvailableWindow(
            session_uuid=session_uuid,
            pianist_person_uuid=pianist_person_uuid,
            jury_date=jury_date,
            start_minute=window.start_minute,
            end_minute=window.end_minute,
        ))
    bump_jury_revision(db)
    db.flush()
    return _availability_out(db, declaration)


def _availability_out(
    db: Session,
    declaration: module_models.JuryPianistAvailabilityDeclaration,
) -> JuryAvailabilityOut:
    windows = db.query(module_models.JuryPianistAvailableWindow).filter_by(
        session_uuid=declaration.session_uuid,
        pianist_person_uuid=declaration.pianist_person_uuid,
        jury_date=declaration.jury_date,
    ).order_by(module_models.JuryPianistAvailableWindow.start_minute).all()
    return JuryAvailabilityOut(
        session_uuid=declaration.session_uuid,
        pianist_person_uuid=declaration.pianist_person_uuid,
        jury_date=declaration.jury_date,
        is_complete=declaration.is_complete,
        declared_at=declaration.declared_at,
        modified_at=declaration.modified_at,
        windows=[AvailableWindowIn(start_minute=row.start_minute, end_minute=row.end_minute) for row in windows],
    )


def get_availability(
    db: Session,
    pianist_person_uuid: str,
    jury_date: date,
) -> JuryAvailabilityOut | None:
    session_uuid = active_session_uuid(db)
    declaration = db.get(
        module_models.JuryPianistAvailabilityDeclaration,
        (session_uuid, pianist_person_uuid, jury_date),
    )
    return _availability_out(db, declaration) if declaration is not None else None


def clear_availability(
    db: Session,
    pianist_person_uuid: str,
    jury_date: date,
) -> int:
    session_uuid = active_session_uuid(db)
    panel_for_date = db.query(module_models.JuryPanelDate.panel_uuid).filter_by(
        session_uuid=session_uuid,
        jury_date=jury_date,
    ).first()
    if panel_for_date is None:
        raise JuryDataError(
            "JURY_DATE_MISMATCH",
            "Jury Availability Windows can only be cleared for a date used by a Jury Panel.",
        )
    if db.get(module_models.PersonIdentity, pianist_person_uuid) is None or db.query(
        module_models.AccompanistPianistIdentity
    ).filter_by(person_uuid=pianist_person_uuid).first() is None:
        raise JuryDataError("PIANIST_NOT_FOUND", "Pianist identity is not available from Accompanist Scheduling.")

    window_count = db.query(module_models.JuryPianistAvailableWindow).filter_by(
        session_uuid=session_uuid,
        pianist_person_uuid=pianist_person_uuid,
        jury_date=jury_date,
    ).delete(synchronize_session=False)
    declaration = db.get(
        module_models.JuryPianistAvailabilityDeclaration,
        (session_uuid, pianist_person_uuid, jury_date),
    )
    if declaration is not None:
        db.delete(declaration)
    if window_count or declaration is not None:
        bump_jury_revision(db)
    db.flush()
    return window_count


def _issue(code: str, severity: str, message: str, *entities: str) -> ReadinessIssue:
    return ReadinessIssue(
        code=code,
        severity=severity,
        message=message,
        entity_uuids=[UUID(entity) for entity in entities],
    )


def readiness(db: Session) -> JuryReadinessOut:
    session_uuid = active_session_uuid(db)
    configuration = db.get(module_models.JuryConfiguration, session_uuid)
    if configuration is None:
        raise JuryDataError("JURY_CONFIGURATION_MISSING", "Jury configuration is unavailable.")

    result = accompanist_results.get_current_accompanist_result(db)
    issues: list[ReadinessIssue] = []
    if result is None:
        latest_result = accompanist_results.get_latest_accompanist_result(db)
        if latest_result is None:
            issues.append(_issue(
                "NO_CURRENT_FINALIZED_RESULT", "error",
                "There is no current finalized Accompanist result.",
            ))
        else:
            issues.append(_issue(
                "ACCOMPANIST_RESULT_STALE", "error",
                "Accompanist source data changed after the last finalized result; finalize the current schedule before continuing.",
                str(latest_result.result_uuid),
            ))
        return JuryReadinessOut(
            ready=False,
            source_result_uuid=latest_result.result_uuid if latest_result else None,
            source_revision=latest_result.source_revision if latest_result else None,
            jury_input_revision=configuration.input_revision,
            issues=issues,
        )

    source_by_uuid = {str(item.source_lesson_uuid): item for item in result.payload.entries}

    if result is not None and configuration.roster_source_result_uuid != str(result.result_uuid):
        issues.append(_issue(
            "ROSTER_SOURCE_CHANGED", "warning",
            "The Jury lesson roster has not been synchronized to the current finalized Accompanist result.",
            str(result.result_uuid),
        ))

    panels = {
        panel.panel_uuid: panel
        for panel in db.query(module_models.JuryPanel).filter_by(session_uuid=session_uuid).all()
    }
    panel_dates = {
        panel_date.panel_uuid: panel_date.jury_date
        for panel_date in db.query(module_models.JuryPanelDate).filter_by(session_uuid=session_uuid).all()
    }
    for panel in sorted(panels.values(), key=lambda item: item.panel_uuid):
        try:
            JuryPanelFields(
                panel_name=panel.panel_name,
                room=panel.room,
                jury_date=panel_dates.get(panel.panel_uuid),
                earliest_start_minute=panel.earliest_start_minute,
                preferred_start_minute=panel.preferred_start_minute,
                jury_length_minutes=panel.jury_length_minutes,
                break_needed=panel.break_needed,
                break_every_x_juries=panel.break_every_x_juries,
                break_length_minutes=panel.break_length_minutes,
                meal_break=panel.meal_break,
                meal_start_minute=panel.meal_start_minute,
                meal_end_minute=panel.meal_end_minute,
            )
        except Exception as error:
            issues.append(_issue(
                "INVALID_PANEL", "error", f"Panel '{panel.panel_name}' is invalid: {error}", panel.panel_uuid
            ))

    entries = db.query(module_models.JuryLessonEntry).filter_by(
        session_uuid=session_uuid
    ).order_by(module_models.JuryLessonEntry.source_lesson_uuid).all()
    required_count = 0
    used_panels: set[str] = set()
    for entry in entries:
        lesson_uuid = entry.source_lesson_uuid
        source = source_by_uuid.get(lesson_uuid)
        if source is None:
            issues.append(_issue(
                "SOURCE_LESSON_NOT_FOUND", "error",
                "Jury entry does not correspond to a lesson in the current finalized Accompanist result.",
                lesson_uuid,
            ))
            continue
        if entry.student_person_uuid != str(source.student_person_uuid):
            issues.append(_issue(
                "STUDENT_IDENTITY_MISMATCH", "error",
                "Jury entry identity does not match the finalized source lesson.",
                lesson_uuid, str(source.student_person_uuid),
            ))
        jury_required, pianist_required = source.jury_required, source.pianist_required
        if not jury_required:
            continue
        required_count += 1
        panel_date = None
        if not entry.panel_uuid or entry.panel_uuid not in panels:
            issues.append(_issue(
                "PANEL_REQUIRED", "error", "Select a valid Jury Panel for this Jury-required lesson.", lesson_uuid
            ))
        else:
            used_panels.add(entry.panel_uuid)
            panel_date = panel_dates.get(entry.panel_uuid)
            if panel_date is None:
                issues.append(_issue(
                    "PANEL_DATE_REQUIRED", "error", "Set a Scheduling Date for this Panel.", entry.panel_uuid
                ))
        if pianist_required:
            assigned = source.assigned_pianist
            if assigned is None:
                issues.append(_issue(
                    "FINALIZED_PIANIST_REQUIRED", "error",
                    "This Jury-required lesson requires a pianist but has no finalized Accompanist assignment.",
                    lesson_uuid,
                ))
                continue
            panel = panels.get(entry.panel_uuid) if entry.panel_uuid else None
            if panel is None or panel_date is None:
                continue
            declaration = db.get(
                module_models.JuryPianistAvailabilityDeclaration,
                (session_uuid, str(assigned.person_uuid), panel_date),
            )
            if declaration is None or not declaration.is_complete:
                pianist_name = assigned.display_name or "The assigned Pianist"
                date_label = panel_date.strftime("%B %d, %Y").replace(" 0", " ")
                issues.append(_issue(
                    "PIANIST_AVAILABILITY_INCOMPLETE", "error",
                    f"{pianist_name} has no Jury Availability Windows for {date_label}.",
                    lesson_uuid, str(assigned.person_uuid), str(panel.panel_uuid),
                ))

    entry_lesson_uuids = {entry.source_lesson_uuid for entry in entries}
    for lesson_uuid, source in source_by_uuid.items():
        jury_required = source.jury_required
        if jury_required and lesson_uuid not in entry_lesson_uuids:
            required_count += 1
            issues.append(_issue(
                "JURY_ENTRY_MISSING", "error",
                "Jury-required source Lesson has no Jury entry. Synchronize the roster and select a Panel.",
                lesson_uuid,
            ))

    if required_count == 0:
        issues.append(_issue("NO_JURY_REQUIRED_ENTRIES", "warning", "No Jury lesson entries are currently required."))
    for panel_uuid in sorted(set(panels) - used_panels):
        issues.append(_issue(
            "UNUSED_PANEL", "warning", f"Panel '{panels[panel_uuid].panel_name}' has no Jury-required lessons.", panel_uuid
        ))

    issues.sort(key=lambda item: (item.severity != "error", item.code, tuple(str(value) for value in item.entity_uuids)))
    return JuryReadinessOut(
        ready=not any(issue.severity == "error" for issue in issues),
        source_result_uuid=result.result_uuid if result else None,
        source_revision=result.source_revision if result else None,
        jury_input_revision=configuration.input_revision,
        issues=issues,
    )