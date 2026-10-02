"""Typed contracts for finalized Accompanist data and Jury-owned inputs."""

from datetime import date, datetime
from typing import Annotated, Literal, Union
from uuid import UUID

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class FinalizedPianist(BaseModel):
    person_uuid: UUID
    display_name: str


class FinalizedLessonEntryV1(BaseModel):
    source_lesson_uuid: UUID
    student_person_uuid: UUID
    student_display_name: str
    instrument: str
    teacher: str
    pianist_required: bool
    assigned_pianist: FinalizedPianist | None


class AccompanistAssignmentPayloadV1(BaseModel):
    entries: list[FinalizedLessonEntryV1]


class FinalizedLessonEntry(BaseModel):
    source_lesson_uuid: UUID
    student_person_uuid: UUID
    student_display_name: str
    instrument: str
    teacher: str
    pianist_required: bool
    assigned_pianist: FinalizedPianist | None
    jury_required: bool


class AccompanistAssignmentPayload(BaseModel):
    entries: list[FinalizedLessonEntry]


class ResultEnvelope(BaseModel):
    result_uuid: UUID
    session_uuid: UUID
    module_id: str
    contract_id: str
    contract_version: int
    source_revision: int
    result_version: int
    state: Literal["draft", "finalized", "superseded"]
    created_at: datetime
    finalized_at: datetime | None
    payload_schema_version: int
    payload: AccompanistAssignmentPayload


class HistoricalResultEnvelopeV1(BaseModel):
    result_uuid: UUID
    session_uuid: UUID
    module_id: str
    contract_id: str
    contract_version: Literal[1]
    source_revision: int
    result_version: int
    state: Literal["draft", "finalized", "superseded"]
    created_at: datetime
    finalized_at: datetime | None
    payload_schema_version: Literal[1]
    payload: AccompanistAssignmentPayloadV1


class FinalizeAccompanistRequest(BaseModel):
    expected_source_revision: int = Field(ge=0)


class AccompanistFinalizationStateOut(BaseModel):
    session_uuid: UUID
    source_revision: int
    current_result_uuid: UUID | None
    current_result_version: int | None


class JuryConfigurationOut(BaseModel):
    session_uuid: UUID
    input_revision: int
    roster_source_result_uuid: UUID | None
    created_at: datetime
    modified_at: datetime

    model_config = ConfigDict(from_attributes=True)


class JuryPanelFields(BaseModel):
    panel_name: str = Field(min_length=1, max_length=200)
    room: str = Field(default="", max_length=200)
    jury_date: date
    earliest_start_minute: int = Field(ge=0, le=1439)
    preferred_start_minute: int | None = Field(default=None, ge=0, le=1439)
    jury_length_minutes: int = Field(gt=0, le=1440)
    break_needed: bool = False
    break_every_x_juries: int | None = Field(default=None, gt=0)
    break_length_minutes: int | None = Field(default=None, gt=0)
    meal_break: bool = False
    meal_start_minute: int | None = Field(default=None, ge=0, le=1439)
    meal_end_minute: int | None = Field(default=None, ge=1, le=1440)

    @field_validator("panel_name")
    @classmethod
    def trim_panel_name(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("Panel Name cannot be blank.")
        return value

    @field_validator("room")
    @classmethod
    def trim_room(cls, value: str) -> str:
        return value.strip()

    @model_validator(mode="after")
    def validate_panel(self):
        if self.preferred_start_minute is None:
            self.preferred_start_minute = self.earliest_start_minute
        if self.preferred_start_minute < self.earliest_start_minute:
            raise ValueError("Preferred Start must be on or after Earliest Start.")
        if self.break_needed:
            if self.break_every_x_juries is None:
                raise ValueError("Break Every X Juries is required when Break Needed is enabled.")
            if self.break_length_minutes is None:
                self.break_length_minutes = self.jury_length_minutes
        else:
            if self.break_every_x_juries is not None or self.break_length_minutes is not None:
                raise ValueError("Periodic break values must be empty when Break Needed is disabled.")
        if self.meal_break:
            if self.meal_start_minute is None or self.meal_end_minute is None:
                raise ValueError("Meal Start and Meal End are required when Meal Break is enabled.")
            if self.meal_end_minute <= self.meal_start_minute:
                raise ValueError("Meal End must be later than Meal Start.")
        else:
            if self.meal_start_minute is not None or self.meal_end_minute is not None:
                raise ValueError("Meal times must be empty when Meal Break is disabled.")
        return self


class JuryPanelOut(BaseModel):
    panel_uuid: UUID
    session_uuid: UUID
    panel_name: str
    room: str
    jury_date: date | None
    earliest_start_minute: int
    preferred_start_minute: int
    jury_length_minutes: int
    break_needed: bool
    break_every_x_juries: int | None
    break_length_minutes: int | None
    meal_break: bool
    meal_start_minute: int | None
    meal_end_minute: int | None

    model_config = ConfigDict(from_attributes=True)


class JuryLessonEntryUpdate(BaseModel):
    panel_uuid: UUID | None = None


class JuryRequiredUpdate(BaseModel):
    jury_required: bool


class JuryLessonEntryOut(BaseModel):
    session_uuid: UUID
    source_lesson_uuid: UUID
    student_person_uuid: UUID
    student_display_name: str
    instrument: str
    teacher: str
    pianist_required: bool
    assigned_pianist: FinalizedPianist | None
    jury_required: bool
    panel_uuid: UUID | None


class JuryPanelAutoAssignmentOut(BaseModel):
    assigned_count: int
    entries: list[JuryLessonEntryOut]


class AvailableWindowIn(BaseModel):
    start_minute: int = Field(ge=0, le=1439)
    end_minute: int = Field(ge=1, le=1440)

    @model_validator(mode="after")
    def validate_interval(self):
        if self.end_minute <= self.start_minute:
            raise ValueError("Availability Window end must be later than its start.")
        return self


class JuryAvailabilityIn(BaseModel):
    windows: list[AvailableWindowIn] = Field(default_factory=list)


class JuryAvailabilityOut(BaseModel):
    session_uuid: UUID
    pianist_person_uuid: UUID
    jury_date: date
    is_complete: bool
    declared_at: datetime | None
    modified_at: datetime
    windows: list[AvailableWindowIn]


class JuryAvailabilityClearOut(BaseModel):
    windows_deleted: int


class ReadinessIssue(BaseModel):
    code: str
    severity: Literal["error", "warning"]
    message: str
    entity_uuids: list[UUID] = Field(default_factory=list)


class JuryReadinessOut(BaseModel):
    ready: bool
    source_result_uuid: UUID | None
    source_revision: int | None
    jury_input_revision: int
    issues: list[ReadinessIssue]


class JurySynchronizationSummary(BaseModel):
    session_uuid: UUID
    accompanist_source_revision: int
    current_accompanist_result_uuid: UUID | None
    lesson_references_checked: int
    lesson_entries_removed: int
    lesson_identity_references_updated: int
    roster_entries_created: int
    pianist_references_checked: int
    availability_records_removed: int
    availability_windows_removed: int
    stale_results_detected: int


class JuryGenerateRequest(BaseModel):
    expected_jury_input_revision: int | None = Field(default=None, ge=0)


class JuryResultPanelSnapshot(BaseModel):
    panel_uuid: UUID
    panel_name: str
    jury_date: date
    earliest_start_minute: int
    preferred_start_minute: int
    jury_length_minutes: int
    break_every_x_juries: int | None
    break_length_minutes: int | None
    meal_start_minute: int | None
    meal_end_minute: int | None


class JuryScheduledLessonEvent(BaseModel):
    kind: Literal["jury"] = "jury"
    source_lesson_uuid: UUID
    student_person_uuid: UUID
    student_display_name: str
    instrument: str
    panel_uuid: UUID
    jury_date: date
    pianist_person_uuid: UUID | None
    pianist_display_name: str | None
    start_minute: int
    end_minute: int


class JuryPeriodicBreakEvent(BaseModel):
    kind: Literal["periodic_break"] = "periodic_break"
    panel_uuid: UUID
    jury_date: date
    start_minute: int
    end_minute: int
    after_jury_count: int


class JuryMealBreakEvent(BaseModel):
    kind: Literal["meal_break"] = "meal_break"
    panel_uuid: UUID
    jury_date: date
    start_minute: int
    end_minute: int


JuryScheduleEventOut = Annotated[
    Union[JuryScheduledLessonEvent, JuryPeriodicBreakEvent, JuryMealBreakEvent],
    Field(discriminator="kind"),
]


class JuryUnscheduledLessonOut(BaseModel):
    source_lesson_uuid: UUID
    student_person_uuid: UUID
    student_display_name: str
    instrument: str
    panel_uuid: UUID
    jury_date: date
    reason_code: str
    explanation: str


class JuryOptimizerDiagnosticOut(BaseModel):
    code: str
    severity: Literal["info", "warning", "error"]
    message: str


class JuryConflictExplanationOut(BaseModel):
    code: str
    message: str
    source_lesson_uuids: list[UUID]
    person_uuid: UUID | None
    panel_uuid: UUID | None


class JuryStoredSchedulePayload(BaseModel):
    session_uuid: UUID
    source_result_uuid: UUID
    source_contract_id: str
    source_contract_version: int
    source_revision: int
    jury_input_revision: int
    required_lesson_uuids: list[UUID]
    events: list[JuryScheduleEventOut]
    unscheduled_lessons: list[JuryUnscheduledLessonOut]
    diagnostics: list[JuryOptimizerDiagnosticOut]
    conflict_explanations: list[JuryConflictExplanationOut]
    panel_snapshots: list[JuryResultPanelSnapshot]
    readiness_warnings: list[ReadinessIssue]


class JuryPanelTimelineOut(BaseModel):
    panel_uuid: UUID
    panel_name: str
    jury_date: date
    events: list[JuryScheduleEventOut]


class JuryScheduleWarningOut(BaseModel):
    code: str
    severity: Literal["warning", "error"]
    message: str


class JuryScheduleViewOut(BaseModel):
    result_uuid: UUID
    session_uuid: UUID
    contract_id: str
    contract_version: int
    result_version: int
    state: Literal["draft", "finalized", "superseded"]
    created_at: datetime
    source_result_uuid: UUID
    source_contract_id: str
    source_contract_version: int
    source_revision: int
    jury_input_revision: int
    stale: bool
    stale_reasons: list[Literal[
        "accompanist_result_changed",
        "accompanist_source_revision_changed",
        "jury_inputs_changed",
    ]]
    panel_timelines: list[JuryPanelTimelineOut]
    unscheduled_lessons: list[JuryUnscheduledLessonOut]
    warnings: list[JuryScheduleWarningOut]


class JuryScheduleHistoryItemOut(BaseModel):
    result_uuid: UUID
    session_uuid: UUID
    result_version: int
    state: Literal["draft", "finalized", "superseded"]
    created_at: datetime
    source_result_uuid: UUID
    source_revision: int
    jury_input_revision: int
    stale: bool
    stale_reasons: list[Literal[
        "accompanist_result_changed",
        "accompanist_source_revision_changed",
        "jury_inputs_changed",
    ]]
    scheduled_count: int
    unscheduled_count: int