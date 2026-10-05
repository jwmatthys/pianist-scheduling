export const DAYS_ORDER = [
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
  "Sunday",
] as const;

export type Day = (typeof DAYS_ORDER)[number];

export type AvailabilityStatus = "Available" | "Tentative" | "Unavailable";

export interface Pianist {
  id: number;
  name: string;
  email: string;
  max_hours_per_week: number | null;
  availability_complete: boolean;
}

export interface AvailabilitySlot {
  day: string;
  slot_start_minute: number;
  status: AvailabilityStatus;
}

export type AvailabilityImportLayout = "normalized" | "wide";

export interface AvailabilityWindowColumns {
  start_column: string | null;
  end_column: string | null;
}

export interface AvailabilityImportMapping {
  layout: AvailabilityImportLayout;
  person_name_column: string | null;
  email_column?: string | null;
  day_column?: string | null;
  start_column?: string | null;
  end_column?: string | null;
  status_column?: string | null;
  wide_status?: AvailabilityStatus;
  wide_windows?: Record<string, AvailabilityWindowColumns[]>;
}

export interface AvailabilityImportInspection {
  upload_token: string;
  sheets: string[];
  selected_sheet: string | null;
  columns: string[];
  sample_rows: Record<string, string>[];
  suggested_normalized: Record<string, string | null>;
  suggested_wide_windows: Record<string, AvailabilityWindowColumns[]>;
}

export interface AvailabilityImportIssue {
  severity: "error" | "warning";
  code: string;
  message: string;
  row_number: number | null;
}

export interface AvailabilityImportWindow {
  pianist_name: string;
  day: string;
  start_minute: number;
  end_minute: number;
  status: AvailabilityStatus;
}

export interface AvailabilityImportPreview {
  preview_token: string;
  sheet_name: string | null;
  rows_processed: number;
  existing_pianist_count: number;
  existing_assignment_count: number;
  incoming_pianist_count: number;
  valid_window_count: number;
  existing_slots_in_scope: number;
  pianists: {
    action: "new" | "invalid";
    pianist_name: string;
    email: string;
    max_hours_per_week: number | null;
    days: string[];
    row_numbers: number[];
  }[];
  windows: AvailabilityImportWindow[];
  errors: AvailabilityImportIssue[];
  warnings: AvailabilityImportIssue[];
  can_apply: boolean;
}

export interface AvailabilityImportApplyResult {
  pianists_removed: number;
  pianists_created: number;
  assignments_cleared: number;
  slots_replaced: number;
  slots_created: number;
  days_replaced: number;
  jury_availability_windows_removed: number;
}

export interface Lesson {
  id: number;
  teacher: string;
  teacher_email: string;
  student: string;
  student_id: string;
  day: string;
  start_minute: number;
  end_minute: number;
  room: string;
  instrument: string;
  required_pianist_name: string;
  need_pianist: boolean;
  jury_required: boolean;
  assigned_pianist_id: number | null;
  fit_quality: string;
  notes: string;
  hours: number;
  manually_edited: boolean;
}

export interface RunAssignmentResult {
  lessons: Lesson[];
  hours_by_pianist: Record<string, number>;
  conflicts: string[];
  unassigned_count: number;
}

export interface ImportPreview {
  columns: string[];
  rows: Record<string, string>[];
  upload_token: string;
}

export interface ImportCommitResult {
  created: number;
  skipped: number;
  warnings: string[];
}

export interface SchedulingSession {
  session_uuid: string;
  institution_name: string;
  program_name: string;
  term_label: string;
  year: number | null;
  start_date: string | null;
  end_date: string | null;
  created_at: string;
  modified_at: string;
}

export interface SchedulingSessionInput {
  institution_name: string;
  program_name: string;
  term_label: string;
  year: number | null;
  start_date: string | null;
  end_date: string | null;
}

export interface AccompanistFinalizationState {
  session_uuid: string;
  source_revision: number;
  current_result_uuid: string | null;
  current_result_version: number | null;
}

export interface JuryConfiguration {
  session_uuid: string;
  input_revision: number;
  roster_source_result_uuid: string | null;
  created_at: string;
  modified_at: string;
}

export interface JuryPanelInput {
  panel_name: string;
  room: string;
  jury_date: string;
  earliest_start_minute: number;
  preferred_start_minute: number;
  jury_length_minutes: number;
  break_needed: boolean;
  break_every_x_juries: number | null;
  break_length_minutes: number | null;
  meal_break: boolean;
  meal_start_minute: number | null;
  meal_end_minute: number | null;
}

export interface JuryPanel extends Omit<JuryPanelInput, "jury_date"> {
  panel_uuid: string;
  session_uuid: string;
  jury_date: string | null;
}

export interface FinalizedPianist {
  person_uuid: string;
  display_name: string;
}

export interface JuryLessonEntry {
  session_uuid: string;
  source_lesson_uuid: string;
  student_person_uuid: string;
  student_display_name: string;
  instrument: string;
  teacher: string;
  pianist_required: boolean;
  assigned_pianist: FinalizedPianist | null;
  jury_required: boolean;
  panel_uuid: string | null;
}

export interface JuryAvailableWindow {
  start_minute: number;
  end_minute: number;
}

export interface JuryAvailabilityInput {
  windows: JuryAvailableWindow[];
}

export interface JuryAvailability extends JuryAvailabilityInput {
  is_complete: boolean;
  session_uuid: string;
  pianist_person_uuid: string;
  jury_date: string;
  declared_at: string | null;
  modified_at: string;
}

export interface JuryAvailabilityClearResult {
  windows_deleted: number;
}

export interface JuryReadinessIssue {
  code: string;
  severity: "error" | "warning";
  message: string;
  entity_uuids: string[];
}

export interface JuryReadiness {
  ready: boolean;
  source_result_uuid: string | null;
  source_revision: number | null;
  jury_input_revision: number;
  issues: JuryReadinessIssue[];
}

export interface JurySynchronizationSummary {
  session_uuid: string;
  accompanist_source_revision: number;
  current_accompanist_result_uuid: string | null;
  lesson_references_checked: number;
  lesson_entries_removed: number;
  lesson_identity_references_updated: number;
  roster_entries_created: number;
  pianist_references_checked: number;
  availability_records_removed: number;
  availability_windows_removed: number;
  stale_results_detected: number;
}

export interface JuryScheduleLessonEvent {
  kind: "jury";
  source_lesson_uuid: string;
  student_person_uuid: string;
  student_display_name: string;
  instrument: string;
  panel_uuid: string;
  jury_date: string;
  pianist_person_uuid: string | null;
  pianist_display_name: string | null;
  start_minute: number;
  end_minute: number;
}

export interface JuryPeriodicBreakEvent {
  kind: "periodic_break";
  panel_uuid: string;
  jury_date: string;
  start_minute: number;
  end_minute: number;
  after_jury_count: number;
}

export interface JuryMealBreakEvent {
  kind: "meal_break";
  panel_uuid: string;
  jury_date: string;
  start_minute: number;
  end_minute: number;
}

export type JuryScheduleEvent = JuryScheduleLessonEvent | JuryPeriodicBreakEvent | JuryMealBreakEvent;

export interface JuryPanelTimeline {
  panel_uuid: string;
  panel_name: string;
  jury_date: string;
  events: JuryScheduleEvent[];
}

export interface JuryUnscheduledLesson {
  source_lesson_uuid: string;
  student_person_uuid: string;
  student_display_name: string;
  instrument: string;
  panel_uuid: string;
  jury_date: string;
  reason_code: string;
  explanation: string;
}

export interface JuryScheduleWarning {
  code: string;
  severity: "warning" | "error";
  message: string;
}

export type JuryScheduleStaleReason =
  | "accompanist_result_changed"
  | "accompanist_source_revision_changed"
  | "jury_inputs_changed";

export type JuryScheduleState = "draft" | "finalized" | "superseded";

export interface JuryScheduleResult {
  result_uuid: string;
  session_uuid: string;
  contract_id: string;
  contract_version: number;
  result_version: number;
  state: JuryScheduleState;
  created_at: string;
  source_result_uuid: string;
  source_contract_id: string;
  source_contract_version: number;
  source_revision: number;
  jury_input_revision: number;
  stale: boolean;
  stale_reasons: JuryScheduleStaleReason[];
  panel_timelines: JuryPanelTimeline[];
  unscheduled_lessons: JuryUnscheduledLesson[];
  warnings: JuryScheduleWarning[];
}

export interface JuryScheduleHistoryItem {
  result_uuid: string;
  session_uuid: string;
  result_version: number;
  state: JuryScheduleState;
  created_at: string;
  source_result_uuid: string;
  source_revision: number;
  jury_input_revision: number;
  stale: boolean;
  stale_reasons: JuryScheduleStaleReason[];
  scheduled_count: number;
  unscheduled_count: number;
}
