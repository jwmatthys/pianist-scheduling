import { useCallback, useEffect, useState } from "react";
import { api } from "../../lib/api";
import { formatMinutes, parseTwelveHourTime } from "../../lib/time";
import { ConfirmDialog } from "../clinical/ClinicalDialogs";
import type {
  LessonInstructor,
  LessonReadiness,
  LessonReadinessCategory,
  LessonReadinessWarning,
  LessonModuleState,
  LessonSchedule,
  LessonScheduleEntry,
} from "../../lib/types";

const CATEGORY_ORDER: LessonReadinessCategory[] = ["data", "students", "instructors", "tags"];
const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const CATEGORY_TITLES: Record<LessonReadinessCategory, string> = {
  data: "Overall",
  students: "Students",
  instructors: "Instructors",
  tags: "Tags",
};

function plural(count: number, singular: string, pluralForm = `${singular}s`): string {
  return count === 1 ? singular : pluralForm;
}

function warningSummary(code: string, count: number): string {
  switch (code) {
    case "NO_STUDENTS":
      return "No students have been entered.";
    case "NO_INSTRUCTORS":
      return "No instructors have been entered.";
    case "NO_FEASIBLE_CANDIDATES":
      return "No student has a feasible instructor and lesson time.";
    case "STUDENT_WITHOUT_AVAILABILITY":
      return `${count} ${plural(count, "student has", "students have")} no availability`;
    case "INSTRUCTOR_WITHOUT_AVAILABILITY":
      return `${count} ${plural(count, "instructor has", "instructors have")} no availability`;
    case "STUDENT_WITHOUT_CANDIDATES":
      return `${count} ${plural(count, "student has", "students have")} no feasible instructors`;
    case "REQUIRED_TAG_WITHOUT_MATCH":
      return `${count} Required ${plural(count, "tag has", "tags have")} no matching instructor`;
    case "ZERO_CAPACITY":
      return `${count} ${plural(count, "instructor has", "instructors have")} zero capacity`;
    case "INVALID_INSTRUCTOR_WINDOW":
      return `${count} ${plural(count, "instructor has", "instructors have")} invalid Availability Windows`;
    case "INSTRUCTOR_WITHOUT_SCHEDULABLE_LESSONS":
      return `${count} ${plural(count, "instructor cannot", "instructors cannot")} fit a lesson inside availability`;
    default:
      return code;
  }
}

function ReadinessWarningGroup({ warning }: { warning: LessonReadinessWarning }) {
  return (
    <li
      className={`clinical-readiness-item clinical-readiness-${warning.severity}`}
      data-code={warning.code}
    >
      <span aria-hidden="true">⚠ </span>
      {warningSummary(warning.code, warning.subjects.length || 1)}
      {warning.subjects.length > 0 && (
        <ul>
          {warning.subjects.map((subject) => (
            <li key={subject.id}>
              {subject.kind === "tag"
                ? `Required tag "${subject.label}" has no matching instructor`
                : subject.label}
            </li>
          ))}
        </ul>
      )}
    </li>
  );
}

const MANUALLY_UNPLACED_OPTION = "__manually_unplaced__";
const UNASSIGNED_OPTION = "__unassigned__";
const UNPLACED_REASONS: Record<string, string> = {
  NO_FEASIBLE_CANDIDATE: "No feasible instructor",
  CAPACITY_LIMIT: "Capacity limit",
  INSTRUCTOR_OVERLAP: "Instructor overlap",
  MANUALLY_UNPLACED: "Manually Unplaced",
};

// Unassigned first, then instructors, then Manually Unplaced.
function InstructorOptions({ instructors }: { instructors: LessonInstructor[] }) {
  return (
    <>
      <option value={UNASSIGNED_OPTION}>Unassigned</option>
      <option disabled>────────────</option>
      {instructors.map((instructor) => (
        <option key={instructor.id} value={instructor.id}>{instructor.name}</option>
      ))}
      <option disabled>────────────</option>
      <option value={MANUALLY_UNPLACED_OPTION}>Manually Unplaced</option>
    </>
  );
}

function displayType(type: LessonScheduleEntry["placement_type"]): string {
  if (type === "automatic") return "Automatic";
  if (type === "manual") return "Manual";
  if (type === "manuallyUnplaced") return "Manually Unplaced";
  return "Unassigned";
}

function LessonPlacementRow({
  entry,
  instructors,
  busy,
  onMove,
  onUnplace,
  onUnassign,
}: {
  entry: LessonScheduleEntry;
  instructors: LessonInstructor[];
  busy: boolean;
  onMove: (instructorId: string, day: string, startMinute: number) => Promise<boolean>;
  onUnplace: () => Promise<boolean>;
  onUnassign: () => Promise<boolean>;
}) {
  const [instructorId, setInstructorId] = useState(entry.instructor_id ?? "");
  const [day, setDay] = useState(entry.day ?? "Monday");
  const [startTime, setStartTime] = useState(entry.start_minute === null ? "" : formatMinutes(entry.start_minute));
  const [saving, setSaving] = useState(false);
  const startMinute = parseTwelveHourTime(startTime);
  const invalidTime = startTime.length > 0 && startMinute === null;
  const lessonLength = entry.start_minute !== null && entry.end_minute !== null
    ? entry.end_minute - entry.start_minute
    : instructors.find((item) => item.id === instructorId)?.lesson_length_minutes ?? 0;

  const reset = () => {
    setInstructorId(entry.instructor_id ?? "");
    setDay(entry.day ?? "Monday");
    setStartTime(entry.start_minute === null ? "" : formatMinutes(entry.start_minute));
  };

  const commit = async (nextInstructorId: string, nextDay: string, nextTime: string) => {
    const minute = parseTwelveHourTime(nextTime);
    if (!nextInstructorId || minute === null) return;
    if (nextInstructorId === entry.instructor_id && nextDay === entry.day && minute === entry.start_minute) {
      setStartTime(formatMinutes(minute));
      return;
    }
    setSaving(true);
    const ok = await onMove(nextInstructorId, nextDay, minute);
    if (!ok) reset();
    setSaving(false);
  };

  const hasWarnings = entry.warnings.length > 0;
  return (
    <tr className={hasWarnings ? "clinical-placement-row-warning" : ""}>
      <td>{entry.student_name}</td>
      <td>
        <select
          aria-label={`Instructor for ${entry.student_name}`}
          value={instructorId}
          disabled={busy || saving}
          onChange={(event) => {
            const value = event.target.value;
            if (value === MANUALLY_UNPLACED_OPTION || value === UNASSIGNED_OPTION) {
              setSaving(true);
              void (value === UNASSIGNED_OPTION ? onUnassign() : onUnplace()).then((ok) => {
                if (!ok) setSaving(false);
              });
              return;
            }
            setInstructorId(value);
            void commit(value, day, startTime);
          }}
        >
          <InstructorOptions instructors={instructors} />
        </select>
      </td>
      <td>
        <select
          aria-label={`Day for ${entry.student_name}`}
          value={day}
          disabled={busy || saving}
          onChange={(event) => {
            setDay(event.target.value);
            void commit(instructorId, event.target.value, startTime);
          }}
        >
          {WEEKDAYS.map((weekday) => <option key={weekday} value={weekday}>{weekday}</option>)}
        </select>
      </td>
      <td>
        <input
          type="text"
          className="clinical-inline-time"
          aria-label={`Start time for ${entry.student_name}`}
          aria-invalid={invalidTime}
          value={startTime}
          disabled={busy || saving}
          placeholder="9:00 AM"
          onChange={(event) => setStartTime(event.target.value)}
          onBlur={() => void commit(instructorId, day, startTime)}
          onKeyDown={(event) => {
            if (event.key === "Enter") event.currentTarget.blur();
          }}
        />
        {invalidTime ? (
          <div className="inline-error">Enter a valid time, such as 9:00 AM.</div>
        ) : (
          startMinute !== null && lessonLength > 0 && (
            <div className="muted">Ends {formatMinutes(startMinute + lessonLength)}</div>
          )
        )}
      </td>
      <td>{displayType(entry.placement_type)}</td>
      <td>
        {hasWarnings ? (
          <>
            <strong>
              {entry.warnings.every((w) => w.code === "TENTATIVE_AVAILABILITY") ? "Placement warning" : "Placement conflict"}
            </strong>
            <ul className="clinical-placement-warnings">
              {entry.warnings.map((warning) => <li key={warning.code}>{warning.message}</li>)}
            </ul>
          </>
        ) : entry.verified ? (
          "Verified"
        ) : (
          "Needs review"
        )}
      </td>
    </tr>
  );
}

export function LessonSchedulePage() {
  const [readiness, setReadiness] = useState<LessonReadiness | null>(null);
  const [moduleState, setModuleState] = useState<LessonModuleState | null>(null);
  const [schedule, setSchedule] = useState<LessonSchedule | null>(null);
  const [instructors, setInstructors] = useState<LessonInstructor[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [actionLoading, setActionLoading] = useState(false);
  const [confirmClear, setConfirmClear] = useState(false);

  const loadReadiness = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const [nextReadiness, nextState, nextSchedule, nextInstructors] = await Promise.all([
        api.getLessonReadiness(),
        api.getLessonModuleState(),
        api.getLessonSchedule(),
        api.listLessonInstructors(),
      ]);
      setReadiness(nextReadiness);
      setModuleState(nextState);
      setSchedule(nextSchedule);
      setInstructors(nextInstructors);
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadReadiness();
  }, [loadReadiness]);

  const runScheduleMutation = async (mutation: () => Promise<LessonSchedule>): Promise<boolean> => {
    setActionLoading(true);
    setError(null);
    try {
      const result = await mutation();
      setSchedule(result);
      await loadReadiness();
      return true;
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
      return false;
    } finally {
      setActionLoading(false);
    }
  };

  const runScheduleAction = (action: "generate" | "clear") =>
    runScheduleMutation(action === "generate" ? api.generateLessonSchedule : api.clearLessonSchedule);

  const placedEntries = (schedule?.entries ?? []).filter(
    (entry) => entry.placement_type === "automatic" || entry.placement_type === "manual",
  );
  const unplacedEntries = (schedule?.entries ?? []).filter(
    (entry) => entry.placement_type === null || entry.placement_type === "manuallyUnplaced",
  );

  const scheduleStale = schedule !== null && moduleState !== null
    && (schedule.is_schedule_stale
      || schedule.source_data_revision !== moduleState.source_data_revision);

  return (
    <section className="clinical-sites clinical-schedule" aria-labelledby="lesson-schedule-heading">
      <div className="clinical-section clinical-schedule-toolbar">
        <div className="clinical-section-heading">
          <div>
            <h2 id="lesson-schedule-heading">Schedule</h2>
            <p>
              Generating replaces automatic placements only. Manual placements and Manually Unplaced decisions are
              preserved.
            </p>
          </div>
        </div>
        <div className="dialog-actions">
          <button
            type="button"
            className="primary"
            disabled={loading || actionLoading || !readiness?.is_ready}
            onClick={() => void runScheduleAction("generate")}
          >
            Generate Schedule
          </button>
          <button
            type="button"
            className="danger-btn"
            disabled={loading || actionLoading || !schedule?.has_schedule}
            onClick={() => setConfirmClear(true)}
          >
            Clear Schedule
          </button>
        </div>
      </div>
      {confirmClear && (
        <ConfirmDialog
          title="Clear Lesson Schedule?"
          message="This removes automatic Lesson assignments and automatic unplaced results. Manual assignments and Manually Unplaced decisions are preserved."
          confirmLabel="Clear Schedule"
          onCancel={() => setConfirmClear(false)}
          onConfirm={() => {
            setConfirmClear(false);
            void runScheduleAction("clear");
          }}
        />
      )}
      {error && <div className="error-banner" role="alert">{error}</div>}
      {loading && !readiness && <p role="status">Checking Lesson Scheduling readiness…</p>}
      {readiness && (
        <div className="clinical-section">
          <h3>Warnings</h3>
          {readiness.warnings.length === 0 ? (
            <p><span aria-hidden="true">✓ </span>No readiness warnings.</p>
          ) : (
            CATEGORY_ORDER.map((category) => {
              const warnings = readiness.warnings.filter((item) => item.category === category);
              if (warnings.length === 0) return null;
              return (
                <div key={category} className="clinical-readiness-category">
                  <h4>{CATEGORY_TITLES[category]}</h4>
                  <ul className="clinical-readiness-list">
                    {warnings.map((warning) => (
                      <ReadinessWarningGroup key={warning.code} warning={warning} />
                    ))}
                  </ul>
                </div>
              );
            })
          )}
        </div>
      )}
      {schedule?.has_schedule && (
        <>
        <div className="clinical-section">
          <h3>Placements</h3>
          <div className="clinical-table-scroll">
            <table className="clinical-placement-table">
              <thead>
                <tr>
                  <th>Student</th>
                  <th>Instructor</th>
                  <th>Day</th>
                  <th>Start Time</th>
                  <th>Placement Type</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {placedEntries.length === 0 ? (
                  <tr><td colSpan={6}>No lessons assigned.</td></tr>
                ) : placedEntries.map((entry) => (
                  <LessonPlacementRow
                    key={`${entry.student_id}-${entry.instructor_id}-${entry.day}-${entry.start_minute}-${entry.placement_type}`}
                    entry={entry}
                    instructors={instructors}
                    busy={loading || actionLoading}
                    onMove={(instructorId, day, startMinute) => runScheduleMutation(() => api.moveLessonManualPlacement(
                      entry.student_id,
                      { student_id: entry.student_id, instructor_id: instructorId, day, start_minute: startMinute },
                    ))}
                    onUnplace={() => runScheduleMutation(() => api.manuallyUnplaceLessonStudent(entry.student_id))}
                    onUnassign={() => runScheduleMutation(() => api.removeLessonPlacement(entry.student_id))}
                  />
                ))}
              </tbody>
            </table>
          </div>
        </div>
        <div className="clinical-section">
          <h3>Unplaced Students</h3>
          {unplacedEntries.length === 0 ? (
            <p>All students currently have a lesson.</p>
          ) : (
            <div className="clinical-table-scroll">
              <table className="clinical-placement-table">
                <thead>
                  <tr><th>Student</th><th>Instructor</th><th>Reason</th><th>Placement Type</th></tr>
                </thead>
                <tbody>
                  {unplacedEntries.map((entry) => (
                    <tr key={entry.student_id} className="clinical-unplaced-row">
                      <td>{entry.student_name}</td>
                      <td>
                        <select
                          aria-label={`Instructor for ${entry.student_name}`}
                          value={entry.placement_state === "manuallyUnplaced" ? MANUALLY_UNPLACED_OPTION : UNASSIGNED_OPTION}
                          disabled={loading || actionLoading}
                          onChange={(event) => {
                            const value = event.target.value;
                            if (value === MANUALLY_UNPLACED_OPTION) {
                              void runScheduleMutation(() => api.manuallyUnplaceLessonStudent(entry.student_id));
                            } else if (value === UNASSIGNED_OPTION) {
                              void runScheduleMutation(() => api.removeLessonPlacement(entry.student_id));
                            } else {
                              void runScheduleMutation(() => api.createLessonManualPlacement({
                                student_id: entry.student_id,
                                instructor_id: value,
                                day: "Monday",
                                start_minute: 540,
                              }));
                            }
                          }}
                        >
                          <InstructorOptions instructors={instructors} />
                        </select>
                      </td>
                      <td>{UNPLACED_REASONS[entry.unplaced_reason ?? "NO_FEASIBLE_CANDIDATE"] ?? entry.unplaced_reason}</td>
                      <td>{displayType(entry.placement_type)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
          {schedule.verification.length > 0 && (
            <div className="clinical-section">
              <h3>Verification Warnings</h3>
              <ul aria-label="Placement verification warnings">
                {schedule.verification.map((item) => (
                  <li key={item.student_id}>
                    {item.student_name}: {item.warnings.map((warning) => warning.message).join(" ")}
                  </li>
                ))}
              </ul>
            </div>
          )}
          <div className="clinical-section">
            {scheduleStale && (
              <div className="warning-banner" role="status" aria-label="Schedule out of date">
                Schedule is out of date. Student or Instructor data has changed since lessons were last generated.
                Existing placements have been preserved. Regenerate or clear the schedule.
              </div>
            )}
            <h3>Schedule Summary</h3>
            <ul className="clinical-readiness-list" aria-label="Schedule lifecycle status">
              <li>Automatic: {schedule.automatic_count}</li>
              <li>Manual: {schedule.manual_count}</li>
              <li>Unassigned: {unplacedEntries.filter((entry) => entry.placement_state === "unassigned").length}</li>
              <li>Manually Unplaced: {schedule.manually_unplaced_count}</li>
              <li>Verified: {schedule.verified_count}</li>
              <li>Unverified: {schedule.unverified_count}</li>
              <li>Stale: {scheduleStale ? "Yes" : "No"}</li>
            </ul>
            <dl aria-label="Schedule summary">
              <dt>Students scheduled</dt>
              <dd>{schedule.students_scheduled}</dd>
              <dt>Students unscheduled</dt>
              <dd>{schedule.students_unscheduled}</dd>
              <dt>Instructor utilization</dt>
              <dd>{(schedule.instructor_utilization * 100).toFixed(1)}%</dd>
              <dt>Preferred-tag matches</dt>
              <dd>{schedule.preferred_tag_matches}</dd>
              <dt>Tentative minutes</dt>
              <dd>{schedule.tentative_minutes}</dd>
              <dt>Instructor gap cost</dt>
              <dd>{schedule.gap_cost_minutes} minutes</dd>
            </dl>
            <p className="muted">
              Schedule revision {schedule.schedule_revision}; based on source revision {schedule.source_data_revision}.
            </p>
          </div>
        </>
      )}
    </section>
  );
}
