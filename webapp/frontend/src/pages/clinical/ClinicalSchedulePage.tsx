import { useCallback, useEffect, useState } from "react";
import { api } from "../../lib/api";
import { formatMinutes, parseTwelveHourTime } from "../../lib/time";
import type {
  ClinicalReadiness,
  ClinicalReadinessCategory,
  ClinicalReadinessWarning,
  ClinicalSchedule,
  ClinicalScheduleEntry,
  PlacementOpportunity,
} from "../../lib/types";
import { ConfirmDialog } from "./ClinicalDialogs";

const UNPLACED_REASONS: Record<string, string> = {
  NO_FEASIBLE_CANDIDATE: "No feasible opportunity",
  CAPACITY_LIMIT: "Capacity limit",
  MANUALLY_UNPLACED: "Manually unplaced",
};

const CATEGORY_TITLES: Record<ClinicalReadinessCategory, string> = {
  data: "Data",
  students: "Students",
  opportunities: "Placement Opportunities",
  tags: "Tags",
};
const CATEGORY_ORDER: ClinicalReadinessCategory[] = ["data", "students", "opportunities", "tags"];

function plural(count: number, one: string, many: string) {
  return count === 1 ? one : many;
}

// Per-code wording; the backend supplies only machine-readable codes and subjects.
function groupSummary(code: string, items: ClinicalReadinessWarning[]): string {
  const n = items.length;
  switch (code) {
    case "NO_STUDENTS":
      return "No students have been entered.";
    case "NO_PLACEMENT_OPPORTUNITIES":
      return "No placement opportunities have been entered.";
    case "NO_FEASIBLE_CANDIDATES":
      return "No student has any feasible placement opportunity.";
    case "STUDENT_WITHOUT_AVAILABILITY":
      return `${n} ${plural(n, "student has", "students have")} no availability`;
    case "STUDENT_WITHOUT_CANDIDATES":
      return `${n} ${plural(n, "student has", "students have")} no feasible opportunities`;
    case "OPPORTUNITY_WITHOUT_AVAILABILITY_MATCH":
      return `${n} ${plural(n, "placement opportunity matches", "placement opportunities match")} no student availability`;
    case "REQUIRED_TAG_WITHOUT_MATCH":
      return `${n} Required ${plural(n, "tag has", "tags have")} no matching site`;
    case "ZERO_CAPACITY":
      return `${n} ${plural(n, "placement opportunity has", "placement opportunities have")} zero capacity`;
    case "INVALID_SITE_WINDOW":
      return `${n} ${plural(n, "placement opportunity has", "placement opportunities have")} an invalid time window`;
    case "CAPACITY_BELOW_STUDENT_COUNT":
      return "Total capacity is lower than the number of students, so some students will remain unplaced.";
    default:
      return code;
  }
}

function WarningGroup({ code, items }: { code: string; items: ClinicalReadinessWarning[] }) {
  const severity = items[0].severity;
  const subjects = items.filter((item) => item.subject);
  return (
    <li className={`clinical-readiness-item clinical-readiness-${severity}`} data-code={code}>
      <span aria-hidden="true">{severity === "info" ? "ℹ" : "⚠"} </span>
      {code === "REQUIRED_TAG_WITHOUT_MATCH" ? (
        <>
          {groupSummary(code, items)}
          <ul>
            {subjects.map((item) => (
              <li key={item.subject!.id}>Required tag &ldquo;{item.subject!.label}&rdquo; has no matching site</li>
            ))}
          </ul>
        </>
      ) : (
        <>
          {groupSummary(code, items)}
          {subjects.length > 0 && (
            <ul>
              {subjects.map((item) => (
                <li key={item.subject!.id}>{item.subject!.label}</li>
              ))}
            </ul>
          )}
        </>
      )}
    </li>
  );
}

function clockLabel(minutes: number): string {
  const dayOffset = Math.floor(minutes / (24 * 60));
  const value = formatMinutes(minutes % (24 * 60));
  return dayOffset > 0 ? `${value} (+${dayOffset} day)` : value;
}

function displayType(kind: ClinicalScheduleEntry["kind"]): string {
  if (kind === "automatic") return "Automatic";
  if (kind === "manual") return "Manual";
  if (kind === "manuallyUnplaced") return "Manually Unplaced";
  return "Unplaced";
}

const MANUALLY_UNPLACED_OPTION = "__manually_unplaced__";

function PlacementRow({
  entry,
  opportunities,
  busy,
  onMove,
  onUnplace,
}: {
  entry: ClinicalScheduleEntry;
  opportunities: PlacementOpportunity[];
  busy: boolean;
  onMove: (opportunityId: string, startMinute: number) => Promise<boolean>;
  onUnplace: () => Promise<boolean>;
}) {
  const [opportunityId, setOpportunityId] = useState(entry.opportunity_id ?? "");
  const [startTime, setStartTime] = useState(entry.start_minute === null ? "" : formatMinutes(entry.start_minute));
  const [saving, setSaving] = useState(false);
  const startMinute = parseTwelveHourTime(startTime);
  const invalidTime = startTime.length > 0 && startMinute === null;
  const opportunity = opportunities.find((item) => item.id === opportunityId);

  const commit = async (nextOpportunityId: string, nextTime: string) => {
    const minute = parseTwelveHourTime(nextTime);
    if (!nextOpportunityId || minute === null) return;
    if (nextOpportunityId === entry.opportunity_id && minute === entry.start_minute) {
      setStartTime(formatMinutes(minute));
      return;
    }
    setSaving(true);
    const ok = await onMove(nextOpportunityId, minute);
    if (!ok) {
      setOpportunityId(entry.opportunity_id ?? "");
      setStartTime(entry.start_minute === null ? "" : formatMinutes(entry.start_minute));
    }
    setSaving(false);
  };

  const hasWarnings = entry.warnings.length > 0;
  return (
    <tr className={hasWarnings ? "clinical-placement-row-warning" : ""}>
      <td>
        {entry.student_name}
        {entry.student_number && <div className="muted">Student ID: {entry.student_number}</div>}
      </td>
      <td>
        <select
          aria-label={`Placement Site for ${entry.student_name}`}
          value={opportunityId}
          disabled={busy || saving}
          onChange={(event) => {
            const value = event.target.value;
            if (value === MANUALLY_UNPLACED_OPTION) {
              setSaving(true);
              void onUnplace().then((ok) => {
                if (!ok) setSaving(false);
              });
              return;
            }
            setOpportunityId(value);
            const option = opportunities.find((item) => item.id === value);
            const nextTime = option && (invalidTime || startMinute === null) ? formatMinutes(option.start_minute) : startTime;
            setStartTime(nextTime);
            void commit(value, nextTime);
          }}
        >
          {opportunities.map((item) => (
            <option key={item.id} value={item.id}>
              {item.site_name} | {item.day_of_week.slice(0, 3)} {formatMinutes(item.start_minute)}–
              {formatMinutes(item.end_minute)}
            </option>
          ))}
          <option value={MANUALLY_UNPLACED_OPTION}>Manually unplaced</option>
        </select>
        <div className="muted">{entry.tags.length ? `Tags: ${entry.tags.join(", ")}` : "No tags"}</div>
      </td>
      <td>{opportunity?.day_of_week ?? entry.day}</td>
      <td>
        <input
          type="text"
          className="clinical-inline-time"
          aria-label={`Start time for ${entry.student_name}`}
          aria-invalid={invalidTime}
          value={startTime}
          disabled={busy || saving}
          placeholder="8:10 AM"
          onChange={(event) => setStartTime(event.target.value)}
          onBlur={() => void commit(opportunityId, startTime)}
          onKeyDown={(event) => {
            if (event.key === "Enter") event.currentTarget.blur();
          }}
        />
        {invalidTime ? (
          <div className="inline-error">Enter a valid time, such as 8:10 AM.</div>
        ) : (
          opportunity &&
          startMinute !== null && (
            <div className="muted">Ends {clockLabel(startMinute + opportunity.session_length_minutes)}</div>
          )
        )}
      </td>
      <td>{displayType(entry.kind)}</td>
      <td>
        {hasWarnings ? (
          <>
            <strong>{entry.warnings.every((w) => w.code === "TENTATIVE_AVAILABILITY") ? "Placement warning" : "Placement conflict"}</strong>
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

export function ClinicalSchedulePage() {
  const [readiness, setReadiness] = useState<ClinicalReadiness | null>(null);
  const [schedule, setSchedule] = useState<ClinicalSchedule | null>(null);
  const [opportunities, setOpportunities] = useState<PlacementOpportunity[]>([]);
  const [confirmClear, setConfirmClear] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    try {
      const [nextReadiness, nextSchedule, nextOpportunities] = await Promise.all([
        api.getClinicalReadiness(),
        api.getClinicalSchedule(),
        api.listClinicalOpportunities(),
      ]);
      setReadiness(nextReadiness);
      setSchedule(nextSchedule);
      setOpportunities(nextOpportunities);
      setError("");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const run = async (action: () => Promise<ClinicalSchedule>) => {
    setBusy(true);
    try {
      const nextSchedule = await action();
      setSchedule(nextSchedule);
      setError("");
      try {
        const [nextReadiness, nextOpportunities] = await Promise.all([
          api.getClinicalReadiness(),
          api.listClinicalOpportunities(),
        ]);
        setReadiness(nextReadiness);
        setOpportunities(nextOpportunities);
      } catch (caught) {
        setError(caught instanceof Error ? caught.message : String(caught));
      }
      return true;
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
      return false;
    } finally {
      setBusy(false);
    }
  };

  const saveManual = (
    studentId: string,
    opportunityId: string,
    startMinute: number,
    isMove: boolean,
  ) =>
    run(() => {
      const data = { student_id: studentId, opportunity_id: opportunityId, start_minute: startMinute };
      return isMove ? api.moveClinicalManualPlacement(studentId, data) : api.createClinicalManualPlacement(data);
    });

  const placedEntries = (schedule?.entries ?? []).filter(
    (entry) => entry.kind === "automatic" || entry.kind === "manual",
  );
  const unplacedEntries = (schedule?.entries ?? []).filter(
    (entry) => entry.kind === null || entry.kind === "manuallyUnplaced",
  );

  const groups = new Map<string, ClinicalReadinessWarning[]>();
  for (const warning of readiness?.warnings ?? []) {
    groups.set(warning.code, [...(groups.get(warning.code) ?? []), warning]);
  }

  return (
    <section className="clinical-sites clinical-schedule" aria-labelledby="clinical-schedule-heading">
      <div className="clinical-section clinical-schedule-toolbar">
        <div className="clinical-section-heading">
          <div>
            <h2 id="clinical-schedule-heading">Schedule</h2>
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
            disabled={busy || !readiness?.is_ready}
            onClick={() => void run(api.generateClinicalSchedule)}
          >
            Generate Schedule
          </button>
          <button
            type="button"
            className="danger-btn"
            disabled={busy || !schedule?.has_placements}
            onClick={() => setConfirmClear(true)}
          >
            Clear Schedule
          </button>
        </div>
      </div>
      {confirmClear && (
        <ConfirmDialog
          title="Clear Schedule?"
          message="This removes all generated and manual placement decisions. Students and Clinical Sites are not changed."
          confirmLabel="Clear Schedule"
          onCancel={() => setConfirmClear(false)}
          onConfirm={() => {
            setConfirmClear(false);
            void run(api.clearClinicalSchedule);
          }}
        />
      )}
      {error && (
        <div className="error-banner" role="alert">
          {error}
        </div>
      )}
      {readiness && (
        <div className="clinical-section">
          <h3>Warnings</h3>
          {groups.size === 0 ? (
            <p>
              <span aria-hidden="true">✓ </span>No readiness warnings.
            </p>
          ) : (
            CATEGORY_ORDER.map((category) => {
              const entries = [...groups.entries()].filter(([, items]) => items[0].category === category);
              if (entries.length === 0) return null;
              return (
                <div key={category} className="clinical-readiness-category">
                  <h4>{CATEGORY_TITLES[category]}</h4>
                  <ul className="clinical-readiness-list">
                    {entries.map(([code, items]) => (
                      <WarningGroup key={code} code={code} items={items} />
                    ))}
                  </ul>
                </div>
              );
            })
          )}
        </div>
      )}
      {schedule && (
        <>
          <div className="clinical-section">
            <h3>Placements</h3>
            <div className="clinical-table-scroll">
              <table className="clinical-placement-table">
                <thead>
                  <tr>
                    <th>Student</th>
                    <th>Placement Site</th>
                    <th>Day</th>
                    <th>Start Time</th>
                    <th>Placement Type</th>
                    <th>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {placedEntries.length === 0 ? (
                    <tr><td colSpan={6}>No placements assigned.</td></tr>
                  ) : placedEntries.map((entry) => (
                    <PlacementRow
                      key={`${entry.student_id}-${entry.opportunity_id}-${entry.start_minute}-${entry.kind}`}
                      entry={entry}
                      opportunities={opportunities}
                      busy={busy}
                      onMove={(opportunityId, startMinute) => saveManual(entry.student_id, opportunityId, startMinute, true)}
                      onUnplace={() => run(() => api.markClinicalStudentManuallyUnplaced(entry.student_id))}
                    />
                  ))}
                </tbody>
              </table>
            </div>
          </div>
          <div className="clinical-section">
            <h3>Unplaced Students</h3>
            {unplacedEntries.length === 0 ? (
              <p>All students currently have a placement.</p>
            ) : (
              <div className="clinical-table-scroll">
                <table className="clinical-placement-table">
                  <thead><tr><th>Student</th><th>Placement Site</th><th>Reason</th><th>Placement Type</th></tr></thead>
                  <tbody>
                    {unplacedEntries.map((entry) => (
                      <tr key={entry.student_id} className="clinical-unplaced-row">
                        <td>
                          {entry.student_name}
                          {entry.student_number && <div className="muted">Student ID: {entry.student_number}</div>}
                        </td>
                        <td>
                          <select
                            aria-label={`Placement Site for ${entry.student_name}`}
                            value={entry.kind === "manuallyUnplaced" ? MANUALLY_UNPLACED_OPTION : ""}
                            disabled={busy}
                            onChange={(event) => {
                              const value = event.target.value;
                              if (value === MANUALLY_UNPLACED_OPTION) {
                                void run(() => api.markClinicalStudentManuallyUnplaced(entry.student_id));
                              } else if (value === "") {
                                void run(() => api.removeClinicalPlacementOverride(entry.student_id));
                              } else {
                                const option = opportunities.find((item) => item.id === value);
                                if (option) void saveManual(entry.student_id, value, option.start_minute, false);
                              }
                            }}
                          >
                            <option value="">Unplaced - select a site</option>
                            {opportunities.map((item) => (
                              <option key={item.id} value={item.id}>
                                {item.site_name} | {item.day_of_week.slice(0, 3)} {formatMinutes(item.start_minute)}–
                                {formatMinutes(item.end_minute)}
                              </option>
                            ))}
                            <option value={MANUALLY_UNPLACED_OPTION}>Manually unplaced</option>
                          </select>
                        </td>
                        <td>{UNPLACED_REASONS[entry.unplaced_reason ?? "NO_FEASIBLE_CANDIDATE"]}</td>
                        <td>{displayType(entry.kind)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
          {schedule.has_placements && (
            <div className="clinical-section">
              {schedule.is_schedule_stale && (
                <div className="warning-banner" role="status" aria-label="Schedule out of date">
                  Schedule is out of date. Student or Clinical Site data has changed since placements were last
                  generated. Existing placements have been preserved. Regenerate or clear the schedule.
                </div>
              )}
              <h3>Schedule Summary</h3>
              <ul className="clinical-readiness-list">
                <li>
                  {schedule.placed_count} {plural(schedule.placed_count, "student", "students")} placed
                </li>
                <li>
                  {schedule.unplaced_count} {plural(schedule.unplaced_count, "student", "students")} unplaced
                </li>
                <li>
                  Utilization: {Math.round(schedule.utilization * 100)}% ({schedule.used_capacity} of{" "}
                  {schedule.total_capacity} {plural(schedule.total_capacity, "seat", "seats")})
                </li>
                <li>Preferred-tag matches: {schedule.preferred_tag_matches}</li>
                <li>Tentative minutes: {schedule.tentative_minutes}</li>
              </ul>
            </div>
          )}
        </>
      )}
    </section>
  );
}
