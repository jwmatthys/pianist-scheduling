import { useCallback, useEffect, useState } from "react";
import { api } from "../../lib/api";
import type {
  ClinicalReadiness,
  ClinicalReadinessCategory,
  ClinicalReadinessWarning,
  ClinicalSchedule,
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

export function ClinicalSchedulePage() {
  const [readiness, setReadiness] = useState<ClinicalReadiness | null>(null);
  const [schedule, setSchedule] = useState<ClinicalSchedule | null>(null);
  const [confirmClear, setConfirmClear] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    try {
      const [nextReadiness, nextSchedule] = await Promise.all([api.getClinicalReadiness(), api.getClinicalSchedule()]);
      setReadiness(nextReadiness);
      setSchedule(nextSchedule);
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
      setSchedule(await action());
      setReadiness(await api.getClinicalReadiness());
      setError("");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
    } finally {
      setBusy(false);
    }
  };

  const groups = new Map<string, ClinicalReadinessWarning[]>();
  for (const warning of readiness?.warnings ?? []) {
    groups.set(warning.code, [...(groups.get(warning.code) ?? []), warning]);
  }

  return (
    <section className="clinical-sites" aria-labelledby="clinical-schedule-heading">
      <div className="dialog-heading">
        <h2 id="clinical-schedule-heading">Schedule</h2>
        <div className="dialog-actions">
          <button type="button" onClick={() => void load()}>
            Refresh
          </button>
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
      <p className="muted">
        Generating replaces automatic placements only. Manual editing of placements is not available yet.
      </p>
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
      {schedule?.has_placements && (
        <div aria-label="Schedule summary">
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
          {schedule.unplaced.length > 0 && (
            <ul className="clinical-readiness-list" aria-label="Unplaced students">
              {schedule.unplaced.map((item) => (
                <li key={item.student_id}>
                  <span aria-hidden="true">⚠ </span>
                  {item.student_name}: {UNPLACED_REASONS[item.reason] ?? item.reason}
                </li>
              ))}
            </ul>
          )}
        </div>
      )}
      {error && (
        <div className="error-banner" role="alert">
          {error}
        </div>
      )}
      {readiness && (
        <>
          <div
            className={readiness.is_ready ? "availability-success" : "error-banner"}
            role="status"
            aria-label="Readiness status"
          >
            {readiness.is_ready ? "Ready for scheduling" : "Not ready for scheduling"}
          </div>
          <h3>Readiness Summary</h3>
          <ul className="clinical-readiness-list" aria-label="Readiness summary">
            <li>
              <span aria-hidden="true">{readiness.student_count > 0 ? "✓" : "⚠"} </span>
              {readiness.student_count} {plural(readiness.student_count, "student", "students")} loaded
            </li>
            <li>
              <span aria-hidden="true">{readiness.opportunity_count > 0 ? "✓" : "⚠"} </span>
              {readiness.opportunity_count}{" "}
              {plural(readiness.opportunity_count, "placement opportunity", "placement opportunities")} loaded
            </li>
          </ul>
          <h3>Candidates</h3>
          <ul className="clinical-readiness-list" aria-label="Candidate counts">
            <li>
              {readiness.feasible_candidate_count} feasible student/opportunity{" "}
              {plural(readiness.feasible_candidate_count, "pairing", "pairings")} of {readiness.total_candidate_count}{" "}
              evaluated
            </li>
            <li>
              {readiness.students_with_candidates} of {readiness.student_count}{" "}
              {plural(readiness.student_count, "student has", "students have")} at least one feasible opportunity
            </li>
          </ul>
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
        </>
      )}
    </section>
  );
}
