import { useEffect, useState } from "react";
import { api } from "../lib/api";
import { formatUtcTimestampLocally } from "../lib/time";
import type {
  JuryReadiness,
  JuryScheduleEvent,
  JuryScheduleResult,
} from "../lib/types";

type Props = {
  readiness: JuryReadiness | null;
  onReadinessChange: (readiness: JuryReadiness) => void;
};

function errorText(error: unknown) {
  return error instanceof Error ? error.message : "The request could not be completed.";
}

function formatTime(value: number) {
  if (value === 0 || value === 1440) return "12:00 AM";
  const hour = Math.floor(value / 60);
  const minute = value % 60;
  return `${hour % 12 || 12}:${String(minute).padStart(2, "0")} ${hour < 12 ? "AM" : "PM"}`;
}

function formatDate(value: string) {
  return new Date(`${value}T12:00:00`).toLocaleDateString(undefined, {
    month: "long",
    day: "numeric",
    year: "numeric",
  });
}

function titleCase(value: string) {
  return value.toLowerCase().replaceAll("_", " ").replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function resultLessonCount(result: JuryScheduleResult) {
  return result.unscheduled_lessons.length + result.panel_timelines.reduce(
    (count, panel) => count + panel.events.filter((event) => event.kind === "jury").length,
    0,
  );
}

function eventTitle(event: JuryScheduleEvent) {
  if (event.kind === "meal_break") return "MEAL BREAK";
  if (event.kind === "periodic_break") return "BREAK";
  return event.student_display_name || "Unnamed student";
}

export function JurySchedulePage({ readiness, onReadinessChange }: Props) {
  const [currentResult, setCurrentResult] = useState<JuryScheduleResult | null>(null);
  const [loading, setLoading] = useState(true);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function refreshResults() {
    try {
      setCurrentResult(await api.getCurrentJurySchedule());
      setError(null);
    } catch (cause) {
      if (cause instanceof Error && cause.message.startsWith("404:")) {
      setCurrentResult(null);
        setError(null);
      } else {
        setError(errorText(cause));
      }
    }
  }

  useEffect(() => {
    setError(null);
    void refreshResults()
      .catch((cause) => setError(errorText(cause)))
      .finally(() => setLoading(false));
  }, []);

  async function generateSchedule() {
    if (!readiness?.ready || generating) return;
    setGenerating(true);
    setError(null);
    try {
      const freshReadiness = await api.getJuryReadiness();
      onReadinessChange(freshReadiness);
      if (!freshReadiness.ready) return;
      await api.generateJurySchedule(freshReadiness.jury_input_revision);
      await refreshResults();
    } catch (cause) {
      setError(errorText(cause));
      try {
        onReadinessChange(await api.getJuryReadiness());
      } catch {
        // Keep the generation error visible if the readiness refresh also fails.
      }
    } finally {
      setGenerating(false);
    }
  }

  const blockers = readiness?.issues.filter((issue) => issue.severity === "error") ?? [];
  const warnings = readiness?.issues.filter((issue) => issue.severity === "warning") ?? [];

  return (
    <div className="jury-schedule-page">
      <section className="jury-section jury-generate-schedule" aria-labelledby="generate-jury-schedule-title">
        <div className="jury-section-heading">
          <div>
            <h3 id="generate-jury-schedule-title">Generate Jury Schedule</h3>
            <p>{readiness
              ? `${blockers.length} blocker${blockers.length === 1 ? "" : "s"} · ${warnings.length} warning${warnings.length === 1 ? "" : "s"}`
              : "Checking readiness…"}</p>
          </div>
        </div>
        {blockers.length > 0 && (
          <div className="jury-schedule-issues jury-schedule-blockers" aria-label="Schedule blockers">
            <ul>
              {blockers.map((issue, index) => (
                <li key={`${issue.code}-${index}`}>
                  <strong>{titleCase(issue.code)}</strong>
                  <span>{issue.message}</span>
                </li>
              ))}
            </ul>
          </div>
        )}
        {warnings.length > 0 && (
          <div className="jury-schedule-issues jury-schedule-warnings" aria-label="Schedule warnings">
            <ul>
              {warnings.map((issue, index) => (
                <li key={`${issue.code}-${index}`}>
                  <strong>{titleCase(issue.code)}</strong>
                  <span>{issue.message}</span>
                </li>
              ))}
            </ul>
          </div>
        )}
        {readiness && blockers.length === 0 && warnings.length === 0 && (
          <p className="jury-clear-state">No readiness blockers or warnings.</p>
        )}
        <button
          type="button"
          className="primary-btn"
          disabled={!readiness?.ready || blockers.length > 0 || generating || loading}
          onClick={() => void generateSchedule()}
        >
          {generating ? "Generating…" : "Generate Schedule"}
        </button>
      </section>

      {error && <div className="jury-alert jury-alert-error" role="alert">{error}</div>}
      {generating && <div className="jury-loading" role="status">Generating Jury schedule…</div>}
      {loading ? (
        <div className="jury-loading" role="status">Loading saved Jury schedules…</div>
      ) : currentResult ? (
        <ScheduleDetails result={currentResult} />
      ) : (
        <section className="jury-section jury-empty-state">
          <strong>No Jury schedule generated yet</strong>
          <p>A current schedule will appear here after successful generation.</p>
        </section>
      )}
    </div>
  );
}

function ScheduleDetails({
  result,
}: {
  result: JuryScheduleResult;
}) {
  const panelNames = new Map(result.panel_timelines.map((timeline) => [timeline.panel_uuid, timeline.panel_name]));
  return (
    <section className="jury-section jury-schedule-result" aria-label="Current generated schedule">
      <div className="jury-section-heading jury-schedule-result-heading">
        <div>
          <h3>Current generated schedule</h3>
          <p>Generated {formatUtcTimestampLocally(result.created_at)}</p>
        </div>
      </div>

      {result.stale && (
        <div className="jury-schedule-stale" role="status">
          This schedule uses older source data. {result.stale_reasons.map(titleCase).join(" · ")}
        </div>
      )}

      {resultLessonCount(result) === 0 && (
        <p className="jury-empty-copy">No Jury-required lessons were included in this result.</p>
      )}

      {result.warnings.length > 0 && (
        <ul className="jury-schedule-warning-list" aria-label="Schedule warnings">
          {result.warnings.map((warning, index) => (
            <li key={`${warning.code}-${index}`} className={`severity-${warning.severity}`}>
              <strong>{titleCase(warning.code)}</strong>
              <span>{warning.message}</span>
            </li>
          ))}
        </ul>
      )}

      {result.panel_timelines.map((panel) => (
        <section className="jury-schedule-panel" key={panel.panel_uuid}>
          <div className="jury-schedule-panel-heading">
            <h4>{panel.panel_name}</h4>
            <span>{formatDate(panel.jury_date)}</span>
          </div>
          {panel.events.length > 0 ? (
            <div className="jury-table-wrap">
              <table className="jury-table jury-schedule-table">
                <thead>
                  <tr>
                    <th scope="col">Start</th>
                    <th scope="col">End</th>
                    <th scope="col">Student / entry</th>
                    <th scope="col">Assigned Pianist</th>
                  </tr>
                </thead>
                <tbody>
                  {panel.events.map((event, index) => (
                    <tr key={event.kind === "jury" ? event.source_lesson_uuid : `${event.kind}-${event.start_minute}-${index}`}>
                      <td>{formatTime(event.start_minute)}</td>
                      <td>{formatTime(event.end_minute)}</td>
                      <td className="jury-schedule-student-cell">
                        <strong>{eventTitle(event)}</strong>
                        {event.kind === "jury" && <small className="jury-schedule-instrument">{event.instrument || "Instrument not supplied"}</small>}
                      </td>
                      <td>{event.kind === "jury" ? event.pianist_display_name || "—" : "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="jury-empty-copy">No scheduled entries for this Panel.</p>
          )}
        </section>
      ))}

      <section className="jury-unscheduled-section" aria-label="Unscheduled lessons">
        <div className="jury-section-heading">
          <div>
            <h4>Unscheduled lessons</h4>
            <p>{result.unscheduled_lessons.length} lesson{result.unscheduled_lessons.length === 1 ? "" : "s"} need attention.</p>
          </div>
        </div>
        {result.unscheduled_lessons.length > 0 ? (
          <div className="jury-table-wrap">
            <table className="jury-table jury-schedule-table">
              <thead>
                <tr>
                  <th scope="col">Student</th>
                  <th scope="col">Panel</th>
                  <th scope="col">Reason</th>
                  <th scope="col">Explanation</th>
                </tr>
              </thead>
              <tbody>
                {result.unscheduled_lessons.map((lesson) => (
                  <tr key={lesson.source_lesson_uuid}>
                    <td><strong>{lesson.student_display_name || "Unnamed student"}</strong><small>{lesson.instrument}</small></td>
                    <td>{panelNames.get(lesson.panel_uuid) ?? "Panel unavailable"}<small>{formatDate(lesson.jury_date)}</small></td>
                    <td>{titleCase(lesson.reason_code)}</td>
                    <td>{lesson.explanation}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : resultLessonCount(result) > 0 ? (
          <p className="jury-clear-state">All Jury-required lessons have a scheduled time.</p>
        ) : null}
      </section>
    </section>
  );
}

