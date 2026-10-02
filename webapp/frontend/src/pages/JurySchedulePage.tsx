import { useEffect, useState } from "react";
import { api } from "../lib/api";
import type {
  JuryReadiness,
  JuryScheduleEvent,
  JuryScheduleHistoryItem,
  JuryScheduleResult,
  JuryScheduleState,
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

function formatTimestamp(value: string) {
  return new Date(value).toLocaleString();
}

function titleCase(value: string) {
  return value.toLowerCase().replaceAll("_", " ").replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function lifecycleTone(state: JuryScheduleState) {
  if (state === "finalized") return "is-finalized";
  if (state === "superseded") return "is-superseded";
  return "is-draft";
}

function resultLessonCount(result: JuryScheduleResult) {
  return result.unscheduled_lessons.length + result.panel_timelines.reduce(
    (count, panel) => count + panel.events.filter((event) => event.kind === "jury").length,
    0,
  );
}

function eventTitle(event: JuryScheduleEvent) {
  if (event.kind === "meal_break") return "Meal Break";
  if (event.kind === "periodic_break") return `Periodic Break · after ${event.after_jury_count} juries`;
  return event.student_display_name || "Unnamed student";
}

function ReadinessPanel({ readiness }: { readiness: JuryReadiness | null }) {
  if (!readiness) {
    return <section className="jury-section jury-schedule-readiness" aria-label="Schedule readiness">
      <h3>Schedule readiness</h3>
      <p className="jury-muted">Readiness is loading.</p>
    </section>;
  }

  const blockers = readiness.issues.filter((issue) => issue.severity === "error");
  const warnings = readiness.issues.filter((issue) => issue.severity === "warning");
  return (
    <section className="jury-section jury-schedule-readiness" aria-label="Schedule readiness">
      <div className="jury-section-heading">
        <div>
          <h3>Schedule readiness</h3>
          <p>{blockers.length} blocking · {warnings.length} warnings</p>
        </div>
        <span className={`jury-schedule-status ${readiness.ready ? "is-current" : "is-stale"}`}>
          {readiness.ready ? "Ready" : "Blocked"}
        </span>
      </div>
      {blockers.length > 0 && (
        <div className="jury-schedule-issues jury-schedule-blockers" aria-label="Blocking readiness issues">
          <h4>Blockers</h4>
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
        <div className="jury-schedule-issues jury-schedule-warnings" aria-label="Readiness warnings">
          <h4>Warnings</h4>
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
      {blockers.length === 0 && warnings.length === 0 && (
        <p className="jury-clear-state">No readiness blockers or warnings.</p>
      )}
    </section>
  );
}

export function JurySchedulePage({ readiness, onReadinessChange }: Props) {
  const [currentResult, setCurrentResult] = useState<JuryScheduleResult | null>(null);
  const [history, setHistory] = useState<JuryScheduleHistoryItem[]>([]);
  const [historicalResult, setHistoricalResult] = useState<JuryScheduleResult | null>(null);
  const [loading, setLoading] = useState(true);
  const [historyLoading, setHistoryLoading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function refreshResults() {
    const [currentResponse, historyResponse] = await Promise.allSettled([
      api.getCurrentJurySchedule(),
      api.getJuryScheduleHistory(),
    ]);
    const errors: string[] = [];
    if (currentResponse.status === "fulfilled") {
      setCurrentResult(currentResponse.value);
    } else if (currentResponse.reason instanceof Error && currentResponse.reason.message.startsWith("404:")) {
      setCurrentResult(null);
    } else {
      errors.push(errorText(currentResponse.reason));
    }
    if (historyResponse.status === "fulfilled") {
      setHistory(historyResponse.value);
    } else {
      errors.push(errorText(historyResponse.reason));
    }
    setError(errors.length ? errors.join(" ") : null);
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
      setHistoricalResult(null);
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

  async function viewHistoricalResult(resultUuid: string) {
    setHistoryLoading(true);
    setError(null);
    try {
      setHistoricalResult(await api.getJuryScheduleResult(resultUuid));
    } catch (cause) {
      setError(errorText(cause));
    } finally {
      setHistoryLoading(false);
    }
  }

  const blockers = readiness?.issues.filter((issue) => issue.severity === "error") ?? [];

  return (
    <div className="jury-schedule-page">
      <ReadinessPanel readiness={readiness} />

      <section className="jury-section jury-schedule-controls">
        <div>
          <h3>Generate schedule</h3>
          <p>{readiness?.ready ? "Readiness checks pass. Generate a new Jury schedule." : "Resolve all readiness blockers before generation."}</p>
        </div>
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
        <ScheduleDetails
          result={currentResult}
          heading="Current generated schedule"
        />
      ) : (
        <section className="jury-section jury-empty-state">
          <strong>No Jury schedule generated yet</strong>
          <p>A current schedule will appear here after successful generation.</p>
        </section>
      )}

      <section className="jury-section jury-schedule-history" aria-label="Schedule history">
        <div className="jury-section-heading">
          <div>
            <h3>Schedule history</h3>
            <p>Saved results retain their original Accompanist and Jury input revisions.</p>
          </div>
        </div>
        {history.length > 0 ? (
          <ul className="jury-schedule-history-list">
            {history.map((item) => (
              <li key={item.result_uuid}>
                <button
                  type="button"
                  className="jury-schedule-history-item"
                  disabled={historyLoading}
                  onClick={() => void viewHistoricalResult(item.result_uuid)}
                >
                  <strong>Result v{item.result_version}</strong>
                  <span>{formatTimestamp(item.created_at)}</span>
                  <span className="jury-schedule-history-meta">
                    Source revision {item.source_revision} · {item.scheduled_count} scheduled · {item.unscheduled_count} unscheduled
                  </span>
                  <span className="jury-schedule-history-status">
                    <span className={`jury-schedule-status ${lifecycleTone(item.state)}`}>{titleCase(item.state)}</span>
                    {item.stale && <span className="jury-schedule-status is-stale">Stale</span>}
                  </span>
                </button>
              </li>
            ))}
          </ul>
        ) : (
          <p className="jury-empty-copy">No generated schedules in this session.</p>
        )}
      </section>

      {historicalResult && (
        <ScheduleDetails
          result={historicalResult}
          heading={`Saved schedule v${historicalResult.result_version}`}
        />
      )}
    </div>
  );
}

function ScheduleDetails({
  result,
  heading,
}: {
  result: JuryScheduleResult;
  heading: string;
}) {
  const panelNames = new Map(result.panel_timelines.map((timeline) => [timeline.panel_uuid, timeline.panel_name]));
  return (
    <section className="jury-section jury-schedule-result" aria-label={heading}>
      <div className="jury-section-heading jury-schedule-result-heading">
        <div>
          <h3>{heading}</h3>
          <p>Result v{result.result_version} · Generated {formatTimestamp(result.created_at)} · Accompanist source revision {result.source_revision}</p>
        </div>
        <span className="jury-schedule-result-status">
          <span className={`jury-schedule-status ${lifecycleTone(result.state)}`}>{titleCase(result.state)}</span>
          {result.stale && <span className="jury-schedule-status is-stale">Stale</span>}
        </span>
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
                      <td>
                        <strong>{eventTitle(event)}</strong>
                        {event.kind === "jury" && <small>{event.instrument || "Instrument not supplied"}</small>}
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

