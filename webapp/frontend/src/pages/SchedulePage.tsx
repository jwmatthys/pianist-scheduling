import { useEffect, useMemo, useState } from "react";
import { api } from "../lib/api";
import type { Lesson, Pianist } from "../lib/types";
import { inputValueToMinutes, minutesToInputValue } from "../lib/time";

const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

function fitClass(lesson: Lesson): string {
  if (lesson.assigned_pianist_id === null) return "row-unassigned";
  if (lesson.notes.includes("CONFLICT") || lesson.notes.includes("double-booked")) return "row-conflict";
  if (lesson.notes.includes("OVER CAP")) return "row-overcap";
  if (lesson.fit_quality === "Near" || lesson.fit_quality === "Overlap" || lesson.notes.includes("TENTATIVE"))
    return "row-warn";
  if (lesson.fit_quality === "None") return "row-warn";
  return "";
}

export function SchedulePage() {
  const [lessons, setLessons] = useState<Lesson[]>([]);
  const [pianists, setPianists] = useState<Pianist[]>([]);
  const [hoursByPianist, setHoursByPianist] = useState<Record<string, number>>({});
  const [conflicts, setConflicts] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<string | null>(null);

  async function refresh() {
    const [lessonList, pianistList] = await Promise.all([api.listLessons(), api.listPianists()]);
    setLessons(lessonList);
    setPianists(pianistList);
    const validation = await api.validate();
    setLessons(validation.lessons);
    setHoursByPianist(validation.hours_by_pianist);
    setConflicts(validation.conflicts);
  }

  useEffect(() => {
    refresh();
  }, []);

  async function runAlgorithm() {
    setBusy(true);
    setMessage(null);
    try {
      const result = await api.runAssignment();
      setLessons(result.lessons);
      setHoursByPianist(result.hours_by_pianist);
      setConflicts(result.conflicts);
      setMessage(
        `Assigned ${result.lessons.length - result.unassigned_count} of ${result.lessons.length} lessons.` +
          (result.unassigned_count ? ` ${result.unassigned_count} still need attention.` : "")
      );
    } catch (e: any) {
      setMessage(e.message ?? String(e));
    } finally {
      setBusy(false);
    }
  }

  async function patchLesson(id: number, data: Partial<Lesson> & { clear_assigned_pianist?: boolean }) {
    await api.updateLesson(id, data);
    const validation = await api.validate();
    setLessons(validation.lessons);
    setHoursByPianist(validation.hours_by_pianist);
    setConflicts(validation.conflicts);
  }

  async function addLesson() {
    await api.createLesson({
      teacher: "",
      student: "New Student",
      day: "Monday",
      start_minute: 9 * 60,
      end_minute: 9 * 60 + 50,
      room: "",
      instrument: "",
      need_pianist: true,
    });
    refresh();
  }

  async function removeLesson(id: number) {
    if (!confirm("Delete this lesson?")) return;
    await api.deleteLesson(id);
    refresh();
  }

  const sortedLessons = useMemo(
    () =>
      [...lessons].sort(
        (a, b) => DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || a.start_minute - b.start_minute
      ),
    [lessons]
  );

  return (
    <div className="page schedule-page">
      <div className="schedule-toolbar">
        <h2>Schedule</h2>
        <button className="primary-btn" onClick={runAlgorithm} disabled={busy}>
          {busy ? "Assigning\u2026" : "Run best-fit assignment"}
        </button>
        <button className="secondary-btn" onClick={addLesson}>
          + Add lesson
        </button>
        {message && <span className="toolbar-message">{message}</span>}
      </div>

      {conflicts.length > 0 && (
        <div className="conflicts-banner">
          <strong>Conflicts detected:</strong>
          <ul>
            {conflicts.map((c, i) => (
              <li key={i}>{c}</li>
            ))}
          </ul>
        </div>
      )}

      <div className="workload-summary">
        {pianists.map((p) => {
          const hours = hoursByPianist[p.name] ?? 0;
          const overCap = p.max_hours_per_week != null && hours > p.max_hours_per_week;
          return (
            <div key={p.id} className={`workload-chip ${overCap ? "over-cap" : ""}`}>
              <strong>{p.name}</strong>
              <span>
                {hours.toFixed(2)}h{p.max_hours_per_week ? ` / ${p.max_hours_per_week}h` : ""}
              </span>
            </div>
          );
        })}
      </div>

      <div className="schedule-table-wrap">
        <table className="schedule-table">
          <thead>
            <tr>
              <th>Day</th>
              <th>Start</th>
              <th>End</th>
              <th>Teacher</th>
              <th>Student</th>
              <th>Room</th>
              <th>Instrument</th>
              <th>Required</th>
              <th>Assigned pianist</th>
              <th>Hours</th>
              <th>Fit</th>
              <th>Notes</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {sortedLessons.map((lesson) => (
              <tr key={lesson.id} className={fitClass(lesson)}>
                <td>
                  <select
                    value={lesson.day}
                    onChange={(e) => patchLesson(lesson.id, { day: e.target.value })}
                  >
                    {DAYS.map((d) => (
                      <option key={d} value={d}>
                        {d}
                      </option>
                    ))}
                  </select>
                </td>
                <td>
                  <input
                    type="time"
                    value={minutesToInputValue(lesson.start_minute)}
                    onChange={(e) =>
                      patchLesson(lesson.id, { start_minute: inputValueToMinutes(e.target.value) })
                    }
                  />
                </td>
                <td>
                  <input
                    type="time"
                    value={minutesToInputValue(lesson.end_minute)}
                    onChange={(e) =>
                      patchLesson(lesson.id, { end_minute: inputValueToMinutes(e.target.value) })
                    }
                  />
                </td>
                <td>
                  <input
                    defaultValue={lesson.teacher}
                    onBlur={(e) => e.target.value !== lesson.teacher && patchLesson(lesson.id, { teacher: e.target.value })}
                  />
                </td>
                <td>
                  <input
                    defaultValue={lesson.student}
                    onBlur={(e) => e.target.value !== lesson.student && patchLesson(lesson.id, { student: e.target.value })}
                  />
                </td>
                <td>
                  <input
                    defaultValue={lesson.room}
                    onBlur={(e) => e.target.value !== lesson.room && patchLesson(lesson.id, { room: e.target.value })}
                  />
                </td>
                <td>
                  <input
                    defaultValue={lesson.instrument}
                    onBlur={(e) =>
                      e.target.value !== lesson.instrument && patchLesson(lesson.id, { instrument: e.target.value })
                    }
                  />
                </td>
                <td>
                  <input
                    defaultValue={lesson.required_pianist_name}
                    placeholder="(none)"
                    onBlur={(e) =>
                      e.target.value !== lesson.required_pianist_name &&
                      patchLesson(lesson.id, { required_pianist_name: e.target.value })
                    }
                  />
                </td>
                <td>
                  <select
                    value={lesson.assigned_pianist_id ?? ""}
                    onChange={(e) =>
                      e.target.value
                        ? patchLesson(lesson.id, { assigned_pianist_id: Number(e.target.value) })
                        : patchLesson(lesson.id, { clear_assigned_pianist: true })
                    }
                  >
                    <option value="">-- unassigned --</option>
                    {pianists.map((p) => (
                      <option key={p.id} value={p.id}>
                        {p.name}
                      </option>
                    ))}
                  </select>
                </td>
                <td>{lesson.hours.toFixed(2)}</td>
                <td>
                  <span className={`fit-badge fit-${lesson.fit_quality.toLowerCase()}`}>
                    {lesson.fit_quality || "\u2014"}
                  </span>
                </td>
                <td className="notes-cell" title={lesson.notes}>
                  {lesson.notes}
                </td>
                <td>
                  <button className="danger-btn small" onClick={() => removeLesson(lesson.id)}>
                    &times;
                  </button>
                </td>
              </tr>
            ))}
            {sortedLessons.length === 0 && (
              <tr>
                <td colSpan={13} className="muted">
                  No lessons yet. Import a file or add one manually.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
