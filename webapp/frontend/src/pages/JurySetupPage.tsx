import { useEffect, useState } from "react";
import { api } from "../lib/api";
import type {
  AccompanistFinalizationState,
  JuryAvailability,
  JuryAvailableWindow,
  JuryLessonEntry,
  JuryPanel,
  JuryPanelInput,
  JuryReadiness,
} from "../lib/types";
import { JurySchedulePage } from "./JurySchedulePage";
import "./JurySetupPage.css";

type SetupView = "overview" | "lessons" | "panels" | "availability" | "schedule";
type AvailabilityDraft = {
  windows: JuryAvailableWindow[];
  dirty: boolean;
};
type PanelDraft = {
  panel_uuid: string | null;
  panel_name: string;
  room: string;
  jury_date: string;
  earliest_start: string;
  preferred_start: string;
  jury_length: string;
  break_needed: boolean;
  break_every: string;
  break_length: string;
  meal_break: boolean;
  meal_start: string;
  meal_end: string;
};

const VIEWS: { id: SetupView; label: string }[] = [
  { id: "overview", label: "Overview" },
  { id: "panels", label: "Panels" },
  { id: "availability", label: "Pianist availability" },
  { id: "lessons", label: "Lessons" },
  { id: "schedule", label: "Schedule" },
];

const EMPTY_PANEL: PanelDraft = {
  panel_uuid: null,
  panel_name: "",
  room: "",
  jury_date: "",
  earliest_start: "09:00",
  preferred_start: "09:00",
  jury_length: "10",
  break_needed: false,
  break_every: "",
  break_length: "",
  meal_break: false,
  meal_start: "12:00",
  meal_end: "13:00",
};

function availabilityKey(pianistUuid: string, juryDate: string) {
  return `${juryDate}:${pianistUuid}`;
}

function toMinutes(value: string): number | null {
  const match = /^(\d{2}):(\d{2})$/.exec(value);
  if (!match) return null;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (hours > 23 || minutes > 59) return null;
  return hours * 60 + minutes;
}

function fromMinutes(value: number): string {
  const hours = Math.floor(value / 60);
  const minutes = value % 60;
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function formatTime(value: number): string {
  if (value === 0 || value === 1440) return "12:00 AM";
  const hours = Math.floor(value / 60);
  const minutes = value % 60;
  const suffix = hours < 12 ? "AM" : "PM";
  const displayHours = hours % 12 || 12;
  return `${displayHours}:${String(minutes).padStart(2, "0")} ${suffix}`;
}

function formatJuryDate(value: string): string {
  return new Date(`${value}T12:00:00`).toLocaleDateString(undefined, { month: "long", day: "numeric" });
}

function formatPanelTime(panel: JuryPanel) {
  const dateLabel = panel.jury_date
    ? new Date(`${panel.jury_date}T12:00:00`).toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" })
    : "Scheduling Date needed";
  return `${dateLabel} · ${formatTime(panel.earliest_start_minute)} earliest · ${formatTime(panel.preferred_start_minute)} preferred`;
}

function panelDraft(panel?: JuryPanel): PanelDraft {
  if (!panel) return { ...EMPTY_PANEL };
  return {
    panel_uuid: panel.panel_uuid,
    panel_name: panel.panel_name,
    room: panel.room,
    jury_date: panel.jury_date ?? "",
    earliest_start: fromMinutes(panel.earliest_start_minute),
    preferred_start: fromMinutes(panel.preferred_start_minute),
    jury_length: String(panel.jury_length_minutes),
    break_needed: panel.break_needed,
    break_every: panel.break_every_x_juries === null ? "" : String(panel.break_every_x_juries),
    break_length: panel.break_length_minutes === null ? String(panel.jury_length_minutes) : String(panel.break_length_minutes),
    meal_break: panel.meal_break,
    meal_start: panel.meal_start_minute === null ? "12:00" : fromMinutes(panel.meal_start_minute),
    meal_end: panel.meal_end_minute === null ? "13:00" : fromMinutes(panel.meal_end_minute),
  };
}

function errorText(error: unknown) {
  return error instanceof Error ? error.message : "The request could not be completed.";
}

export function JurySetupPage() {
  const [view, setView] = useState<SetupView>("overview");
  const [finalization, setFinalization] = useState<AccompanistFinalizationState | null>(null);
  const [panels, setPanels] = useState<JuryPanel[]>([]);
  const [entries, setEntries] = useState<JuryLessonEntry[]>([]);
  const [readiness, setReadiness] = useState<JuryReadiness | null>(null);
  const [availability, setAvailability] = useState<Record<string, JuryAvailability | null>>({});
  const [availabilityDrafts, setAvailabilityDrafts] = useState<Record<string, AvailabilityDraft>>({});
  const [selectedPianistUuid, setSelectedPianistUuid] = useState<string | null>(null);
  const [selectedAvailabilityDate, setSelectedAvailabilityDate] = useState<string | null>(null);
  const [availabilityStart, setAvailabilityStart] = useState("09:00");
  const [availabilityEnd, setAvailabilityEnd] = useState("12:00");
  const [panelDialog, setPanelDialog] = useState<PanelDraft | null>(null);
  const [entryFilter, setEntryFilter] = useState<"all" | "required">("all");
  const [busy, setBusy] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  const panelDates = [...new Set(panels.map((panel) => panel.jury_date).filter((value): value is string => Boolean(value)))].sort();
  const assignedPianists = [...new Map(
    entries
      .filter((entry) => entry.jury_required && entry.pianist_required && entry.assigned_pianist)
      .map((entry) => [entry.assigned_pianist!.person_uuid, entry.assigned_pianist!] as const),
  ).values()].sort((left, right) => left.display_name.localeCompare(right.display_name));
  const requiredCount = entries.filter((entry) => entry.jury_required).length;
  const unresolvedCount = entries.filter((entry) => entry.jury_required && !entry.panel_uuid).length;
  const selectedPianist = assignedPianists.find((pianist) => pianist.person_uuid === selectedPianistUuid) ?? null;
  const selectedPianistRequiredOnDate = Boolean(selectedPianist && selectedAvailabilityDate && entries.some((entry) => {
    if (!entry.jury_required || !entry.pianist_required || entry.assigned_pianist?.person_uuid !== selectedPianist.person_uuid) {
      return false;
    }
    return panels.some((panel) => panel.panel_uuid === entry.panel_uuid && panel.jury_date === selectedAvailabilityDate);
  }));
  const selectedAvailabilityKey = selectedPianist && selectedAvailabilityDate
    ? availabilityKey(selectedPianist.person_uuid, selectedAvailabilityDate)
    : null;
  const selectedAvailability = selectedAvailabilityKey ? availabilityDrafts[selectedAvailabilityKey] : undefined;

  async function loadAll(synchronizeRoster = false) {
    setLoading(true);
    setError(null);
    try {
      if (synchronizeRoster) await api.synchronizeJuryWithAccompanist();
      const nextFinalization = await api.getAccompanistFinalizationState();
      const nextPanels = await api.getJuryPanels();
      const nextEntries = await api.getJuryEntries();
      const nextReadiness = await api.getJuryReadiness();
      setFinalization(nextFinalization);
      setPanels(nextPanels);
      setEntries(nextEntries);
      setReadiness(nextReadiness);

      const juryPianists = [...new Map(
        nextEntries
          .filter((entry) => entry.jury_required && entry.pianist_required && entry.assigned_pianist)
          .map((entry) => [entry.assigned_pianist!.person_uuid, entry.assigned_pianist!] as const),
      ).values()];
      if (!selectedPianistUuid || !juryPianists.some((pianist) => pianist.person_uuid === selectedPianistUuid)) {
        setSelectedPianistUuid(juryPianists[0]?.person_uuid ?? null);
      }

      const dates = [...new Set(nextPanels.map((panel) => panel.jury_date).filter((value): value is string => Boolean(value)))].sort();
      setSelectedAvailabilityDate((current) => current && dates.includes(current) ? current : dates[0] ?? null);
      const nextAvailability: Record<string, JuryAvailability | null> = {};
      const nextDrafts: Record<string, AvailabilityDraft> = {};
      for (const juryDate of dates) {
        for (const pianist of juryPianists) {
          const key = availabilityKey(pianist.person_uuid, juryDate);
          try {
            const record = await api.getJuryAvailability(pianist.person_uuid, juryDate);
            nextAvailability[key] = record;
            nextDrafts[key] = { windows: record.windows.map((window) => ({ ...window })), dirty: false };
          } catch (requestError) {
            if (!(requestError instanceof Error && requestError.message.startsWith("404:"))) throw requestError;
            nextAvailability[key] = null;
            nextDrafts[key] = { windows: [], dirty: false };
          }
        }
      }
      setAvailability((current) => ({ ...current, ...nextAvailability }));
      setAvailabilityDrafts((current) => {
        const merged = { ...current };
        for (const [key, draft] of Object.entries(nextDrafts)) {
          if (!merged[key]?.dirty) merged[key] = draft;
        }
        return merged;
      });
    } catch (loadError) {
      setError(errorText(loadError));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    void loadAll(true);
  }, []);

  async function refreshReadiness() {
    try {
      setReadiness(await api.getJuryReadiness());
    } catch (requestError) {
      setError(errorText(requestError));
    }
  }

  async function activateView(nextView: SetupView) {
    setView(nextView);
    if (nextView !== "lessons" || loading || busy || !finalization?.current_result_uuid) return;

    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      const result = await api.assignJuryPanelsByInstrument();
      setEntries(result.entries);
      await refreshReadiness();
      if (result.assigned_count > 0) {
        setNotice(`Assigned ${result.assigned_count} lesson${result.assigned_count === 1 ? "" : "s"} to matching Panels.`);
      }
    } catch (requestError) {
      setError(errorText(requestError));
    } finally {
      setBusy(false);
    }
  }

  async function finalizeAndSync() {
    if (!finalization) return;
    const approved = window.confirm(
      "Approve the current Accompanist schedule for downstream Jury use? Existing solver warnings and manual assignments may be included."
    );
    if (!approved) return;
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      await api.finalizeAccompanist(finalization.source_revision);
      await api.synchronizeJuryRoster();
      setNotice("Accompanist schedule finalized and Jury lesson entries synchronized.");
      await loadAll();
    } catch (requestError) {
      setError(errorText(requestError));
      await loadAll();
    } finally {
      setBusy(false);
    }
  }

  async function synchronizeRoster() {
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      const nextEntries = await api.synchronizeJuryRoster();
      setEntries(nextEntries);
      setNotice(`Jury roster synchronized: ${nextEntries.length} lesson entries.`);
      await loadAll();
    } catch (requestError) {
      setError(errorText(requestError));
    } finally {
      setBusy(false);
    }
  }

  async function updateEntry(entry: JuryLessonEntry, panelUuid: string | null) {
    setBusy(true);
    setError(null);
    try {
      const updated = await api.updateJuryEntry(entry.source_lesson_uuid, panelUuid);
      setEntries((current) => current.map((item) =>
        item.source_lesson_uuid === updated.source_lesson_uuid ? updated : item
      ));
      await refreshReadiness();
    } catch (requestError) {
      setError(errorText(requestError));
      await loadAll();
    } finally {
      setBusy(false);
    }
  }

  async function updateJuryRequired(entry: JuryLessonEntry, juryRequired: boolean) {
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      const updated = await api.updateJuryRequired(entry.source_lesson_uuid, juryRequired);
      setEntries((current) => current.map((item) => (
        item.source_lesson_uuid === updated.source_lesson_uuid ? updated : item
      )));
      const [nextFinalization] = await Promise.all([
        api.getAccompanistFinalizationState(),
        refreshReadiness(),
      ]);
      setFinalization(nextFinalization);
      setNotice("Jury Required updated on the source Lesson. Finalize the current Accompanist revision before scheduling.");
    } catch (requestError) {
      setError(errorText(requestError));
    } finally {
      setBusy(false);
    }
  }

  function openNewPanel() {
    setPanelDialog({ ...EMPTY_PANEL });
  }

  function openEditPanel(panel: JuryPanel) {
    setPanelDialog(panelDraft(panel));
  }

  async function savePanel(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!panelDialog) return;
    const earliest = toMinutes(panelDialog.earliest_start);
    const preferred = toMinutes(panelDialog.preferred_start);
    const juryLength = Number(panelDialog.jury_length);
    const mealStart = panelDialog.meal_break ? toMinutes(panelDialog.meal_start) : null;
    const mealEnd = panelDialog.meal_break ? toMinutes(panelDialog.meal_end) : null;
    if (!panelDialog.jury_date) {
      setError("Set a Scheduling Date for this Panel.");
      return;
    }
    if (earliest === null || preferred === null || juryLength <= 0) {
      setError("Enter valid Panel start times and a positive Jury Length.");
      return;
    }
    if (preferred < earliest) {
      setError("Preferred Start must be on or after Earliest Start.");
      return;
    }
    if (panelDialog.break_needed && Number(panelDialog.break_every) <= 0) {
      setError("Break Every X Juries must be a positive whole number.");
      return;
    }
    if (panelDialog.meal_break && (mealStart === null || mealEnd === null || mealEnd <= mealStart)) {
      setError("Meal End must be later than Meal Start.");
      return;
    }
    const payload: JuryPanelInput = {
      panel_name: panelDialog.panel_name.trim(),
      room: panelDialog.room.trim(),
      jury_date: panelDialog.jury_date,
      earliest_start_minute: earliest,
      preferred_start_minute: preferred,
      jury_length_minutes: juryLength,
      break_needed: panelDialog.break_needed,
      break_every_x_juries: panelDialog.break_needed ? Number(panelDialog.break_every) : null,
      break_length_minutes: panelDialog.break_needed
        ? Number(panelDialog.break_length || juryLength)
        : null,
      meal_break: panelDialog.meal_break,
      meal_start_minute: mealStart,
      meal_end_minute: mealEnd,
    };
    setBusy(true);
    setError(null);
    try {
      if (panelDialog.panel_uuid) {
        await api.updateJuryPanel(panelDialog.panel_uuid, payload);
      } else {
        await api.createJuryPanel(payload);
      }
      setPanelDialog(null);
      await loadAll();
    } catch (requestError) {
      setError(errorText(requestError));
    } finally {
      setBusy(false);
    }
  }

  async function removePanel(panel: JuryPanel) {
    if (!window.confirm(`Delete ${panel.panel_name}? Jury entries assigned to it will need another Panel.`)) return;
    setBusy(true);
    setError(null);
    try {
      await api.deleteJuryPanel(panel.panel_uuid);
      await loadAll();
    } catch (requestError) {
      setError(errorText(requestError));
    } finally {
      setBusy(false);
    }
  }

  async function persistAvailabilityWindows(
    pianistUuid: string,
    juryDate: string,
    windows: JuryAvailableWindow[],
  ) {
    const key = availabilityKey(pianistUuid, juryDate);
    setAvailabilityDrafts((current) => ({
      ...current,
      [key]: { windows, dirty: true },
    }));
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      const saved = await api.saveJuryAvailability(pianistUuid, juryDate, { windows });
      setAvailability((current) => ({ ...current, [key]: saved }));
      setAvailabilityDrafts((current) => ({
        ...current,
        [key]: { windows: saved.windows, dirty: false },
      }));
      setNotice("Jury Availability Windows saved.");
      await refreshReadiness();
    } catch (requestError) {
      setError(errorText(requestError));
    } finally {
      setBusy(false);
    }
  }

  async function clearSelectedAvailability() {
    if (!selectedPianist || !selectedAvailabilityDate || !selectedAvailability) return;
    if (selectedAvailability.windows.length === 0) return;
    const dateLabel = formatJuryDate(selectedAvailabilityDate);
    if (!window.confirm(
      `Clear Jury Availability?\n\nThis will remove all Jury Availability Windows for ${selectedPianist.display_name} on ${dateLabel}.`
    )) return;

    const key = availabilityKey(selectedPianist.person_uuid, selectedAvailabilityDate);
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      await api.clearJuryAvailability(selectedPianist.person_uuid, selectedAvailabilityDate);
      setAvailability((current) => ({ ...current, [key]: null }));
      setAvailabilityDrafts((current) => ({
        ...current,
        [key]: { windows: [], dirty: false },
      }));
      setNotice("Jury Availability Windows cleared.");
      await refreshReadiness();
    } catch (requestError) {
      setError(errorText(requestError));
    } finally {
      setBusy(false);
    }
  }

  function addAvailabilityWindow(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!selectedPianist || !selectedAvailabilityDate || !selectedAvailabilityKey) return;
    const start = toMinutes(availabilityStart);
    const end = toMinutes(availabilityEnd);
    if (start === null || end === null || end <= start) {
      setError("Availability Window end must be later than its start.");
      return;
    }
    const windows = selectedAvailability?.windows ?? [];
    if (windows.some((window) => start < window.end_minute && end > window.start_minute)) {
      setError("Jury Availability Windows cannot overlap.");
      return;
    }
    setError(null);
    const nextWindows = [...windows, { start_minute: start, end_minute: end }]
      .sort((left, right) => left.start_minute - right.start_minute);
    void persistAvailabilityWindows(selectedPianist.person_uuid, selectedAvailabilityDate, nextWindows);
  }

  function editAvailabilityWindow(index: number, field: "start_minute" | "end_minute", value: string) {
    if (!selectedAvailabilityKey || !selectedAvailability) return;
    const minute = toMinutes(value);
    if (minute === null) return;
    const nextWindows = selectedAvailability.windows.map((window, windowIndex) => (
      windowIndex === index ? { ...window, [field]: minute } : window
    ));
    setAvailabilityDrafts((current) => ({
      ...current,
      [selectedAvailabilityKey]: { windows: nextWindows, dirty: true },
    }));
  }

  function saveEditedAvailabilityWindow() {
    if (!selectedPianist || !selectedAvailabilityDate || !selectedAvailability) return;
    if (!selectedAvailability.dirty) return;
    const nextWindows = [...selectedAvailability.windows].sort((left, right) => left.start_minute - right.start_minute);
    if (nextWindows.some((window) => window.end_minute <= window.start_minute)) {
      setError("Availability Window end must be later than its start.");
      return;
    }
    if (nextWindows.some((window, index) => index > 0 && window.start_minute < nextWindows[index - 1].end_minute)) {
      setError("Jury Availability Windows cannot overlap.");
      return;
    }
    setError(null);
    void persistAvailabilityWindows(selectedPianist.person_uuid, selectedAvailabilityDate, nextWindows);
  }

  const visibleEntries = entries.filter((entry) => entryFilter === "all" || entry.jury_required);

  return (
    <section className="jury-setup" aria-labelledby="jury-setup-title">
      <div className="jury-setup-heading">
        <div>
          <p className="jury-eyebrow">PERFORMANCE JURIES · SETUP</p>
          <h2 id="jury-setup-title">Jury setup</h2>
        </div>
        <span className="jury-scope-note">Jury schedule</span>
      </div>

      {error && <div className="jury-alert jury-alert-error" role="alert">{error}</div>}
      {notice && <div className="jury-alert jury-alert-success" role="status">{notice}</div>}

      <section className={`jury-readiness-strip ${readiness?.ready ? "is-ready" : "is-blocked"}`} aria-live="polite">
        <div className="jury-readiness-state">
          <span className="jury-readiness-mark" aria-hidden="true">{readiness?.ready ? "✓" : "!"}</span>
          <div>
            <strong>{readiness?.ready ? "Ready for scheduling" : "Readiness needs attention"}</strong>
            <span>
              {readiness
                ? `${readiness.issues.filter((issue) => issue.severity === "error").length} blocking · ${readiness.issues.filter((issue) => issue.severity === "warning").length} warnings`
                : "Checking current setup"}
            </span>
          </div>
        </div>
        <div className="jury-readiness-meta">
          <span>{requiredCount} required lesson{requiredCount === 1 ? "" : "s"}</span>
          <span>{unresolvedCount} without Panel</span>
          <button type="button" className="jury-text-button" disabled={busy || loading} onClick={() => void refreshReadiness()}>
            Recheck
          </button>
        </div>
      </section>

      <nav className="jury-view-nav" aria-label="Jury setup views">
        {VIEWS.map((item) => (
          <button
            key={item.id}
            type="button"
            className={view === item.id ? "is-active" : ""}
            aria-current={view === item.id ? "page" : undefined}
            disabled={loading || busy}
            onClick={() => void activateView(item.id)}
          >
            {item.label}
            {item.id === "lessons" && <span>{entries.length}</span>}
            {item.id === "panels" && <span>{panels.length}</span>}
          </button>
        ))}
      </nav>

      {loading ? (
        <div className="jury-loading" role="status">Loading Jury setup…</div>
      ) : (
        <>
          {view === "overview" && (
            <div className="jury-overview-grid">
              <section className="jury-section jury-date-section">
                <div className="jury-section-heading">
                  <div>
                    <h3>Panel dates</h3>
                    <p>Each Panel has its own one-day schedule date.</p>
                  </div>
                </div>
                {panels.length === 0 ? (
                  <p className="jury-field-note">Add a Panel and set its Scheduling Date before collecting availability.</p>
                ) : (
                  <ul className="jury-panel-date-list">
                    {panels.map((panel) => (
                      <li key={panel.panel_uuid}>
                        <strong>{panel.panel_name}</strong>
                        {" "}
                        <span>{panel.jury_date ? new Date(`${panel.jury_date}T12:00:00`).toLocaleDateString(undefined, { dateStyle: "long" }) : "Date required"}</span>
                      </li>
                    ))}
                  </ul>
                )}
              </section>

              <section className="jury-section jury-source-section">
                <div className="jury-section-heading">
                  <div>
                    <h3>Accompanist source</h3>
                    <p>Jury entries use only the human-approved finalized lesson result.</p>
                  </div>
                </div>
                {finalization?.current_result_uuid ? (
                  <div className="jury-source-current">
                    <span className="jury-state-dot" aria-hidden="true" />
                    <div>
                      <strong>Finalized result v{finalization.current_result_version}</strong>
                      <span>Source revision {finalization.source_revision}</span>
                    </div>
                    <button type="button" className="secondary-btn" disabled={busy} onClick={() => void synchronizeRoster()}>
                      Sync roster
                    </button>
                  </div>
                ) : (
                  <div className="jury-source-empty">
                    <strong>No current finalized result</strong>
                    <p>Review the manual Accompanist schedule before approving it for Jury use.</p>
                    <button type="button" className="primary-btn" disabled={busy} onClick={() => void finalizeAndSync()}>
                      {busy ? "Working…" : `Finalize revision ${finalization?.source_revision ?? "—"} & sync roster`}
                    </button>
                  </div>
                )}
                <p className="jury-field-note">Finalization records your approved assignments, including ordinary solver warnings and manual overrides.</p>
              </section>

              <section className="jury-section jury-readiness-detail">
                <div className="jury-section-heading">
                  <div>
                    <h3>Readiness issues</h3>
                    <p>Blocking issues must be resolved before any future schedule generation.</p>
                  </div>
                </div>
                {readiness?.issues.length ? (
                  <ul className="jury-issue-list">
                    {readiness.issues.map((issue, index) => (
                      <li key={`${issue.code}-${issue.entity_uuids.join("-")}-${index}`} className={`severity-${issue.severity}`}>
                        <span>{issue.severity === "error" ? "Blocking" : "Warning"}</span>
                        <div>
                          <strong>{issue.code.replaceAll("_", " ").toLowerCase()}</strong>
                          <p>{issue.message}</p>
                          {issue.entity_uuids.length > 0 && (
                            <small>{issue.entity_uuids.map((uuid) => uuid.slice(0, 8)).join(" · ")}</small>
                          )}
                        </div>
                      </li>
                    ))}
                  </ul>
                ) : (
                  <div className="jury-clear-state"><span aria-hidden="true">✓</span> No readiness issues.</div>
                )}
              </section>
            </div>
          )}

          {view === "lessons" && (
            <section className="jury-section jury-entries-section">
              <div className="jury-section-heading jury-entries-heading">
                <div>
                  <h3>Lesson entries</h3>
                  <p>Each finalized Accompanist lesson has independent Jury participation and Panel selection.</p>
                </div>
                <div className="jury-filter" role="group" aria-label="Filter lesson entries">
                  <button type="button" className={entryFilter === "all" ? "is-active" : ""} onClick={() => setEntryFilter("all")}>All</button>
                  <button type="button" className={entryFilter === "required" ? "is-active" : ""} onClick={() => setEntryFilter("required")}>Jury required</button>
                </div>
              </div>
              {entries.length === 0 ? (
                <div className="jury-empty-state">
                  <strong>Roster not synchronized</strong>
                  <p>Finalize the Accompanist schedule, then sync its lesson entries here.</p>
                  {!finalization?.current_result_uuid && (
                    <button type="button" className="primary-btn" disabled={busy} onClick={() => void finalizeAndSync()}>
                      Finalize Accompanist schedule
                    </button>
                  )}
                </div>
              ) : (
                <div className="jury-table-wrap">
                  <table className="jury-table">
                    <thead>
                      <tr>
                        <th scope="col">Student</th>
                        <th scope="col">Lesson / instrument</th>
                        <th scope="col">Teacher</th>
                        <th scope="col">Pianist</th>
                        <th scope="col">Jury Required</th>
                        <th scope="col">Panel</th>
                      </tr>
                    </thead>
                    <tbody>
                      {visibleEntries.map((entry) => (
                        <tr key={entry.source_lesson_uuid} className={entry.jury_required ? "jury-entry-required" : ""}>
                          <td>
                            <strong>{entry.student_display_name || "Unnamed student"}</strong>
                          </td>
                          <td>
                            <span>{entry.instrument || "Instrument not supplied"}</span>
                            <small>{entry.pianist_required ? "Needs pianist? Yes" : "Needs pianist? No"}</small>
                          </td>
                          <td>{entry.teacher || <span className="jury-muted">Not supplied</span>}</td>
                          <td>
                            {entry.assigned_pianist?.display_name ?? (
                              entry.pianist_required
                                ? <span className="jury-missing-value">Unassigned</span>
                                : <span className="jury-muted">Not required</span>
                            )}
                          </td>
                          <td>
                            <label className="jury-entry-flag">
                              <input
                                type="checkbox"
                                checked={entry.jury_required}
                                disabled={busy}
                                aria-label={`Jury Required for ${entry.student_display_name}, ${entry.instrument}`}
                                onChange={(event) => void updateJuryRequired(entry, event.target.checked)}
                              />
                              <span>{entry.jury_required ? "Required" : "Not required"}</span>
                            </label>
                          </td>
                          <td>
                            {entry.jury_required ? (
                              <select
                                aria-label={`Panel for ${entry.student_display_name}, ${entry.instrument}`}
                                value={entry.panel_uuid ?? ""}
                                disabled={busy}
                                onChange={(event) => void updateEntry(entry, event.target.value || null)}
                              >
                                <option value="">Select Panel</option>
                                {panels.map((panel) => <option key={panel.panel_uuid} value={panel.panel_uuid}>{panel.panel_name}</option>)}
                              </select>
                            ) : <span className="jury-muted">Not required</span>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          )}

          {view === "panels" && (
            <section className="jury-section jury-panels-section">
              <div className="jury-section-heading jury-entries-heading">
                <div>
                  <h3>Panels</h3>
                  <p>Define start preferences, Jury duration, periodic breaks, and meal intervals.</p>
                </div>
                <button type="button" className="primary-btn" onClick={openNewPanel}>Add Panel</button>
              </div>
              {panels.length === 0 ? (
                <div className="jury-empty-state">
                  <strong>No Panels defined</strong>
                  <p>Panels are assigned explicitly to individual lesson entries.</p>
                </div>
              ) : (
                <div className="jury-panel-list">
                  {panels.map((panel) => (
                    <article className="jury-panel-row" key={panel.panel_uuid}>
                      <div className="jury-panel-main">
                        <div className="jury-panel-name-line">
                          <h4>{panel.panel_name}</h4>
                          <span className="jury-room-tag">{panel.room || "Room not specified"}</span>
                        </div>
                        <p>{formatPanelTime(panel)}</p>
                        <div className="jury-panel-details">
                          <span>{panel.jury_length_minutes} min Jury</span>
                          {panel.break_needed && <span>Break every {panel.break_every_x_juries} · {panel.break_length_minutes} min</span>}
                          {panel.meal_break && panel.meal_start_minute !== null && panel.meal_end_minute !== null && (
                            <span>Meal {formatTime(panel.meal_start_minute)}–{formatTime(panel.meal_end_minute)}</span>
                          )}
                        </div>
                      </div>
                      <div className="jury-panel-actions">
                        <button type="button" className="secondary-btn" disabled={busy} onClick={() => openEditPanel(panel)}>Edit</button>
                        <button type="button" className="jury-delete-button" disabled={busy} aria-label={`Delete ${panel.panel_name}`} onClick={() => void removePanel(panel)}>Delete</button>
                      </div>
                    </article>
                  ))}
                </div>
              )}
            </section>
          )}

          {view === "availability" && (
            <section className="jury-availability-layout">
              <aside className="jury-section jury-pianist-picker">
                <div className="jury-section-heading">
                  <div>
                    <h3>Pianists</h3>
                    <p>{selectedAvailabilityDate ? formatJuryDate(selectedAvailabilityDate) : "Select a Panel date"}</p>
                  </div>
                </div>
                {assignedPianists.length === 0 ? (
                  <p className="jury-empty-copy">No assigned pianists in the current Jury roster.</p>
                ) : (
                  <ul>
                    {assignedPianists.map((pianist) => {
                      const key = selectedAvailabilityDate ? availabilityKey(pianist.person_uuid, selectedAvailabilityDate) : "";
                      const record = key ? availability[key] : null;
                      const draft = key ? availabilityDrafts[key] : undefined;
                      const windows = draft?.windows ?? record?.windows ?? [];
                      return (
                        <li key={pianist.person_uuid}>
                          <button
                            type="button"
                            className={selectedPianistUuid === pianist.person_uuid ? "is-selected" : ""}
                            onClick={() => setSelectedPianistUuid(pianist.person_uuid)}
                          >
                            <strong>{pianist.display_name}</strong>
                            <span className="jury-pianist-window-summary">
                              {!selectedAvailabilityDate ? "Select a Panel date"
                                : windows.length
                                  ? windows.map((item) => (
                                    <span key={`${item.start_minute}-${item.end_minute}`}>
                                      {formatTime(item.start_minute)} - {formatTime(item.end_minute)}
                                    </span>
                                  ))
                                  : "No Jury Availability Windows"}
                            </span>
                          </button>
                        </li>
                      );
                    })}
                  </ul>
                )}
              </aside>

              <section className="jury-section jury-availability-editor">
                {panelDates.length === 0 ? (
                  <div className="jury-empty-state"><strong>Set a Scheduling Date on a Panel first</strong><p>Each Panel specifies its own date; enter Availability Windows for each date.</p></div>
                ) : !selectedAvailabilityDate ? (
                  <div className="jury-empty-state"><strong>Select a Panel date</strong><p>Availability is stored independently for each date used by a Panel.</p></div>
                ) : !selectedPianist ? (
                  <div className="jury-empty-state"><strong>No pianist selected</strong><p>Assigned pianists from the finalized result appear here.</p></div>
                ) : selectedAvailability ? (
                  <>
                    <div className="jury-section-heading">
                      <div>
                        <h3>Jury Availability Windows</h3>
                        <p>{selectedPianist.display_name}</p>
                      </div>
                      <button
                        type="button"
                        className="jury-delete-button"
                        disabled={busy || selectedAvailability.windows.length === 0}
                        onClick={() => void clearSelectedAvailability()}
                      >Clear Availability</button>
                    </div>
                    <label className="jury-date-control">
                      <span>Panel Scheduling Date</span>
                      <select value={selectedAvailabilityDate} onChange={(event) => setSelectedAvailabilityDate(event.target.value)}>
                        {panelDates.map((juryDateValue) => (
                          <option key={juryDateValue} value={juryDateValue}>
                            {new Date(`${juryDateValue}T12:00:00`).toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric", year: "numeric" })}
                          </option>
                        ))}
                      </select>
                    </label>
                    {!selectedPianistRequiredOnDate && (
                      <p className="jury-field-note">No Jury-required lesson assigns this Pianist to a Panel on this date; Availability Windows are not required.</p>
                    )}
                    <form className="jury-window-form" onSubmit={addAvailabilityWindow}>
                      <label>Available from<input type="time" value={availabilityStart} onChange={(event) => setAvailabilityStart(event.target.value)} required /></label>
                      <label>Available until<input type="time" value={availabilityEnd} onChange={(event) => setAvailabilityEnd(event.target.value)} required /></label>
                      <button type="submit" className="secondary-btn" disabled={busy}>Add Availability Window</button>
                    </form>
                    {selectedAvailability.windows.length > 0 ? (
                      <ul className="jury-window-list">
                        {selectedAvailability.windows.map((window, index) => (
                          <li key={index}>
                            <span className="jury-window-swatch" aria-hidden="true" />
                              <label className="jury-window-time">From<input aria-label={`Availability Window ${index + 1} start`} type="time" value={fromMinutes(window.start_minute)} disabled={busy} onChange={(event) => editAvailabilityWindow(index, "start_minute", event.target.value)} onBlur={saveEditedAvailabilityWindow} /></label>
                              <label className="jury-window-time">Until<input aria-label={`Availability Window ${index + 1} end`} type="time" value={fromMinutes(window.end_minute)} disabled={busy} onChange={(event) => editAvailabilityWindow(index, "end_minute", event.target.value)} onBlur={saveEditedAvailabilityWindow} /></label>
                            <button
                              type="button"
                              className="jury-delete-button"
                              aria-label={`Remove Available window ${index + 1}`}
                              disabled={busy}
                              onClick={() => void persistAvailabilityWindows(
                                selectedPianist.person_uuid,
                                selectedAvailabilityDate,
                                selectedAvailability.windows.filter((_, itemIndex) => itemIndex !== index),
                              )}
                            >Remove</button>
                          </li>
                        ))}
                      </ul>
                    ) : (
                      <p className="jury-empty-copy">No Jury Availability Windows</p>
                    )}
                    <p className="jury-field-note">Each Panel has its own Scheduling Date. Availability Windows apply only to the selected date; times outside these windows are Unavailable.</p>
                  </>
                ) : (
                  <div className="jury-empty-state" role="status">Loading Jury Availability Windows…</div>
                )}
              </section>
            </section>
          )}

          {view === "schedule" && (
            <JurySchedulePage readiness={readiness} onReadinessChange={setReadiness} />
          )}
        </>
      )}

      {panelDialog && (
        <dialog className="jury-panel-dialog" open aria-labelledby="jury-panel-dialog-title">
          <form onSubmit={(event) => void savePanel(event)}>
            <div className="jury-dialog-heading">
              <div><p className="jury-eyebrow">PANEL SETTINGS</p><h2 id="jury-panel-dialog-title">{panelDialog.panel_uuid ? "Edit Panel" : "Add Panel"}</h2></div>
              <button type="button" className="jury-icon-button" aria-label="Close Panel editor" onClick={() => setPanelDialog(null)}>×</button>
            </div>
            <div className="jury-panel-form-grid">
              <label className="jury-form-wide">Scheduling Date<input type="date" required value={panelDialog.jury_date} onChange={(event) => setPanelDialog({ ...panelDialog, jury_date: event.target.value })} /></label>
              <label className="jury-form-wide">Panel Name<input required maxLength={200} value={panelDialog.panel_name} onChange={(event) => setPanelDialog({ ...panelDialog, panel_name: event.target.value })} /></label>
              <label className="jury-form-wide">Room <span>Descriptive only</span><input maxLength={200} value={panelDialog.room} onChange={(event) => setPanelDialog({ ...panelDialog, room: event.target.value })} /></label>
              <label>Earliest Start<input type="time" value={panelDialog.earliest_start} required onChange={(event) => setPanelDialog({ ...panelDialog, earliest_start: event.target.value })} /></label>
              <label>Preferred Start<input type="time" value={panelDialog.preferred_start} required onChange={(event) => setPanelDialog({ ...panelDialog, preferred_start: event.target.value })} /></label>
              <label className="jury-form-wide">Jury Length <span>Minutes</span><input type="number" min="1" max="1440" step="1" required value={panelDialog.jury_length} onChange={(event) => setPanelDialog({ ...panelDialog, jury_length: event.target.value })} /></label>
              <label className="jury-check-field jury-form-wide">
                <input
                  type="checkbox"
                  checked={panelDialog.break_needed}
                  onChange={(event) => setPanelDialog({
                    ...panelDialog,
                    break_needed: event.target.checked,
                    break_length: event.target.checked ? panelDialog.jury_length : "",
                  })}
                />
                <span>Periodic break needed</span>
              </label>
              {panelDialog.break_needed && <>
                <label>Break Every X Juries<input type="number" min="1" step="1" required value={panelDialog.break_every} onChange={(event) => setPanelDialog({ ...panelDialog, break_every: event.target.value })} /></label>
                <label>Break Length <span>Minutes</span><input type="number" min="1" step="1" required value={panelDialog.break_length} onChange={(event) => setPanelDialog({ ...panelDialog, break_length: event.target.value })} /></label>
              </>}
              <label className="jury-check-field jury-form-wide"><input type="checkbox" checked={panelDialog.meal_break} onChange={(event) => setPanelDialog({ ...panelDialog, meal_break: event.target.checked })} /><span>Meal break needed</span></label>
              {panelDialog.meal_break && <>
                <label>Meal Start<input type="time" required value={panelDialog.meal_start} onChange={(event) => setPanelDialog({ ...panelDialog, meal_start: event.target.value })} /></label>
                <label>Meal End<input type="time" required value={panelDialog.meal_end} onChange={(event) => setPanelDialog({ ...panelDialog, meal_end: event.target.value })} /></label>
              </>}
            </div>
            <div className="jury-dialog-actions">
              <button type="button" className="secondary-btn" disabled={busy} onClick={() => setPanelDialog(null)}>Cancel</button>
              <button type="submit" className="primary-btn" disabled={busy}>{busy ? "Saving…" : "Save Panel"}</button>
            </div>
          </form>
        </dialog>
      )}
    </section>
  );
}