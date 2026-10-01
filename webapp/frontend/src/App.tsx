import { useEffect, useRef, useState } from "react";
import type { FormEvent } from "react";
import "./App.css";
import { api } from "./lib/api";
import { chooseSessionArchive, saveSessionArchive } from "./lib/platform";
import type { SchedulingSession, SchedulingSessionInput } from "./lib/types";
import { ImportPage } from "./pages/ImportPage";
import { PianistsPage, type PianistsPageHandle } from "./pages/PianistsPage";
import { SchedulePage } from "./pages/SchedulePage";
import { ReportsPage } from "./pages/ReportsPage";

type Tab = "import" | "pianists" | "schedule" | "reports";

const TABS: { id: Tab; label: string }[] = [
  { id: "import", label: "1. Import Lessons" },
  { id: "pianists", label: "2. Pianists & Availability" },
  { id: "schedule", label: "3. Schedule" },
  { id: "reports", label: "4. Reports" },
];

function App() {
  const [tab, setTab] = useState<Tab>("import");
  const [session, setSession] = useState<SchedulingSession | null>(null);
  const [editingSession, setEditingSession] = useState(false);
  const [sessionFormKey, setSessionFormKey] = useState(0);
  const [busy, setBusy] = useState(false);
  const [sessionError, setSessionError] = useState<string | null>(null);
  const pianistsPageRef = useRef<PianistsPageHandle>(null);
  const newSessionDialogRef = useRef<HTMLDialogElement>(null);

  useEffect(() => {
    void api.getSession().then(setSession).catch((error: unknown) => {
      setSessionError(error instanceof Error ? error.message : "Could not load the active session.");
    });
  }, []);

  async function changeTab(nextTab: Tab) {
    if (tab === "pianists" && nextTab !== "pianists") {
      await pianistsPageRef.current?.saveAvailabilityBeforeLeaving();
    }
    setTab(nextTab);
  }

  async function handleNewSession(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!editingSession && !window.confirm("Replace the active session? A local recovery snapshot will be created before replacement.")) return;
    const form = new FormData(event.currentTarget);
    const year = String(form.get("year") ?? "").trim();
    const data: SchedulingSessionInput = {
      institution_name: String(form.get("institution_name") ?? "").trim(),
      program_name: String(form.get("program_name") ?? "").trim(),
      term_label: String(form.get("term_label") ?? "").trim(),
      year: year ? Number(year) : null,
      start_date: String(form.get("start_date") ?? "") || null,
      end_date: String(form.get("end_date") ?? "") || null,
    };
    setBusy(true);
    setSessionError(null);
    try {
      if (tab === "pianists") await pianistsPageRef.current?.saveAvailabilityBeforeLeaving();
      const nextSession = editingSession
        ? await api.updateSession(data)
        : await api.createSession(data);
      setSession(nextSession);
      setTab("import");
      newSessionDialogRef.current?.close();
    } catch (error) {
      setSessionError(error instanceof Error ? error.message : "Could not create the new session.");
    } finally {
      setBusy(false);
    }
  }

  async function handleExportSession() {
    setBusy(true);
    setSessionError(null);
    try {
      if (tab === "pianists") await pianistsPageRef.current?.saveAvailabilityBeforeLeaving();
      await saveSessionArchive(await api.exportSession());
    } catch (error) {
      setSessionError(error instanceof Error ? error.message : "Could not export the session.");
    } finally {
      setBusy(false);
    }
  }

  async function handleOpenSession() {
    setSessionError(null);
    try {
      const archive = await chooseSessionArchive();
      if (!archive) return;
      if (!window.confirm("Replace the active session with this archive? A local recovery snapshot will be created before replacement.")) return;
      setBusy(true);
      if (tab === "pianists") await pianistsPageRef.current?.saveAvailabilityBeforeLeaving();
      const nextSession = await api.restoreSession(archive);
      setSession(nextSession);
      setTab("import");
    } catch (error) {
      setSessionError(error instanceof Error ? error.message : "Could not open the session archive.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="app-shell">
      <header className="app-header">
        <div className="product-header-row">
          <div>
            <h1>Music Program Scheduler</h1>
            <div className="app-module-label">Accompanist Scheduling</div>
          </div>
          <div className="session-header-actions">
            <div className="session-identity" aria-live="polite">
              {session
                ? `${session.institution_name} | ${session.program_name} | ${session.term_label}`
                : "Loading active session"}
            </div>
            <div className="session-action-buttons">
                <button className="header-action" disabled={busy || !session} onClick={() => {
                  setEditingSession(true);
                  setSessionFormKey((key) => key + 1);
                  newSessionDialogRef.current?.showModal();
                }}>
                  Edit Session
                </button>
                <button className="header-action" disabled={busy} onClick={() => {
                  setEditingSession(false);
                  setSessionFormKey((key) => key + 1);
                  newSessionDialogRef.current?.showModal();
                }}>
                New Session
              </button>
              <button className="header-action" disabled={busy} onClick={() => void handleExportSession()}>
                Export Session
              </button>
              <button className="header-action" disabled={busy} onClick={() => void handleOpenSession()}>
                Open Session
              </button>
            </div>
          </div>
        </div>
        <nav className="tab-nav">
          {TABS.map((t) => (
            <button
              key={t.id}
              className={`tab-button ${tab === t.id ? "active" : ""}`}
              onClick={() => changeTab(t.id)}
            >
              {t.label}
            </button>
          ))}
        </nav>
      </header>
      <main className="app-main">
        {sessionError && <div className="error-banner session-error" role="alert">{sessionError}</div>}
        {tab === "import" && <ImportPage key={session?.session_uuid} onImported={() => {}} />}
        {tab === "pianists" && <PianistsPage key={session?.session_uuid} ref={pianistsPageRef} />}
        {tab === "schedule" && <SchedulePage key={session?.session_uuid} />}
        {tab === "reports" && <ReportsPage key={session?.session_uuid} />}
      </main>
      <dialog className="session-dialog" ref={newSessionDialogRef}>
        <form key={sessionFormKey} onSubmit={(event) => void handleNewSession(event)}>
          <div className="dialog-heading">
            <h2>{editingSession ? "Edit Scheduling Session" : "New Scheduling Session"}</h2>
            <button type="button" className="dialog-close" aria-label="Close" onClick={() => newSessionDialogRef.current?.close()}>
              Close
            </button>
          </div>
          <label>
            Institution
            <input name="institution_name" required maxLength={200} defaultValue={editingSession ? session?.institution_name : undefined} />
          </label>
          <label>
            Program
            <input name="program_name" required maxLength={200} defaultValue={editingSession ? session?.program_name : undefined} />
          </label>
          <label>
            Term label
            <input name="term_label" required maxLength={100} placeholder="Fall 2027" defaultValue={editingSession ? session?.term_label : undefined} />
          </label>
          <label>
            Year <span>(optional)</span>
            <input name="year" type="number" min={1000} max={9999} defaultValue={editingSession ? session?.year ?? undefined : undefined} />
          </label>
          <div className="session-date-fields">
            <label>
              Start date <span>(optional)</span>
              <input name="start_date" type="date" defaultValue={editingSession ? session?.start_date ?? "" : ""} />
            </label>
            <label>
              End date <span>(optional)</span>
              <input name="end_date" type="date" defaultValue={editingSession ? session?.end_date ?? "" : ""} />
            </label>
          </div>
          <p className="dialog-note">
            {editingSession
              ? "Session identity and scheduling data will be preserved."
              : "Your current working session will be replaced. A local recovery snapshot is created first."}
          </p>
          <div className="dialog-actions">
            <button type="button" className="secondary-btn" disabled={busy} onClick={() => newSessionDialogRef.current?.close()}>
              Cancel
            </button>
            <button type="submit" className="primary-btn" disabled={busy}>
              {busy ? "Working…" : editingSession ? "Update Session" : "Create Session"}
            </button>
          </div>
        </form>
      </dialog>
    </div>
  );
}

export default App;
