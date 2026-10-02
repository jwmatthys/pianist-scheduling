import { useEffect, useRef, useState } from "react";
import type { FormEvent } from "react";
import "./App.css";
import { api } from "./lib/api";
import { chooseSessionArchive, saveSessionArchive } from "./lib/platform";
import type { SchedulingSession, SchedulingSessionInput } from "./lib/types";
import { SCHEDULING_MODULES, type ModuleKey } from "./moduleRegistry";
import { ModuleShell } from "./components/ModuleShell";
import { AccompanistModule, type AccompanistModuleHandle } from "./pages/AccompanistModule";
import { DashboardPage } from "./pages/DashboardPage";
import { JurySetupPage } from "./pages/JurySetupPage";

function App() {
  const [activeModule, setActiveModule] = useState<ModuleKey | null>(null);
  const [session, setSession] = useState<SchedulingSession | null>(null);
  const [editingSession, setEditingSession] = useState(false);
  const [sessionFormKey, setSessionFormKey] = useState(0);
  const [busy, setBusy] = useState(false);
  const [sessionError, setSessionError] = useState<string | null>(null);
  const accompanistModuleRef = useRef<AccompanistModuleHandle>(null);
  const newSessionDialogRef = useRef<HTMLDialogElement>(null);

  useEffect(() => {
    void api.getSession().then(setSession).catch((error: unknown) => {
      setSessionError(error instanceof Error ? error.message : "Could not load the active session.");
    });
  }, []);

  async function returnToDashboard() {
    await accompanistModuleRef.current?.saveAvailabilityBeforeLeaving();
    setActiveModule(null);
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
      const nextSession = editingSession
        ? await api.updateSession(data)
        : await api.createSession(data);
      setSession(nextSession);
      setActiveModule(null);
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
      const nextSession = await api.restoreSession(archive);
      setSession(nextSession);
      setActiveModule(null);
    } catch (error) {
      setSessionError(error instanceof Error ? error.message : "Could not open the session archive.");
    } finally {
      setBusy(false);
    }
  }

  const selectedModule = SCHEDULING_MODULES.find((module) => module.key === activeModule) ?? null;

  return (
    <div className="app-shell">
      {sessionError && <div className="error-banner session-error" role="alert">{sessionError}</div>}
      {!selectedModule && (
        <DashboardPage
          session={session}
          busy={busy}
          onEditSession={() => {
            setEditingSession(true);
            setSessionFormKey((key) => key + 1);
            newSessionDialogRef.current?.showModal();
          }}
          onNewSession={() => {
            setEditingSession(false);
            setSessionFormKey((key) => key + 1);
            newSessionDialogRef.current?.showModal();
          }}
          onOpenSession={() => void handleOpenSession()}
          onExportSession={() => void handleExportSession()}
          onOpenModule={setActiveModule}
        />
      )}
      {selectedModule && (
        <ModuleShell module={selectedModule} session={session} onBack={() => void returnToDashboard()}>
          {selectedModule.key === "accompanist" ? (
            <AccompanistModule key={session?.session_uuid} ref={accompanistModuleRef} />
          ) : selectedModule.key === "juries" ? (
            <JurySetupPage key={session?.session_uuid} />
          ) : (
            <section className="module-landing" aria-labelledby="module-landing-heading">
              <h2 id="module-landing-heading">{selectedModule.name}</h2>
              <p>{selectedModule.description}</p>
              <p className="muted">This module shell is in place; its workflow is outside the current release.</p>
            </section>
          )}
        </ModuleShell>
      )}
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
