import { useEffect, useRef, useState } from "react";
import type { FormEvent } from "react";
import "./App.css";
import { api } from "./lib/api";
import { chooseSessionArchive, saveSessionArchive } from "./lib/platform";
import * as startupProfile from "./lib/startupProfile";
import type { LicenseStatus, SchedulingSession, SchedulingSessionInput, SessionArchiveInspection } from "./lib/types";
import { formatUtcTimestampLocally } from "./lib/time";
import { Modal } from "./components/dialogs/Modal";
import { SCHEDULING_MODULES, type ModuleKey } from "./moduleRegistry";
import { ModuleShell } from "./components/ModuleShell";
import { ClinicalModule } from "./pages/ClinicalModule";
import { LessonModule } from "./pages/LessonModule";
import { TutorModule } from "./pages/TutorModule";
import { AccompanistModule, type AccompanistModuleHandle } from "./pages/AccompanistModule";
import { DashboardPage } from "./pages/DashboardPage";
import { JurySetupPage } from "./pages/JurySetupPage";
import { useDismissOnBackdrop } from "./lib/useDismissOnBackdrop";
import { ThemeProvider } from "./theme/ThemeContext";
import { TrialExpiredDialog } from "./components/LicenseDialog";

function AppContent() {
  useDismissOnBackdrop();
  const [activeModule, setActiveModule] = useState<ModuleKey | null>(null);
  const [session, setSession] = useState<SchedulingSession | null>(null);
  const [editingSession, setEditingSession] = useState(false);
  const [sessionFormKey, setSessionFormKey] = useState(0);
  const [busy, setBusy] = useState(false);
  const [sessionError, setSessionError] = useState<string | null>(null);
  const [licenseStatus, setLicenseStatus] = useState<LicenseStatus | null>(null);
  const [trialError, setTrialError] = useState<string | null>(null);
  const [trialClock, setTrialClock] = useState<{ seconds: number; started: number } | null>(null);
  const [clockNow, setClockNow] = useState(() => performance.now());
  const [trialLocked, setTrialLocked] = useState(false);
  const licenseRevision = useRef(0);
  const [olderArchive, setOlderArchive] = useState<{ archive: Uint8Array; inspection: SessionArchiveInspection } | null>(null);
  const accompanistModuleRef = useRef<AccompanistModuleHandle>(null);
  const newSessionDialogRef = useRef<HTMLDialogElement>(null);

  function applyLicenseStatus(status: LicenseStatus) {
    if (status.status !== "licensed" && (
      typeof status.trial_seconds_remaining !== "number"
      || !Number.isFinite(status.trial_seconds_remaining)
      || status.trial_seconds_remaining < 0
    )) {
      throw new Error("The trial time returned by the local service is invalid.");
    }
    const now = performance.now();
    setLicenseStatus(status);
    if (status.status === "licensed") setTrialLocked(false);
    else if (status.trial_seconds_remaining === 0) setTrialLocked(true);
    setClockNow(now);
    setTrialClock(status.status !== "licensed" && status.trial_seconds_remaining !== null
      ? { seconds: status.trial_seconds_remaining, started: now }
      : null);
  }

  const trialRemaining = trialClock === null
    ? null
    : Math.max(0, trialClock.seconds - Math.min(10, Math.floor(Math.max(0, clockNow - trialClock.started) / 1000)));
  const trialExpired = licenseStatus !== null && licenseStatus.status !== "licensed" && (trialLocked || trialRemaining === 0);
  const canUseApp = licenseStatus !== null && !trialExpired && (licenseStatus.status === "licensed" || (trialRemaining !== null && trialRemaining > 0));

  function onLicenseStatus(status: LicenseStatus) {
    licenseRevision.current += 1;
    applyLicenseStatus(status);
    setTrialError(null);
  }

  useEffect(() => {
    if (!trialClock) return;
    const timer = window.setInterval(() => {
      const now = performance.now();
      setClockNow(now);
      if (trialClock.seconds <= Math.min(10, Math.floor(Math.max(0, now - trialClock.started) / 1000))) {
        setTrialLocked(true);
      }
    }, 1000);
    return () => window.clearInterval(timer);
  }, [trialClock]);

  useEffect(() => {
    let active = true;
    let pending = false;
    async function heartbeat() {
      if (pending) return;
      pending = true;
      const revision = licenseRevision.current;
      try {
        const status = await api.trialHeartbeat();
        if (active && revision === licenseRevision.current) {
          applyLicenseStatus(status);
          setTrialError(null);
        }
      } catch (error) {
        if (active && revision === licenseRevision.current) {
          setTrialError(error instanceof Error ? error.message : "Could not update trial time.");
        }
      } finally {
        pending = false;
      }
    }
    void heartbeat();
    const timer = window.setInterval(() => void heartbeat(), 5000);
    return () => {
      active = false;
      window.clearInterval(timer);
    };
  }, []);

  useEffect(() => {
    if (!canUseApp || session) return;
    let active = true;
    startupProfile.mark("Session load begin");
    void api.getSession().then((loaded) => {
      if (!active) return;
      startupProfile.mark("Session load complete");
      setSession(loaded);
      setSessionError(null);
    }).catch((error: unknown) => {
      if (!active) return;
      startupProfile.mark("Session load failed");
      setSessionError(error instanceof Error ? error.message : "Could not load the active session.");
    });
    return () => { active = false; };
  }, [canUseApp, session]);

  if (session) startupProfile.mark("Dashboard render begin");
  useEffect(() => {
    if (!session) return;
    startupProfile.mark("Dashboard render complete");
    const frame = requestAnimationFrame(() => {
      startupProfile.mark("First usable UI rendered");
      startupProfile.logSummary();
    });
    return () => cancelAnimationFrame(frame);
  }, [session]);

  async function returnToDashboard() {
    await accompanistModuleRef.current?.saveAvailabilityBeforeLeaving();
    setActiveModule(null);
  }

  async function handleNewSession(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!editingSession && !window.confirm("Replace the active session? A local recovery snapshot will be created before replacement.")) return;
    const form = new FormData(event.currentTarget);
    const data: SchedulingSessionInput = {
      institution_name: String(form.get("institution_name") ?? "").trim(),
      program_name: String(form.get("program_name") ?? "").trim(),
      term_label: String(form.get("term_label") ?? "").trim(),
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
      const inspection = await api.inspectSessionArchive(archive);
      if (inspection.is_older) {
        setOlderArchive({ archive, inspection });
        return;
      }
      await confirmAndRestore(archive);
    } catch (error) {
      setSessionError(error instanceof Error ? error.message : "Could not open the session archive.");
    } finally {
      setBusy(false);
    }
  }

  async function confirmAndRestore(archive: Uint8Array) {
    if (!window.confirm("Replace the active session with this archive? A local recovery snapshot will be created before replacement.")) return;
    setBusy(true);
    const nextSession = await api.restoreSession(archive);
    setSession(nextSession);
    setActiveModule(null);
  }

  async function continueOpeningOlderArchive() {
    const pending = olderArchive;
    if (!pending) return;
    setOlderArchive(null);
    setSessionError(null);
    try {
      await confirmAndRestore(pending.archive);
    } catch (error) {
      setSessionError(error instanceof Error ? error.message : "Could not open the session archive.");
    } finally {
      setBusy(false);
    }
  }

  const selectedModule = SCHEDULING_MODULES.find((module) => module.key === activeModule) ?? null;

  return (
    <>
    {!licenseStatus && (
      <div className="app-shell" role="status">
        {trialError ? `Could not check license status: ${trialError}` : "Checking license status…"}
      </div>
    )}
    {licenseStatus && (
    <div className="app-shell" inert={trialExpired} aria-hidden={trialExpired || undefined}>
      {sessionError && <div className="error-banner session-error" role="alert">{sessionError}</div>}
      {!selectedModule && (
        <DashboardPage
          session={session}
          busy={busy}
          licenseStatus={licenseStatus}
          trialRemaining={trialRemaining}
          trialError={trialError}
          onLicenseStatus={onLicenseStatus}
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
          ) : selectedModule.key === "clinical" ? (
            <ClinicalModule key={session?.session_uuid} />
          ) : selectedModule.key === "lesson" ? (
            <LessonModule key={session?.session_uuid} />
          ) : selectedModule.key === "tutor" ? (
            <TutorModule key={session?.session_uuid} />
          ) : (
            <section className="module-landing" aria-labelledby="module-landing-heading">
              <h2 id="module-landing-heading">{selectedModule.name}</h2>
              <p>{selectedModule.description}</p>
              <p className="muted">This module shell is in place; its workflow is outside the current release.</p>
            </section>
          )}
        </ModuleShell>
      )}
      {olderArchive && (
        <Modal title="Open Older Session Archive?" onClose={() => setOlderArchive(null)}>
          <p className="dialog-note">
            This session archive was saved earlier than the session currently open in Scheduling Maestro.
          </p>
          <p className="dialog-note">
            Current session:
            <br />
            {formatUtcTimestampLocally(olderArchive.inspection.current_activity_utc ?? "")}
          </p>
          <p className="dialog-note">
            Archive being opened:
            <br />
            {formatUtcTimestampLocally(olderArchive.inspection.archive_exported_utc ?? "")}
          </p>
          <p className="dialog-note">
            You may be intentionally opening an older archive, but recent changes from the currently open session may not be present.
          </p>
          <p className="dialog-note">Do you want to continue opening this archive?</p>
          <div className="dialog-actions">
            <button type="button" onClick={() => setOlderArchive(null)}>Cancel</button>
            <button type="button" onClick={() => void continueOpeningOlderArchive()}>Continue Opening</button>
          </div>
        </Modal>
      )}
      <dialog className="session-dialog" ref={newSessionDialogRef}>
        <form key={sessionFormKey} onSubmit={(event) => void handleNewSession(event)}>
          <div className="dialog-heading">
            <h2>{editingSession ? "Edit Scheduling Session" : "New Scheduling Session"}</h2>
            <button type="button" className="dialog-close" aria-label="Close" onClick={() => newSessionDialogRef.current?.close()}>
              ×
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
          <p className="dialog-note">
            {editingSession
              ? "Session identity and scheduling data will be preserved."
              : "Your current working session will be replaced. A local recovery snapshot is created first."}
          </p>
          <div className="dialog-actions">
            <button type="button" className="secondary-btn" disabled={busy} onClick={() => newSessionDialogRef.current?.close()}>
              Cancel
            </button>
            <button type="submit" className={editingSession ? "" : "primary-btn"} disabled={busy}>
              {busy ? "Working…" : editingSession ? "Update Session" : "Create Session"}
            </button>
          </div>
        </form>
      </dialog>
    </div>
    )}
    {trialExpired && <TrialExpiredDialog onLicenseStatus={onLicenseStatus} />}
    </>
  );
}

export default function App() {
  return (
    <ThemeProvider>
      <AppContent />
    </ThemeProvider>
  );
}
