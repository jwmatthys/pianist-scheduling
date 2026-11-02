import { useEffect, useRef } from "react";
import { SCHEDULING_MODULES, formatSessionIdentity, type ModuleKey } from "../moduleRegistry";
import type { SchedulingSession } from "../lib/types";
import { ThemeDialog } from "../components/ThemeDialog";

const INDEPENDENT_MODULE_KEYS: ModuleKey[] = ["lesson", "tutor", "clinical"];

type Props = {
  session: SchedulingSession | null;
  busy: boolean;
  onEditSession: () => void;
  onNewSession: () => void;
  onOpenSession: () => void;
  onExportSession: () => void;
  onOpenModule: (key: ModuleKey) => void;
};

export function DashboardPage({
  session,
  busy,
  onEditSession,
  onNewSession,
  onOpenSession,
  onExportSession,
  onOpenModule,
}: Props) {
  const headingRef = useRef<HTMLHeadingElement>(null);

  useEffect(() => {
    headingRef.current?.focus();
  }, []);

  function renderCard(key: ModuleKey) {
    const module = SCHEDULING_MODULES.find((candidate) => candidate.key === key);
    if (!module) return null;
    return (
      <button
        type="button"
        key={module.key}
        className={`module-card ${module.availability === "shell" ? "module-card-shell" : ""}`}
        onClick={() => onOpenModule(module.key)}
      >
        <span className="module-card-title">{module.name}</span>
        <span className="module-card-description">{module.description}</span>
        <span className="module-card-action">
          {module.availability === "available" ? "Open module" : "View module shell"}
        </span>
      </button>
    );
  }

  return (
    <main className="app-dashboard">
      <ThemeDialog />
      <header className="dashboard-header">
        <h1 ref={headingRef} tabIndex={-1}>Scheduling Maestro</h1>
      </header>

      <section className="dashboard-session" aria-labelledby="active-session-heading">
        <div className="dashboard-session-identity">
          <p className="dashboard-eyebrow">Active Scheduling Session</p>
          <h2 id="active-session-heading">{session?.institution_name ?? "Loading session"}</h2>
          <p className="dashboard-session-context">
            {session
              ? `${session.program_name} · ${formatSessionIdentity(session).split(" · ").at(-1)}`
              : "Session details are loading."}
          </p>
        </div>
        <nav className="dashboard-session-actions" aria-label="Session actions">
          <button type="button" className="secondary-btn" disabled={busy || !session} onClick={onEditSession}>
            Edit Session
          </button>
          <button type="button" className="secondary-btn" disabled={busy} onClick={onNewSession}>
            New Session
          </button>
          <button type="button" className="secondary-btn" disabled={busy} onClick={onOpenSession}>
            Open Session
          </button>
          <button type="button" className="secondary-btn" disabled={busy || !session} onClick={onExportSession}>
            Export Session
          </button>
        </nav>
      </section>

      <section className="dashboard-modules" aria-labelledby="modules-heading">
        <div className="dashboard-section-heading">
          <h2 id="modules-heading">Scheduling Modules</h2>
        </div>
        <div className="module-layout">
          <nav className="module-card-grid" aria-label="Independent scheduling modules">
            {INDEPENDENT_MODULE_KEYS.map(renderCard)}
          </nav>
          <nav className="module-dependency-row" aria-label="Accompanist to jury scheduling workflow">
            {renderCard("accompanist")}
            <div className="module-dependency" role="note">
              <svg className="module-dependency-arrow" viewBox="0 0 64 12" aria-hidden="true" focusable="false">
                <path d="M2 6h54" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                <path d="m50 1 8 5-8 5" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
              <span className="module-dependency-caption">Uses finalized accompanist assignments</span>
            </div>
            {renderCard("juries")}
          </nav>
        </div>
      </section>
    </main>
  );
}