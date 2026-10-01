import { useEffect, useRef } from "react";
import { SCHEDULING_MODULES, formatSessionIdentity, type ModuleKey } from "../moduleRegistry";
import type { SchedulingSession } from "../lib/types";

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

  return (
    <main className="app-dashboard">
      <header className="dashboard-header">
        <h1 ref={headingRef} tabIndex={-1}>Music Program Scheduler</h1>
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
        <nav className="module-card-grid" aria-label="Scheduling modules">
          {SCHEDULING_MODULES.map((module) => (
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
          ))}
        </nav>
      </section>
    </main>
  );
}