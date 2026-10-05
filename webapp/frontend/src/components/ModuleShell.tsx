import { useEffect, useRef, type ReactNode } from "react";
import { formatSessionIdentity, type SchedulingModule } from "../moduleRegistry";
import type { SchedulingSession } from "../lib/types";

type Props = {
  module: SchedulingModule;
  session: SchedulingSession | null;
  onBack: () => void;
  children: ReactNode;
};

export function ModuleShell({ module, session, onBack, children }: Props) {
  const headingRef = useRef<HTMLHeadingElement>(null);

  useEffect(() => {
    headingRef.current?.focus();
  }, [module.key]);

  return (
    <div className="module-shell">
      <header className="module-header">
        <button
          type="button"
          className="module-back-button"
          aria-label="Home"
          onClick={onBack}
        >
          <span aria-hidden="true">←</span>
          <span>Home</span>
        </button>
        <div className="module-heading-block">
          <h1 ref={headingRef} tabIndex={-1}>{module.name}</h1>
          <p>{formatSessionIdentity(session)}</p>
        </div>
        {module.availability === "shell" && <span className="module-shell-label">Module shell</span>}
      </header>
      <main className="app-main module-main">{children}</main>
    </div>
  );
}