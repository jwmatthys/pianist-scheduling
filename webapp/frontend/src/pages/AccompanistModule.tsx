import { forwardRef, useImperativeHandle, useRef, useState } from "react";
import { ImportPage } from "./ImportPage";
import { PianistsPage, type PianistsPageHandle } from "./PianistsPage";
import { SchedulePage } from "./SchedulePage";
import { ReportsPage } from "./ReportsPage";

type Tab = "pianists" | "schedule" | "reports";

const TABS: { id: Tab; label: string }[] = [
  { id: "pianists", label: "Pianist Availability" },
  { id: "schedule", label: "Lesson Roster" },
  { id: "reports", label: "Reports" },
];

export type AccompanistModuleHandle = {
  saveAvailabilityBeforeLeaving: () => Promise<void>;
};

export const AccompanistModule = forwardRef<AccompanistModuleHandle>(function AccompanistModule(_, ref) {
  const [tab, setTab] = useState<Tab>("schedule");
  const [rosterRevision, setRosterRevision] = useState(0);
  const pianistsPageRef = useRef<PianistsPageHandle>(null);
  const lessonImportDialogRef = useRef<HTMLDialogElement>(null);

  useImperativeHandle(ref, () => ({
    saveAvailabilityBeforeLeaving: async () => {
      if (tab === "pianists") await pianistsPageRef.current?.saveAvailabilityBeforeLeaving();
    },
  }), [tab]);

  async function changeTab(nextTab: Tab) {
    if (tab === "pianists" && nextTab !== "pianists") {
      await pianistsPageRef.current?.saveAvailabilityBeforeLeaving();
    }
    setTab(nextTab);
  }

  return (
    <>
      <nav className="tab-nav module-tab-nav" aria-label="Accompanist Scheduling views">
        {TABS.map((item) => (
          <button
            key={item.id}
            type="button"
            className={`tab-button ${tab === item.id ? "active" : ""}`}
            aria-current={tab === item.id ? "page" : undefined}
            onClick={() => void changeTab(item.id)}
          >
            {item.label}
          </button>
        ))}
      </nav>
      <div className="app-main accompanist-main">
        {tab === "pianists" && <PianistsPage ref={pianistsPageRef} />}
        {tab === "schedule" && (
          <SchedulePage
            key={rosterRevision}
            onImportLessons={() => lessonImportDialogRef.current?.showModal()}
          />
        )}
        {tab === "reports" && <ReportsPage />}
      </div>
      <dialog className="availability-import-dialog" ref={lessonImportDialogRef}>
        <div className="availability-import-content">
          <header className="availability-import-heading">
            <div>
              <h2>Import Lessons</h2>
              <p className="muted">CSV, XLSX, or XLS</p>
            </div>
            <button type="button" className="dialog-close" aria-label="Close" onClick={() => lessonImportDialogRef.current?.close()}>
              ×
            </button>
          </header>
          <ImportPage onImported={() => {
            setRosterRevision((revision) => revision + 1);
            lessonImportDialogRef.current?.close();
          }} />
        </div>
      </dialog>
    </>
  );
});