import { forwardRef, useImperativeHandle, useRef, useState } from "react";
import { ImportPage } from "./ImportPage";
import { PianistsPage, type PianistsPageHandle } from "./PianistsPage";
import { SchedulePage } from "./SchedulePage";
import { ReportsPage } from "./ReportsPage";

type Tab = "import" | "pianists" | "schedule" | "reports";

const TABS: { id: Tab; label: string }[] = [
  { id: "import", label: "1. Import Lessons" },
  { id: "pianists", label: "2. Pianists & Availability" },
  { id: "schedule", label: "3. Schedule" },
  { id: "reports", label: "4. Reports" },
];

export type AccompanistModuleHandle = {
  saveAvailabilityBeforeLeaving: () => Promise<void>;
};

export const AccompanistModule = forwardRef<AccompanistModuleHandle>(function AccompanistModule(_, ref) {
  const [tab, setTab] = useState<Tab>("import");
  const pianistsPageRef = useRef<PianistsPageHandle>(null);

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
        {tab === "import" && <ImportPage onImported={() => setTab("schedule")} />}
        {tab === "pianists" && <PianistsPage ref={pianistsPageRef} />}
        {tab === "schedule" && <SchedulePage />}
        {tab === "reports" && <ReportsPage />}
      </div>
    </>
  );
});