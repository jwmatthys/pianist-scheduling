import { useRef, useState } from "react";
import "./App.css";
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
  const pianistsPageRef = useRef<PianistsPageHandle>(null);

  async function changeTab(nextTab: Tab) {
    if (tab === "pianists" && nextTab !== "pianists") {
      await pianistsPageRef.current?.saveAvailabilityBeforeLeaving();
    }
    setTab(nextTab);
  }

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1>🎹 Pianist Scheduling</h1>
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
        {tab === "import" && <ImportPage onImported={() => {}} />}
        {tab === "pianists" && <PianistsPage ref={pianistsPageRef} />}
        {tab === "schedule" && <SchedulePage />}
        {tab === "reports" && <ReportsPage />}
      </main>
    </div>
  );
}

export default App;
