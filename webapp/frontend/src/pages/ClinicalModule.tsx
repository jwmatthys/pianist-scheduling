import { useState } from "react";
import { ClinicalReportsPage } from "./clinical/ClinicalReportsPage";
import { ClinicalSchedulePage } from "./clinical/ClinicalSchedulePage";
import { ClinicalSitesPage } from "./clinical/ClinicalSitesPage";
import { ClinicalStudentRosterPage } from "./clinical/ClinicalStudentRosterPage";

type Tab = "sites" | "roster" | "schedule" | "reports";

const TABS: { id: Tab; label: string }[] = [
  { id: "sites", label: "Clinical Sites" },
  { id: "roster", label: "Student Roster" },
  { id: "schedule", label: "Schedule" },
  { id: "reports", label: "Reports" },
];

export function ClinicalModule() {
  const [tab, setTab] = useState<Tab>("sites");

  return (
    <>
      <nav className="tab-nav module-tab-nav" aria-label="Clinical Placements views">
        {TABS.map((item) => (
          <button
            key={item.id}
            type="button"
            className={`tab-button ${tab === item.id ? "active" : ""}`}
            aria-current={tab === item.id ? "page" : undefined}
            onClick={() => setTab(item.id)}
          >
            {item.label}
          </button>
        ))}
      </nav>
      <div className="app-main clinical-main">
        {tab === "sites" && <ClinicalSitesPage />}
        {tab === "roster" && <ClinicalStudentRosterPage />}
        {tab === "schedule" && <ClinicalSchedulePage />}
        {tab === "reports" && <ClinicalReportsPage />}
      </div>
    </>
  );
}
