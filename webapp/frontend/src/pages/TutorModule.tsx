import { useState } from "react";
import { TutorTutorRosterPage } from "./tutor/TutorTutorRosterPage";
import { TutorReportsPage } from "./tutor/TutorReportsPage";
import { TutorSchedulePage } from "./tutor/TutorSchedulePage";
import { TutorStudentRosterPage } from "./tutor/TutorStudentRosterPage";

type Tab = "tutors" | "students" | "schedule" | "reports";

const TABS: { id: Tab; label: string }[] = [
  { id: "tutors", label: "Tutor Roster" },
  { id: "students", label: "Student Roster" },
  { id: "schedule", label: "Schedule" },
  { id: "reports", label: "Reports" },
];

export function TutorModule() {
  const [tab, setTab] = useState<Tab>("tutors");

  return (
    <>
      <nav className="tab-nav module-tab-nav" aria-label="Tutor Scheduling views">
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
        {tab === "tutors" && <TutorTutorRosterPage />}
        {tab === "students" && <TutorStudentRosterPage />}
        {tab === "schedule" && <TutorSchedulePage />}
        {tab === "reports" && <TutorReportsPage />}
      </div>
    </>
  );
}
