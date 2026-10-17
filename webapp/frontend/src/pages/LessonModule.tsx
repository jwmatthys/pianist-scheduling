import { useState } from "react";
import { LessonInstructorRosterPage } from "./lesson/LessonInstructorRosterPage";
import { LessonReportsPage } from "./lesson/LessonReportsPage";
import { LessonSchedulePage } from "./lesson/LessonSchedulePage";
import { LessonStudentRosterPage } from "./lesson/LessonStudentRosterPage";

type Tab = "instructors" | "students" | "schedule" | "reports";

const TABS: { id: Tab; label: string }[] = [
  { id: "instructors", label: "Instructor Roster" },
  { id: "students", label: "Student Roster" },
  { id: "schedule", label: "Schedule" },
  { id: "reports", label: "Reports" },
];

export function LessonModule() {
  const [tab, setTab] = useState<Tab>("instructors");

  return (
    <>
      <nav className="tab-nav module-tab-nav" aria-label="Lesson Scheduling views">
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
        {tab === "instructors" && <LessonInstructorRosterPage />}
        {tab === "students" && <LessonStudentRosterPage />}
        {tab === "schedule" && <LessonSchedulePage />}
        {tab === "reports" && <LessonReportsPage />}
      </div>
    </>
  );
}
