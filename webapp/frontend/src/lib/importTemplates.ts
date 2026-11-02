import type { ReportCell } from "./reportTable";

export type ImportTemplateKey =
  | "accompanistLessons"
  | "pianistAvailability"
  | "juryPanels"
  | "clinicalSites"
  | "clinicalStudents"
  | "lessonStudents"
  | "lessonInstructors"
  | "tutorStudents"
  | "tutorTutors";

export interface ImportTemplate {
  /** Base file name without extension. */
  fileName: string;
  description: string;
  sheetName: string;
  columns: string[];
  rows: ReportCell[][];
}

const WEEKDAY_COLUMNS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const WEEKDAY_SAMPLE = ["8:00 AM-12:00 PM", "T 1:00 PM-4:00 PM", "", "9:00 AM-11:00 AM; T 2:00 PM-3:00 PM", "", "", ""];

// Headers use the product's field labels, which the importers recognize when suggesting a column mapping.
export const IMPORT_TEMPLATES: Record<ImportTemplateKey, ImportTemplate> = {
  accompanistLessons: {
    fileName: "accompanist-lessons-template",
    description: "Accompanist lessons",
    sheetName: "Lessons",
    columns: [
      "Teacher", "Instructor Email", "Student", "Student ID", "Day", "Start Time", "End Time", "Room",
      "Instrument", "Needs Pianist", "Jury Required", "Specific Pianist",
    ],
    rows: [[
      "Taylor Example", "taylor.example@example.edu", "Sam Sample", "S0001", "Monday", "10:00 AM", "10:50 AM",
      "Room 101", "Violin", "Yes", "No", "",
    ]],
  },
  pianistAvailability: {
    fileName: "pianist-availability-template",
    description: "Pianist availability",
    sheetName: "Availability",
    columns: ["Pianist Name", "Pianist Email", "Max Hours Per Week", ...WEEKDAY_COLUMNS],
    rows: [["Avery Example", "avery.example@example.edu", 20, ...WEEKDAY_SAMPLE]],
  },
  juryPanels: {
    fileName: "jury-panels-template",
    description: "Jury panels",
    sheetName: "Jury Panels",
    columns: [
      "Schedule Date", "Panel Name", "Room", "Earliest Start", "Preferred Start", "Jury Length",
      "Break Needed", "Break Every X Juries", "Break Length", "Meal Break Needed", "Meal Start", "Meal End",
    ],
    rows: [["2026-12-07", "Violin", "Recital Hall", "9:00 AM", "9:30 AM", 15, "Yes", 4, 10, "Yes", "12:00 PM", "1:00 PM"]],
  },
  clinicalSites: {
    fileName: "clinical-sites-template",
    description: "Clinical sites",
    sheetName: "Clinical Sites",
    columns: [
      "Site Name", "Site Address", "Contact Name", "Contact Email", "Contact Phone", "Day of Week",
      "Start Time", "End Time", "Session Length", "Travel Time", "Capacity", "Tags",
    ],
    rows: [[
      "Example Community Center", "100 Main Street", "Jamie Sample", "jamie.sample@example.org", "555-0100",
      "Tuesday", "9:00 AM", "12:00 PM", 45, 15, 2, "Pediatric; Community",
    ]],
  },
  clinicalStudents: {
    fileName: "clinical-student-roster-template",
    description: "Clinical student roster",
    sheetName: "Student Roster",
    columns: ["Student ID", "Name", "Email"],
    rows: [["S0001", "Sam Sample", "sam.sample@example.edu"]],
  },
  lessonStudents: {
    fileName: "lesson-student-roster-template",
    description: "Lesson student roster",
    sheetName: "Student Roster",
    columns: ["Student ID", "Name", "Email", ...WEEKDAY_COLUMNS],
    rows: [["S0001", "Sam Sample", "sam.sample@example.edu", ...WEEKDAY_SAMPLE]],
  },
  lessonInstructors: {
    fileName: "lesson-instructor-roster-template",
    description: "Lesson instructor roster",
    sheetName: "Instructor Roster",
    columns: ["Instructor Name", "Instructor Email", "Room", "Lesson Length", "Capacity", ...WEEKDAY_COLUMNS, "Tags"],
    rows: [["Taylor Example", "taylor.example@example.edu", "Room 101", 30, 6, ...WEEKDAY_SAMPLE, "Strings; Violin"]],
  },
  tutorStudents: {
    fileName: "tutor-student-roster-template",
    description: "Tutor student roster",
    sheetName: "Student Roster",
    columns: ["Student ID", "Name", "Email", ...WEEKDAY_COLUMNS],
    rows: [["S0001", "Sam Sample", "sam.sample@example.edu", ...WEEKDAY_SAMPLE]],
  },
  tutorTutors: {
    fileName: "tutor-roster-template",
    description: "Tutor roster",
    sheetName: "Tutor Roster",
    columns: ["Tutor Name", "Tutor Email", "Room", "Lesson Length", "Capacity", ...WEEKDAY_COLUMNS, "Tags"],
    rows: [["Riley Example", "riley.example@example.edu", "Room 202", 30, 6, ...WEEKDAY_SAMPLE, "Theory; Piano"]],
  },
};

function escapeCsv(value: ReportCell): string {
  const text = value === null ? "" : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

/** Column headers and sample rows for a template. */
export function getImportTemplateTable(key: ImportTemplateKey): { columns: string[]; rows: ReportCell[][] } {
  const { columns, rows } = IMPORT_TEMPLATES[key];
  return { columns, rows };
}

export function generateImportTemplateCsv(key: ImportTemplateKey): string {
  const { columns, rows } = IMPORT_TEMPLATES[key];
  return [columns, ...rows].map((row) => row.map(escapeCsv).join(",")).join("\r\n");
}
