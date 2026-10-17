import { formatMinutes } from "./time";
import type { ReportTable } from "./reportTable";
import type { Lesson, Pianist } from "./types";

const DAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

/** Accompanist-owned: the Lesson Roster (Schedule tab), one row per Lesson with every roster column. */
export function buildAccompanistReportTable(lessons: Lesson[], pianists: Pianist[]): ReportTable {
  const pianistsById = new Map(pianists.map((pianist) => [pianist.id, pianist]));
  const ordered = [...lessons].sort(
    (a, b) => DAY_ORDER.indexOf(a.day) - DAY_ORDER.indexOf(b.day) || a.start_minute - b.start_minute,
  );
  return {
    name: "Accompanist Schedule",
    columns: [
      "Day", "Start", "End", "Teacher", "Teacher Email", "Student", "Student ID", "Room", "Instrument",
      "Needs pianist?", "Jury required?", "Specific pianist", "Assigned pianist", "Assigned pianist Email",
      "Hours", "Fit", "Manually Edited", "Notes",
    ],
    rows: ordered.map((lesson) => {
      const pianist = lesson.assigned_pianist_id === null ? undefined : pianistsById.get(lesson.assigned_pianist_id);
      return [
        lesson.day, formatMinutes(lesson.start_minute), formatMinutes(lesson.end_minute), lesson.teacher,
        lesson.teacher_email, lesson.student, lesson.student_id, lesson.room, lesson.instrument,
        lesson.need_pianist, lesson.jury_required, lesson.required_pianist_name, pianist?.name ?? "",
        pianist?.email ?? "", lesson.hours, lesson.fit_quality, lesson.manually_edited, lesson.notes,
      ];
    }),
  };
}
