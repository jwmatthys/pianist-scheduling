import { describe, expect, it } from "vitest";
import {
  INSTRUCTOR_ROSTER_HEADING,
  STALE_NOTE,
  STUDENT_ROSTER_HEADING,
  UNPLACED_HEADING,
  WARNINGS_HEADING,
  buildLessonReport,
  buildLessonReportTable,
} from "./lessonReport";
import type { LessonInstructor, LessonSchedule, LessonScheduleEntry } from "./types";

function splitSections(markdown: string) {
  const sections = new Map<string, string>();
  for (const part of markdown.split(/^(?=## )/m).slice(1)) sections.set(part.split("\n")[0].slice(3), part);
  return { sections };
}

function entry(overrides: Partial<LessonScheduleEntry>): LessonScheduleEntry {
  return {
    student_id: "s",
    student_name: "Student",
    placement_type: "automatic",
    placement_state: "assigned",
    instructor_id: "i1",
    instructor_name: "Mina Teacher",
    day: "Monday",
    start_minute: 540,
    end_minute: 600,
    preferred_tag_match_count: 0,
    tentative_minutes: 0,
    verified: true,
    warnings: [],
    unplaced_reason: null,
    ...overrides,
  };
}

function instructor(overrides: Partial<LessonInstructor>): LessonInstructor {
  return {
    id: "i1",
    name: "Mina Teacher",
    email: "",
    room: "Studio 1",
    lesson_length_minutes: 60,
    capacity: 3,
    tags: [],
    availability_window_count: 1,
    created_at: "",
    updated_at: "",
    ...overrides,
  };
}

function schedule(entries: LessonScheduleEntry[], overrides: Partial<LessonSchedule> = {}): LessonSchedule {
  return {
    schedule_revision: 1,
    source_data_revision: 1,
    is_schedule_stale: false,
    has_schedule: true,
    students_scheduled: 0,
    students_unscheduled: 0,
    instructor_utilization: 0,
    preferred_tag_matches: 0,
    tentative_minutes: 0,
    gap_cost_minutes: 0,
    automatic_count: 0,
    manual_count: 0,
    manually_unplaced_count: 0,
    verified_count: 0,
    unverified_count: 0,
    entries,
    verification: [],
    ...overrides,
  };
}

const instructors = [
  instructor({}),
  instructor({ id: "i2", name: "Jo Instructor", room: "", lesson_length_minutes: 30 }),
  instructor({ id: "i3", name: "Idle Instructor" }),
];

const entries = [
  entry({ student_id: "s2", student_name: "Zed Example", day: "Tuesday", start_minute: 600, end_minute: 660 }),
  entry({ student_id: "s1", student_name: "Ada Example", day: "Monday", start_minute: 660, end_minute: 720, placement_type: "manual" }),
  entry({ student_id: "s3", student_name: "Bo Example", instructor_id: "i2", instructor_name: "Jo Instructor", start_minute: 540, end_minute: 570 }),
  entry({ student_id: "s4", student_name: "Cy Example", placement_type: null, placement_state: "unassigned", instructor_id: null, instructor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: "CAPACITY_LIMIT", verified: null }),
  entry({ student_id: "s5", student_name: "Di Example", placement_type: "manuallyUnplaced", placement_state: "manuallyUnplaced", instructor_id: null, instructor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: "MANUALLY_UNPLACED", verified: null }),
  entry({ student_id: "s6", student_name: "Ed Example", placement_type: null, placement_state: "unassigned", instructor_id: null, instructor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: "INSTRUCTOR_OVERLAP", verified: null }),
  entry({ student_id: "s7", student_name: "Fay Example", placement_type: null, placement_state: "unassigned", instructor_id: null, instructor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: null, verified: null }),
  // The API reports an automatically unplaced student as type "automatic" with no instructor or time.
  entry({ student_id: "s8", student_name: "Gus Example", placement_type: "automatic", placement_state: "assigned", instructor_id: null, instructor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: "CAPACITY_LIMIT", verified: null }),
];

describe("buildLessonReport", () => {
  const report = buildLessonReport(schedule(entries), instructors);
  const { sections } = splitSections(report.markdown);

  it("lists placed students alphabetically with instructor, room, day and time", () => {
    const text = sections.get(STUDENT_ROSTER_HEADING)!;
    const names = [...text.matchAll(/\*\*([^*]+)\*\*/g)].map((match) => match[1]);
    expect(names).toEqual(["Ada Example", "Bo Example", "Zed Example"]);
    expect(text).toContain("Ada Example** - Mina Teacher - Studio 1 - Monday 11:00 AM-12:00 PM (Manual)");
    expect(text).toContain("Bo Example** - Jo Instructor - Monday 9:00 AM-9:30 AM");
  });

  it("groups the instructor roster by instructor, then day and start time", () => {
    const text = sections.get(INSTRUCTOR_ROSTER_HEADING)!;
    expect(text.indexOf("### Idle Instructor")).toBeLessThan(text.indexOf("### Jo Instructor"));
    expect(text.indexOf("### Jo Instructor")).toBeLessThan(text.indexOf("### Mina Teacher"));
    expect(text).toContain("### Mina Teacher - Room Studio 1, 60-minute lessons");
    expect(text).toContain("### Jo Instructor - 30-minute lessons");
    expect(text).toContain("No students assigned.");
    expect(text.indexOf("Monday 11:00 AM-12:00 PM")).toBeLessThan(text.indexOf("Tuesday 10:00 AM-11:00 AM"));
  });

  it("maps every unplaced reason to a user-facing label", () => {
    const text = sections.get(UNPLACED_HEADING)!;
    expect(text).toContain("**Cy Example** - Capacity limit");
    expect(text).toContain("**Di Example** - Manually unplaced");
    expect(text).toContain("**Ed Example** - Instructor overlap");
    expect(text).toContain("**Fay Example** - No feasible instructor");
    expect(text).toContain("**Gus Example** - Capacity limit");
    expect(sections.get(STUDENT_ROSTER_HEADING)).not.toContain("Gus Example");
  });

  it("reports no warnings when verification is clean and the schedule is fresh", () => {
    expect(sections.get(WARNINGS_HEADING)).toContain("No placement warnings.");
    expect(report.markdown).not.toContain(STALE_NOTE);
  });

  it("lists verification warnings and flags manual overrides", () => {
    const warned = buildLessonReport(
      schedule(entries, {
        verification: [
          { student_id: "s1", student_name: "Ada Example", instructor_id: "i1", instructor_name: "Mina Teacher", day: "Monday", start_minute: 660, warnings: [{ code: "INSTRUCTOR_OVERLAP", message: "The lesson overlaps another lesson." }] },
          { student_id: "s2", student_name: "Zed Example", instructor_id: "i1", instructor_name: "Mina Teacher", day: "Tuesday", start_minute: 600, warnings: [{ code: "STUDENT_UNAVAILABLE", message: "The student is unavailable." }] },
        ],
      }),
      instructors,
    );
    const text = splitSections(warned.markdown).sections.get(WARNINGS_HEADING)!;
    expect(text).toContain("**Ada Example** - Mina Teacher Monday 11:00 AM (Manual override)");
    expect(text).toContain("  - The lesson overlaps another lesson.");
    expect(text).toContain("**Zed Example** - Mina Teacher Tuesday 10:00 AM\n");
    expect(text).toContain("  - The student is unavailable.");
  });

  it("includes the stale warning in the report when the schedule or module state is stale", () => {
    for (const built of [
      buildLessonReport(schedule(entries, { is_schedule_stale: true }), instructors),
      buildLessonReport(schedule(entries), instructors, { stale: true }),
    ]) {
      expect(built.markdown).toContain(STALE_NOTE);
      expect(splitSections(built.markdown).sections.get(WARNINGS_HEADING)).toContain("The schedule is stale");
    }
  });

  it("does not expose internal identifiers", () => {
    expect(report.markdown).not.toMatch(/\bi[123]\b|\bs[1-8]\b/);
    expect(report.pageBreakBeforeHeadings).toEqual([INSTRUCTOR_ROSTER_HEADING, UNPLACED_HEADING]);
  });
});

describe("buildLessonReportTable", () => {
  it("builds one alphabetical row per student with full fields and no identifiers", () => {
    const table = buildLessonReportTable(schedule(entries), instructors);
    expect(table.columns).toEqual(expect.arrayContaining(["Student", "Instructor", "Room", "Day", "Start Time", "Unplaced Reason"]));
    expect(table.rows.map((row) => row[0])).toEqual([
      "Ada Example", "Bo Example", "Cy Example", "Di Example", "Ed Example", "Fay Example", "Gus Example", "Zed Example",
    ]);
    expect(table.rows[0]).toEqual([
      "Ada Example", "Mina Teacher", "Studio 1", 60, "Monday", "11:00 AM", "12:00 PM", "Manual", 0, 0, "Verified", "", "", "Current",
    ]);
    const unplaced = table.rows.find((row) => row[0] === "Cy Example")!;
    expect(unplaced[12]).toBe("Capacity limit");
    expect(unplaced[7]).toBe("Unassigned");
  });

  it("marks every row stale when the schedule or module state is stale", () => {
    for (const table of [
      buildLessonReportTable(schedule(entries, { is_schedule_stale: true }), instructors),
      buildLessonReportTable(schedule(entries), instructors, { stale: true }),
    ]) {
      expect(table.columns.at(-1)).toBe("Schedule Status");
      expect(table.rows.every((row) => String(row.at(-1)).startsWith("Stale"))).toBe(true);
    }
  });
});
