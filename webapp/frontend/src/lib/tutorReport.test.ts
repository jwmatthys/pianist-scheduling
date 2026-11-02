import { describe, expect, it } from "vitest";
import {
  TUTOR_ROSTER_HEADING,
  STALE_NOTE,
  STUDENT_ROSTER_HEADING,
  UNPLACED_HEADING,
  WARNINGS_HEADING,
  buildTutorReport,
  buildTutorReportTable,
} from "./tutorReport";
import type { TutorTutor, TutorSchedule, TutorScheduleEntry } from "./types";

function splitSections(markdown: string) {
  const sections = new Map<string, string>();
  for (const part of markdown.split(/^(?=## )/m).slice(1)) sections.set(part.split("\n")[0].slice(3), part);
  return { sections };
}

function entry(overrides: Partial<TutorScheduleEntry>): TutorScheduleEntry {
  return {
    student_id: "s",
    student_name: "Student",
    placement_type: "automatic",
    placement_state: "assigned",
    tutor_id: "i1",
    tutor_name: "Mina Teacher",
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

function tutor(overrides: Partial<TutorTutor>): TutorTutor {
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

function schedule(entries: TutorScheduleEntry[], overrides: Partial<TutorSchedule> = {}): TutorSchedule {
  return {
    schedule_revision: 1,
    source_data_revision: 1,
    is_schedule_stale: false,
    has_schedule: true,
    students_scheduled: 0,
    students_unscheduled: 0,
    tutor_utilization: 0,
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

const tutors = [
  tutor({}),
  tutor({ id: "i2", name: "Jo Tutor", room: "", lesson_length_minutes: 30 }),
  tutor({ id: "i3", name: "Idle Tutor" }),
];

const entries = [
  entry({ student_id: "s2", student_name: "Zed Example", day: "Tuesday", start_minute: 600, end_minute: 660 }),
  entry({ student_id: "s1", student_name: "Ada Example", day: "Monday", start_minute: 660, end_minute: 720, placement_type: "manual" }),
  entry({ student_id: "s3", student_name: "Bo Example", tutor_id: "i2", tutor_name: "Jo Tutor", start_minute: 540, end_minute: 570 }),
  entry({ student_id: "s4", student_name: "Cy Example", placement_type: null, placement_state: "unassigned", tutor_id: null, tutor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: "CAPACITY_LIMIT", verified: null }),
  entry({ student_id: "s5", student_name: "Di Example", placement_type: "manuallyUnplaced", placement_state: "manuallyUnplaced", tutor_id: null, tutor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: "MANUALLY_UNPLACED", verified: null }),
  entry({ student_id: "s6", student_name: "Ed Example", placement_type: null, placement_state: "unassigned", tutor_id: null, tutor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: "TUTOR_OVERLAP", verified: null }),
  entry({ student_id: "s7", student_name: "Fay Example", placement_type: null, placement_state: "unassigned", tutor_id: null, tutor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: null, verified: null }),
  // The API reports an automatically unplaced student as type "automatic" with no tutor or time.
  entry({ student_id: "s8", student_name: "Gus Example", placement_type: "automatic", placement_state: "assigned", tutor_id: null, tutor_name: null, day: null, start_minute: null, end_minute: null, unplaced_reason: "CAPACITY_LIMIT", verified: null }),
];

describe("buildTutorReport", () => {
  const report = buildTutorReport(schedule(entries), tutors);
  const { sections } = splitSections(report.markdown);

  it("lists placed students alphabetically with tutor, room, day and time", () => {
    const text = sections.get(STUDENT_ROSTER_HEADING)!;
    const names = [...text.matchAll(/\*\*([^*]+)\*\*/g)].map((match) => match[1]);
    expect(names).toEqual(["Ada Example", "Bo Example", "Zed Example"]);
    expect(text).toContain("Ada Example** - Mina Teacher - Studio 1 - Monday 11:00 AM-12:00 PM (Manual)");
    expect(text).toContain("Bo Example** - Jo Tutor - Monday 9:00 AM-9:30 AM");
  });

  it("groups the tutor roster by tutor, then day and start time", () => {
    const text = sections.get(TUTOR_ROSTER_HEADING)!;
    expect(text.indexOf("### Idle Tutor")).toBeLessThan(text.indexOf("### Jo Tutor"));
    expect(text.indexOf("### Jo Tutor")).toBeLessThan(text.indexOf("### Mina Teacher"));
    expect(text).toContain("### Mina Teacher - Room Studio 1, 60-minute sessions");
    expect(text).toContain("### Jo Tutor - 30-minute sessions");
    expect(text).toContain("No students assigned.");
    expect(text.indexOf("Monday 11:00 AM-12:00 PM")).toBeLessThan(text.indexOf("Tuesday 10:00 AM-11:00 AM"));
  });

  it("maps every unplaced reason to a user-facing label", () => {
    const text = sections.get(UNPLACED_HEADING)!;
    expect(text).toContain("**Cy Example** - Capacity limit");
    expect(text).toContain("**Di Example** - Manually unplaced");
    expect(text).toContain("**Ed Example** - Tutor overlap");
    expect(text).toContain("**Fay Example** - No feasible tutor");
    expect(text).toContain("**Gus Example** - Capacity limit");
    expect(sections.get(STUDENT_ROSTER_HEADING)).not.toContain("Gus Example");
  });

  it("reports no warnings when verification is clean and the schedule is fresh", () => {
    expect(sections.get(WARNINGS_HEADING)).toContain("No placement warnings.");
    expect(report.markdown).not.toContain(STALE_NOTE);
  });

  it("lists verification warnings and flags manual overrides", () => {
    const warned = buildTutorReport(
      schedule(entries, {
        verification: [
          { student_id: "s1", student_name: "Ada Example", tutor_id: "i1", tutor_name: "Mina Teacher", day: "Monday", start_minute: 660, warnings: [{ code: "TUTOR_OVERLAP", message: "The session overlaps another session." }] },
          { student_id: "s2", student_name: "Zed Example", tutor_id: "i1", tutor_name: "Mina Teacher", day: "Tuesday", start_minute: 600, warnings: [{ code: "STUDENT_UNAVAILABLE", message: "The student is unavailable." }] },
        ],
      }),
      tutors,
    );
    const text = splitSections(warned.markdown).sections.get(WARNINGS_HEADING)!;
    expect(text).toContain("**Ada Example** - Mina Teacher Monday 11:00 AM (Manual override)");
    expect(text).toContain("  - The session overlaps another session.");
    expect(text).toContain("**Zed Example** - Mina Teacher Tuesday 10:00 AM\n");
    expect(text).toContain("  - The student is unavailable.");
  });

  it("includes the stale warning in the report when the schedule or module state is stale", () => {
    for (const built of [
      buildTutorReport(schedule(entries, { is_schedule_stale: true }), tutors),
      buildTutorReport(schedule(entries), tutors, { stale: true }),
    ]) {
      expect(built.markdown).toContain(STALE_NOTE);
      expect(splitSections(built.markdown).sections.get(WARNINGS_HEADING)).toContain("The schedule is stale");
    }
  });

  it("does not expose internal identifiers", () => {
    expect(report.markdown).not.toMatch(/\bi[123]\b|\bs[1-8]\b/);
    expect(report.pageBreakBeforeHeadings).toEqual([TUTOR_ROSTER_HEADING, UNPLACED_HEADING]);
  });
});

describe("buildTutorReportTable", () => {
  it("builds one alphabetical row per student with full fields and no identifiers", () => {
    const table = buildTutorReportTable(schedule(entries), tutors);
    expect(table.columns).toEqual(expect.arrayContaining(["Student", "Tutor", "Room", "Day", "Start Time", "Unplaced Reason"]));
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
      buildTutorReportTable(schedule(entries, { is_schedule_stale: true }), tutors),
      buildTutorReportTable(schedule(entries), tutors, { stale: true }),
    ]) {
      expect(table.columns.at(-1)).toBe("Schedule Status");
      expect(table.rows.every((row) => String(row.at(-1)).startsWith("Stale"))).toBe(true);
    }
  });
});
