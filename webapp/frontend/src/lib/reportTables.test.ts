import { describe, expect, it } from "vitest";
import { buildAccompanistReportTable } from "./accompanistReportTable";
import { buildClinicalReportTable } from "./clinicalReportTable";
import { buildJuryReportTable } from "./juryReportTable";
import type { ClinicalSchedule, JuryScheduleResult, Lesson } from "./types";

describe("module report tables include all fields", () => {
  it("Accompanist: every Lesson field plus assigned Pianist details", () => {
    const lesson: Lesson = {
      id: 1, teacher: "Synthetic Teacher", teacher_email: "t@example.test", student: "Synthetic Student",
      student_id: "S1", day: "Monday", start_minute: 540, end_minute: 570, room: "R1", instrument: "Voice",
      required_pianist_name: "Synthetic Pianist", need_pianist: true, jury_required: false, assigned_pianist_id: 7,
      fit_quality: "Manual", notes: "A note", hours: 0.5, manually_edited: true,
    };
    const table = buildAccompanistReportTable([lesson], [
      { id: 7, name: "Synthetic Pianist", email: "p@example.test", max_hours_per_week: 12, availability_complete: true },
    ]);
    expect(table.columns).toEqual([
      "Day", "Start", "End", "Teacher", "Teacher Email", "Student", "Student ID", "Room", "Instrument",
      "Needs pianist?", "Jury required?", "Specific pianist", "Assigned pianist", "Assigned pianist Email",
      "Hours", "Fit", "Manually Edited", "Notes",
    ]);
    expect(table.rows[0]).toEqual([
      "Monday", "9:00 AM", "9:30 AM", "Synthetic Teacher", "t@example.test", "Synthetic Student", "S1", "R1", "Voice",
      true, false, "Synthetic Pianist", "Synthetic Pianist", "p@example.test", 0.5, "Manual", true, "A note",
    ]);
  });

  it("Jury: Panel settings, entry fields, breaks, and unscheduled reasons", () => {
    const result = {
      panel_timelines: [{
        panel_uuid: "p1", panel_name: "Voice Panel", jury_date: "2026-10-20",
        events: [
          { kind: "jury", source_lesson_uuid: "l1", student_person_uuid: "s", student_display_name: "Synthetic Student",
            instrument: "Voice", panel_uuid: "p1", jury_date: "2026-10-20", pianist_person_uuid: "x",
            pianist_display_name: "Synthetic Pianist", start_minute: 540, end_minute: 570 },
          { kind: "periodic_break", panel_uuid: "p1", jury_date: "2026-10-20", start_minute: 570, end_minute: 580, after_jury_count: 3 },
        ],
      }],
      unscheduled_lessons: [{
        source_lesson_uuid: "l2", student_person_uuid: "s2", student_display_name: "Other Student", instrument: "Voice",
        panel_uuid: "p1", jury_date: "2026-10-20", reason_code: "NO_SLOT", explanation: "No slot available",
      }],
    } as unknown as JuryScheduleResult;
    const table = buildJuryReportTable(result, [
      { source_lesson_uuid: "l1", teacher: "Synthetic Instructor", pianist_required: true, assigned_pianist: null },
      { source_lesson_uuid: "l2", teacher: "Other Instructor", pianist_required: false, assigned_pianist: null },
    ] as never);
    const rows = table.rows.map((r) => Object.fromEntries(table.columns.map((c, i) => [c, r[i]])));
    expect(table.columns).toEqual([
      "Panel", "Jury Date", "Start", "End", "Student / entry", "Instrument", "Instructor", "Assigned Pianist",
      "Pianist Required", "Status", "Reason", "Explanation",
    ]);
    expect(rows).toHaveLength(3);
    expect(rows[0]).toMatchObject({ Panel: "Voice Panel", "Jury Date": "2026-10-20", Start: "9:00 AM", End: "9:30 AM",
      "Student / entry": "Synthetic Student", Instrument: "Voice", Instructor: "Synthetic Instructor", "Assigned Pianist": "Synthetic Pianist", "Pianist Required": true, Status: "Scheduled" });
    expect(rows[1]).toMatchObject({ "Student / entry": "BREAK", Start: "9:30 AM", End: "9:40 AM" });
    expect(rows[2]).toMatchObject({ Panel: "Voice Panel", "Student / entry": "Other Student", Instructor: "Other Instructor", "Pianist Required": false, Status: "Unscheduled",
      Reason: "No Slot", Explanation: "No slot available" });
  });

  it("Clinical: placed and unplaced students with site, contact, tag, and warning fields", () => {
    const schedule = {
      entries: [
        { student_id: "s1", student_name: "Ada Example", student_number: "A1", kind: "manual", opportunity_id: "o1",
          site_name: "Mercy Clinic", opportunity_label: "AM", day: "Tuesday", start_minute: 540, end_minute: 600,
          tags: ["adult", "hospital"], preferred_tag_match_count: 1, tentative_minutes: 15, verified: true,
          warnings: [{ code: "TENTATIVE_AVAILABILITY", message: "Tentative" }], unplaced_reason: null },
        { student_id: "s2", student_name: "Bo Example", student_number: "B2", kind: null, opportunity_id: null, site_name: null,
          opportunity_label: null, day: null, start_minute: null, end_minute: null, tags: [], preferred_tag_match_count: 0,
          tentative_minutes: null, verified: null, warnings: [], unplaced_reason: "CAPACITY_LIMIT" },
      ],
    } as unknown as ClinicalSchedule;
    const table = buildClinicalReportTable(
      schedule,
      [{ id: "o1", day_of_week: "Tuesday", session_length_minutes: 60, travel_minutes: 20, site_address: "1 Main St",
        contact_name: "C", contact_email: "c@example.test", contact_phone: "555" }] as never,
      [{ id: "s1", email: "ada@example.test" }] as never,
    );
    const rows = table.rows.map((r) => Object.fromEntries(table.columns.map((c, i) => [c, r[i]])));
    expect(table.columns).toEqual([
      "Student", "Student ID", "Student Email", "Placement Site", "Site Address", "Contact Name", "Contact Email",
      "Contact Phone", "Opportunity Label", "Tags", "Day", "Start Time", "End Time", "Session Length (minutes)",
      "Travel (minutes)", "Placement Type", "Preferred Tag Matches", "Tentative Minutes", "Status", "Warnings",
      "Unplaced Reason",
    ]);
    expect(rows[0]).toMatchObject({ Student: "Ada Example", "Student ID": "A1", "Student Email": "ada@example.test", "Placement Site": "Mercy Clinic",
      "Site Address": "1 Main St", "Contact Name": "C", "Contact Email": "c@example.test", "Contact Phone": "555",
      "Opportunity Label": "AM", "Session Length (minutes)": 60, "Travel (minutes)": 20, "Preferred Tag Matches": 1,
      "Tentative Minutes": 15,
      Tags: "adult, hospital", Day: "Tuesday", "Start Time": "9:00 AM", "End Time": "10:00 AM",
      "Placement Type": "Manual", Status: "Placement warning", Warnings: "Tentative", "Unplaced Reason": "" });
    expect(rows[1]).toMatchObject({ Student: "Bo Example", "Placement Type": "Unplaced", Status: "",
      "Unplaced Reason": "Capacity limit" });
  });
});
