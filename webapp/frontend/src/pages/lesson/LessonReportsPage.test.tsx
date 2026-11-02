import { render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  getLessonSchedule: vi.fn(),
  getLessonModuleState: vi.fn(),
  listLessonInstructors: vi.fn(),
  getSavedReport: vi.fn(),
  saveReport: vi.fn(),
  createMarkdownPdf: vi.fn(),
  saveBinaryFile: vi.fn(),
  saveTextFile: vi.fn(),
}));

vi.mock("../../lib/api", () => ({ api: mocks }));
vi.mock("../../lib/markdownPdf", () => ({ createMarkdownPdf: mocks.createMarkdownPdf }));
vi.mock("../../lib/platform", () => ({
  saveBinaryFile: mocks.saveBinaryFile,
  saveTextFile: mocks.saveTextFile,
}));

import { LessonReportsPage } from "./LessonReportsPage";

const entry = {
  student_id: "s1",
  student_name: "Ada Example",
  placement_type: "automatic",
  placement_state: "assigned",
  instructor_id: "i1",
  instructor_name: "Mina Teacher",
  day: "Tuesday",
  start_minute: 540,
  end_minute: 600,
  preferred_tag_match_count: 0,
  tentative_minutes: 0,
  verified: true,
  warnings: [],
  unplaced_reason: null,
};
const unplacedEntry = {
  ...entry,
  student_id: "s2",
  student_name: "Bo Example",
  placement_type: null,
  placement_state: "unassigned",
  instructor_id: null,
  instructor_name: null,
  day: null,
  start_minute: null,
  end_minute: null,
  unplaced_reason: "CAPACITY_LIMIT",
};

function schedule(overrides: object = {}) {
  return {
    schedule_revision: 2,
    source_data_revision: 3,
    is_schedule_stale: false,
    entries: [entry, unplacedEntry],
    verification: [],
    ...overrides,
  };
}

describe("LessonReportsPage", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mocks.getSavedReport.mockResolvedValue(null);
    mocks.saveReport.mockImplementation(async (_id: string, body: object) => ({ ...body, generated_at: "2026-10-02T14:00:00Z" }));
    mocks.getLessonSchedule.mockResolvedValue(schedule());
    mocks.getLessonModuleState.mockResolvedValue({ schedule_revision: 2, source_data_revision: 3, is_schedule_stale: false });
    mocks.listLessonInstructors.mockResolvedValue([
      { id: "i1", name: "Mina Teacher", room: "Studio 1", lesson_length_minutes: 60 },
    ]);
    mocks.createMarkdownPdf.mockResolvedValue(new Uint8Array([1]));
    mocks.saveBinaryFile.mockResolvedValue(true);
    mocks.saveTextFile.mockResolvedValue(true);
  });

  it("disables export until a report is generated, then previews every section", async () => {
    render(<LessonReportsPage />);
    expect((screen.getByRole("button", { name: "Export..." }) as HTMLButtonElement).disabled).toBe(true);
    await userEvent.click(screen.getByRole("button", { name: "Generate report" }));
    for (const name of ["Student Placement Roster", "Instructor Placement Roster", "Unplaced Students", "Placement Warnings"]) {
      expect(await screen.findByRole("heading", { name })).toBeTruthy();
    }
    expect(screen.getByText(/Capacity limit/)).toBeTruthy();
    expect(mocks.saveReport).toHaveBeenCalledWith("lesson", expect.objectContaining({ source_stamp: "2:3" }));
  });

  it("has no report selector", async () => {
    render(<LessonReportsPage />);
    await userEvent.click(screen.getByRole("button", { name: "Generate report" }));
    await screen.findByRole("heading", { name: "Student Placement Roster" });
    expect(screen.queryByRole("combobox")).toBeNull();
  });

  it("exports the full report as Markdown and PDF", async () => {
    render(<LessonReportsPage />);
    await userEvent.click(screen.getByRole("button", { name: "Generate report" }));
    await screen.findByRole("heading", { name: "Student Placement Roster" });
    await userEvent.click(screen.getByRole("button", { name: "Export..." }));
    await userEvent.click(screen.getByRole("button", { name: "PDF" }));
    await waitFor(() => expect(mocks.createMarkdownPdf).toHaveBeenCalledWith(
      expect.stringContaining("Instructor Placement Roster"),
      { pageBreakBeforeHeadings: ["Instructor Placement Roster", "Unplaced Students"] }));
    await waitFor(() => expect(mocks.saveBinaryFile).toHaveBeenCalledWith(
      expect.anything(), "Lesson Schedule.pdf", "PDF report", "pdf"));

    await userEvent.click(screen.getByRole("button", { name: "Export..." }));
    await userEvent.click(screen.getByRole("button", { name: "Markdown" }));
    await waitFor(() => expect(mocks.saveTextFile).toHaveBeenCalledWith(
      expect.stringContaining("Bo Example"), "Lesson Schedule.md", "Markdown report", "md"));
    const markdown = mocks.saveTextFile.mock.calls[0][0] as string;
    for (const name of ["Student Placement Roster", "Instructor Placement Roster", "Unplaced Students", "Placement Warnings"]) {
      expect(markdown).toContain(`## ${name}`);
    }
  });

  it("offers Excel and CSV exports from the shared dialog", async () => {
    render(<LessonReportsPage />);
    await userEvent.click(screen.getByRole("button", { name: "Generate report" }));
    await screen.findByRole("heading", { name: "Student Placement Roster" });
    await userEvent.click(screen.getByRole("button", { name: "Export..." }));
    expect(screen.getByRole("button", { name: "Excel Workbook" })).toBeTruthy();
    expect(screen.getByRole("button", { name: "CSV" })).toBeTruthy();
  });

  it("generates a stale schedule with a warning in the report and keeps export enabled", async () => {
    mocks.getLessonSchedule.mockResolvedValue(schedule({ is_schedule_stale: true }));
    render(<LessonReportsPage />);
    await userEvent.click(screen.getByRole("button", { name: "Generate report" }));
    expect((await screen.findAllByText(/has changed since this schedule was generated/)).length).toBeGreaterThan(0);
    expect(screen.getByText(/This report is out of date/)).toBeTruthy();
    expect((screen.getByRole("button", { name: "Export..." }) as HTMLButtonElement).disabled).toBe(false);
    await userEvent.click(screen.getByRole("button", { name: "Export..." }));
    await userEvent.click(screen.getByRole("button", { name: "Markdown" }));
    await waitFor(() => expect(mocks.saveTextFile).toHaveBeenCalledWith(
      expect.stringContaining("has changed since this schedule was generated"), expect.any(String), "Markdown report", "md"));
  });

  it("marks a saved report stale when the schedule revision has moved on", async () => {
    mocks.getSavedReport.mockResolvedValue({
      markdown: "# Lesson Schedule\n\n## Student Placement Roster\n\nNo students have been placed.\n",
      page_break_headings: [],
      table: { name: "Lesson Schedule", columns: [], rows: [] },
      source_stamp: "1:1",
      generated_at: "2026-10-01T14:00:00Z",
    });
    render(<LessonReportsPage />);
    expect(await screen.findByText(/This report is out of date/)).toBeTruthy();
  });

  it("shows an error if the schedule cannot be loaded", async () => {
    mocks.getLessonSchedule.mockRejectedValue(new Error("boom"));
    render(<LessonReportsPage />);
    await userEvent.click(screen.getByRole("button", { name: "Generate report" }));
    expect(within(await screen.findByRole("alert")).getByText(/boom/)).toBeTruthy();
  });
});
