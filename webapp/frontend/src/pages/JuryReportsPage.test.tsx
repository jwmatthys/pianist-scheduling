import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { JuryScheduleResult } from "../lib/types";

const mocks = vi.hoisted(() => ({
  getCurrentJurySchedule: vi.fn(),
  getJuryEntries: vi.fn(),
  getJuryPanels: vi.fn(),
  createMarkdownPdf: vi.fn(),
  saveBinaryFile: vi.fn(),
  saveTextFile: vi.fn(),
}));

vi.mock("../lib/api", () => ({
  api: {
    getCurrentJurySchedule: mocks.getCurrentJurySchedule,
    getJuryEntries: mocks.getJuryEntries,
    getJuryPanels: mocks.getJuryPanels,
  },
}));
vi.mock("../lib/markdownPdf", () => ({ createMarkdownPdf: mocks.createMarkdownPdf }));
vi.mock("../lib/platform", () => ({
  saveBinaryFile: mocks.saveBinaryFile,
  saveTextFile: mocks.saveTextFile,
}));

import { JuryReportsPage } from "./JuryReportsPage";

function currentSchedule(overrides: Partial<JuryScheduleResult> = {}): JuryScheduleResult {
  return {
    result_uuid: "jury-result-current",
    session_uuid: "session-1",
    contract_id: "jury.schedule-result",
    contract_version: 1,
    result_version: 3,
    state: "draft",
    created_at: "2026-10-02T14:00:00Z",
    source_result_uuid: "accompanist-result-4",
    source_contract_id: "accompanist.assignment-result",
    source_contract_version: 2,
    source_revision: 4,
    jury_input_revision: 6,
    stale: false,
    stale_reasons: [],
    panel_timelines: [
      {
        panel_uuid: "panel-voice",
        panel_name: "Voice Panel",
        jury_date: "2026-10-20",
        events: [{
          kind: "jury",
          source_lesson_uuid: "lesson-voice",
          student_person_uuid: "student-voice",
          student_display_name: "Synthetic Voice Student (A.B!)",
          instrument: "Voice",
          panel_uuid: "panel-voice",
          jury_date: "2026-10-20",
          pianist_person_uuid: "pianist-1",
          pianist_display_name: "Synthetic Pianist",
          start_minute: 540,
          end_minute: 570,
        }],
      },
      {
        panel_uuid: "panel-cello",
        panel_name: "Cello Panel",
        jury_date: "2026-10-21",
        events: [{
          kind: "jury",
          source_lesson_uuid: "lesson-cello",
          student_person_uuid: "student-cello",
          student_display_name: "Synthetic Cello Student",
          instrument: "Cello",
          panel_uuid: "panel-cello",
          jury_date: "2026-10-21",
          pianist_person_uuid: null,
          pianist_display_name: null,
          start_minute: 600,
          end_minute: 660,
        }],
      },
    ],
    unscheduled_lessons: [],
    warnings: [],
    ...overrides,
  };
}

const currentEntries = [
  {
    session_uuid: "session-1",
    source_lesson_uuid: "lesson-voice",
    student_person_uuid: "student-voice",
    student_display_name: "Synthetic Voice Student",
    instrument: "Voice",
    teacher: "Synthetic Instructor One",
    pianist_required: true,
    assigned_pianist: { person_uuid: "pianist-1", display_name: "Synthetic Pianist" },
    jury_required: true,
    panel_uuid: "panel-voice",
  },
  {
    session_uuid: "session-1",
    source_lesson_uuid: "lesson-cello",
    student_person_uuid: "student-cello",
    student_display_name: "Synthetic Cello Student",
    instrument: "Cello",
    teacher: "Synthetic Instructor Two",
    pianist_required: false,
    assigned_pianist: null,
    jury_required: true,
    panel_uuid: "panel-cello",
  },
];

async function generate() {
  fireEvent.click(screen.getByRole("button", { name: "Generate report" }));
  await screen.findByRole("heading", { name: /Voice Panel/ });
}

describe("Jury reports", () => {
  beforeEach(() => {
    vi.resetAllMocks();
    mocks.getCurrentJurySchedule.mockResolvedValue(currentSchedule());
    mocks.getJuryEntries.mockResolvedValue(currentEntries);
    mocks.getJuryPanels.mockResolvedValue([
      { panel_uuid: "panel-voice", room: "Voice Studio" },
      { panel_uuid: "panel-cello", room: "" },
    ]);
    mocks.createMarkdownPdf.mockResolvedValue(new Blob(["pdf"], { type: "application/pdf" }));
    mocks.saveBinaryFile.mockResolvedValue(true);
    mocks.saveTextFile.mockResolvedValue(true);
  });

  it("reports Panel-grouped lessons with time, name, instrument, pianist, and instructor", async () => {
    render(<JuryReportsPage />);
    await generate();

    expect(screen.getByRole("heading", { name: /Voice Panel/ })).toBeTruthy();
    expect(screen.getByRole("heading", { name: /Cello Panel/ })).toBeTruthy();
    const lessonRows = screen.getAllByRole("listitem");
    expect(lessonRows[0].textContent).toContain("Synthetic Voice Student (A.B!)");
    expect(lessonRows[0].textContent).toContain("Synthetic Pianist");
    expect(lessonRows[0].textContent).toContain("Synthetic Instructor One");
    expect(lessonRows[0].textContent).toContain("9:00 AM");
    expect(lessonRows[1].textContent).toContain("Synthetic Cello Student");
    expect(lessonRows[1].textContent).toContain("Synthetic Instructor Two");
  });

  it("adds a pianist-grouped schedule after the Panel listings", async () => {
    render(<JuryReportsPage />);
    await generate();
    fireEvent.click(screen.getByRole("button", { name: "Download .md" }));

    await waitFor(() => expect(mocks.saveTextFile).toHaveBeenCalled());
    const markdown = mocks.saveTextFile.mock.calls[0][0] as string;
    expect(markdown.indexOf("# Schedule by Pianist")).toBeGreaterThan(markdown.indexOf("## Cello Panel"));
    expect(markdown).toContain("## Synthetic Pianist\n\n### October 20, 2026\n\n- **9:00 AM-9:30 AM** - Synthetic Voice Student (A.B!) - Voice - Voice Studio");
    expect(markdown).not.toContain("No Pianist Assigned");
    expect(markdown).toContain("## Voice Panel - October 20, 2026 - Voice Studio");
    expect(markdown).toContain("## Cello Panel - October 21, 2026\n");
    expect(markdown).toContain("Synthetic Voice Student (A.B!) - Voice - Pianist: Synthetic Pianist");
  });

  it("exports only the local generated timestamp and leaves punctuation unescaped", async () => {
    mocks.getCurrentJurySchedule.mockResolvedValue(currentSchedule({ stale: true, stale_reasons: ["jury_inputs_changed"] }));
    render(<JuryReportsPage />);
    await generate();

    fireEvent.click(screen.getByRole("button", { name: "Download .md" }));
    await waitFor(() => expect(mocks.saveTextFile).toHaveBeenCalled());
    const markdown = mocks.saveTextFile.mock.calls[0][0] as string;
    expect(markdown).toContain(`Generated: ${new Date(currentSchedule().created_at).toLocaleString()}`);
    expect(markdown).toContain("Synthetic Voice Student (A.B!)");
    expect(markdown).not.toMatch(/\\[.!()]/);
    expect(markdown).not.toMatch(/^(Result:|Accompanist source revision:|Jury input revision:|Stale:|Stale reasons:)/m);
  });

  it("exports meal and periodic breaks with uppercase labels", async () => {
    mocks.getCurrentJurySchedule.mockResolvedValue(currentSchedule({
      panel_timelines: [{
        panel_uuid: "panel-breaks",
        panel_name: "Break Panel",
        jury_date: "2026-10-22",
        events: [
          { kind: "meal_break", panel_uuid: "panel-breaks", jury_date: "2026-10-22", start_minute: 720, end_minute: 750 },
          { kind: "periodic_break", panel_uuid: "panel-breaks", jury_date: "2026-10-22", start_minute: 800, end_minute: 810, after_jury_count: 3 },
        ],
      }],
    }));
    render(<JuryReportsPage />);
    fireEvent.click(screen.getByRole("button", { name: "Generate report" }));
    await screen.findByRole("heading", { name: /Break Panel/ });
    fireEvent.click(screen.getByRole("button", { name: "Download .md" }));

    await waitFor(() => expect(mocks.saveTextFile).toHaveBeenCalled());
    const markdown = mocks.saveTextFile.mock.calls[0][0] as string;
    expect(markdown).toContain("MEAL BREAK");
    expect(markdown).toContain("BREAK");
    expect(markdown).not.toContain("Periodic Break");
  });

  it("requests a PDF page break before each Panel after the first", async () => {
    render(<JuryReportsPage />);
    await generate();
    fireEvent.click(screen.getByRole("button", { name: "Download .pdf" }));

    await waitFor(() => expect(mocks.createMarkdownPdf).toHaveBeenCalled());
    const [markdown, options] = mocks.createMarkdownPdf.mock.calls[0];
    expect(markdown).toContain("## Voice Panel");
    expect(markdown).toContain("## Cello Panel");
    expect(options.pageBreakBeforeHeadings).toEqual([
      expect.stringContaining("Cello Panel"),
      "Schedule by Pianist",
    ]);
    expect(mocks.saveBinaryFile).toHaveBeenCalledWith(expect.any(Blob), "jury_schedule.pdf", "PDF report", "pdf");
  });

  it("saves the Markdown report through the Save As helper", async () => {
    render(<JuryReportsPage />);
    await generate();
    fireEvent.click(screen.getByRole("button", { name: "Download .md" }));

    await waitFor(() => expect(mocks.saveTextFile).toHaveBeenCalled());
    expect(mocks.saveTextFile).toHaveBeenCalledWith(
      expect.stringContaining("Synthetic Instructor Two"),
      "jury_schedule.md",
      "Markdown report",
      "md",
    );
  });

  it("surfaces missing current schedule and API errors", async () => {
    mocks.getCurrentJurySchedule.mockRejectedValue(new Error("404: No current Jury schedule."));
    render(<JuryReportsPage />);
    fireEvent.click(screen.getByRole("button", { name: "Generate report" }));

    expect(await screen.findByRole("alert")).toBeTruthy();
    expect(screen.getByRole("alert").textContent).toContain("404: No current Jury schedule.");
  });
});
