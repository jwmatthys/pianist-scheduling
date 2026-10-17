import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  getClinicalSchedule: vi.fn(),
  getSavedReport: vi.fn(),
  saveReport: vi.fn(),
  listClinicalOpportunities: vi.fn(),
  listClinicalStudents: vi.fn(),
  createMarkdownPdf: vi.fn(),
  saveBinaryFile: vi.fn(),
  saveTextFile: vi.fn(),
}));

vi.mock("../../lib/api", () => ({
  api: {
    getClinicalSchedule: mocks.getClinicalSchedule,
    getSavedReport: mocks.getSavedReport,
    saveReport: mocks.saveReport,
    listClinicalOpportunities: mocks.listClinicalOpportunities,
    listClinicalStudents: mocks.listClinicalStudents,
  },
}));
vi.mock("../../lib/markdownPdf", () => ({ createMarkdownPdf: mocks.createMarkdownPdf }));
vi.mock("../../lib/platform", () => ({
  saveBinaryFile: mocks.saveBinaryFile,
  saveTextFile: mocks.saveTextFile,
}));

import { ClinicalReportsPage } from "./ClinicalReportsPage";

const entry = {
  student_id: "s1",
  student_name: "Ada Example",
  student_number: "",
  kind: "automatic",
  opportunity_id: "o1",
  site_name: "Mercy Clinic",
  opportunity_label: null,
  day: "Tuesday",
  start_minute: 540,
  end_minute: 600,
  tags: [],
  preferred_tag_match_count: 0,
  tentative_minutes: 0,
  verified: true,
  warnings: [],
  unplaced_reason: null,
};

describe("ClinicalReportsPage", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mocks.getSavedReport.mockResolvedValue(null);
    mocks.saveReport.mockImplementation(async (_id: string, body: object) => ({ ...body, generated_at: "2026-10-02T14:00:00Z" }));
    mocks.getClinicalSchedule.mockResolvedValue({ is_schedule_stale: false, entries: [entry] });
    mocks.listClinicalOpportunities.mockResolvedValue([]);
    mocks.listClinicalStudents.mockResolvedValue([{ id: "s1", email: "ada@example.test" }]);
    mocks.createMarkdownPdf.mockResolvedValue(new Uint8Array([1]));
    mocks.saveBinaryFile.mockResolvedValue(true);
    mocks.saveTextFile.mockResolvedValue(true);
  });

  it("disables export until a report is generated, then previews it", async () => {
    render(<ClinicalReportsPage />);
    expect((screen.getByRole("button", { name: "Export..." }) as HTMLButtonElement).disabled).toBe(true);
    await userEvent.click(screen.getByRole("button", { name: "Generate report" }));
    expect(await screen.findByRole("heading", { name: "Student Placement Roster" })).toBeTruthy();
    expect(screen.getByRole("heading", { name: "Site Placement Roster" })).toBeTruthy();
    expect(screen.getByRole("heading", { name: "Unplaced Students" })).toBeTruthy();
    expect(screen.getAllByText(/ada@example.test/).length).toBe(2);
  });

  it("exports Markdown and PDF locally", async () => {
    render(<ClinicalReportsPage />);
    await userEvent.click(screen.getByRole("button", { name: "Generate report" }));
    await screen.findByRole("heading", { name: "Student Placement Roster" });
    await userEvent.click(screen.getByRole("button", { name: "Export..." }));
    await userEvent.click(screen.getByRole("button", { name: "Markdown" }));
    await waitFor(() => expect(mocks.saveTextFile).toHaveBeenCalledWith(
      expect.stringContaining("Ada Example"), "Clinical Placements.md", "Markdown report", "md"));
    await userEvent.click(screen.getByRole("button", { name: "Export..." }));
    await userEvent.click(screen.getByRole("button", { name: "PDF" }));
    await waitFor(() => expect(mocks.createMarkdownPdf).toHaveBeenCalledWith(
      expect.stringContaining("Site Placement Roster"),
      { pageBreakBeforeHeadings: ["Site Placement Roster", "Unplaced Students"] }));
    await waitFor(() => expect(mocks.saveBinaryFile).toHaveBeenCalledWith(
      expect.anything(), "Clinical Placements.pdf", "PDF report", "pdf"));
  });

  it("shows an error if the schedule cannot be loaded", async () => {
    mocks.getClinicalSchedule.mockRejectedValue(new Error("boom"));
    render(<ClinicalReportsPage />);
    await userEvent.click(screen.getByRole("button", { name: "Generate report" }));
    expect((await screen.findByRole("alert")).textContent).toContain("boom");
  });
});
