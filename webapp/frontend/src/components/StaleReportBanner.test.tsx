import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  getSavedReport: vi.fn(),
  saveReport: vi.fn(),
  getAccompanistFinalizationState: vi.fn(),
  markdownReport: vi.fn(),
  listLessons: vi.fn(),
  listPianists: vi.fn(),
  getCurrentJurySchedule: vi.fn(),
  getClinicalSchedule: vi.fn(),
}));

vi.mock("../lib/api", () => ({ api: mocks }));

import { ReportsPage } from "../pages/ReportsPage";
import { JuryReportsPage } from "../pages/JuryReportsPage";
import { ClinicalReportsPage } from "../pages/clinical/ClinicalReportsPage";

const saved = (stamp: string) => ({
  markdown: "# Saved Synthetic Report",
  page_break_headings: [],
  table: { name: "Saved", columns: ["A"], rows: [["x"]] },
  source_stamp: stamp,
  generated_at: "2026-10-02T14:00:00Z",
});

describe("persisted reports and stale warning", () => {
  beforeEach(() => {
    vi.resetAllMocks();
    mocks.getAccompanistFinalizationState.mockResolvedValue({ source_revision: 3 });
    mocks.getCurrentJurySchedule.mockResolvedValue({ result_uuid: "r1", stale: false });
    mocks.getClinicalSchedule.mockResolvedValue({ schedule_revision: 2, source_data_revision: 5, is_schedule_stale: false });
    mocks.saveReport.mockImplementation(async (_id: string, body: object) => ({ ...body, generated_at: "2026-10-03T10:00:00Z" }));
  });

  it("shows the saved report on open without regenerating and without a warning when current", async () => {
    mocks.getSavedReport.mockResolvedValue(saved("3"));
    render(<ReportsPage />);
    expect(await screen.findByText("Saved Synthetic Report")).toBeTruthy();
    expect(screen.getByText(/^Generated /)).toBeTruthy();
    expect(screen.queryByText(/out of date/)).toBeNull();
    expect(mocks.getSavedReport).toHaveBeenCalledWith("accompanists");
    expect(mocks.markdownReport).not.toHaveBeenCalled();
  });

  it("warns when the Accompanist source revision changed", async () => {
    mocks.getSavedReport.mockResolvedValue(saved("2"));
    render(<ReportsPage />);
    expect((await screen.findByText(/out of date/)).closest("[role=status]")).toBeTruthy();
  });

  it("warns when the Jury result changed or is flagged stale", async () => {
    mocks.getSavedReport.mockResolvedValue(saved("other-result"));
    const { unmount } = render(<JuryReportsPage />);
    expect(await screen.findByText(/out of date/)).toBeTruthy();
    unmount();
    mocks.getSavedReport.mockResolvedValue(saved("r1"));
    mocks.getCurrentJurySchedule.mockResolvedValue({ result_uuid: "r1", stale: true });
    render(<JuryReportsPage />);
    expect(await screen.findByText(/out of date/)).toBeTruthy();
  });

  it("warns for Clinical only when the schedule or its source data changed", async () => {
    mocks.getSavedReport.mockResolvedValue(saved("2:5"));
    const { unmount } = render(<ClinicalReportsPage />);
    await screen.findByText("Saved Synthetic Report");
    expect(screen.queryByText(/out of date/)).toBeNull();
    unmount();
    mocks.getSavedReport.mockResolvedValue(saved("2:4"));
    render(<ClinicalReportsPage />);
    expect(await screen.findByText(/out of date/)).toBeTruthy();
  });

  it("regenerating saves the report with the current source stamp and clears the warning", async () => {
    mocks.getSavedReport.mockResolvedValue(saved("2"));
    mocks.markdownReport.mockResolvedValue("# Fresh Synthetic Report");
    mocks.listLessons.mockResolvedValue([]);
    mocks.listPianists.mockResolvedValue([]);
    render(<ReportsPage />);
    await screen.findByText(/out of date/);
    fireEvent.click(screen.getByRole("button", { name: "Generate report" }));
    await screen.findByText("Fresh Synthetic Report");
    await waitFor(() => expect(screen.queryByText(/out of date/)).toBeNull());
    expect(mocks.saveReport).toHaveBeenCalledWith(
      "accompanists",
      expect.objectContaining({ markdown: "# Fresh Synthetic Report", source_stamp: "3" }),
    );
  });
});
