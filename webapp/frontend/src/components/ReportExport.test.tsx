import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  exportReport: vi.fn(),
  markdownReport: vi.fn(),
  listLessons: vi.fn(),
  listPianists: vi.fn(),
  getSavedReport: vi.fn(),
  saveReport: vi.fn(),
  getAccompanistFinalizationState: vi.fn(),
  getCurrentJurySchedule: vi.fn(),
  getJuryEntries: vi.fn(),
  getJuryPanels: vi.fn(),
  getClinicalSchedule: vi.fn(),
  listClinicalOpportunities: vi.fn(),
  listClinicalStudents: vi.fn(),
}));

vi.mock("../lib/reportExport", async (importOriginal) => {
  const original = await importOriginal<typeof import("../lib/reportExport")>();
  return { ...original, reportExportService: { ...original.reportExportService, export: mocks.exportReport } };
});
vi.mock("../lib/api", () => ({
  api: {
    markdownReport: mocks.markdownReport,
    listLessons: mocks.listLessons,
    listPianists: mocks.listPianists,
    getSavedReport: mocks.getSavedReport,
    saveReport: mocks.saveReport,
    getAccompanistFinalizationState: mocks.getAccompanistFinalizationState,
    getCurrentJurySchedule: mocks.getCurrentJurySchedule,
    getJuryEntries: mocks.getJuryEntries,
    getJuryPanels: mocks.getJuryPanels,
    getClinicalSchedule: mocks.getClinicalSchedule,
    listClinicalOpportunities: mocks.listClinicalOpportunities,
    listClinicalStudents: mocks.listClinicalStudents,
  },
}));

import { ReportExportButton } from "./ReportExport";
import { ReportsPage } from "../pages/ReportsPage";
import { JuryReportsPage } from "../pages/JuryReportsPage";
import { ClinicalReportsPage } from "../pages/clinical/ClinicalReportsPage";

describe("shared Export dialog", () => {
  beforeEach(() => {
    vi.resetAllMocks();
    mocks.exportReport.mockResolvedValue({ saved: true, message: "CSV report saved." });
    mocks.getSavedReport.mockResolvedValue(null);
    mocks.saveReport.mockImplementation(async (_id: string, body: object) => ({ ...body, generated_at: "2026-10-02T14:00:00Z" }));
    mocks.getAccompanistFinalizationState.mockResolvedValue({ source_revision: 3 });
  });

  const payload = { markdown: "# My Report", table: { name: "My Report", columns: ["A"], rows: [["x"]] } };

  function renderButton(onStatus = vi.fn()) {
    render(<ReportExportButton reportId="r.id" title="My Report" payload={payload} onStatus={onStatus} />);
    return onStatus;
  }

  it("opens a dialog with stacked format choices and Cancel", () => {
    renderButton();
    expect(screen.queryByRole("dialog")).toBeNull();
    fireEvent.click(screen.getByRole("button", { name: "Export..." }));
    expect(screen.getByRole("dialog").textContent).toContain("Export Report");
    expect(screen.getByText("Choose format:")).toBeTruthy();
    const labels = screen.getAllByRole("button").map((button) => button.textContent);
    expect(labels).toEqual(expect.arrayContaining(["PDF", "Markdown", "Excel Workbook", "CSV", "Cancel"]));
  });

  it("dispatches the selected format with report id, title, and payload", async () => {
    const onStatus = renderButton();
    fireEvent.click(screen.getByRole("button", { name: "Export..." }));
    fireEvent.click(screen.getByRole("button", { name: "Excel Workbook" }));
    await waitFor(() => expect(mocks.exportReport).toHaveBeenCalledWith({
      reportId: "r.id",
      title: "My Report",
      format: "xlsx",
      payload,
    }));
    await waitFor(() => expect(onStatus).toHaveBeenCalledWith({ notice: "CSV report saved." }));
    expect(screen.queryByRole("dialog")).toBeNull();
  });

  it("closes on Cancel without exporting", () => {
    renderButton();
    fireEvent.click(screen.getByRole("button", { name: "Export..." }));
    fireEvent.click(screen.getByRole("button", { name: "Cancel" }));
    expect(screen.queryByRole("dialog")).toBeNull();
    expect(mocks.exportReport).not.toHaveBeenCalled();
  });

  it("surfaces export failures", async () => {
    mocks.exportReport.mockRejectedValue(new Error("disk full"));
    const onStatus = renderButton();
    fireEvent.click(screen.getByRole("button", { name: "Export..." }));
    fireEvent.click(screen.getByRole("button", { name: "CSV" }));
    await waitFor(() => expect(onStatus).toHaveBeenCalledWith({ error: "disk full" }));
  });

  it("is disabled without a payload", () => {
    render(<ReportExportButton reportId="r" title="t" payload={null} onStatus={vi.fn()} />);
    expect((screen.getByRole("button", { name: "Export..." }) as HTMLButtonElement).disabled).toBe(true);
  });

  it("is the same dialog from every module's report page", async () => {
    mocks.markdownReport.mockResolvedValue("# Lesson Schedules\n\n- one");
    mocks.listLessons.mockResolvedValue([]);
    mocks.listPianists.mockResolvedValue([]);
    mocks.getCurrentJurySchedule.mockResolvedValue({
      created_at: "2026-10-01T12:00:00Z",
      panel_timelines: [],
      unscheduled_lessons: [],
    });
    mocks.getJuryEntries.mockResolvedValue([]);
    mocks.getJuryPanels.mockResolvedValue([]);
    mocks.getClinicalSchedule.mockResolvedValue({ entries: [], unplaced: [], warnings: [] });
    mocks.listClinicalOpportunities.mockResolvedValue([]);
    mocks.listClinicalStudents.mockResolvedValue([]);

    const expected = [
      [ReportsPage, "Accompanist Schedule"],
      [JuryReportsPage, "Jury Schedule"],
      [ClinicalReportsPage, "Clinical Placements"],
    ] as const;
    for (const [Page, title] of expected) {
      const { unmount } = render(<Page />);
      fireEvent.click(screen.getByRole("button", { name: "Generate report" }));
      await waitFor(() => expect((screen.getByRole("button", { name: "Export..." }) as HTMLButtonElement).disabled).toBe(false));
      fireEvent.click(screen.getByRole("button", { name: "Export..." }));
      expect(screen.getByRole("dialog").textContent).toContain("Export Report");
      expect(screen.getAllByRole("button").map((b) => b.textContent)).toEqual(expect.arrayContaining(["PDF", "Markdown", "Excel Workbook", "CSV"]));
      fireEvent.click(screen.getByRole("button", { name: "CSV" }));
      await waitFor(() => expect(mocks.exportReport).toHaveBeenLastCalledWith(expect.objectContaining({ title, format: "csv" })));
      unmount();
    }
  });
});
