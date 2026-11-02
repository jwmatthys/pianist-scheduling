// @vitest-environment jsdom
import { cleanup, render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../lib/api";
import { chooseLocalFile } from "../lib/platform";
import type { JuryPanelImportInspection, JuryPanelImportMapping } from "../lib/types";
import { JuryPanelImportDialog } from "./JuryPanelImportDialog";

vi.mock("../lib/api", () => ({
  api: {
    inspectJuryPanelFile: vi.fn(),
    previewJuryPanelImport: vi.fn(),
    applyJuryPanelImport: vi.fn(),
  },
}));
vi.mock("../lib/platform", () => ({ chooseLocalFile: vi.fn() }));

const mapping: JuryPanelImportMapping = {
  schedule_date: "Date", panel_name: "Panel", room: null, earliest_start: "Start", preferred_start: null,
  jury_length: "Length", break_needed: null, break_every: null, break_length: null,
  meal_break: null, meal_start: null, meal_end: null,
};
const inspection: JuryPanelImportInspection = {
  sheets: [], selected_sheet: null, columns: ["Date", "Panel", "Start", "Length"],
  sample_rows: [{ Date: "2026-12-01", Panel: "Voice", Start: "9:00", Length: "20" }],
  suggested_mapping: mapping,
};

describe("Jury Panel import preview", () => {
  const scroll = vi.fn();

  beforeEach(() => {
    vi.clearAllMocks();
    Object.defineProperty(HTMLDialogElement.prototype, "showModal", { configurable: true, value(this: HTMLDialogElement) { this.setAttribute("open", ""); } });
    Object.defineProperty(HTMLElement.prototype, "scrollIntoView", { configurable: true, value: scroll });
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "panels.csv"));
    vi.mocked(api.inspectJuryPanelFile).mockResolvedValue(inspection);
    vi.mocked(api.previewJuryPanelImport).mockResolvedValue({
      panels: [{
        panel_name: "Voice", room: "", jury_date: "2026-12-01", earliest_start_minute: 540,
        preferred_start_minute: null, jury_length_minutes: 20,
      }],
      panels_to_remove: 2,
      assignments_to_clear: 3,
    });
  });

  afterEach(cleanup);

  it("requires Preview Import before applying, then scrolls to the preview", async () => {
    const user = userEvent.setup();
    render(<JuryPanelImportDialog onApplied={vi.fn()} />);
    await user.click(screen.getByRole("button", { name: "Import jury panels" }));
    await user.click(screen.getByRole("button", { name: "Choose spreadsheet" }));
    await screen.findByRole("button", { name: "Preview Import" });
    expect(screen.getAllByRole("button", { name: "Import jury panels" })).toHaveLength(1);

    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    const region = await screen.findByRole("region", { name: "Import preview" });
    expect(region.textContent).toContain("1 Panel will be imported");
    expect(region.textContent).toContain("2 existing Panels and 3 Lesson Panel assignments");
    expect(scroll).toHaveBeenCalled();
    expect(api.applyJuryPanelImport).not.toHaveBeenCalled();
  });
});
