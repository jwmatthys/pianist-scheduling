import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { JuryLessonEntry, JuryReadiness } from "../lib/types";

const apiMocks = vi.hoisted(() => ({
  getAccompanistFinalizationState: vi.fn(),
  finalizeAccompanist: vi.fn(),
  getJuryPanels: vi.fn(),
  synchronizeJuryRoster: vi.fn(),
  getJuryReadiness: vi.fn(),
}));

vi.mock("../lib/api", () => ({ api: apiMocks }));
vi.mock("./JuryReportsPage", () => ({
  JuryReportsPage: () => <div data-testid="jury-reports-panel">Jury reports content</div>,
}));

import { JurySetupPage } from "./JurySetupPage";

function sourceEntry(overrides: Partial<JuryLessonEntry> = {}): JuryLessonEntry {
  return {
    session_uuid: "session-1",
    source_lesson_uuid: "lesson-1",
    student_person_uuid: "student-1",
    student_display_name: "Synthetic Student",
    instrument: "Voice",
    teacher: "Synthetic Teacher",
    pianist_required: false,
    assigned_pianist: null,
    jury_required: true,
    panel_uuid: null,
    ...overrides,
  };
}

function blockedReadiness(): JuryReadiness {
  return {
    ready: false,
    source_result_uuid: "result-1",
    source_revision: 1,
    jury_input_revision: 1,
    issues: [{
      code: "PANEL_REQUIRED",
      severity: "error",
      message: "Select a Jury Panel for this Jury-required lesson.",
      entity_uuids: ["lesson-1"],
    }],
  };
}

describe("Jury Setup workflow", () => {
  beforeEach(() => {
    vi.resetAllMocks();
    apiMocks.getAccompanistFinalizationState
      .mockResolvedValueOnce({
        session_uuid: "session-1",
        source_revision: 1,
        current_result_uuid: null,
        current_result_version: null,
      })
      .mockResolvedValue({
        session_uuid: "session-1",
        source_revision: 1,
        current_result_uuid: "result-1",
        current_result_version: 1,
      });
    apiMocks.finalizeAccompanist.mockResolvedValue({});
    apiMocks.getJuryPanels.mockResolvedValue([]);
    apiMocks.synchronizeJuryRoster.mockResolvedValue([sourceEntry()]);
    apiMocks.getJuryReadiness.mockResolvedValue(blockedReadiness());
  });

  it("opens Jury Panels first and silently finalizes/synchronizes when no current source exists", async () => {
    const confirm = vi.spyOn(window, "confirm").mockReturnValue(false);
    render(<JurySetupPage />);

    const tabs = screen.getByRole("navigation", { name: "Jury Scheduling views" });
    expect(Array.from(tabs.querySelectorAll("button")).map((button) => button.textContent)).toEqual([
      "Jury Panels", "Lesson Roster", "Pianist Availability", "Schedule", "Reports",
    ]);
    expect(screen.getByRole("button", { name: "Jury Panels" }).getAttribute("aria-current")).toBe("page");
    expect(screen.queryByRole("button", { name: "Overview" })).toBeNull();

    await waitFor(() => expect(apiMocks.finalizeAccompanist).toHaveBeenCalledWith(1));
    expect(apiMocks.synchronizeJuryRoster).toHaveBeenCalledTimes(1);
    expect(confirm).not.toHaveBeenCalled();
    expect(screen.queryByRole("button", { name: /Finalize Accompanist schedule/i })).toBeNull();
  });

  it("orders roster columns and highlights Jury-required lessons without a Panel", async () => {
    render(<JurySetupPage />);
    await waitFor(() => expect(apiMocks.synchronizeJuryRoster).toHaveBeenCalled());
    fireEvent.click(screen.getByRole("button", { name: "Lesson Roster" }));

    expect(screen.getAllByRole("columnheader").map((header) => header.textContent)).toEqual([
      "Student", "Teacher", "Pianist", "Jury Required", "Lesson / instrument", "Panel",
    ]);
    const lessonRow = screen.getByRole("row", { name: /Synthetic Student/ }) as HTMLTableRowElement;
    expect(lessonRow.classList.contains("jury-entry-unassigned-panel")).toBe(true);
    expect(lessonRow.cells[4].textContent).toBe("Voice");
  });

  it("keeps Jury Reports available as a module tab", async () => {
    render(<JurySetupPage />);
    fireEvent.click(screen.getByRole("button", { name: "Reports" }));
    await waitFor(() => expect(apiMocks.synchronizeJuryRoster).toHaveBeenCalled());
    expect(screen.getByTestId("jury-reports-panel")).toBeTruthy();
  });
});