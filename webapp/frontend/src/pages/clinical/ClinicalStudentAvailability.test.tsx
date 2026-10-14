import { fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import type { ClinicalStudent } from "../../lib/types";
import { ClinicalStudentRosterPage } from "./ClinicalStudentRosterPage";

vi.mock("../../lib/api", () => ({
  api: {
    listClinicalStudents: vi.fn(),
    listClinicalTags: vi.fn(),
    getClinicalStudentAvailability: vi.fn(),
    setClinicalStudentAvailability: vi.fn(),
  },
}));
vi.mock("../../lib/platform", () => ({ chooseLocalFile: vi.fn() }));

const student = (count = 0): ClinicalStudent => ({
  id: "s1",
  student_id: "S100",
  name: "Ada Example",
  email: "",
  availability_window_count: count,
  tag_states: {},
  created_at: "2026-01-01T00:00:00",
  updated_at: "2026-01-01T00:00:00",
});

const cell = (dialog: HTMLElement, name: string) => within(dialog).getByRole("button", { name, hidden: true });

describe("Clinical student availability editing", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.listClinicalStudents).mockResolvedValue([student()]);
    vi.mocked(api.listClinicalTags).mockResolvedValue([]);
    vi.mocked(api.getClinicalStudentAvailability).mockResolvedValue({ windows: [] });
  });

  it("loads existing windows, saves recurring windows with unmarked as Unavailable, and refreshes the count", async () => {
    vi.mocked(api.getClinicalStudentAvailability).mockResolvedValue({
      windows: [{ day: "Monday", start_minute: 540, end_minute: 600, status: "Available" }],
    });
    vi.mocked(api.setClinicalStudentAvailability).mockResolvedValue({ windows: [] });
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    expect(await screen.findByText("No availability configured")).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "Edit availability for Ada Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Availability: Ada Example", hidden: true });
    await waitFor(() => expect(cell(dialog, "Monday 9:00 AM: Available")).toBeTruthy());

    fireEvent.mouseDown(cell(dialog, "Monday 10:00 AM: Unmarked"), { button: 0 });
    fireEvent.mouseUp(window);
    vi.mocked(api.listClinicalStudents).mockResolvedValue([student(5)]);
    await user.click(within(dialog).getByRole("button", { name: "Save Availability", hidden: true }));

    await waitFor(() => expect(api.setClinicalStudentAvailability).toHaveBeenCalled());
    const [id, windows] = vi.mocked(api.setClinicalStudentAvailability).mock.calls[0];
    expect(id).toBe("s1");
    expect(windows.filter((item) => item.day === "Monday")).toEqual([
      { day: "Monday", start_minute: 420, end_minute: 540, status: "Unavailable" },
      { day: "Monday", start_minute: 540, end_minute: 630, status: "Available" },
      { day: "Monday", start_minute: 630, end_minute: 1260, status: "Unavailable" },
    ]);
    expect(windows.find((item) => item.day === "Tuesday")).toEqual({
      day: "Tuesday", start_minute: 420, end_minute: 1260, status: "Unavailable",
    });
    expect(await screen.findByText("Availability entered")).toBeTruthy();
  });

  it("clears the grid and saves no windows so availability returns to not configured", async () => {
    vi.mocked(api.getClinicalStudentAvailability).mockResolvedValue({
      windows: [{ day: "Monday", start_minute: 540, end_minute: 600, status: "Available" }],
    });
    vi.mocked(api.setClinicalStudentAvailability).mockResolvedValue({ windows: [] });
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Edit availability for Ada Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Availability: Ada Example", hidden: true });
    await waitFor(() => expect(cell(dialog, "Monday 9:00 AM: Available")).toBeTruthy());
    await user.click(within(dialog).getByRole("button", { name: "Clear Availability", hidden: true }));
    expect(cell(dialog, "Monday 9:00 AM: Unmarked")).toBeTruthy();
    await user.click(within(dialog).getByRole("button", { name: "Save Availability", hidden: true }));
    await waitFor(() => expect(api.setClinicalStudentAvailability).toHaveBeenCalledWith("s1", []));
  });

  it("offers Edit Availability in the student dialog for saved students", async () => {
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Edit Ada Example" }));
    const studentDialog = screen.getByRole("dialog", { name: "Edit Student", hidden: true });
    await user.click(within(studentDialog).getByRole("button", { name: "Edit Availability", hidden: true }));
    expect(await screen.findByRole("dialog", { name: "Availability: Ada Example", hidden: true })).toBeTruthy();
  });

  it("disables Edit Availability for a student not yet saved", async () => {
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Add Student" }));
    const dialog = screen.getByRole("dialog", { name: "Add Student", hidden: true });
    expect((within(dialog).getByRole("button", { name: "Edit Availability", hidden: true }) as HTMLButtonElement).disabled).toBe(true);
  });
});
