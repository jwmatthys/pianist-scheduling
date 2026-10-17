import { fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import type {
  LessonStudent,
  LessonStudentImportInspection,
  LessonStudentImportPreview,
  LessonTag,
} from "../../lib/types";
import { LessonStudentRosterPage } from "./LessonStudentRosterPage";

vi.mock("../../lib/api", () => ({
  api: {
    listLessonStudents: vi.fn(),
    listLessonTags: vi.fn(),
    createLessonStudent: vi.fn(),
    updateLessonStudent: vi.fn(),
    deleteLessonStudent: vi.fn(),
    getLessonStudentAvailability: vi.fn(),
    setLessonStudentAvailability: vi.fn(),
    inspectLessonStudentFile: vi.fn(),
    previewLessonStudentImport: vi.fn(),
    applyLessonStudentImport: vi.fn(),
  },
}));
vi.mock("../../lib/platform", () => ({ chooseLocalFile: vi.fn() }));

const tag = (id: string, name: string): LessonTag => ({
  id,
  name,
  instructor_count: 0,
  student_count: 0,
  is_unused: true,
});

const student = (overrides: Partial<LessonStudent> = {}): LessonStudent => ({
  id: "s1",
  student_id: "S-001",
  name: "Sam Example",
  email: "sam@example.edu",
  availability_window_count: 0,
  tag_states: {},
  created_at: "2026-01-01T00:00:00",
  updated_at: "2026-01-01T00:00:00",
  ...overrides,
});

const inspection: LessonStudentImportInspection = {
  sheets: [],
  selected_sheet: null,
  columns: ["ID", "Name", "Mon"],
  sample_rows: [],
  suggested_mapping: { student_id: "ID", name: "Name", email: null, days: { Monday: "Mon" } },
};

const preview: LessonStudentImportPreview = {
  selected_sheet: null,
  total_rows: 1,
  valid_rows: 1,
  invalid_rows: 0,
  availability_mapped: true,
  rows: [
    {
      row_number: 2,
      student_id: "S-100",
      name: "Imported Student",
      email: "",
      availability_window_count: 1,
      availability: ["Mon 9:00 AM-12:00 PM"],
      warnings: [],
      errors: [],
    },
  ],
  existing_student_count: 5,
  existing_placement_count: 3,
  can_apply: true,
};

const cell = (dialog: HTMLElement, name: string) => within(dialog).getByRole("button", { name, hidden: true });

describe("Lesson Student Roster", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.listLessonStudents).mockResolvedValue([]);
    vi.mocked(api.listLessonTags).mockResolvedValue([]);
    vi.mocked(api.getLessonStudentAvailability).mockResolvedValue({ windows: [] });
  });

  it("validates then creates a student with cycled tag states in one request", async () => {
    vi.mocked(api.listLessonTags).mockResolvedValue([tag("t1", "VOICE")]);
    vi.mocked(api.createLessonStudent).mockResolvedValue(student());
    const user = userEvent.setup();
    render(<LessonStudentRosterPage />);
    expect(await screen.findByText("No students yet.")).toBeTruthy();
    await user.click(screen.getByRole("button", { name: "Add Student" }));
    const dialog = await screen.findByRole("dialog", { name: "Add Student" });

    await user.click(within(dialog).getByRole("button", { name: "Add Student" }));
    expect(within(dialog).getByText("Student ID is required.")).toBeTruthy();
    expect(within(dialog).getByText("Name is required.")).toBeTruthy();
    expect(api.createLessonStudent).not.toHaveBeenCalled();

    expect(within(dialog).getByLabelText("Tag state legend")).toBeTruthy();
    await user.type(within(dialog).getByLabelText("Student ID"), "S-009");
    await user.type(within(dialog).getByLabelText("Name"), "New Student");
    const tagButton = () => within(dialog).getByRole("button", { name: /^VOICE:/ });
    expect(tagButton().getAttribute("aria-label")).toBe("VOICE: Neutral");
    await user.click(tagButton());
    expect(tagButton().getAttribute("aria-label")).toBe("VOICE: Required");
    await user.click(tagButton());
    expect(tagButton().getAttribute("aria-label")).toBe("VOICE: Preferred");
    await user.click(tagButton());
    expect(tagButton().getAttribute("aria-label")).toBe("VOICE: Restricted");
    await user.click(within(dialog).getByRole("button", { name: "Add Student" }));

    await waitFor(() =>
      expect(api.createLessonStudent).toHaveBeenCalledWith({
        student_id: "S-009",
        name: "New Student",
        email: "",
        tag_states: { t1: "restricted" },
      }),
    );
  });

  it("edits a student and sends tag states with the update", async () => {
    vi.mocked(api.listLessonStudents).mockResolvedValue([student({ tag_states: { t1: "required" } })]);
    vi.mocked(api.listLessonTags).mockResolvedValue([tag("t1", "VOICE")]);
    vi.mocked(api.updateLessonStudent).mockResolvedValue(student());
    const user = userEvent.setup();
    render(<LessonStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Edit Sam Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Edit Student" });
    await user.click(within(dialog).getByRole("button", { name: "VOICE: Required" }));
    await user.click(within(dialog).getByRole("button", { name: "Save Student" }));
    await waitFor(() =>
      expect(api.updateLessonStudent).toHaveBeenCalledWith("s1", {
        student_id: "S-001",
        name: "Sam Example",
        email: "sam@example.edu",
        tag_states: { t1: "preferred" },
      }),
    );
  });

  it("deletes a student only after confirmation", async () => {
    vi.mocked(api.listLessonStudents).mockResolvedValue([student()]);
    vi.mocked(api.deleteLessonStudent).mockResolvedValue(undefined);
    const user = userEvent.setup();
    render(<LessonStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Delete Sam Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Delete Sam Example?" });
    await user.click(within(dialog).getByRole("button", { name: "Cancel" }));
    expect(api.deleteLessonStudent).not.toHaveBeenCalled();

    await user.click(screen.getByRole("button", { name: "Delete Sam Example" }));
    await user.click(
      within(await screen.findByRole("dialog", { name: "Delete Sam Example?" })).getByRole("button", {
        name: "Delete Student",
      }),
    );
    await waitFor(() => expect(api.deleteLessonStudent).toHaveBeenCalledWith("s1"));
  });

  it("edits availability from the row and from the edit dialog, saving recurring windows", async () => {
    vi.mocked(api.listLessonStudents).mockResolvedValue([student()]);
    vi.mocked(api.getLessonStudentAvailability).mockResolvedValue({
      windows: [{ day: "Monday", start_minute: 540, end_minute: 600, status: "Available" }],
    });
    vi.mocked(api.setLessonStudentAvailability).mockResolvedValue({ windows: [] });
    const user = userEvent.setup();
    render(<LessonStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Edit availability for Sam Example" }));
    const dialog = await screen.findByRole("dialog", { name: "Availability: Sam Example", hidden: true });
    await waitFor(() => expect(cell(dialog, "Monday 9:00 AM: Available")).toBeTruthy());
    fireEvent.mouseDown(cell(dialog, "Monday 10:00 AM: Unmarked"), { button: 0 });
    fireEvent.mouseUp(window);
    await user.click(within(dialog).getByRole("button", { name: "Save Availability", hidden: true }));
    await waitFor(() => expect(api.setLessonStudentAvailability).toHaveBeenCalled());
    const [id, windows] = vi.mocked(api.setLessonStudentAvailability).mock.calls[0];
    expect(id).toBe("s1");
    expect(windows.filter((item) => item.day === "Monday")).toEqual([
      { day: "Monday", start_minute: 420, end_minute: 540, status: "Unavailable" },
      { day: "Monday", start_minute: 540, end_minute: 630, status: "Available" },
      { day: "Monday", start_minute: 630, end_minute: 1260, status: "Unavailable" },
    ]);

    await user.click(screen.getByRole("button", { name: "Edit Sam Example" }));
    const edit = await screen.findByRole("dialog", { name: "Edit Student" });
    await user.click(within(edit).getByRole("button", { name: "Edit Availability" }));
    expect(await screen.findByRole("dialog", { name: "Availability: Sam Example", hidden: true })).toBeTruthy();
  });

  it("imports: choose file, map, preview, confirm replacement text, then apply", async () => {
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "s.csv"));
    vi.mocked(api.inspectLessonStudentFile).mockResolvedValue(inspection);
    vi.mocked(api.previewLessonStudentImport).mockResolvedValue(preview);
    vi.mocked(api.applyLessonStudentImport).mockResolvedValue({
      students_removed: 5,
      placements_removed: 3,
      students_created: 1,
    });
    const user = userEvent.setup();
    render(<LessonStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Import Student Roster" }));
    const dialog = await screen.findByRole("dialog", { name: "Import Student Roster" });
    await user.click(within(dialog).getByRole("button", { name: "Choose spreadsheet" }));
    expect(await within(dialog).findByLabelText("Student ID")).toBeTruthy();
    await user.click(within(dialog).getByRole("button", { name: "Preview Import" }));
    expect(await within(dialog).findByText("Imported Student")).toBeTruthy();
    expect(within(dialog).getByText("Mon 9:00 AM-12:00 PM")).toBeTruthy();
    expect(api.applyLessonStudentImport).not.toHaveBeenCalled();

    await user.click(within(dialog).getByRole("button", { name: "Import Students" }));
    const confirm = await screen.findByRole("dialog", { name: "Replace Students?" });
    expect(within(confirm).getByText(/Delete all current students \(5 currently\)/)).toBeTruthy();
    expect(within(confirm).getByText(/Remove any existing lesson placements \(3 currently\)/)).toBeTruthy();
    expect(within(confirm).getByText("Preserve Instructor records and Instructor availability")).toBeTruthy();
    expect(within(confirm).getByText("Preserve the Lesson Tag registry")).toBeTruthy();
    expect(api.applyLessonStudentImport).not.toHaveBeenCalled();

    await user.click(within(confirm).getByRole("button", { name: "Replace Students" }));
    await waitFor(() => expect(api.applyLessonStudentImport).toHaveBeenCalled());
    expect(await screen.findByRole("status")).toBeTruthy();
  });

  it("blocks apply when the preview has errors", async () => {
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "s.csv"));
    vi.mocked(api.inspectLessonStudentFile).mockResolvedValue(inspection);
    vi.mocked(api.previewLessonStudentImport).mockResolvedValue({
      ...preview,
      valid_rows: 0,
      invalid_rows: 1,
      can_apply: false,
      rows: [{ ...preview.rows[0], errors: ["Student ID is required."] }],
    });
    const user = userEvent.setup();
    render(<LessonStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Import Student Roster" }));
    const dialog = await screen.findByRole("dialog", { name: "Import Student Roster" });
    await user.click(within(dialog).getByRole("button", { name: "Choose spreadsheet" }));
    await user.click(await within(dialog).findByRole("button", { name: "Preview Import" }));
    const apply = await within(dialog).findByRole("button", { name: "Import Students" });
    expect((apply as HTMLButtonElement).disabled).toBe(true);
  });
});
