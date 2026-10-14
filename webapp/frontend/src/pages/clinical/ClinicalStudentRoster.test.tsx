import { render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import type {
  ClinicalStudent,
  ClinicalStudentsImportInspection,
  ClinicalStudentsImportPreview,
  ClinicalTag,
} from "../../lib/types";
import { ClinicalStudentRosterPage } from "./ClinicalStudentRosterPage";

vi.mock("../../lib/api", () => ({
  api: {
    listClinicalStudents: vi.fn(),
    listClinicalTags: vi.fn(),
    createClinicalStudent: vi.fn(),
    updateClinicalStudent: vi.fn(),
    setClinicalStudentTagStates: vi.fn(),
    deleteClinicalStudent: vi.fn(),
    clearClinicalStudents: vi.fn(),
    inspectClinicalStudentsFile: vi.fn(),
    previewClinicalStudentsImport: vi.fn(),
    applyClinicalStudentsImport: vi.fn(),
  },
}));
vi.mock("../../lib/platform", () => ({ chooseLocalFile: vi.fn() }));

const tag = (id: string, name: string): ClinicalTag => ({
  id,
  name,
  opportunity_count: 1,
  student_count: 0,
  is_unused: false,
});
const kids = tag("t1", "KIDS");
const adults = tag("t2", "ADULTS");

const student = (overrides: Partial<ClinicalStudent> = {}): ClinicalStudent => ({
  id: "s1",
  student_id: "S100",
  name: "Ada Example",
  email: "ada@example.test",
  availability_window_count: 0,
  tag_states: {},
  created_at: "2026-01-01T00:00:00",
  updated_at: "2026-01-01T00:00:00",
  ...overrides,
});

describe("Student Roster page", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.listClinicalStudents).mockResolvedValue([student()]);
    vi.mocked(api.listClinicalTags).mockResolvedValue([kids, adults]);
  });

  it("lists students with availability placeholder and tag summary", async () => {
    vi.mocked(api.listClinicalStudents).mockResolvedValue([
      student({ tag_states: { t1: "required" } }),
      student({ id: "s2", student_id: "S200", name: "Ben Sample", availability_window_count: 2 }),
    ]);
    render(<ClinicalStudentRosterPage />);
    expect(await screen.findByText("Ada Example")).toBeTruthy();
    expect(screen.getByText("No availability configured")).toBeTruthy();
    expect(screen.getByText("Availability entered")).toBeTruthy();
    expect(screen.getByText("No availability configured").closest("td")?.className).toContain("clinical-availability-missing");
    expect(screen.getByText("Availability entered").closest("td")?.className ?? "").not.toContain("clinical-availability-missing");
    expect(screen.getByText("KIDS: Required")).toBeTruthy();
  });

  it("adds a student with validation and persists tag states", async () => {
    vi.mocked(api.createClinicalStudent).mockResolvedValue(student({ id: "s9", student_id: "S9", name: "New One" }));
    vi.mocked(api.setClinicalStudentTagStates).mockResolvedValue(student({ id: "s9" }));
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Add Student" }));
    const dialog = screen.getByRole("dialog", { name: "Add Student", hidden: true });
    await user.click(within(dialog).getByRole("button", { name: "Add Student", hidden: true }));
    expect(within(dialog).getByText("Student ID is required.")).toBeTruthy();
    expect(within(dialog).getByText("Name is required.")).toBeTruthy();
    expect(api.createClinicalStudent).not.toHaveBeenCalled();

    await user.type(within(dialog).getByLabelText("Student ID"), "S9");
    await user.type(within(dialog).getByLabelText("Name"), "New One");
    await user.click(within(dialog).getByRole("button", { name: "KIDS: Neutral", hidden: true }));
    await user.click(within(dialog).getByRole("button", { name: "Add Student", hidden: true }));
    await waitFor(() => expect(api.createClinicalStudent).toHaveBeenCalledWith({ student_id: "S9", name: "New One", email: "" }));
    expect(api.setClinicalStudentTagStates).toHaveBeenCalledWith("s9", { t1: "required" });
  });

  it("shows the legend and cycles tag states Neutral, Required, Preferred, Restricted, Neutral", async () => {
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Edit Ada Example" }));
    const dialog = screen.getByRole("dialog", { name: "Edit Student", hidden: true });
    const legend = within(dialog).getByRole("list", { name: "Tag state legend", hidden: true });
    for (const label of ["Neutral", "Required", "Preferred", "Restricted"]) {
      expect(within(legend).getByText(label)).toBeTruthy();
    }
    for (const expected of ["Required", "Preferred", "Restricted", "Neutral"]) {
      const current = within(dialog).getByRole("button", { name: /^KIDS: /, hidden: true });
      await user.click(current);
      expect(within(dialog).getByRole("button", { name: `KIDS: ${expected}`, hidden: true }).className).toContain(
        `clinical-state-${expected.toLowerCase()}`,
      );
    }
  });

  it("edits a student and saves changed tag states only when changed", async () => {
    vi.mocked(api.updateClinicalStudent).mockResolvedValue(student({ name: "Ada Renamed" }));
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Edit Ada Example" }));
    const dialog = screen.getByRole("dialog", { name: "Edit Student", hidden: true });
    const name = within(dialog).getByLabelText("Name");
    await user.clear(name);
    await user.type(name, "Ada Renamed");
    await user.click(within(dialog).getByRole("button", { name: "Save Student", hidden: true }));
    await waitFor(() =>
      expect(api.updateClinicalStudent).toHaveBeenCalledWith("s1", {
        student_id: "S100",
        name: "Ada Renamed",
        email: "ada@example.test",
      }),
    );
    expect(api.setClinicalStudentTagStates).not.toHaveBeenCalled();
  });

  it("shows server errors such as duplicate Student IDs", async () => {
    vi.mocked(api.updateClinicalStudent).mockRejectedValue(
      new Error("409: DUPLICATE_STUDENT_ID: A student with that Student ID already exists."),
    );
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Edit Ada Example" }));
    const dialog = screen.getByRole("dialog", { name: "Edit Student", hidden: true });
    await user.click(within(dialog).getByRole("button", { name: "Save Student", hidden: true }));
    expect(await within(dialog).findByText(/already exists/)).toBeTruthy();
  });

  it("deletes all students only after confirmation", async () => {
    vi.mocked(api.clearClinicalStudents).mockResolvedValue({ removed_count: 1, removed_placement_count: 0 });
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Delete all students" }));
    const first = screen.getByRole("dialog", { name: "Delete all students?", hidden: true });
    await user.click(within(first).getByRole("button", { name: "Cancel", hidden: true }));
    expect(api.clearClinicalStudents).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "Delete all students" }));
    const confirm = screen.getByRole("dialog", { name: "Delete all students?", hidden: true });
    await user.click(within(confirm).getByRole("button", { name: "Delete all students", hidden: true }));
    await waitFor(() => expect(api.clearClinicalStudents).toHaveBeenCalledTimes(1));
    expect(await screen.findByText("Deleted 1 student.")).toBeTruthy();
  });

  it("deletes a student after confirmation", async () => {
    vi.mocked(api.deleteClinicalStudent).mockResolvedValue(undefined);
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Delete Ada Example" }));
    const confirm = screen.getByRole("dialog", { name: "Delete Ada Example?", hidden: true });
    await user.click(within(confirm).getByRole("button", { name: "Delete Student", hidden: true }));
    await waitFor(() => expect(api.deleteClinicalStudent).toHaveBeenCalledWith("s1"));
  });
});

const inspection: ClinicalStudentsImportInspection = {
  sheets: [],
  selected_sheet: null,
  columns: ["ID", "Who", "Mail", "Mon", "Tue"],
  sample_rows: [],
  suggested_mapping: { student_id: "ID", name: null, email: "Mail", days: { Monday: "Mon", Tuesday: "Tue" } },
};

const previewOf = (errors: string[]): ClinicalStudentsImportPreview => ({
  selected_sheet: null,
  total_rows: 2,
  valid_rows: errors.length ? 1 : 2,
  invalid_rows: errors.length ? 1 : 0,
  students_to_create: 1,
  students_to_update: 1,
  availability_mapped: true,
  rows: [
    {
      row_number: 2, student_id: "S1", name: "Ada", email: "", action: "update",
      student_change: "Update email", availability_change: "Replace 5 availability windows with 8",
      warnings: [], errors: [],
    },
    {
      row_number: 3, student_id: "S2", name: "Ben", email: "", action: "create",
      student_change: "Create student", availability_change: "Clear availability",
      warnings: [], errors,
    },
  ],
  can_apply: errors.length === 0,
});

async function openMapping(user: ReturnType<typeof userEvent.setup>) {
  render(<ClinicalStudentRosterPage />);
  await user.click(await screen.findByRole("button", { name: "Import Student Roster" }));
  await user.click(screen.getByRole("button", { name: "Choose spreadsheet" }));
  await screen.findByRole("heading", { name: "Column mapping" });
}

describe("Student Roster import dialog", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.listClinicalStudents).mockResolvedValue([]);
    vi.mocked(api.listClinicalTags).mockResolvedValue([]);
    vi.mocked(api.inspectClinicalStudentsFile).mockResolvedValue(inspection);
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "roster.csv"));
  });

  it("no longer offers a standalone Import Availability button", async () => {
    render(<ClinicalStudentRosterPage />);
    await screen.findByRole("button", { name: "Import Student Roster" });
    expect(screen.queryByRole("button", { name: "Import Availability" })).toBeNull();
  });

  it("maps identity and weekday columns in one dialog, gating Preview on Student ID only", async () => {
    const user = userEvent.setup();
    await openMapping(user);
    expect((screen.getByLabelText("Student ID") as HTMLSelectElement).value).toBe("ID");
    expect((screen.getByLabelText("Monday") as HTMLSelectElement).value).toBe("Mon");
    expect((screen.getByLabelText("Wednesday") as HTMLSelectElement).value).toBe("");
    const preview = screen.getByRole("button", { name: "Preview Import" }) as HTMLButtonElement;
    expect(preview.disabled).toBe(false);
    await user.selectOptions(screen.getByLabelText("Student ID"), "");
    expect(preview.disabled).toBe(true);
  });

  it("sends the unified mapping including weekday columns", async () => {
    vi.mocked(api.previewClinicalStudentsImport).mockResolvedValue(previewOf([]));
    const user = userEvent.setup();
    await openMapping(user);
    await user.selectOptions(screen.getByLabelText("Name"), "Who");
    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    await screen.findByRole("region", { name: "Import preview" });
    expect(api.previewClinicalStudentsImport).toHaveBeenCalledWith(expect.anything(), null, {
      student_id: "ID", name: "Who", email: "Mail", days: { Monday: "Mon", Tuesday: "Tue" },
    });
  });

  it("renders roster and availability effects plus row-level validation, blocking import", async () => {
    vi.mocked(api.previewClinicalStudentsImport).mockResolvedValue(previewOf(["Student ID is duplicated in this file (also on row 2)."]));
    const user = userEvent.setup();
    await openMapping(user);
    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    const region = await screen.findByRole("region", { name: "Import preview" });
    expect(within(region).getByText("2 rows: 1 valid, 1 with errors.")).toBeTruthy();
    expect(within(region).getByText("Update email")).toBeTruthy();
    expect(within(region).getByText("Replace 5 availability windows with 8")).toBeTruthy();
    expect(within(region).getByText("Create student")).toBeTruthy();
    expect(within(region).getByText("Clear availability")).toBeTruthy();
    expect(within(region).getByText(/duplicated in this file/)).toBeTruthy();
    expect((within(region).getByRole("button", { name: "Import Student Roster" }) as HTMLButtonElement).disabled).toBe(true);
  });

  it("shows a non-destructive confirmation, can be cancelled, then imports", async () => {
    vi.mocked(api.previewClinicalStudentsImport).mockResolvedValue(previewOf([]));
    vi.mocked(api.applyClinicalStudentsImport).mockResolvedValue({
      students_created: 1,
      students_updated: 1,
      availability_updated: 2,
      windows_imported: 8,
    });
    const user = userEvent.setup();
    await openMapping(user);
    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    const region = await screen.findByRole("region", { name: "Import preview" });
    await user.click(within(region).getByRole("button", { name: "Import Student Roster" }));
    const confirm = screen.getByRole("dialog", { name: "Import Student Roster?", hidden: true });
    for (const text of [
      "Create 1 new student(s)",
      "Update 1 existing student(s) matched by Student ID",
      "Replace weekly availability for the students listed in the file",
      "Preserve students not listed in the file",
      "Preserve Clinical Sites and the Clinical Tag registry",
      "Preserve tag selections and placements",
      "Mark any generated schedule as out of date",
    ]) {
      expect(within(confirm).getByText(text)).toBeTruthy();
    }
    expect(within(confirm).queryByText(/Delete/)).toBeNull();
    await user.click(within(confirm).getByRole("button", { name: "Cancel" }));
    expect(api.applyClinicalStudentsImport).not.toHaveBeenCalled();

    await user.click(within(region).getByRole("button", { name: "Import Student Roster" }));
    const again = screen.getByRole("dialog", { name: "Import Student Roster?", hidden: true });
    await user.click(within(again).getByRole("button", { name: "Import Student Roster" }));
    await waitFor(() => expect(api.applyClinicalStudentsImport).toHaveBeenCalledTimes(1));
    expect(
      await screen.findByText(/Imported 2 students \(1 created, 1 updated\); replaced availability for 2 students\./),
    ).toBeTruthy();
  });
});
