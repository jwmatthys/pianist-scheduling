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
    expect(screen.getByText("Not yet configured")).toBeTruthy();
    expect(screen.getByText("2 Availability Windows")).toBeTruthy();
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
  columns: ["ID", "Who", "Mail"],
  sample_rows: [],
  suggested_mapping: { student_id: "ID", name: null, email: "Mail" },
};

const previewOf = (errors: string[]): ClinicalStudentsImportPreview => ({
  selected_sheet: null,
  total_rows: 2,
  valid_rows: errors.length ? 1 : 2,
  invalid_rows: errors.length ? 1 : 0,
  rows: [
    { row_number: 2, student_id: "S1", name: "Ada", email: "", errors: [] },
    { row_number: 3, student_id: "S2", name: "Ben", email: "", errors },
  ],
  existing_student_count: 4,
  existing_placement_count: 3,
  can_apply: errors.length === 0,
});

async function openMapping(user: ReturnType<typeof userEvent.setup>) {
  render(<ClinicalStudentRosterPage />);
  await user.click(await screen.findByRole("button", { name: "Import Student Roster" }));
  await user.click(screen.getByRole("button", { name: "Choose spreadsheet" }));
  await screen.findByRole("heading", { name: "Column mapping" });
  await user.selectOptions(screen.getByLabelText("Name"), "Who");
}

describe("Student Roster import dialog", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.listClinicalStudents).mockResolvedValue([]);
    vi.mocked(api.listClinicalTags).mockResolvedValue([]);
    vi.mocked(api.inspectClinicalStudentsFile).mockResolvedValue(inspection);
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "roster.csv"));
  });

  it("suggests an editable mapping and gates Preview on required fields", async () => {
    const user = userEvent.setup();
    render(<ClinicalStudentRosterPage />);
    await user.click(await screen.findByRole("button", { name: "Import Student Roster" }));
    await user.click(screen.getByRole("button", { name: "Choose spreadsheet" }));
    await screen.findByRole("heading", { name: "Column mapping" });
    expect((screen.getByLabelText("Student ID") as HTMLSelectElement).value).toBe("ID");
    const preview = screen.getByRole("button", { name: "Preview Import" }) as HTMLButtonElement;
    expect(preview.disabled).toBe(true);
    await user.selectOptions(screen.getByLabelText("Name"), "Who");
    expect(preview.disabled).toBe(false);
  });

  it("renders preview rows and row-level validation, blocking import", async () => {
    vi.mocked(api.previewClinicalStudentsImport).mockResolvedValue(previewOf(["Student ID is duplicated in this file (also on row 2)."]));
    const user = userEvent.setup();
    await openMapping(user);
    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    const region = await screen.findByRole("region", { name: "Import preview" });
    expect(within(region).getByText("2 rows: 1 valid, 1 with errors.")).toBeTruthy();
    expect(within(region).getByText(/duplicated in this file/)).toBeTruthy();
    expect((within(region).getByRole("button", { name: "Import Student Roster" }) as HTMLButtonElement).disabled).toBe(true);
  });

  it("shows the replacement warning, can be cancelled, then imports on confirmation", async () => {
    vi.mocked(api.previewClinicalStudentsImport).mockResolvedValue(previewOf([]));
    vi.mocked(api.applyClinicalStudentsImport).mockResolvedValue({
      students_removed: 4,
      placements_removed: 3,
      students_created: 2,
    });
    const user = userEvent.setup();
    await openMapping(user);
    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    const region = await screen.findByRole("region", { name: "Import preview" });
    await user.click(within(region).getByRole("button", { name: "Import Student Roster" }));
    const confirm = screen.getByRole("dialog", { name: "Replace Student Roster?", hidden: true });
    for (const text of [
      /Delete all existing Clinical Students \(4 currently\)/,
      "Delete all student availability",
      "Delete all Required / Preferred / Restricted selections",
      /Delete all Clinical Placements \(3 currently\)/,
      "Replace the roster with imported Students",
      "Preserve Clinical Sites",
      "Preserve Clinical Tag registry",
    ]) {
      expect(within(confirm).getByText(text)).toBeTruthy();
    }
    await user.click(within(confirm).getByRole("button", { name: "Cancel" }));
    expect(api.applyClinicalStudentsImport).not.toHaveBeenCalled();

    await user.click(within(region).getByRole("button", { name: "Import Student Roster" }));
    const again = screen.getByRole("dialog", { name: "Replace Student Roster?", hidden: true });
    await user.click(within(again).getByRole("button", { name: "Replace Student Roster" }));
    await waitFor(() => expect(api.applyClinicalStudentsImport).toHaveBeenCalledTimes(1));
    expect(await screen.findByText(/Imported 2 students; replaced 4 and removed 3 placement\(s\)/)).toBeTruthy();
  });
});
