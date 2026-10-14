import { render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { api } from "../../lib/api";
import { chooseLocalFile } from "../../lib/platform";
import type {
  ClinicalSitesImportInspection,
  ClinicalSitesImportPreview,
  ClinicalSitesImportPreviewRow,
} from "../../lib/types";
import { ClinicalSitesPage } from "./ClinicalSitesPage";

vi.mock("../../lib/api", () => ({
  api: {
    listClinicalOpportunities: vi.fn(),
    listClinicalTags: vi.fn(),
    inspectClinicalSitesFile: vi.fn(),
    previewClinicalSitesImport: vi.fn(),
    applyClinicalSitesImport: vi.fn(),
  },
}));
vi.mock("../../lib/platform", () => ({ chooseLocalFile: vi.fn() }));

const inspection: ClinicalSitesImportInspection = {
  sheets: [],
  selected_sheet: null,
  columns: ["Site", "Day", "Begin", "Finish", "Mins", "Focus"],
  sample_rows: [{ Site: "North Clinic", Day: "Mon", Begin: "9:00", Finish: "12:00", Mins: "45", Focus: "kids" }],
  suggested_mapping: {
    site_name: "Site",
    site_address: null,
    contact_name: null,
    contact_email: null,
    contact_phone: null,
    day_of_week: "Day",
    start_time: null,
    end_time: null,
    session_length: null,
    travel_time: null,
    capacity: null,
  },
};

const goodRow: ClinicalSitesImportPreviewRow = {
  row_number: 2,
  site_name: "North Clinic",
  day_of_week: "Monday",
  start_minute: 540,
  end_minute: 720,
  session_length_minutes: 45,
  travel_minutes: 0,
  capacity: 1,
  tags: ["KIDS"],
  errors: [],
};

function previewOf(rows: ClinicalSitesImportPreviewRow[]): ClinicalSitesImportPreview {
  const invalid = rows.filter((row) => row.errors.length).length;
  return {
    selected_sheet: null,
    tag_columns: ["Focus"],
    total_rows: rows.length,
    valid_rows: rows.length - invalid,
    invalid_rows: invalid,
    rows,
    new_tags: ["KIDS"],
    reused_tags: [],
    existing_opportunity_count: 3,
    existing_placement_count: 2,
    can_apply: invalid === 0,
  };
}

async function openMapping(user: ReturnType<typeof userEvent.setup>) {
  render(<ClinicalSitesPage />);
  await user.click(await screen.findByRole("button", { name: "Import Clinical Sites" }));
  await user.click(screen.getByRole("button", { name: "Choose spreadsheet" }));
  await screen.findByRole("heading", { name: "Column mapping" });
}

async function mapRequired(user: ReturnType<typeof userEvent.setup>) {
  await user.selectOptions(screen.getByLabelText("Start Time"), "Begin");
  await user.selectOptions(screen.getByLabelText("End Time"), "Finish");
  await user.selectOptions(screen.getByLabelText("Session Length"), "Mins");
}

describe("Clinical Sites import dialog", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(api.listClinicalOpportunities).mockResolvedValue([]);
    vi.mocked(api.listClinicalTags).mockResolvedValue([]);
    vi.mocked(api.inspectClinicalSitesFile).mockResolvedValue(inspection);
    vi.mocked(chooseLocalFile).mockResolvedValue(new File(["x"], "sites.csv"));
  });

  it("shows suggested column mapping, editable, and gates Preview on required fields", async () => {
    const user = userEvent.setup();
    await openMapping(user);
    expect((screen.getByLabelText("Site Name") as HTMLSelectElement).value).toBe("Site");
    expect((screen.getByLabelText("Start Time") as HTMLSelectElement).value).toBe("");
    const preview = screen.getByRole("button", { name: "Preview Import" }) as HTMLButtonElement;
    expect(preview.disabled).toBe(true);
    await mapRequired(user);
    expect(preview.disabled).toBe(false);
    expect(screen.getByText(/Every column not mapped above is read as tags/)).toBeTruthy();
  });

  it("renders the preview with 12-hour times, tags, and tag registry impact", async () => {
    vi.mocked(api.previewClinicalSitesImport).mockResolvedValue(previewOf([goodRow]));
    const user = userEvent.setup();
    await openMapping(user);
    await mapRequired(user);
    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    const region = await screen.findByRole("region", { name: "Import preview" });
    expect(within(region).getByText("1 row: 1 valid, 0 with errors.")).toBeTruthy();
    expect(within(region).getByText("9:00 AM")).toBeTruthy();
    expect(within(region).getByText("12:00 PM")).toBeTruthy();
    expect(within(region).getAllByText(/KIDS/).length).toBeGreaterThan(0);
    expect((within(region).getByRole("button", { name: "Import Clinical Sites" }) as HTMLButtonElement).disabled).toBe(false);
    const sent = vi.mocked(api.previewClinicalSitesImport).mock.calls[0][2];
    expect(sent.start_time).toBe("Begin");
  });

  it("shows row-level validation messages and blocks import", async () => {
    const bad = { ...goodRow, row_number: 3, day_of_week: null, errors: ["Day of Week 'Funday' is not a recognized weekday."] };
    vi.mocked(api.previewClinicalSitesImport).mockResolvedValue(previewOf([goodRow, bad]));
    const user = userEvent.setup();
    await openMapping(user);
    await mapRequired(user);
    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    const region = await screen.findByRole("region", { name: "Import preview" });
    expect(within(region).getByText("Day of Week 'Funday' is not a recognized weekday.")).toBeTruthy();
    expect(within(region).getByText("2 rows: 1 valid, 1 with errors.")).toBeTruthy();
    expect((within(region).getByRole("button", { name: "Import Clinical Sites" }) as HTMLButtonElement).disabled).toBe(true);
    expect(screen.getByText(/Fix the rows marked below in the source file/)).toBeTruthy();
  });

  it("requires explicit confirmation showing the replacement warning, and can be cancelled", async () => {
    vi.mocked(api.previewClinicalSitesImport).mockResolvedValue(previewOf([goodRow]));
    const user = userEvent.setup();
    await openMapping(user);
    await mapRequired(user);
    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    const region = await screen.findByRole("region", { name: "Import preview" });
    await user.click(within(region).getByRole("button", { name: "Import Clinical Sites" }));
    const confirm = screen.getByRole("dialog", { name: "Replace Clinical Sites?", hidden: true });
    for (const text of [
      /Delete all current Placement Opportunities \(3 currently\)/,
      "Replace all Placement Opportunities with imported data",
      "Preserve the Clinical Tag registry",
      "Preserve Student records",
      "Preserve Student tag selections",
      /Remove any existing Clinical Placements \(2 currently\)/,
      "Mark schedules as cleared/not generated",
    ]) {
      expect(within(confirm).getByText(text)).toBeTruthy();
    }
    await user.click(within(confirm).getByRole("button", { name: "Cancel" }));
    expect(api.applyClinicalSitesImport).not.toHaveBeenCalled();
  });

  it("commits after confirmation, reports the result, and refreshes the table", async () => {
    vi.mocked(api.previewClinicalSitesImport).mockResolvedValue(previewOf([goodRow]));
    vi.mocked(api.applyClinicalSitesImport).mockResolvedValue({
      opportunities_removed: 3,
      placements_removed: 2,
      opportunities_created: 1,
      tags_created: 1,
    });
    const user = userEvent.setup();
    await openMapping(user);
    await mapRequired(user);
    await user.click(screen.getByRole("button", { name: "Preview Import" }));
    const region = await screen.findByRole("region", { name: "Import preview" });
    await user.click(within(region).getByRole("button", { name: "Import Clinical Sites" }));
    const confirm = screen.getByRole("dialog", { name: "Replace Clinical Sites?", hidden: true });
    await user.click(within(confirm).getByRole("button", { name: "Replace Clinical Sites" }));
    await waitFor(() => expect(api.applyClinicalSitesImport).toHaveBeenCalledTimes(1));
    expect(await screen.findByText(/Imported 1 Placement Opportunity; replaced 3, removed 2 placement\(s\)/)).toBeTruthy();
    expect(vi.mocked(api.listClinicalOpportunities).mock.calls.length).toBeGreaterThan(1);
  });
});
