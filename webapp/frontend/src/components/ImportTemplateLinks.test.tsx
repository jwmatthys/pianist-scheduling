// @vitest-environment jsdom
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";

const saveTextFile = vi.fn();
const saveBinaryFile = vi.fn();
vi.mock("../lib/platform", () => ({
  saveTextFile: (...args: unknown[]) => saveTextFile(...args),
  saveBinaryFile: (...args: unknown[]) => saveBinaryFile(...args),
}));

import { ImportTemplateLinks } from "./ImportTemplateLinks";

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe("ImportTemplateLinks", () => {
  it("saves a CSV template with the header and sample row", async () => {
    saveTextFile.mockResolvedValue(true);
    render(<ImportTemplateLinks template="clinicalStudents" />);
    fireEvent.click(screen.getByRole("button", { name: "Download CSV template" }));
    await waitFor(() => expect(saveTextFile).toHaveBeenCalled());
    const [contents, name, , extension] = saveTextFile.mock.calls[0];
    expect(name).toBe("clinical-student-roster-template.csv");
    expect(extension).toBe("csv");
    expect(contents.split("\r\n")).toHaveLength(2);
    await screen.findByText("CSV template saved.");
  });

  it("saves an XLSX template", async () => {
    saveBinaryFile.mockResolvedValue(true);
    render(<ImportTemplateLinks template="juryPanels" />);
    fireEvent.click(screen.getByRole("button", { name: "Download XLSX template" }));
    await waitFor(() => expect(saveBinaryFile).toHaveBeenCalled());
    expect(saveBinaryFile.mock.calls[0][1]).toBe("jury-panels-template.xlsx");
    expect(saveBinaryFile.mock.calls[0][3]).toBe("xlsx");
  });

  it("reports save failures", async () => {
    saveTextFile.mockRejectedValue(new Error("denied"));
    render(<ImportTemplateLinks template="tutorTutors" />);
    fireEvent.click(screen.getByRole("button", { name: "Download CSV template" }));
    await screen.findByRole("alert");
  });
});
