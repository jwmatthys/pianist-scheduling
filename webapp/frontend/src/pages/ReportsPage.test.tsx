import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  markdownReport: vi.fn(),
  saveTextFile: vi.fn(),
  saveBinaryFile: vi.fn(),
  createMarkdownPdf: vi.fn(),
}));

vi.mock("../lib/api", () => ({ api: { markdownReport: mocks.markdownReport } }));
vi.mock("../lib/platform", () => ({
  saveTextFile: mocks.saveTextFile,
  saveBinaryFile: mocks.saveBinaryFile,
}));
vi.mock("../lib/markdownPdf", () => ({ createMarkdownPdf: mocks.createMarkdownPdf }));

import { ReportsPage } from "./ReportsPage";

const report = "# Lesson Schedules\n\n- **Monday 9:00 AM** - Synthetic Student";

async function generateReport() {
  fireEvent.click(screen.getByRole("button", { name: "Generate report" }));
  await screen.findByText("Lesson Schedules");
}

describe("Accompanist report exports", () => {
  beforeEach(() => {
    vi.resetAllMocks();
    mocks.markdownReport.mockResolvedValue(report);
    mocks.saveTextFile.mockResolvedValue(true);
    mocks.saveBinaryFile.mockResolvedValue(true);
    mocks.createMarkdownPdf.mockResolvedValue(new Blob(["synthetic PDF bytes"], { type: "application/pdf" }));
  });

  it("opens the Markdown save flow with the generated Markdown and suggested filename", async () => {
    render(<ReportsPage />);
    await generateReport();
    fireEvent.click(screen.getByRole("button", { name: "Download .md" }));

    await waitFor(() => expect(mocks.saveTextFile).toHaveBeenCalledWith(
      report,
      "lesson_pianists.md",
      "Markdown report",
      "md",
    ));
    expect((await screen.findByRole("status")).textContent).toContain("Markdown report saved.");
  });

  it("generates the PDF from Markdown and opens the PDF save flow", async () => {
    const pdfBlob = new Blob(["synthetic PDF bytes"], { type: "application/pdf" });
    mocks.createMarkdownPdf.mockResolvedValue(pdfBlob);
    render(<ReportsPage />);
    await generateReport();
    fireEvent.click(screen.getByRole("button", { name: "Download .pdf" }));

    await waitFor(() => expect(mocks.createMarkdownPdf).toHaveBeenCalledWith(report));
    expect(mocks.saveBinaryFile).toHaveBeenCalledWith(pdfBlob, "lesson_pianists.pdf", "PDF report", "pdf");
    expect((await screen.findByRole("status")).textContent).toContain("PDF report saved.");
  });

  it("surfaces save-dialog failures rather than silently doing nothing", async () => {
    mocks.saveTextFile.mockRejectedValue(new Error("Could not write the selected file."));
    render(<ReportsPage />);
    await generateReport();
    fireEvent.click(screen.getByRole("button", { name: "Download .md" }));

    expect((await screen.findByRole("alert")).textContent).toContain("Could not write the selected file.");
  });
});
