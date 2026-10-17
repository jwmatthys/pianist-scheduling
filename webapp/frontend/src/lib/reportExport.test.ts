import { unzipSync, strFromU8 } from "fflate";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  createMarkdownPdf: vi.fn(),
  saveBinaryFile: vi.fn(),
  saveTextFile: vi.fn(),
}));

vi.mock("./markdownPdf", () => ({ createMarkdownPdf: mocks.createMarkdownPdf }));
vi.mock("./platform", () => ({ saveBinaryFile: mocks.saveBinaryFile, saveTextFile: mocks.saveTextFile }));

import { ExportFormat, reportExportService } from "./reportExport";

const markdown = "# Jury Schedule\n\n## Voice Panel\n\n- **9:00 AM** - Synthetic Student - Voice\n- **9:30 AM** - Other, \"Quoted\" - Cello\n";
const table = {
  name: "Jury Schedule",
  columns: ["Student", "Instrument", "Hours", "Needs Pianist", "Notes"],
  rows: [
    ["Synthetic Student", "Voice", 1.5, true, ""],
    ['Other, "Quoted"', "Cello", 0, false, "Hidden field not in Markdown"],
  ],
};
const base = { reportId: "test.report", title: "Jury Schedule", payload: { markdown, table } };

describe("reportExportService", () => {
  beforeEach(() => {
    vi.resetAllMocks();
    mocks.createMarkdownPdf.mockResolvedValue(new Blob(["pdf"]));
    mocks.saveBinaryFile.mockResolvedValue(true);
    mocks.saveTextFile.mockResolvedValue(true);
  });

  it("offers PDF, Markdown, Excel Workbook, and CSV", () => {
    expect(reportExportService.formats().map((format) => format.label)).toEqual(["PDF", "Markdown", "Excel Workbook", "CSV"]);
  });

  it("owns filename extensions based on the report title", () => {
    expect(reportExportService.fileName("Accompanist Schedule", ExportFormat.PDF)).toBe("Accompanist Schedule.pdf");
    expect(reportExportService.fileName("Accompanist Schedule", ExportFormat.XLSX)).toBe("Accompanist Schedule.xlsx");
    expect(reportExportService.fileName("Jury Schedule", ExportFormat.CSV)).toBe("Jury Schedule.csv");
    expect(reportExportService.fileName("Jury Schedule", ExportFormat.MARKDOWN)).toBe("Jury Schedule.md");
    expect(reportExportService.fileName(' a/b:c?. ', ExportFormat.CSV)).toBe("a b c.csv");
    expect(reportExportService.fileName("  ", ExportFormat.CSV)).toBe("Report.csv");
  });

  it("exports PDF with module page-break options", async () => {
    const result = await reportExportService.export({
      ...base,
      format: ExportFormat.PDF,
      payload: { markdown, table, pageBreakBeforeHeadings: ["Voice Panel"] },
    });
    expect(mocks.createMarkdownPdf).toHaveBeenCalledWith(markdown, { pageBreakBeforeHeadings: ["Voice Panel"] });
    expect(mocks.saveBinaryFile).toHaveBeenCalledWith(expect.any(Blob), "Jury Schedule.pdf", "PDF report", "pdf");
    expect(result).toEqual({ saved: true, message: "PDF report saved." });
  });

  it("exports Markdown unchanged", async () => {
    await reportExportService.export({ ...base, format: ExportFormat.MARKDOWN });
    expect(mocks.saveTextFile).toHaveBeenCalledWith(markdown, "Jury Schedule.md", "Markdown report", "md");
  });

  it("exports CSV rows projected from the report, with quoting", async () => {
    await reportExportService.export({ ...base, format: ExportFormat.CSV });
    const [csv, name, , extension] = mocks.saveTextFile.mock.calls[0];
    expect(name).toBe("Jury Schedule.csv");
    expect(extension).toBe("csv");
    expect(csv).toContain("Student,Instrument,Hours,Needs Pianist,Notes\r\n");
    expect(csv).toContain("Synthetic Student,Voice,1.5,Yes,\r\n");
    expect(csv).toContain('"Other, ""Quoted""",Cello,0,No,Hidden field not in Markdown');
  });

  it("exports an XLSX workbook containing the report text", async () => {
    await reportExportService.export({ ...base, format: ExportFormat.XLSX });
    const [blob, name, , extension] = mocks.saveBinaryFile.mock.calls[0];
    expect(name).toBe("Jury Schedule.xlsx");
    expect(extension).toBe("xlsx");
    const bytes = new Uint8Array(await new Promise<ArrayBuffer>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.readAsArrayBuffer(blob as Blob);
    }));
    const files = unzipSync(bytes);
    expect(Object.keys(files)).toContain("xl/worksheets/sheet1.xml");
    const sheet = strFromU8(files["xl/worksheets/sheet1.xml"]);
    expect(sheet).toContain("Synthetic Student");
    expect(sheet).toContain("Hidden field not in Markdown");
    expect(sheet).toContain("<v>1.5</v>");
    expect(sheet).toContain("Needs Pianist");
  });

  it("reports a canceled save", async () => {
    mocks.saveTextFile.mockResolvedValue(false);
    const result = await reportExportService.export({ ...base, format: ExportFormat.MARKDOWN });
    expect(result).toEqual({ saved: false, message: "Markdown save canceled." });
  });
});
