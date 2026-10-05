import { describe, expect, it, vi } from "vitest";
import type { jsPDF } from "jspdf";
import { createMarkdownPdf, renderMarkdownToPdf } from "./markdownPdf";

type RecordedLine = { text: string; x: number; y: number };

function fakePdf(pageWidth = 612, pageHeight = 792) {
  let currentPage = 0;
  const pages: RecordedLine[][] = [[]];
  const pdf = {
    internal: {
      pageSize: {
        getWidth: () => pageWidth,
        getHeight: () => pageHeight,
      },
    },
    addPage: vi.fn(() => {
      currentPage += 1;
      pages[currentPage] = [];
    }),
    setFont: vi.fn(),
    getTextWidth: (text: string) => text.length * 5,
    setFontSize: vi.fn(),
    setTextColor: vi.fn(),
    splitTextToSize: (text: string, width: number) => {
      const maxLength = Math.max(1, Math.floor(width / 5));
      const words = text.split(/\s+/);
      const wrapped: string[] = [];
      let current = "";
      for (const word of words) {
        const candidate = current ? `${current} ${word}` : word;
        if (candidate.length > maxLength && current) {
          wrapped.push(current);
          current = word;
        } else {
          current = candidate;
        }
      }
      if (current) wrapped.push(current);
      return wrapped;
    },
    text: (text: string, x: number, y: number) => pages[currentPage].push({ text, x, y }),
  };
  return { pdf, pages };
}

describe("Markdown PDF renderer", () => {
  it("wraps the full Markdown report within Letter page bounds without blank pages", () => {
    const { pdf, pages } = fakePdf();
    const report = [
      "# Lesson Schedules",
      "",
      "Source: `Synthetic source`",
      "",
      "## Pianist Schedules",
      ...Array.from({ length: 110 }, (_, index) => `- **Monday 9:00 AM-9:50 AM** - Synthetic student ${index + 1} - Studio ${index + 1}`),
      "",
      "## Unassigned Lessons",
      "- Final marker from the report.",
    ].join("\n");

    renderMarkdownToPdf(report, pdf as unknown as jsPDF);

    expect(pages.length).toBeGreaterThan(1);
    expect(pages.every((page) => page.length > 0)).toBe(true);
    expect(pages.flat().every((line) => line.y >= 42 && line.y < 750)).toBe(true);
    expect(pages.flat().map((line) => line.text).join(" ")).toContain("Lesson Schedules");
    expect(pages.flat().map((line) => line.text).join(" ")).toContain("Final marker from the report.");
    expect(pdf.addPage).toHaveBeenCalledTimes(pages.length - 1);
  });

  it("normalizes Markdown formatting without truncating inline report values", () => {
    const { pdf, pages } = fakePdf();
    renderMarkdownToPdf("# Title\n\n- **Student Name** - `Student ID` — Room", pdf as unknown as jsPDF);
    const output = pages.flat().map((line) => line.text).join(" ");

    expect(output).toContain("Student Name");
    expect(output).toContain("Student ID");
    expect(output).toContain("- Room");
    expect(output).not.toContain("**");
    expect(output).not.toContain("`");
  });

  it("starts the three Accompanist report sections on separate pages without empty pages", () => {
    const { pdf, pages } = fakePdf();
    renderMarkdownToPdf([
      "# Lesson Schedules",
      "## Pianist Schedules",
      "Pianist section body.",
      "## Students by Instructor",
      "Instructor section body.",
      "## Students by Name",
      "Student section body.",
    ].join("\n\n"), pdf as unknown as jsPDF);

    const sectionPage = (heading: string) => pages.findIndex((page) => page.some((line) => line.text === heading));
    const pianistPage = sectionPage("Pianist Schedules");
    const instructorPage = sectionPage("Students by Instructor");
    const studentPage = sectionPage("Students by Name");

    expect(pages).toHaveLength(3);
    expect(pages.every((page) => page.length > 0)).toBe(true);
    expect(new Set([pianistPage, instructorPage, studentPage]).size).toBe(3);
    expect(pianistPage).toBe(0);
    expect(instructorPage).toBe(1);
    expect(studentPage).toBe(2);
    expect(pdf.addPage).toHaveBeenCalledTimes(2);
  });

  it("creates a Letter-size multi-page PDF from a long report", async () => {
    const report = [
      "# Lesson Schedules",
      ...Array.from({ length: 120 }, (_, index) => `- **Monday 9:00 AM-9:50 AM** - Synthetic student ${index + 1} - Studio ${index + 1}`),
      "- Final report entry.",
    ].join("\n");

    const blob = await createMarkdownPdf(report);
    const pdfText = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = () => reject(reader.error);
      reader.readAsText(blob);
    });
    const pageCount = (pdfText.match(/\/Type \/Page\b/g) ?? []).length;

    expect(blob.type).toBe("application/pdf");
    expect(pdfText).toContain("/MediaBox [0 0 612. 792.");
    expect(pageCount).toBeGreaterThan(1);
  });
});