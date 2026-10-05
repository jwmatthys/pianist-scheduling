import type { jsPDF } from "jspdf";

const PAGE_MARGIN = 42;
const BODY_FONT_SIZE = 10;

type MarkdownBlock = {
  text: string;
  boldPrefix: string;
  fontSize: number;
  bold: boolean;
  color: [number, number, number];
  indent: number;
  gapBefore: number;
  gapAfter: number;
  pageBreakBefore: boolean;
};

const SECTION_PAGE_BREAKS = new Set(["Students by Instructor", "Students by Name"]);
export type MarkdownPdfOptions = { pageBreakBeforeHeadings?: string[] };

function plainText(markdown: string) {
  return markdown
    .replace(/!\[([^\]]*)\]\([^)]+\)/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/__(.*?)__/g, "$1")
    .replace(/\*([^*]+)\*/g, "$1")
    .replace(/_([^_]+)_/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/[\u2012-\u2015]/g, "-")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201c\u201d]/g, '"');
}

function markdownBlocks(markdown: string, extraPageBreakHeadings: Set<string>): MarkdownBlock[] {
  const blocks: MarkdownBlock[] = [];
  let hasContent = false;
  let pendingGap = 0;

  for (const rawLine of markdown.replace(/\r\n/g, "\n").split("\n")) {
    const line = rawLine.trim();
    if (!line) {
      if (hasContent) pendingGap = Math.min(pendingGap + 4, 8);
      continue;
    }
    if (/^([-*_]\s*){3,}$/.test(line)) {
      pendingGap = Math.max(pendingGap, 8);
      continue;
    }

    const heading = /^(#{1,4})\s+(.+)$/.exec(line);
    const bullet = /^[-*+]\s+(.+)$/.exec(line);
    let text = line;
    let boldPrefix = "";
    let fontSize = BODY_FONT_SIZE;
    let bold = false;
    let color: [number, number, number] = [31, 41, 51];
    let indent = 0;
    let gapBefore = pendingGap;
    let gapAfter = 0;
    let pageBreakBefore = false;

    if (heading) {
      text = plainText(heading[2]);
      const level = heading[1].length;
      fontSize = level === 1 ? 18 : level === 2 ? 14 : 11;
      pageBreakBefore = level <= 2 && (SECTION_PAGE_BREAKS.has(text) || extraPageBreakHeadings.has(text));
      bold = true;
      color = [31, 78, 121];
      gapBefore = Math.max(gapBefore, level === 1 ? 8 : 10);
      gapAfter = level === 1 ? 8 : 5;
    } else {
      const content = bullet ? bullet[1] : line;
      const lead = /^\*\*(.+?)\*\*(.*)$/.exec(content);
      if (lead) {
        boldPrefix = plainText(lead[1]);
        text = plainText(lead[2]).trim();
      } else {
        text = plainText(content);
      }
      gapAfter = 2;
    }

    blocks.push({ text, boldPrefix, fontSize, bold, color, indent, gapBefore, gapAfter, pageBreakBefore });
    pendingGap = 0;
    hasContent = true;
  }

  return blocks;
}

export function renderMarkdownToPdf(markdown: string, pdf: jsPDF, options: MarkdownPdfOptions = {}) {
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const textWidth = pageWidth - PAGE_MARGIN * 2;
  const bottom = pageHeight - PAGE_MARGIN;
  let y = PAGE_MARGIN;
  let pageHasContent = false;

  function ensurePage(lineHeight: number) {
    if (y + lineHeight <= bottom) return;
    if (pageHasContent) pdf.addPage();
    y = PAGE_MARGIN;
    pageHasContent = false;
  }

  const extraPageBreakHeadings = new Set(options.pageBreakBeforeHeadings ?? []);
  for (const block of markdownBlocks(markdown, extraPageBreakHeadings)) {
    if (block.pageBreakBefore && pageHasContent) {
      pdf.addPage();
      y = PAGE_MARGIN;
      pageHasContent = false;
    }
    const lineHeight = block.fontSize * 1.4;
    y += block.gapBefore;
    pdf.setFontSize(block.fontSize);
    pdf.setTextColor(...block.color);
    const x = PAGE_MARGIN + block.indent;
    const width = textWidth - block.indent;

    if (block.boldPrefix) {
      pdf.setFont("helvetica", "bold");
      const prefixWidth = pdf.getTextWidth(`${block.boldPrefix} `);
      let remaining = block.text;
      const firstLine: string = remaining
        ? (pdf.splitTextToSize(remaining, Math.max(width - prefixWidth, 20)) as string[])[0] ?? ""
        : "";
      remaining = remaining.slice(firstLine.length).trim();
      ensurePage(lineHeight);
      pdf.setFont("helvetica", "bold");
      pdf.text(block.boldPrefix, x, y);
      if (firstLine) {
        pdf.setFont("helvetica", "normal");
        pdf.text(firstLine, x + prefixWidth, y);
      }
      y += lineHeight;
      pageHasContent = true;
      if (remaining) {
        pdf.setFont("helvetica", "normal");
        for (const line of pdf.splitTextToSize(remaining, width)) {
          ensurePage(lineHeight);
          pdf.text(line, x, y);
          y += lineHeight;
        }
      }
      y += block.gapAfter;
      continue;
    }

    const wrapped = pdf.splitTextToSize(block.text, width);
    pdf.setFont("helvetica", block.bold ? "bold" : "normal");

    for (const line of wrapped) {
      ensurePage(lineHeight);
      pdf.text(line, x, y);
      y += lineHeight;
      pageHasContent = true;
    }
    y += block.gapAfter;
  }
}

export async function createMarkdownPdf(markdown: string, options: MarkdownPdfOptions = {}): Promise<Blob> {
  const { jsPDF } = await import("jspdf");
  const pdf = new jsPDF({ orientation: "portrait", unit: "pt", format: "letter" });
  renderMarkdownToPdf(markdown, pdf, options);
  return pdf.output("blob");
}
