import type { jsPDF } from "jspdf";

const PAGE_MARGIN = 42;
const BODY_FONT_SIZE = 10;

type MarkdownBlock = {
  text: string;
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

    const heading = /^(#{1,3})\s+(.+)$/.exec(line);
    const bullet = /^[-*+]\s+(.+)$/.exec(line);
    let text = line;
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
      pageBreakBefore = level === 2 && (SECTION_PAGE_BREAKS.has(text) || extraPageBreakHeadings.has(text));
      bold = true;
      color = [31, 78, 121];
      gapBefore = Math.max(gapBefore, level === 1 ? 8 : 10);
      gapAfter = level === 1 ? 8 : 5;
    } else if (bullet) {
      text = `- ${plainText(bullet[1])}`;
      indent = 12;
      gapAfter = 2;
    } else {
      text = plainText(line);
      gapAfter = 2;
    }

    blocks.push({ text, fontSize, bold, color, indent, gapBefore, gapAfter, pageBreakBefore });
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
    const wrapped = pdf.splitTextToSize(block.text, textWidth - block.indent);
    pdf.setFont("helvetica", block.bold ? "bold" : "normal");
    pdf.setFontSize(block.fontSize);
    pdf.setTextColor(...block.color);

    for (const line of wrapped) {
      ensurePage(lineHeight);
      pdf.text(line, PAGE_MARGIN + block.indent, y);
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
