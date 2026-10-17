import { createMarkdownPdf } from "./markdownPdf";
import { saveBinaryFile, saveTextFile } from "./platform";
import { tableToCsv, type ReportTable } from "./reportTable";
import { createXlsxWorkbook } from "./xlsxWriter";

export const ExportFormat = {
  PDF: "pdf",
  MARKDOWN: "markdown",
  XLSX: "xlsx",
  CSV: "csv",
} as const;
export type ExportFormat = (typeof ExportFormat)[keyof typeof ExportFormat];

/** Module-owned report content handed to the shared export service. */
export interface ReportPayload {
  markdown: string;
  pageBreakBeforeHeadings?: string[];
  /** Full-field rows for CSV and Excel; owned and built by the module. */
  table: ReportTable;
}

export interface ReportExportRequest {
  reportId: string;
  title: string;
  format: ExportFormat;
  payload: ReportPayload;
}

export interface ReportExportResult {
  saved: boolean;
  message: string;
}

interface ExportFormatDefinition {
  id: ExportFormat;
  label: string;
  extension: string;
  save(request: ReportExportRequest, fileName: string, definition: ExportFormatDefinition): Promise<boolean>;
}

// Adding a format means adding one entry here; module pages and the dialog need no changes.
const FORMATS: ExportFormatDefinition[] = [
  {
    id: ExportFormat.PDF,
    label: "PDF",
    extension: "pdf",
    async save({ payload }, fileName, { label, extension }) {
      const options = payload.pageBreakBeforeHeadings ? { pageBreakBeforeHeadings: payload.pageBreakBeforeHeadings } : undefined;
      const pdf = options ? await createMarkdownPdf(payload.markdown, options) : await createMarkdownPdf(payload.markdown);
      return saveBinaryFile(pdf, fileName, `${label} report`, extension);
    },
  },
  {
    id: ExportFormat.MARKDOWN,
    label: "Markdown",
    extension: "md",
    save: ({ payload }, fileName, { label, extension }) => saveTextFile(payload.markdown, fileName, `${label} report`, extension),
  },
  {
    id: ExportFormat.XLSX,
    label: "Excel Workbook",
    extension: "xlsx",
    save: ({ payload }, fileName, { label, extension }) =>
      saveBinaryFile(createXlsxWorkbook(payload.table.name, payload.table.columns, payload.table.rows), fileName, label, extension),
  },
  {
    id: ExportFormat.CSV,
    label: "CSV",
    extension: "csv",
    save: ({ payload }, fileName, { label, extension }) =>
      saveTextFile(`\uFEFF${tableToCsv(payload.table)}`, fileName, `${label} report`, extension),
  },
];

function definitionFor(format: ExportFormat) {
  const definition = FORMATS.find((item) => item.id === format);
  if (!definition) throw new Error(`Unsupported export format: ${format}`);
  return definition;
}

function fileNameFor(title: string, format: ExportFormat) {
  const { extension } = definitionFor(format);
  const base = title.replace(/[\\/:*?"<>|\u0000-\u001F]/g, " ").replace(/\s+/g, " ").trim().replace(/[.\s]+$/, "") || "Report";
  return `${base}.${extension}`;
}

export const reportExportService = {
  formats: (): { id: ExportFormat; label: string; extension: string }[] =>
    FORMATS.map(({ id, label, extension }) => ({ id, label, extension })),
  fileName: fileNameFor,
  async export(request: ReportExportRequest): Promise<ReportExportResult> {
    const definition = definitionFor(request.format);
    const saved = await definition.save(request, fileNameFor(request.title, request.format), definition);
    return {
      saved,
      message: saved ? `${definition.label} report saved.` : `${definition.label} save canceled.`,
    };
  },
};
