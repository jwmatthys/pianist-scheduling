/**
 * Module-owned tabular projection of a report's underlying data, used for CSV and
 * Excel export. Modules must include every field of the underlying records (one row
 * per record), not just what the Markdown/PDF presentation shows. Internal identifiers
 * (database ids, UUIDs) are not user-facing and are omitted.
 */
export type ReportCell = string | number | boolean | null;

export interface ReportTable {
  name: string;
  columns: string[];
  rows: ReportCell[][];
}

export function cellText(value: ReportCell): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "Yes" : "No";
  return String(value);
}

function csvCell(value: ReportCell) {
  const text = cellText(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

export function tableToCsv(table: ReportTable) {
  return [table.columns, ...table.rows].map((row) => row.map(csvCell).join(",")).join("\r\n") + "\r\n";
}
