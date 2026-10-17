import { strToU8, zipSync } from "fflate";
import { cellText, type ReportCell } from "./reportTable";

export const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    // Characters XML 1.0 forbids would corrupt the workbook.
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
}

function columnName(index: number) {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}

function sheetName(value: string) {
  const cleaned = value.replace(/[\\/?*[\]:]/g, " ").trim().slice(0, 31);
  return cleaned || "Report";
}

/** Builds a single-sheet workbook locally with inline strings (no server, no extra spreadsheet library). */
export function createXlsxWorkbook(name: string, columns: string[], dataRows: ReportCell[][]): Blob {
  const rows: ReportCell[][] = [columns, ...dataRows];
  const sheetRows = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((cell, columnIndex) => {
          const ref = `${columnName(columnIndex)}${rowIndex + 1}`;
          if (cell === null || cell === "") return "";
          if (typeof cell === "number" && Number.isFinite(cell)) return `<c r="${ref}"><v>${cell}</v></c>`;
          return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(cellText(cell))}</t></is></c>`;
        })
        .join("");
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join("");
  const header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const files = {
    "[Content_Types].xml": strToU8(
      `${header}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      "</Types>",
    ),
    "_rels/.rels": strToU8(
      `${header}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      "</Relationships>",
    ),
    "xl/workbook.xml": strToU8(
      `${header}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets><sheet name="${escapeXml(sheetName(name))}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `${header}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      "</Relationships>",
    ),
    "xl/worksheets/sheet1.xml": strToU8(
      `${header}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetRows}</sheetData></worksheet>`,
    ),
  };
  return new Blob([zipSync(files)], { type: XLSX_MIME_TYPE });
}
