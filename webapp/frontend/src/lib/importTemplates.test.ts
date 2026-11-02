import { unzipSync, strFromU8 } from "fflate";
import { describe, expect, it } from "vitest";
import { generateImportTemplateCsv, getImportTemplateTable, IMPORT_TEMPLATES, type ImportTemplateKey } from "./importTemplates";
import { createXlsxWorkbook } from "./xlsxWriter";

const KEYS = Object.keys(IMPORT_TEMPLATES) as ImportTemplateKey[];

describe("import templates", () => {
  it("covers every import", () => {
    expect(KEYS).toHaveLength(9);
  });

  it.each(KEYS)("%s has unique headers and a complete sample row", (key) => {
    const { columns, rows } = getImportTemplateTable(key);
    expect(new Set(columns).size).toBe(columns.length);
    expect(rows.length).toBeGreaterThanOrEqual(1);
    for (const row of rows) expect(row).toHaveLength(columns.length);
  });

  it.each(KEYS)("%s CSV has a header followed by sample data", (key) => {
    const lines = generateImportTemplateCsv(key).split("\r\n");
    expect(lines[0].split(",")).toEqual(getImportTemplateTable(key).columns);
    expect(lines.length).toBeGreaterThanOrEqual(2);
  });

  it.each(KEYS)("%s XLSX contains the same headers", async (key) => {
    const { columns, rows } = getImportTemplateTable(key);
    const blob = createXlsxWorkbook(IMPORT_TEMPLATES[key].sheetName, columns, rows);
    const buffer = await new Promise<ArrayBuffer>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.readAsArrayBuffer(blob);
    });
    const files = unzipSync(new Uint8Array(buffer));
    const sheet = strFromU8(files["xl/worksheets/sheet1.xml"]);
    for (const column of columns) expect(sheet).toContain(column);
  });

  it.each(["clinicalSites", "lessonInstructors", "tutorTutors"] as ImportTemplateKey[])(
    "%s includes a Tags column with a semicolon-separated sample",
    (key) => {
      const { columns, rows } = getImportTemplateTable(key);
      expect(String(rows[0][columns.indexOf("Tags")])).toContain(";");
    },
  );

  it("uses synthetic example data only", () => {
    for (const key of KEYS) expect(generateImportTemplateCsv(key)).not.toMatch(/@(?!example\.)/);
  });
});
