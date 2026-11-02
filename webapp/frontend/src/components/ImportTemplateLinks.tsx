import { useState } from "react";
import { generateImportTemplateCsv, getImportTemplateTable, IMPORT_TEMPLATES, type ImportTemplateKey } from "../lib/importTemplates";
import { saveBinaryFile, saveTextFile } from "../lib/platform";
import { createXlsxWorkbook } from "../lib/xlsxWriter";

type Props = {
  template: ImportTemplateKey;
  disabled?: boolean;
};

/** Saves a local CSV or XLSX template with every import field and one sample row. */
export function ImportTemplateLinks({ template, disabled = false }: Props) {
  const [message, setMessage] = useState<{ kind: "success" | "error"; text: string } | null>(null);
  const definition = IMPORT_TEMPLATES[template];
  const label = definition.description;

  async function save(format: "csv" | "xlsx") {
    setMessage(null);
    try {
      const saved = format === "csv"
        ? await saveTextFile(generateImportTemplateCsv(template), `${definition.fileName}.csv`, `${label} CSV`, "csv")
        : await saveBinaryFile(
          (() => {
            const { columns, rows } = getImportTemplateTable(template);
            return createXlsxWorkbook(definition.sheetName, columns, rows);
          })(),
          `${definition.fileName}.xlsx`,
          `${label} workbook`,
          "xlsx",
        );
      if (saved) setMessage({ kind: "success", text: `${format.toUpperCase()} template saved.` });
    } catch {
      setMessage({ kind: "error", text: "Could not save the template. Check the selected destination and try again." });
    }
  }

  return (
    <span className="import-template-links">
      <button type="button" className="template-download-button" disabled={disabled} onClick={() => void save("csv")}>
        Download CSV template
      </button>
      <button type="button" className="template-download-button" disabled={disabled} onClick={() => void save("xlsx")}>
        Download XLSX template
      </button>
      {message && <span className={message.kind === "error" ? "import-template-error" : "muted"} role={message.kind === "error" ? "alert" : "status"}>{message.text}</span>}
    </span>
  );
}
