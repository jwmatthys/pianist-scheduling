import { useState } from "react";
import { reportExportService, type ExportFormat, type ReportPayload } from "../lib/reportExport";

interface ExportDialogProps {
  onSelect: (format: ExportFormat) => void;
  onCancel: () => void;
}

export function ExportDialog({ onSelect, onCancel }: ExportDialogProps) {
  return (
    <div className="export-dialog-backdrop" onKeyDown={(event) => { if (event.key === "Escape") onCancel(); }}>
      <div className="export-dialog" role="dialog" aria-modal="true" aria-labelledby="export-dialog-title">
        <h2 id="export-dialog-title">Export Report</h2>
        <p className="dialog-note">Choose format:</p>
        <div className="export-dialog-formats">
          {reportExportService.formats().map((format) => (
            <button key={format.id} className="secondary-btn" type="button" onClick={() => onSelect(format.id)}>
              {format.label}
            </button>
          ))}
        </div>
        <div className="dialog-actions">
          <button className="secondary-btn" type="button" autoFocus onClick={onCancel}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

export interface ExportStatus {
  error?: string;
  notice?: string;
}

interface ReportExportButtonProps {
  reportId: string;
  title: string;
  payload: ReportPayload | null;
  disabled?: boolean;
  onStatus: (status: ExportStatus) => void;
}

/** The single "Export..." control shared by every module's report page. */
export function ReportExportButton({ reportId, title, payload, disabled = false, onStatus }: ReportExportButtonProps) {
  const [open, setOpen] = useState(false);
  const [exporting, setExporting] = useState(false);

  async function exportAs(format: ExportFormat) {
    if (!payload) return;
    setOpen(false);
    setExporting(true);
    onStatus({});
    try {
      const result = await reportExportService.export({ reportId, title, format, payload });
      onStatus({ notice: result.message });
    } catch (cause) {
      onStatus({ error: cause instanceof Error ? cause.message : "Could not export the report." });
    } finally {
      setExporting(false);
    }
  }

  return (
    <>
      <button
        className="secondary-btn"
        type="button"
        onClick={() => setOpen(true)}
        disabled={disabled || !payload || exporting}
      >
        {exporting ? "Exporting…" : "Export..."}
      </button>
      {open && <ExportDialog onSelect={(format) => void exportAs(format)} onCancel={() => setOpen(false)} />}
    </>
  );
}
