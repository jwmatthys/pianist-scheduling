import type { ReactNode, RefObject } from "react";
import { Modal } from "../dialogs/Modal";

// Presentation-only shells for import dialogs. They own the dialog frame, heading, close control and
// the shared error/success banners; mappings, previews, footers and API calls stay in each dialog.

const FILE_TYPES_HINT = "CSV, XLSX, or XLS";

export function ImportErrorBanner({ children, preserveLineBreaks = false }: { children: ReactNode; preserveLineBreaks?: boolean }) {
  return (
    <div className="error-banner" role="alert" style={preserveLineBreaks ? { whiteSpace: "pre-line" } : undefined}>
      {children}
    </div>
  );
}

export function ImportSuccessBanner({ children }: { children: ReactNode }) {
  return <div className="availability-success" role="status">{children}</div>;
}

/** Native `<dialog>` shell. The owner opens it with `dialogRef.current?.showModal()`. */
export function ImportDialogShell({
  title,
  dialogRef,
  error,
  success,
  preserveErrorLineBreaks = false,
  children,
}: {
  title: string;
  dialogRef: RefObject<HTMLDialogElement | null>;
  error?: string | null;
  success?: string | null;
  preserveErrorLineBreaks?: boolean;
  children: ReactNode;
}) {
  return (
    <dialog className="availability-import-dialog" ref={dialogRef}>
      <div className="availability-import-content">
        <header className="availability-import-heading">
          <div>
            <h2>{title}</h2>
            <p className="muted">{FILE_TYPES_HINT}</p>
          </div>
          <button type="button" className="dialog-close" aria-label="Close" onClick={() => dialogRef.current?.close()}>
            ×
          </button>
        </header>
        {error && <ImportErrorBanner preserveLineBreaks={preserveErrorLineBreaks}>{error}</ImportErrorBanner>}
        {success && <ImportSuccessBanner>{success}</ImportSuccessBanner>}
        {children}
      </div>
    </dialog>
  );
}

/** Shell built on the shared Modal used by the Clinical, Lesson and Tutor imports. */
export function ImportModalShell({ title, onClose, children }: { title: string; onClose: () => void; children: ReactNode }) {
  return (
    <Modal title={title} onClose={onClose} className="clinical-wide-dialog clinical-import-dialog">
      {children}
    </Modal>
  );
}

/** Body of an ImportModalShell: file-type hint and error banner above the dialog's own content. */
export function ImportModalBody({ error, children }: { error?: string | null; children: ReactNode }) {
  return (
    <div className="clinical-tag-manager">
      <p className="muted">{FILE_TYPES_HINT}</p>
      {error && <ImportErrorBanner preserveLineBreaks>{error}</ImportErrorBanner>}
      {children}
    </div>
  );
}
