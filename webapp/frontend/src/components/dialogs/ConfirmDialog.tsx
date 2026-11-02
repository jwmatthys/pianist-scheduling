import type { ReactNode } from "react";
import { Modal } from "./Modal";

export function ConfirmDialog({
  title,
  message,
  children,
  confirmLabel,
  onConfirm,
  onCancel,
}: {
  title: string;
  message: string;
  children?: ReactNode;
  confirmLabel: string;
  onConfirm: () => void;
  onCancel: () => void;
}) {
  return (
    <Modal title={title} onClose={onCancel}>
      {message && <p className="dialog-note">{message}</p>}
      {children}
      <div className="dialog-actions">
        <button type="button" onClick={onCancel}>
          Cancel
        </button>
        <button type="button" className="danger-btn" onClick={onConfirm}>
          {confirmLabel}
        </button>
      </div>
    </Modal>
  );
}
