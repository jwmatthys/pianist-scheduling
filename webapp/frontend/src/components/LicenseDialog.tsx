import { useRef, useState } from "react";
import type { FormEvent } from "react";
import { api } from "../lib/api";
import { chooseLocalFile } from "../lib/platform";
import type { LicenseStatus } from "../lib/types";

function statusLabel(status: LicenseStatus["status"]): string {
  if (status === "licensed") return "Active";
  if (status === "invalid") return "Invalid";
  return "Not activated";
}

export function LicenseDialog() {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [status, setStatus] = useState<LicenseStatus | null>(null);
  const [licenseKey, setLicenseKey] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function loadStatus() {
    setError(null);
    setBusy(true);
    try {
      setStatus(await api.getLicenseStatus());
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not load license status.");
    } finally {
      setBusy(false);
    }
  }

  async function importLicenseFile() {
    setError(null);
    setBusy(true);
    try {
      const file = await chooseLocalFile("Scheduling Maestro License", ["license"]);
      if (!file) return;
      setStatus(await api.activateLicenseFile(file));
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not activate the selected license file.");
    } finally {
      setBusy(false);
    }
  }

  async function activateKey(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setError(null);
    setBusy(true);
    try {
      setStatus(await api.activateLicenseKey(licenseKey));
      setLicenseKey("");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not activate this license key.");
    } finally {
      setBusy(false);
    }
  }

  function openDialog() {
    setError(null);
    const dialog = dialogRef.current;
    if (dialog && !dialog.open) {
      if (typeof dialog.showModal === "function") dialog.showModal();
      else dialog.setAttribute("open", "");
    }
    void loadStatus();
  }

  function closeDialog() {
    const dialog = dialogRef.current;
    if (dialog && typeof dialog.close === "function") dialog.close();
    else dialog?.removeAttribute("open");
    setError(null);
  }

  return (
    <>
      <button type="button" className="secondary-btn" onClick={openDialog}>
        License
      </button>
      <dialog
        className="session-dialog license-dialog"
        ref={dialogRef}
        aria-labelledby="license-dialog-title"
        aria-modal="true"
        onCancel={() => setError(null)}
      >
        <div className="license-dialog-content">
          <header className="dialog-heading">
            <h2 id="license-dialog-title">License</h2>
            <button type="button" className="dialog-close" aria-label="Close" onClick={closeDialog}>
              ×
            </button>
          </header>

          <section className="license-status" aria-label="License status">
            <p><strong>Status:</strong> {status ? statusLabel(status.status) : "Loading…"}</p>
            {status?.payload && (
              <>
                <p><strong>Licensed to:</strong> {status.payload.licensed_to}</p>
                <p><strong>License type:</strong> {status.payload.license_type}</p>
                <p><strong>Issue date:</strong> {status.payload.issued}</p>
              </>
            )}
            {status?.activated_at && (
              <p><strong>Activation date:</strong> {new Date(status.activated_at).toLocaleDateString()}</p>
            )}
            {status?.error && <p className="dialog-note">{status.error}</p>}
          </section>

          <button type="button" className="secondary-btn" disabled={busy} onClick={() => void importLicenseFile()}>
            {busy ? "Working…" : "Import License File"}
          </button>

          <form onSubmit={(event) => void activateKey(event)}>
            <label htmlFor="license-key">Enter License Key</label>
            <input
              id="license-key"
              type="text"
              autoComplete="off"
              value={licenseKey}
              onChange={(event) => setLicenseKey(event.currentTarget.value)}
              disabled={busy}
            />
            <button type="submit" className="primary-btn" disabled={busy || !licenseKey.trim()}>
              Activate License Key
            </button>
          </form>

          {error && <div className="error-banner" role="alert">{error}</div>}
          <p className="dialog-note">License validation and storage are handled locally.</p>
        </div>
      </dialog>
    </>
  );
}
