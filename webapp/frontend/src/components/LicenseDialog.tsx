import { useEffect, useId, useRef, useState } from "react";
import type { FormEvent } from "react";
import { api } from "../lib/api";
import { chooseLocalFile } from "../lib/platform";
import type { LicenseStatus } from "../lib/types";

function statusLabel(status: LicenseStatus["status"]): string {
  if (status === "licensed") return "Active";
  if (status === "invalid") return "Invalid";
  return "Not activated";
}

type Props = {
  licenseStatus: LicenseStatus | null;
  remaining: number | null;
  trialError: string | null;
  onLicenseStatus: (status: LicenseStatus) => void;
};

function LicenseActivationFields({ onLicenseStatus, disabled = false }: {
  onLicenseStatus: (status: LicenseStatus) => void;
  disabled?: boolean;
}) {
  const keyId = useId();
  const [licenseKey, setLicenseKey] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function importLicenseFile() {
    setError(null);
    setBusy(true);
    try {
      const file = await chooseLocalFile("Scheduling Maestro License", ["license", "txt"]);
      if (!file) return;
      onLicenseStatus(await api.activateLicenseFile(file));
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
      onLicenseStatus(await api.activateLicenseKey(licenseKey));
      setLicenseKey("");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not activate this license key.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <>
      <button type="button" className="secondary-btn" disabled={busy || disabled} onClick={() => void importLicenseFile()}>
        {busy ? "Working…" : "Import License File"}
      </button>
      <form onSubmit={(event) => void activateKey(event)}>
        <label htmlFor={keyId}>Enter License Key</label>
        <input
          id={keyId}
          type="text"
          autoComplete="off"
          value={licenseKey}
          onChange={(event) => setLicenseKey(event.currentTarget.value)}
          disabled={busy || disabled}
        />
        <button type="submit" className="primary-btn" disabled={busy || disabled || !licenseKey.trim()}>
          Activate License Key
        </button>
      </form>
      {error && <div className="error-banner" role="alert">{error}</div>}
      <p className="dialog-note">
        Import the generated .license file or license-key.txt, or paste the license key above.
        Validation and storage are handled locally.
      </p>
    </>
  );
}

export function TrialExpiredDialog({ onLicenseStatus }: { onLicenseStatus: (status: LicenseStatus) => void }) {
  const dialogRef = useRef<HTMLDialogElement>(null);

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog || dialog.open) return;
    if (typeof dialog.showModal === "function") dialog.showModal();
    else dialog.setAttribute("open", "");
  }, []);

  return (
    <dialog
      className="session-dialog license-dialog"
      ref={dialogRef}
      aria-labelledby="trial-expired-title"
      aria-modal="true"
      onCancel={(event) => event.preventDefault()}
    >
      <div className="license-dialog-content">
        <h2 id="trial-expired-title" className="trial-expired-title">
          Scheduling Maestro trial expired. Upload or enter your license key to activate.
        </h2>
        <LicenseActivationFields onLicenseStatus={onLicenseStatus} />
      </div>
    </dialog>
  );
}

export function LicenseDialog({ licenseStatus, remaining, trialError, onLicenseStatus }: Props) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [status, setStatus] = useState<LicenseStatus | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const statusRequest = useRef(0);

  function updateStatus(next: LicenseStatus) {
    statusRequest.current += 1;
    setBusy(false);
    setStatus(next);
    onLicenseStatus(next);
  }

  async function loadStatus() {
    const request = ++statusRequest.current;
    setError(null);
    setStatus(null);
    setBusy(true);
    try {
      const next = await api.getLicenseStatus();
      if (request === statusRequest.current) setStatus(next);
    } catch (cause) {
      if (request === statusRequest.current) {
        setError(cause instanceof Error ? cause.message : "Could not load license status.");
      }
    } finally {
      if (request === statusRequest.current) setBusy(false);
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
    statusRequest.current += 1;
    setBusy(false);
    const dialog = dialogRef.current;
    if (dialog && typeof dialog.close === "function") dialog.close();
    else dialog?.removeAttribute("open");
    setError(null);
  }

  return (
    <>
      <div className="license-control">
        <button type="button" className="secondary-btn license-button" onClick={openDialog}>
          License
        </button>
        {trialError && <span className="license-countdown" role="status" title={trialError}>Trial timer unavailable</span>}
        {!trialError && licenseStatus?.status !== "licensed" && remaining !== null && (
          <span className="license-countdown">
            Trial: {Math.floor(remaining / 60)}:{String(remaining % 60).padStart(2, "0")} remaining
          </span>
        )}
      </div>
      <dialog
        className="session-dialog license-dialog"
        ref={dialogRef}
        aria-labelledby="license-dialog-title"
        aria-modal="true"
        onCancel={closeDialog}
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
            {status?.signing_key_id && (
              <p><strong>Signing key:</strong> {status.signing_key_id}</p>
            )}
            {status?.error && <p className="dialog-note">{status.error}</p>}
          </section>

          {error && <div className="error-banner" role="alert">{error}</div>}
          <LicenseActivationFields onLicenseStatus={updateStatus} disabled={busy} />
        </div>
      </dialog>
    </>
  );
}
