import { useRef, useState } from "react";
import { chooseLocalFile } from "../lib/platform";
import { DARK_THEME_PALETTE, DEFAULT_THEME_PALETTE } from "../theme/ThemePalette";
import { parseThemePalettePng } from "../theme/pngPalette";
import { useTheme } from "../theme/ThemeContext";
import { validateThemeAccessibility, validateThemePalette } from "../theme/themeValidation";
import type { ThemePalette } from "../theme/ThemePalette";
import type { ThemeSelection } from "../theme/themeRuntime";

const PALETTE_KEYS = [
  "primary",
  "accent",
  "background",
  "surface",
  "text",
  "textMuted",
  "border",
  "success",
  "warning",
  "error",
] as const satisfies readonly (keyof ThemePalette)[];

export function ThemeDialog() {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const { selection, palette, activate } = useTheme();
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function selectTheme(next: ThemeSelection) {
    setError(null);
    setBusy(true);
    try {
      await activate(next);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not apply the selected theme.");
    } finally {
      setBusy(false);
    }
  }

  async function uploadCustomTheme() {
    setError(null);
    setBusy(true);
    try {
      const file = await chooseLocalFile("Theme palette PNG", ["png"]);
      if (!file) return;
      const parsed = parseThemePalettePng(await file.arrayBuffer());
      const customPalette = validateThemePalette(parsed);
      const issues = validateThemeAccessibility(customPalette);
      if (issues.length > 0) {
        setError(issues.map((issue) => issue.message).join(" "));
        return;
      }
      await selectTheme({ id: "custom", palette: customPalette });
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not load the selected theme PNG.");
    } finally {
      setBusy(false);
    }
  }

  const customPalette = selection.id === "custom" ? palette : undefined;

  return (
    <>
      <button
        type="button"
        className="theme-switcher-button"
        aria-label="Choose theme"
        title="Choose theme"
        onClick={() => {
          setError(null);
          dialogRef.current?.showModal();
        }}
      >
        <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
          <path d="M12 12V2a10 10 0 0 1 8.66 5L12 12Z" fill="var(--theme-primary)" />
          <path d="m12 12 8.66-5a10 10 0 0 1 0 10L12 12Z" fill="var(--theme-accent)" />
          <path d="m12 12 8.66 5a10 10 0 0 1-8.66 5V12Z" fill="var(--theme-success)" />
          <path d="M12 12v10a10 10 0 0 1-8.66-5L12 12Z" fill="var(--theme-warning)" />
          <path d="m12 12-8.66 5a10 10 0 0 1 0-10L12 12Z" fill="var(--theme-error)" />
          <path d="m12 12-8.66-5A10 10 0 0 1 12 2v10Z" fill="var(--theme-text-muted)" />
          <circle cx="12" cy="12" r="3.2" fill="var(--theme-surface)" stroke="var(--theme-border)" />
        </svg>
      </button>
      <dialog
        className="session-dialog theme-dialog"
        ref={dialogRef}
        aria-labelledby="theme-dialog-title"
        aria-modal="true"
        onCancel={() => setError(null)}
      >
        <div className="theme-dialog-content">
          <header className="dialog-heading">
            <h2 id="theme-dialog-title">Theme</h2>
            <button type="button" className="dialog-close" aria-label="Close" onClick={() => dialogRef.current?.close()}>
              ×
            </button>
          </header>

          <div className="theme-options" aria-label="Theme options">
            <ThemeOption
              name="Default Theme"
              palette={DEFAULT_THEME_PALETTE}
              active={selection.id === "default"}
              disabled={busy}
              onClick={() => void selectTheme({ id: "default" })}
            />
            <ThemeOption
              name="Dark Theme"
              palette={DARK_THEME_PALETTE}
              active={selection.id === "dark"}
              disabled={busy}
              onClick={() => void selectTheme({ id: "dark" })}
            />
            <ThemeOption
              name="Custom Theme"
              palette={customPalette}
              active={selection.id === "custom"}
              disabled={busy}
              onClick={() => {
                if (selection.id === "custom") return;
                void uploadCustomTheme();
              }}
            />
          </div>

          <button type="button" className="secondary-btn" disabled={busy} onClick={() => void uploadCustomTheme()}>
            {busy ? "Loading theme…" : "Upload Custom Theme"}
          </button>

          {error && <div className="error-banner" role="alert">{error}</div>}
          <p className="dialog-note">Theme files are validated and stored locally on this device.</p>
        </div>
      </dialog>
    </>
  );
}

function ThemeOption({
  name,
  palette,
  active,
  disabled,
  onClick,
}: {
  name: string;
  palette?: Readonly<ThemePalette>;
  active: boolean;
  disabled: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      className={`theme-option${active ? " active" : ""}`}
      aria-pressed={active}
      disabled={disabled}
      onClick={onClick}
    >
      <span className="theme-option-heading">
        <span className="theme-option-name">{name}</span>
        {active && <span className="theme-active-indicator">Active</span>}
      </span>
      {palette && (
        <span className="theme-preview" aria-hidden="true">
          {PALETTE_KEYS.map((key) => (
            <span key={key} style={{ backgroundColor: palette[key] }} />
          ))}
        </span>
      )}
    </button>
  );
}
