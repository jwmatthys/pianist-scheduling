export interface ThemePalette {
  primary: string;
  accent: string;
  background: string;
  surface: string;
  text: string;
  textMuted: string;
  border: string;
  success: string;
  warning: string;
  error: string;
}

export const DEFAULT_THEME_PALETTE: Readonly<ThemePalette> = Object.freeze({
  primary: "#1F4E79",
  accent: "#2E75B6",
  background: "#F4F6F9",
  surface: "#FFFFFF",
  text: "#1F2933",
  textMuted: "#6B7280",
  border: "#D1D5DB",
  success: "#2E7D32",
  warning: "#B7791F",
  error: "#C53030",
});

export type ThemeCssVariable =
  | "--theme-primary"
  | "--theme-accent"
  | "--theme-background"
  | "--theme-surface"
  | "--theme-text"
  | "--theme-text-muted"
  | "--theme-border"
  | "--theme-success"
  | "--theme-warning"
  | "--theme-error";

export type ThemeCssVariables = Record<ThemeCssVariable, string>;

export interface ThemeLoadWarning {
  code: "THEME_LOAD_FALLBACK";
  message: string;
  cause: string;
}

export type ThemeLoadResult =
  | { palette: Readonly<ThemePalette>; source: "default" }
  | { palette: ThemePalette; source: "loaded" }
  | { palette: Readonly<ThemePalette>; source: "fallback"; warning: ThemeLoadWarning };

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

export function validateThemePalette(value: unknown): ThemePalette {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new TypeError("Theme palette must be an object.");
  }

  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (keys.length !== PALETTE_KEYS.length || PALETTE_KEYS.some((key) => !Object.hasOwn(record, key))) {
    throw new TypeError("Theme palette must contain exactly the ten approved base colors.");
  }

  const palette = {} as ThemePalette;
  for (const key of PALETTE_KEYS) {
    const color = record[key];
    if (typeof color !== "string" || !/^#[\da-fA-F]{6}$/.test(color)) {
      throw new TypeError(`Theme palette color "${key}" must be a six-digit hex value.`);
    }
    palette[key] = color.toUpperCase();
  }
  return palette;
}

export function createThemeCssVariables(palette: ThemePalette): ThemeCssVariables {
  return {
    "--theme-primary": palette.primary,
    "--theme-accent": palette.accent,
    "--theme-background": palette.background,
    "--theme-surface": palette.surface,
    "--theme-text": palette.text,
    "--theme-text-muted": palette.textMuted,
    "--theme-border": palette.border,
    "--theme-success": palette.success,
    "--theme-warning": palette.warning,
    "--theme-error": palette.error,
  };
}

export function applyThemeCssVariables(
  palette: ThemePalette,
  target: Pick<CSSStyleDeclaration, "setProperty">,
): void {
  for (const [name, value] of Object.entries(createThemeCssVariables(palette))) {
    target.setProperty(name, value);
  }
}

export async function loadThemePalette(
  loadCandidate?: () => Promise<unknown>,
  reportWarning: (warning: ThemeLoadWarning) => void = (warning) => console.warn(warning),
): Promise<ThemeLoadResult> {
  if (!loadCandidate) {
    return { palette: DEFAULT_THEME_PALETTE, source: "default" };
  }

  try {
    return { palette: validateThemePalette(await loadCandidate()), source: "loaded" };
  } catch (error) {
    const warning: ThemeLoadWarning = {
      code: "THEME_LOAD_FALLBACK",
      message: "Theme loading failed; using the built-in default palette.",
      cause: error instanceof Error ? error.message : "Unknown theme loading error.",
    };
    reportWarning(warning);
    return { palette: DEFAULT_THEME_PALETTE, source: "fallback", warning };
  }
}
