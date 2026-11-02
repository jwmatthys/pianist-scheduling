import { DEFAULT_THEME_PALETTE, type ThemePalette } from "./ThemePalette";
import {
  ThemeAccessibilityError,
  validateThemeAccessibility,
  validateThemePalette,
} from "./themeValidation";
import type { ThemeValidationIssue } from "./themeValidation";

export { DEFAULT_THEME_PALETTE } from "./ThemePalette";
export type { ThemePalette } from "./ThemePalette";
export {
  contrastRatio,
  validateErrorContrast,
  validateFocusVisibility,
  validateSuccessContrast,
  validateTextContrast,
  validateThemeAccessibility,
  validateThemePalette,
  validateWarningContrast,
  ThemeAccessibilityError,
  ThemePaletteValidationError,
  type ThemeValidationErrorCode,
  type ThemeValidationIssue,
} from "./themeValidation";

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
  causeCode?: string;
  issues?: ThemeValidationIssue[];
}

export type ThemeLoadResult =
  | { palette: Readonly<ThemePalette>; source: "default" }
  | { palette: ThemePalette; source: "loaded" }
  | { palette: Readonly<ThemePalette>; source: "fallback"; warning: ThemeLoadWarning };

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
    const palette = validateThemePalette(await loadCandidate());
    const accessibilityIssues = validateThemeAccessibility(palette);
    if (accessibilityIssues.length > 0) {
      throw new ThemeAccessibilityError(accessibilityIssues);
    }
    return { palette, source: "loaded" };
  } catch (error) {
    const causeCode = getErrorCode(error);
    const warning: ThemeLoadWarning = {
      code: "THEME_LOAD_FALLBACK",
      message: "Theme loading failed; using the built-in default palette.",
      cause: error instanceof Error ? error.message : "Unknown theme loading error.",
      ...(causeCode ? { causeCode } : {}),
      ...(error instanceof ThemeAccessibilityError ? { issues: error.issues } : {}),
    };
    reportWarning(warning);
    return { palette: DEFAULT_THEME_PALETTE, source: "fallback", warning };
  }
}

function getErrorCode(error: unknown): string | undefined {
  if (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    typeof error.code === "string"
  ) {
    return error.code;
  }
  return undefined;
}
