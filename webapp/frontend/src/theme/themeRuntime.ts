import defaultThemePngUrl from "./default-theme.png?url";
import darkThemePngUrl from "./dark-theme.png?url";
import { getLocalPreference, setLocalPreference } from "../lib/platform";
import { parseThemePalettePng } from "./pngPalette";
import { DEFAULT_THEME_PALETTE, type ThemePalette } from "./ThemePalette";
import {
  applyThemeCssVariables,
  type ThemeLoadWarning,
} from "./themeLoader";
import {
  contrastRatio,
  ThemeAccessibilityError,
  validateThemeAccessibility,
  validateThemePalette,
} from "./themeValidation";

const THEME_PREFERENCE_KEY = "scheduling-maestro.theme";

export type ThemeSelection =
  | { id: "default" | "dark" }
  | { id: "custom"; palette: ThemePalette };

export interface ActiveTheme {
  selection: ThemeSelection;
  palette: ThemePalette;
}

export interface ThemeInitializationResult extends ActiveTheme {
  source: "loaded" | "fallback";
  warning?: ThemeLoadWarning;
}

export interface InitializeThemeOptions {
  loadPng?: (themeId: "default" | "dark") => Promise<ArrayBuffer | Uint8Array>;
  target?: Pick<CSSStyleDeclaration, "setProperty">;
  root?: Pick<HTMLElement, "setAttribute">;
  readPreference?: () => string | null;
  writePreference?: (value: string) => void;
  reportWarning?: (warning: ThemeLoadWarning) => void;
}

export interface ActivateThemeOptions {
  loadPng?: (themeId: "default" | "dark") => Promise<ArrayBuffer | Uint8Array>;
  target?: Pick<CSSStyleDeclaration, "setProperty">;
  root?: Pick<HTMLElement, "setAttribute">;
  writePreference?: (value: string) => void;
}

export async function initializeTheme(options: InitializeThemeOptions = {}): Promise<ThemeInitializationResult> {
  const target = options.target ?? document.documentElement.style;
  const root = options.root ?? document.documentElement;
  let selection: ThemeSelection = { id: "default" };
  let preferenceWarning: ThemeLoadWarning | undefined;

  try {
    const stored = (options.readPreference ?? (() => getLocalPreference(THEME_PREFERENCE_KEY)))();
    if (stored) selection = parseStoredSelection(stored);
  } catch (error) {
    preferenceWarning = createWarning(error, "Saved theme preference could not be restored.");
  }

  try {
    const palette = await loadSelectionPalette(selection, options.loadPng);
    applyPalette(palette, target, root);
    if (preferenceWarning) {
      (options.reportWarning ?? console.warn)(preferenceWarning);
      return {
        palette,
        source: "fallback",
        selection: { id: "default" },
        warning: preferenceWarning,
      };
    }
    return { palette, source: "loaded", selection };
  } catch (error) {
    const warning = createWarning(error, "Theme restoration failed; using the built-in default palette.");
    (options.reportWarning ?? console.warn)(warning);
    applyPalette(DEFAULT_THEME_PALETTE, target, root);
    return {
      palette: DEFAULT_THEME_PALETTE,
      source: "fallback",
      selection: { id: "default" },
      warning,
    };
  }
}

export async function activateTheme(
  selection: ThemeSelection,
  options: ActivateThemeOptions = {},
): Promise<ActiveTheme> {
  const palette = await loadSelectionPalette(selection, options.loadPng);
  const normalizedSelection: ThemeSelection = selection.id === "custom"
    ? { id: "custom", palette }
    : selection;
  const serialized = JSON.stringify({ version: 1, ...normalizedSelection });
  (options.writePreference ?? ((value) => setLocalPreference(THEME_PREFERENCE_KEY, value)))(serialized);
  applyPalette(
    palette,
    options.target ?? document.documentElement.style,
    options.root ?? document.documentElement,
  );
  return { selection: normalizedSelection, palette };
}

export function parseStoredSelection(value: string): ThemeSelection {
  const parsed: unknown = JSON.parse(value);
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    throw new TypeError("Saved theme preference is not a valid object.");
  }
  const stored = parsed as Record<string, unknown>;
  if (stored.version !== 1 || (stored.id !== "default" && stored.id !== "dark" && stored.id !== "custom")) {
    throw new TypeError("Saved theme preference has an unsupported format.");
  }
  if (stored.id === "custom") {
    if (Object.keys(stored).some((key) => !["version", "id", "palette"].includes(key))) {
      throw new TypeError("Saved custom theme preference contains unexpected data.");
    }
    return { id: "custom", palette: validateAccessiblePalette(stored.palette) };
  }
  if (Object.keys(stored).some((key) => key !== "version" && key !== "id")) {
    throw new TypeError("Saved built-in theme preference contains unexpected data.");
  }
  return { id: stored.id };
}

async function loadSelectionPalette(
  selection: ThemeSelection,
  loadPng?: (themeId: "default" | "dark") => Promise<ArrayBuffer | Uint8Array>,
): Promise<ThemePalette> {
  if (selection.id === "custom") return validateAccessiblePalette(selection.palette);

  const png = loadPng
    ? await loadPng(selection.id)
    : await fetchThemePng(selection.id === "dark" ? darkThemePngUrl : defaultThemePngUrl);
  return validateAccessiblePalette(parseThemePalettePng(png));
}

function validateAccessiblePalette(value: unknown): ThemePalette {
  const palette = validateThemePalette(value);
  const issues = validateThemeAccessibility(palette);
  if (issues.length > 0) throw new ThemeAccessibilityError(issues);
  return palette;
}

function applyPalette(
  palette: ThemePalette,
  target: Pick<CSSStyleDeclaration, "setProperty">,
  root: Pick<HTMLElement, "setAttribute">,
): void {
  applyThemeCssVariables(palette, target);
  root.setAttribute("data-theme-scheme", isDarkPalette(palette) ? "dark" : "light");
}

function isDarkPalette(palette: ThemePalette): boolean {
  return contrastRatio("#FFFFFF", palette.background) >= 3;
}

function createWarning(error: unknown, fallbackMessage: string): ThemeLoadWarning {
  const causeCode = getErrorCode(error);
  return {
    code: "THEME_LOAD_FALLBACK",
    message: fallbackMessage,
    cause: error instanceof Error ? error.message : "Unknown theme loading error.",
    ...(causeCode ? { causeCode } : {}),
    ...(error instanceof ThemeAccessibilityError ? { issues: error.issues } : {}),
  };
}

async function fetchThemePng(url: string): Promise<ArrayBuffer> {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Unable to load the built-in theme PNG (HTTP ${response.status}).`);
  }
  return response.arrayBuffer();
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
