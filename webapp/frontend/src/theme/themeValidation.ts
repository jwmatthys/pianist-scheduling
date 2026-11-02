import type { ThemePalette } from "./ThemePalette";

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

export type ThemeValidationErrorCode =
  | "PALETTE_STRUCTURE_INVALID"
  | "PALETTE_COLOR_INVALID"
  | "TEXT_CONTRAST_INSUFFICIENT"
  | "SUCCESS_CONTRAST_INSUFFICIENT"
  | "WARNING_CONTRAST_INSUFFICIENT"
  | "ERROR_CONTRAST_INSUFFICIENT"
  | "FOCUS_CONTRAST_INSUFFICIENT"
  | "BORDER_CONTRAST_INSUFFICIENT";

export interface ThemeValidationIssue {
  code: ThemeValidationErrorCode;
  message: string;
  contrast?: number;
  requiredContrast: number;
}

export class ThemeAccessibilityError extends Error {
  readonly code = "THEME_ACCESSIBILITY_INVALID";

  constructor(readonly issues: ThemeValidationIssue[]) {
    super(issues.map((issue) => issue.message).join(" "));
    this.name = "ThemeAccessibilityError";
  }
}

export class ThemePaletteValidationError extends TypeError {
  constructor(message: string, readonly code: "PALETTE_STRUCTURE_INVALID" | "PALETTE_COLOR_INVALID") {
    super(message);
    this.name = "ThemePaletteValidationError";
  }
}

export function validateThemePalette(value: unknown): ThemePalette {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new ThemePaletteValidationError("Theme palette must be an object.", "PALETTE_STRUCTURE_INVALID");
  }

  const record = value as Record<string, unknown>;
  const keys = Reflect.ownKeys(record);
  if (
    keys.length !== PALETTE_KEYS.length ||
    keys.some((key) => typeof key !== "string" || !PALETTE_KEYS.some((paletteKey) => paletteKey === key))
  ) {
    throw new ThemePaletteValidationError(
      "Theme palette must contain exactly the ten approved base colors.",
      "PALETTE_STRUCTURE_INVALID",
    );
  }

  const palette = {} as ThemePalette;
  for (const key of PALETTE_KEYS) {
    const color = record[key];
    if (typeof color !== "string" || !/^#[\da-fA-F]{6}$/.test(color)) {
      throw new ThemePaletteValidationError(
        `Theme palette color "${key}" must be a six-digit hex value.`,
        "PALETTE_COLOR_INVALID",
      );
    }
    palette[key] = color.toUpperCase();
  }
  return palette;
}

export function contrastRatio(foreground: string, background: string): number {
  const foregroundLuminance = relativeLuminance(foreground);
  const backgroundLuminance = relativeLuminance(background);
  const [lighter, darker] = [foregroundLuminance, backgroundLuminance].sort((a, b) => b - a);
  return (lighter + 0.05) / (darker + 0.05);
}

export function validateTextContrast(palette: ThemePalette): ThemeValidationIssue[] {
  const issues: ThemeValidationIssue[] = [];
  const mutedText = mixHex(palette.textMuted, palette.text, 0.97);

  checkContrast(issues, "TEXT_CONTRAST_INSUFFICIENT", "Body text", palette.text, palette.background, 4.5);
  checkContrast(issues, "TEXT_CONTRAST_INSUFFICIENT", "Body text", palette.text, palette.surface, 4.5);
  checkContrast(issues, "TEXT_CONTRAST_INSUFFICIENT", "Primary text", palette.primary, palette.background, 4.5);
  checkContrast(issues, "TEXT_CONTRAST_INSUFFICIENT", "Primary text", palette.primary, palette.surface, 4.5);
  checkContrast(issues, "TEXT_CONTRAST_INSUFFICIENT", "Inverse text", palette.surface, palette.primary, 4.5);
  checkContrast(issues, "TEXT_CONTRAST_INSUFFICIENT", "Muted text", mutedText, palette.background, 4.5);
  checkContrast(issues, "TEXT_CONTRAST_INSUFFICIENT", "Muted text", mutedText, palette.surface, 4.5);
  return issues;
}

export function validateSuccessContrast(palette: ThemePalette): ThemeValidationIssue[] {
  const successSoft = mixHex(palette.success, palette.surface, 0.08);
  const issues: ThemeValidationIssue[] = [];
  checkContrast(issues, "SUCCESS_CONTRAST_INSUFFICIENT", "Success text", palette.success, successSoft, 4.5);
  checkContrast(issues, "SUCCESS_CONTRAST_INSUFFICIENT", "Success text", palette.success, palette.surface, 4.5);
  return issues;
}

export function validateWarningContrast(palette: ThemePalette): ThemeValidationIssue[] {
  const warningForeground = mixHex(palette.warning, palette.text, 0.7);
  const warningSoft = mixHex(palette.warning, palette.surface, 0.12);
  const issues: ThemeValidationIssue[] = [];
  checkContrast(issues, "WARNING_CONTRAST_INSUFFICIENT", "Warning text", warningForeground, warningSoft, 4.5);
  checkContrast(issues, "WARNING_CONTRAST_INSUFFICIENT", "Warning text", warningForeground, palette.surface, 4.5);
  checkContrast(issues, "WARNING_CONTRAST_INSUFFICIENT", "Warning text", warningForeground, palette.background, 4.5);
  checkContrast(issues, "WARNING_CONTRAST_INSUFFICIENT", "Warning indicator", palette.warning, palette.background, 3);
  checkContrast(issues, "WARNING_CONTRAST_INSUFFICIENT", "Warning indicator", palette.warning, palette.surface, 3);
  return issues;
}

export function validateErrorContrast(palette: ThemePalette): ThemeValidationIssue[] {
  const errorSoft = mixHex(palette.error, palette.surface, 0.12);
  const issues: ThemeValidationIssue[] = [];
  checkContrast(issues, "ERROR_CONTRAST_INSUFFICIENT", "Error text", palette.error, errorSoft, 4.5);
  checkContrast(issues, "ERROR_CONTRAST_INSUFFICIENT", "Error text", palette.error, palette.surface, 4.5);
  checkContrast(issues, "ERROR_CONTRAST_INSUFFICIENT", "Error text", palette.error, palette.background, 4.5);
  return issues;
}

export function validateFocusVisibility(palette: ThemePalette): ThemeValidationIssue[] {
  const issues: ThemeValidationIssue[] = [];
  checkContrast(issues, "FOCUS_CONTRAST_INSUFFICIENT", "Focus indicator", palette.primary, palette.background, 3);
  checkContrast(issues, "FOCUS_CONTRAST_INSUFFICIENT", "Focus indicator", palette.primary, palette.surface, 3);
  return issues;
}

export function validateThemeAccessibility(palette: ThemePalette): ThemeValidationIssue[] {
  const issues = [
    ...validateTextContrast(palette),
    ...validateSuccessContrast(palette),
    ...validateWarningContrast(palette),
    ...validateErrorContrast(palette),
    ...validateFocusVisibility(palette),
  ];
  const strongBorder = mixHex(palette.border, palette.textMuted, 0.3);
  checkContrast(issues, "BORDER_CONTRAST_INSUFFICIENT", "Border", palette.border, palette.background, 3);
  checkContrast(issues, "BORDER_CONTRAST_INSUFFICIENT", "Border", palette.border, palette.surface, 3);
  checkContrast(issues, "BORDER_CONTRAST_INSUFFICIENT", "Strong border", strongBorder, palette.background, 3);
  checkContrast(issues, "BORDER_CONTRAST_INSUFFICIENT", "Strong border", strongBorder, palette.surface, 3);
  return issues;
}

function checkContrast(
  issues: ThemeValidationIssue[],
  code: ThemeValidationErrorCode,
  label: string,
  foreground: string,
  background: string,
  requiredContrast: number,
): void {
  const contrast = contrastRatio(foreground, background);
  if (contrast < requiredContrast) {
    issues.push({
      code,
      message: `${label} contrast is ${contrast.toFixed(2)}:1; at least ${requiredContrast}:1 is required.`,
      contrast,
      requiredContrast,
    });
  }
}

function mixHex(first: string, second: string, firstRatio: number): string {
  const channels = [0, 2, 4].map((offset) => {
    const firstChannel = Number.parseInt(first.slice(offset + 1, offset + 3), 16);
    const secondChannel = Number.parseInt(second.slice(offset + 1, offset + 3), 16);
    return Math.round(firstChannel * firstRatio + secondChannel * (1 - firstRatio));
  });
  return `#${channels.map((channel) => channel.toString(16).padStart(2, "0")).join("")}`;
}

function relativeLuminance(color: string): number {
  const channels = [0, 2, 4].map((offset) => Number.parseInt(color.slice(offset + 1, offset + 3), 16) / 255);
  const linear = channels.map((channel) => (
    channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4
  ));
  return linear[0] * 0.2126 + linear[1] * 0.7152 + linear[2] * 0.0722;
}
