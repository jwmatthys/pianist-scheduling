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
  border: "#858B96",
  success: "#2E7D32",
  warning: "#B7791F",
  error: "#C53030",
});

export const DARK_THEME_PALETTE: Readonly<ThemePalette> = Object.freeze({
  primary: "#80BFFF",
  accent: "#62A3E8",
  background: "#111820",
  surface: "#1B2633",
  text: "#F2F4F7",
  textMuted: "#B4BFCC",
  border: "#78889B",
  success: "#67D391",
  warning: "#FFD166",
  error: "#FF7B7B",
});
