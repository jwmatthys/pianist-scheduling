import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { DARK_THEME_PALETTE, DEFAULT_THEME_PALETTE } from "./ThemePalette";
import { contrastRatio } from "./themeValidation";

const css = readFileSync(resolve(process.cwd(), "src/theme/theme.css"), "utf8");

function cellPercent(token: string): number {
  const match = css.match(new RegExp(`--${token}: color-mix\\(in srgb, var\\(--theme-\\w+\\) (\\d+)%, var\\(--theme-surface\\)\\)`));
  if (!match) throw new Error(`Missing ${token}`);
  return Number(match[1]) / 100;
}

function mix(foreground: string, background: string, amount: number): string {
  const channel = (hex: string, index: number) => parseInt(hex.slice(1 + index * 2, 3 + index * 2), 16);
  const value = [0, 1, 2]
    .map((i) => Math.round(channel(foreground, i) * amount + channel(background, i) * (1 - amount)))
    .map((n) => n.toString(16).padStart(2, "0"))
    .join("");
  return `#${value}`;
}

describe.each([
  ["default", DEFAULT_THEME_PALETTE],
  ["dark", DARK_THEME_PALETTE],
])("availability cells in the %s theme", (_name, palette) => {
  it.each([
    ["status-available-cell", palette.success],
    ["status-tentative-cell", palette.warning],
  ])("%s keeps text readable and its status outline visible", (token, status) => {
    const fill = mix(status, palette.surface, cellPercent(token));
    expect(contrastRatio(palette.text, fill)).toBeGreaterThanOrEqual(4.5);
    // The 2px status-colored outline is the 3:1 non-text boundary against the surface.
    expect(contrastRatio(status, palette.surface)).toBeGreaterThanOrEqual(3);
    expect(contrastRatio(fill, palette.surface)).toBeGreaterThan(1.5);
  });
});
