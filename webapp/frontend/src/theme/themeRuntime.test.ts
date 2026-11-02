import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { DARK_THEME_PALETTE, DEFAULT_THEME_PALETTE } from "./ThemePalette";
import { parseThemePalettePng } from "./pngPalette";
import { activateTheme, initializeTheme } from "./themeRuntime";
import { validateThemeAccessibility } from "./themeValidation";

const defaultPng = new Uint8Array(readFileSync(join(process.cwd(), "src", "theme", "default-theme.png")));
const darkPng = new Uint8Array(readFileSync(join(process.cwd(), "src", "theme", "dark-theme.png")));

describe("theme runtime", () => {
  it("loads both built-in PNG themes and validates their accessibility", () => {
    expect(parseThemePalettePng(defaultPng)).toEqual(DEFAULT_THEME_PALETTE);
    expect(parseThemePalettePng(darkPng)).toEqual(DARK_THEME_PALETTE);
    expect(validateThemeAccessibility(DEFAULT_THEME_PALETTE)).toEqual([]);
    expect(validateThemeAccessibility(DARK_THEME_PALETTE)).toEqual([]);
  });

  it("activates and persists the dark theme using only the ten base variables", async () => {
    const setProperty = vi.fn();
    const setAttribute = vi.fn();
    const writePreference = vi.fn();
    const loadPng = vi.fn(async (id: "default" | "dark") => id === "dark" ? darkPng : defaultPng);

    await expect(activateTheme(
      { id: "dark" },
      { loadPng, target: { setProperty }, root: { setAttribute }, writePreference },
    )).resolves.toEqual({ selection: { id: "dark" }, palette: DARK_THEME_PALETTE });

    expect(loadPng).toHaveBeenCalledWith("dark");
    expect(setProperty).toHaveBeenCalledTimes(10);
    expect(setProperty).toHaveBeenCalledWith("--theme-background", DARK_THEME_PALETTE.background);
    expect(setAttribute).toHaveBeenCalledWith("data-theme-scheme", "dark");
    expect(JSON.parse(writePreference.mock.calls[0][0])).toEqual({ version: 1, id: "dark" });
  });

  it("restores the saved built-in theme after a new runtime initialization", async () => {
    localStorage.clear();
    const loadPng = vi.fn(async (id: "default" | "dark") => id === "dark" ? darkPng : defaultPng);
    await activateTheme({ id: "dark" }, { loadPng });

    const setProperty = vi.fn();
    const setAttribute = vi.fn();
    const restored = await initializeTheme({ loadPng, target: { setProperty }, root: { setAttribute } });

    expect(restored.selection).toEqual({ id: "dark" });
    expect(restored.palette).toEqual(DARK_THEME_PALETTE);
    expect(setAttribute).toHaveBeenCalledWith("data-theme-scheme", "dark");
  });

  it("restores a custom palette from local persistence without reopening its PNG", async () => {
    localStorage.clear();
    await activateTheme({ id: "custom", palette: DEFAULT_THEME_PALETTE });

    const loadPng = vi.fn();
    const restored = await initializeTheme({ loadPng });

    expect(restored.selection).toEqual({ id: "custom", palette: DEFAULT_THEME_PALETTE });
    expect(restored.palette).toEqual(DEFAULT_THEME_PALETTE);
    expect(loadPng).not.toHaveBeenCalled();
  });

  it("rejects invalid PNG data without persisting or applying it", async () => {
    const setProperty = vi.fn();
    const writePreference = vi.fn();
    await expect(activateTheme(
      { id: "dark" },
      {
        loadPng: async () => new Uint8Array([1, 2, 3]),
        target: { setProperty },
        writePreference,
      },
    )).rejects.toThrow(/valid PNG/);
    expect(setProperty).not.toHaveBeenCalled();
    expect(writePreference).not.toHaveBeenCalled();
  });

  it("falls back to the default palette when a saved custom palette is invalid", async () => {
    const reportWarning = vi.fn();
    const restored = await initializeTheme({
      readPreference: () => JSON.stringify({
        version: 1,
        id: "custom",
        palette: { ...DEFAULT_THEME_PALETTE, text: "#FFFFFF" },
      }),
      loadPng: async () => defaultPng,
      reportWarning,
    });

    expect(restored.source).toBe("fallback");
    expect(restored.selection).toEqual({ id: "default" });
    expect(restored.palette).toEqual(DEFAULT_THEME_PALETTE);
    expect(reportWarning).toHaveBeenCalledWith(expect.objectContaining({
      code: "THEME_LOAD_FALLBACK",
      causeCode: "THEME_ACCESSIBILITY_INVALID",
    }));
  });
});
