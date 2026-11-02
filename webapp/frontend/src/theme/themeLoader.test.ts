import { readFileSync } from "node:fs";
import { join } from "node:path";
import { zlibSync } from "fflate";
import { describe, expect, it, vi } from "vitest";
import { parseThemePalettePng, ThemePngError, THEME_PNG_CELL_SIZE, THEME_PNG_HEIGHT, THEME_PNG_WIDTH } from "./pngPalette";
import { DEFAULT_THEME_PALETTE, type ThemePalette } from "./ThemePalette";
import {
  applyThemeCssVariables,
  createThemeCssVariables,
  loadThemePalette,
  validateThemePalette,
} from "./themeLoader";
import {
  contrastRatio,
  validateErrorContrast,
  validateFocusVisibility,
  validateSuccessContrast,
  validateTextContrast,
  validateThemeAccessibility,
  validateWarningContrast,
} from "./themeValidation";
import { initializeTheme } from "./themeRuntime";

const themeCss = readFileSync(join(process.cwd(), "src", "theme", "theme.css"), "utf8");
const defaultThemePng = new Uint8Array(readFileSync(join(process.cwd(), "src", "theme", "default-theme.png")));

const palette: ThemePalette = {
  primary: "#112233",
  accent: "#223344",
  background: "#334455",
  surface: "#445566",
  text: "#556677",
  textMuted: "#667788",
  border: "#778899",
  success: "#8899AA",
  warning: "#99AABB",
  error: "#AABBCC",
};

describe("theme palette", () => {
  it("validates exactly ten six-digit color values", () => {
    expect(Object.keys(palette)).toHaveLength(10);
    expect(validateThemePalette(palette)).toEqual(palette);
    expect(() => validateThemePalette({ ...palette, other: "#FFFFFF" })).toThrow(/exactly the ten/);
    const { error: _error, ...withoutError } = palette;
    expect(() => validateThemePalette(withoutError)).toThrow(/exactly the ten/);
    expect(() => validateThemePalette({ ...palette, warning: "red" })).toThrow(/six-digit hex/);
  });

  it("matches the built-in palette values to the CSS defaults", () => {
    const keys = {
      primary: "primary",
      accent: "accent",
      background: "background",
      surface: "surface",
      text: "text",
      textMuted: "text-muted",
      border: "border",
      success: "success",
      warning: "warning",
      error: "error",
    } as const;

    for (const [paletteKey, cssKey] of Object.entries(keys)) {
      const declaration = new RegExp(`--theme-${cssKey}:\\s*(#[\\da-fA-F]{6});`);
      expect(themeCss.match(declaration)?.[1]?.toUpperCase()).toBe(DEFAULT_THEME_PALETTE[paletteKey as keyof ThemePalette]);
    }
  });

  it("keeps readable muted text at WCAG AA contrast on the app background", () => {
    expect(themeCss).toContain(
      "--theme-text-muted-readable: color-mix(in srgb, var(--theme-text-muted) 97%, var(--theme-text));",
    );
    const readableMuted = mixHex(DEFAULT_THEME_PALETTE.textMuted, DEFAULT_THEME_PALETTE.text, 0.03);
    expect(contrastRatio(readableMuted, DEFAULT_THEME_PALETTE.background)).toBeGreaterThanOrEqual(4.5);
  });

  it("meets WCAG AA text contrast for semantic status colors", () => {
    const { surface, text, primary, success, warning, error } = DEFAULT_THEME_PALETTE;
    const successSoft = mixHex(surface, success, 0.08);
    const warningSoft = mixHex(surface, warning, 0.12);
    const errorSoft = mixHex(surface, error, 0.12);
    const warningForeground = mixHex(text, warning, 0.7);

    expect(contrastRatio(text, surface)).toBeGreaterThanOrEqual(4.5);
    expect(contrastRatio(surface, primary)).toBeGreaterThanOrEqual(4.5);
    expect(contrastRatio(success, successSoft)).toBeGreaterThanOrEqual(4.5);
    expect(contrastRatio(warningForeground, warningSoft)).toBeGreaterThanOrEqual(4.5);
    expect(contrastRatio(error, errorSoft)).toBeGreaterThanOrEqual(4.5);
  });

  it("meets WCAG non-text contrast for borders against app surfaces", () => {
    expect(themeCss).toContain(
      "--theme-border-strong: color-mix(in srgb, var(--theme-border) 30%, var(--theme-text-muted));",
    );
    const { border, textMuted, background, surface } = DEFAULT_THEME_PALETTE;
    const strongBorder = mixHex(textMuted, border, 0.3);

    for (const color of [border, strongBorder]) {
      expect(contrastRatio(color, background)).toBeGreaterThanOrEqual(3);
      expect(contrastRatio(color, surface)).toBeGreaterThanOrEqual(3);
    }
  });

  it("passes the accessibility validator for the default palette", () => {
    expect(validateThemeAccessibility(DEFAULT_THEME_PALETTE)).toEqual([]);
    expect(contrastRatio(DEFAULT_THEME_PALETTE.primary, DEFAULT_THEME_PALETTE.surface)).toBeGreaterThanOrEqual(3);
  });

  it("reports insufficient text, warning, error, and focus contrast", () => {
    const unreadableText = { ...DEFAULT_THEME_PALETTE, text: "#FFFFFF" };
    const unreadableSuccess = { ...DEFAULT_THEME_PALETTE, success: "#FFFFFF" };
    const unreadableWarning = { ...DEFAULT_THEME_PALETTE, warning: "#FFFFFF" };
    const unreadableError = { ...DEFAULT_THEME_PALETTE, error: "#FFFFFF" };
    const invisibleFocus = { ...DEFAULT_THEME_PALETTE, primary: "#FFFFFF" };

    expect(validateTextContrast(unreadableText)).toContainEqual(
      expect.objectContaining({ code: "TEXT_CONTRAST_INSUFFICIENT" }),
    );
    expect(validateSuccessContrast(unreadableSuccess)).toContainEqual(
      expect.objectContaining({ code: "SUCCESS_CONTRAST_INSUFFICIENT" }),
    );
    expect(validateWarningContrast(unreadableWarning)).toContainEqual(
      expect.objectContaining({ code: "WARNING_CONTRAST_INSUFFICIENT" }),
    );
    expect(validateErrorContrast(unreadableError)).toContainEqual(
      expect.objectContaining({ code: "ERROR_CONTRAST_INSUFFICIENT" }),
    );
    expect(validateFocusVisibility(invisibleFocus)).toContainEqual(
      expect.objectContaining({ code: "FOCUS_CONTRAST_INSUFFICIENT" }),
    );
  });

  it("generates and applies only the ten base CSS variables", () => {
    const variables = createThemeCssVariables(palette);
    const setProperty = vi.fn();
    expect(Object.keys(variables)).toHaveLength(10);
    expect(variables["--theme-text-muted"]).toBe(palette.textMuted);

    applyThemeCssVariables(palette, { setProperty });

    expect(setProperty).toHaveBeenCalledTimes(10);
    expect(setProperty).toHaveBeenCalledWith("--theme-warning", palette.warning);
  });

  it("uses the built-in palette when no candidate loader is supplied", async () => {
    await expect(loadThemePalette()).resolves.toEqual({
      palette: DEFAULT_THEME_PALETTE,
      source: "default",
    });
  });

  it("returns a loaded palette when the candidate is valid", async () => {
    await expect(loadThemePalette(async () => DEFAULT_THEME_PALETTE)).resolves.toEqual({
      palette: DEFAULT_THEME_PALETTE,
      source: "loaded",
    });
  });

  it.each([
    ["load failure", async () => { throw new Error("Unreadable image."); }],
    ["invalid palette", async () => ({ ...palette, error: "not-a-color" })],
  ])("falls back to the built-in palette for %s and reports a warning", async (_label, loadCandidate) => {
    const reportWarning = vi.fn();
    const result = await loadThemePalette(loadCandidate, reportWarning);

    expect(result.source).toBe("fallback");
    expect(result.palette).toBe(DEFAULT_THEME_PALETTE);
    expect(reportWarning).toHaveBeenCalledWith(expect.objectContaining({
      code: "THEME_LOAD_FALLBACK",
      message: "Theme loading failed; using the built-in default palette.",
    }));
  });

  it("falls back when a structurally valid palette fails accessibility checks", async () => {
    const reportWarning = vi.fn();
    const result = await loadThemePalette(
      async () => ({ ...DEFAULT_THEME_PALETTE, text: "#FFFFFF" }),
      reportWarning,
    );
    expect(result.source).toBe("fallback");
    expect(reportWarning).toHaveBeenCalledWith(expect.objectContaining({
      causeCode: "THEME_ACCESSIBILITY_INVALID",
      issues: expect.arrayContaining([expect.objectContaining({ code: "TEXT_CONTRAST_INSUFFICIENT" })]),
    }));
    if (result.source === "fallback") {
      expect(result.warning.causeCode).toBe("THEME_ACCESSIBILITY_INVALID");
      expect(result.warning.cause).toContain("Body text contrast");
    }
  });

  it("loads the default PNG through the runtime pipeline and applies exactly ten variables", async () => {
    expect(parseThemePalettePng(defaultThemePng)).toEqual(DEFAULT_THEME_PALETTE);
    const setProperty = vi.fn();
    const fetchMock = vi.fn<typeof fetch>().mockResolvedValue(new Response(toArrayBuffer(defaultThemePng)));
    vi.stubGlobal("fetch", fetchMock);
    try {
      const result = await initializeTheme({ target: { setProperty } });

      expect(result.source).toBe("loaded");
      expect(result.palette).toEqual(DEFAULT_THEME_PALETTE);
      expect(fetchMock).toHaveBeenCalledOnce();
      expect(setProperty).toHaveBeenCalledTimes(10);
      expect(setProperty).toHaveBeenCalledWith("--theme-border", DEFAULT_THEME_PALETTE.border);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("continues with the built-in palette when the default PNG cannot be decoded", async () => {
    const setProperty = vi.fn();
    const reportWarning = vi.fn();
    const result = await initializeTheme({
      loadPng: async () => new Uint8Array([1, 2, 3]),
      target: { setProperty },
      reportWarning,
    });

    expect(result.source).toBe("fallback");
    expect(reportWarning).toHaveBeenCalledWith(expect.objectContaining({
      code: "THEME_LOAD_FALLBACK",
      causeCode: "PNG_INVALID",
    }));
    expect(setProperty).toHaveBeenCalledWith("--theme-primary", DEFAULT_THEME_PALETTE.primary);
  });
});

describe("theme PNG parser", () => {
  it("extracts the ten center-pixel colors in the documented cell order", () => {
    expect(parseThemePalettePng(createPalettePng())).toEqual(palette);
  });

  it("accepts non-interlaced RGB PNGs", () => {
    expect(parseThemePalettePng(createPalettePng({ colorType: 2 }))).toEqual(palette);
  });

  it("rejects images with incorrect dimensions", () => {
    try {
      parseThemePalettePng(createPalettePng({ width: THEME_PNG_WIDTH - 1 }));
      throw new Error("Expected invalid dimensions to be rejected.");
    } catch (error) {
      expect(error).toBeInstanceOf(ThemePngError);
      expect(error).toMatchObject({ code: "PNG_DIMENSIONS_INVALID" });
    }
  });

  it("rejects a missing transparent palette cell", () => {
    expect(() => parseThemePalettePng(createPalettePng({ transparentCell: 8 })))
      .toThrow(/palette cell "warning" must be fully opaque/);
  });

  it("rejects corrupt and truncated PNG data", () => {
    const valid = createPalettePng();
    const truncated = valid.slice(0, -5);
    expect(() => parseThemePalettePng(truncated)).toThrow(/truncated chunk|incomplete/);

    const corrupt = valid.slice();
    corrupt[corrupt.length - 1] ^= 0xff;
    expect(() => parseThemePalettePng(corrupt)).toThrow(/corrupt IEND chunk/);
  });

  it("returns a structured error when image data or palette cells are missing", () => {
    expect(() => parseThemePalettePng(new Uint8Array()))
      .toThrow(expect.objectContaining({ code: "PNG_INVALID" }));
    expect(() => parseThemePalettePng(createPalettePng({ includeImageData: false })))
      .toThrow(expect.objectContaining({ code: "PNG_IMAGE_DATA_MISSING" }));
    expect(() => parseThemePalettePng(createPalettePng({ transparentCell: 8 })))
      .toThrow(expect.objectContaining({ code: "PALETTE_INVALID" }));
  });

  it("returns a structured error for corrupt compressed image data", () => {
    expect(() => parseThemePalettePng(createPalettePng({ invalidImageData: true })))
      .toThrow(expect.objectContaining({ code: "PNG_IMAGE_DATA_INVALID" }));
  });
});

function createPalettePng(options: {
  width?: number;
  transparentCell?: number;
  colorType?: 2 | 6;
  includeImageData?: boolean;
  invalidImageData?: boolean;
} = {}): Uint8Array {
  const width = options.width ?? THEME_PNG_WIDTH;
  const height = THEME_PNG_HEIGHT;
  const colorType = options.colorType ?? 6;
  const channels = colorType === 6 ? 4 : 3;
  const colors = Object.values(palette).map((value) => (
    [0, 2, 4].map((index) => Number.parseInt(value.slice(index + 1, index + 3), 16))
  ));
  const scanlines = new Uint8Array((width * channels + 1) * height);

  for (let y = 0; y < height; y += 1) {
    const rowStart = y * (width * channels + 1);
    scanlines[rowStart] = 0;
    for (let x = 0; x < width; x += 1) {
      const cellColumn = Math.floor(x / THEME_PNG_CELL_SIZE);
      const cellRow = Math.floor(y / THEME_PNG_CELL_SIZE);
      const colorIndex = cellRow * 5 + cellColumn;
      const color = colors[colorIndex] ?? [0, 0, 0];
      const pixelOffset = rowStart + 1 + x * channels;
      scanlines[pixelOffset] = color[0];
      scanlines[pixelOffset + 1] = color[1];
      scanlines[pixelOffset + 2] = color[2];
      if (channels === 4) {
        scanlines[pixelOffset + 3] = colorIndex === options.transparentCell && x % THEME_PNG_CELL_SIZE === 4 && y % THEME_PNG_CELL_SIZE === 4
          ? 0
          : 255;
      }
    }
  }

  const header = new Uint8Array(13);
  writeUint32(header, 0, width);
  writeUint32(header, 4, height);
  header[8] = 8;
  header[9] = colorType;

  return concatenate([
    new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    createPngChunk("IHDR", header),
    ...(options.includeImageData === false
      ? []
      : [createPngChunk("IDAT", options.invalidImageData ? new Uint8Array([1, 2, 3]) : zlibSync(scanlines))]),
    createPngChunk("IEND", new Uint8Array()),
  ]);
}

function createPngChunk(type: string, data: Uint8Array): Uint8Array {
  const typeAndData = new Uint8Array(4 + data.length);
  for (let index = 0; index < 4; index += 1) typeAndData[index] = type.charCodeAt(index);
  typeAndData.set(data, 4);

  const chunk = new Uint8Array(data.length + 12);
  writeUint32(chunk, 0, data.length);
  chunk.set(typeAndData, 4);
  writeUint32(chunk, chunk.length - 4, crc32(typeAndData));
  return chunk;
}

function concatenate(chunks: Uint8Array[]): Uint8Array {
  const result = new Uint8Array(chunks.reduce((length, chunk) => length + chunk.length, 0));
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

function writeUint32(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = value >>> 24;
  bytes[offset + 1] = value >>> 16;
  bytes[offset + 2] = value >>> 8;
  bytes[offset + 3] = value;
}

function mixHex(base: string, mix: string, mixRatio: number): string {
  const channels = [0, 2, 4].map((offset) => {
    const baseChannel = Number.parseInt(base.slice(offset + 1, offset + 3), 16);
    const mixChannel = Number.parseInt(mix.slice(offset + 1, offset + 3), 16);
    return Math.round(baseChannel * (1 - mixRatio) + mixChannel * mixRatio);
  });
  return `#${channels.map((channel) => channel.toString(16).padStart(2, "0")).join("")}`;
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const buffer = new ArrayBuffer(bytes.length);
  new Uint8Array(buffer).set(bytes);
  return buffer;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}
