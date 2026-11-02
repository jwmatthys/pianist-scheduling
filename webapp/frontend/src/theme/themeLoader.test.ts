import { readFileSync } from "node:fs";
import { join } from "node:path";
import { zlibSync } from "fflate";
import { describe, expect, it, vi } from "vitest";
import { parseThemePalettePng, ThemePngError, THEME_PNG_CELL_SIZE, THEME_PNG_HEIGHT, THEME_PNG_WIDTH } from "./pngPalette";
import {
  applyThemeCssVariables,
  createThemeCssVariables,
  DEFAULT_THEME_PALETTE,
  loadThemePalette,
  validateThemePalette,
  type ThemePalette,
} from "./themeLoader";

const themeCss = readFileSync(join(process.cwd(), "src", "theme", "theme.css"), "utf8");

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
    await expect(loadThemePalette(async () => palette)).resolves.toEqual({
      palette,
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
});

describe("theme PNG parser", () => {
  it("extracts the ten center-pixel colors in the documented cell order", () => {
    expect(parseThemePalettePng(createPalettePng())).toEqual(palette);
  });

  it("accepts non-interlaced RGB PNGs", () => {
    expect(parseThemePalettePng(createPalettePng({ colorType: 2 }))).toEqual(palette);
  });

  it("rejects images with incorrect dimensions", () => {
    expect(() => parseThemePalettePng(createPalettePng({ width: THEME_PNG_WIDTH - 1 })))
      .toThrow(new ThemePngError("Theme PNG must be exactly 40x16 pixels."));
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
});

function createPalettePng(options: { width?: number; transparentCell?: number; colorType?: 2 | 6 } = {}): Uint8Array {
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
    createPngChunk("IDAT", zlibSync(scanlines)),
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

function contrastRatio(foreground: string, background: string): number {
  const luminance = (color: string): number => {
    const channels = [0, 2, 4].map((offset) => Number.parseInt(color.slice(offset + 1, offset + 3), 16) / 255);
    const linear = channels.map((channel) => (
      channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4
    ));
    return linear[0] * 0.2126 + linear[1] * 0.7152 + linear[2] * 0.0722;
  };
  const foregroundLuminance = luminance(foreground);
  const backgroundLuminance = luminance(background);
  return (Math.max(foregroundLuminance, backgroundLuminance) + 0.05)
    / (Math.min(foregroundLuminance, backgroundLuminance) + 0.05);
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
