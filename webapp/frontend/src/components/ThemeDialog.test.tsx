import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import { zlibSync } from "fflate";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { DEFAULT_THEME_PALETTE } from "../theme/ThemePalette";
import { ThemeDialog } from "./ThemeDialog";

const mocks = vi.hoisted(() => ({
  activate: vi.fn(),
  chooseLocalFile: vi.fn(),
  selectionId: "default" as "default" | "dark" | "custom",
  palette: {
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
  },
}));

vi.mock("../lib/platform", () => ({
  chooseLocalFile: mocks.chooseLocalFile,
}));

vi.mock("../theme/ThemeContext", () => ({
  useTheme: () => ({
    selection: { id: mocks.selectionId },
    palette: mocks.palette,
    activate: mocks.activate,
  }),
}));

describe("ThemeDialog", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mocks.selectionId = "default";
    Object.defineProperty(HTMLDialogElement.prototype, "showModal", {
      configurable: true,
      value(this: HTMLDialogElement) {
        this.setAttribute("open", "");
      },
    });
    Object.defineProperty(HTMLDialogElement.prototype, "close", {
      configurable: true,
      value(this: HTMLDialogElement) {
        this.removeAttribute("open");
      },
    });
  });

  it("opens an accessible dialog with built-in choices, previews, and an active indicator", () => {
    render(<ThemeDialog />);
    fireEvent.click(screen.getByRole("button", { name: "Choose theme" }));

    expect(screen.getByRole("dialog", { name: "Theme" }).hasAttribute("open")).toBe(true);
    expect(screen.getByRole("button", { name: "Default Theme Active" }).getAttribute("aria-pressed")).toBe("true");
    expect(screen.getByRole("button", { name: "Dark Theme" }).getAttribute("aria-pressed")).toBe("false");
    expect(screen.getByRole("button", { name: "Custom Theme" })).not.toBeNull();
    expect(screen.getByRole("button", { name: "Upload Custom Theme" })).not.toBeNull();
    expect(screen.getAllByLabelText("Theme options")).toHaveLength(1);
    expect(screen.getByRole("button", { name: "Custom Theme" }).querySelector(".theme-preview")).toBeNull();
    expect(screen.getByRole("button", { name: "Choose theme" }).querySelector("svg")).not.toBeNull();
  });

  it("activates the selected built-in theme", async () => {
    render(<ThemeDialog />);
    fireEvent.click(screen.getByRole("button", { name: "Choose theme" }));
    fireEvent.click(screen.getByRole("button", { name: "Dark Theme" }));

    await waitFor(() => expect(mocks.activate).toHaveBeenCalledWith({ id: "dark" }));
  });

  it("shows a custom palette preview only after a custom theme is active", () => {
    mocks.selectionId = "custom";
    render(<ThemeDialog />);
    fireEvent.click(screen.getByRole("button", { name: "Choose theme" }));

    expect(screen.getByRole("button", { name: "Custom Theme Active" }).querySelectorAll(".theme-preview span")).toHaveLength(10);
  });

  it("validates and applies a custom PNG locally", async () => {
    const png = createPalettePng();
    const file = new File([png], "custom-theme.png", { type: "image/png" });
    Object.defineProperty(file, "arrayBuffer", {
      value: async () => png.buffer.slice(png.byteOffset, png.byteOffset + png.byteLength),
    });
    mocks.chooseLocalFile.mockResolvedValue(file);

    render(<ThemeDialog />);
    fireEvent.click(screen.getByRole("button", { name: "Choose theme" }));
    fireEvent.click(screen.getByRole("button", { name: "Upload Custom Theme" }));

    await waitFor(() => expect(mocks.activate).toHaveBeenCalledWith({
      id: "custom",
      palette: DEFAULT_THEME_PALETTE,
    }));
    expect(mocks.chooseLocalFile).toHaveBeenCalledWith("Theme palette PNG", ["png"]);
    expect(screen.queryByRole("alert")).toBeNull();
  });

  it("rejects invalid PNG data and leaves the active theme unchanged", async () => {
    const invalidFile = new File([new Uint8Array([1, 2, 3])], "invalid.png", { type: "image/png" });
    Object.defineProperty(invalidFile, "arrayBuffer", {
      value: async () => new Uint8Array([1, 2, 3]).buffer,
    });
    mocks.chooseLocalFile.mockResolvedValue(invalidFile);

    render(<ThemeDialog />);
    fireEvent.click(screen.getByRole("button", { name: "Choose theme" }));
    fireEvent.click(screen.getByRole("button", { name: "Upload Custom Theme" }));

    expect((await screen.findByRole("alert")).textContent).toMatch(/valid PNG/);
    expect(mocks.activate).not.toHaveBeenCalled();
  });
});

function createPalettePng(): Uint8Array {
  const colors = Object.values(DEFAULT_THEME_PALETTE).map((value) => (
    [0, 2, 4].map((index) => Number.parseInt(value.slice(index + 1, index + 3), 16))
  ));
  const scanlines = new Uint8Array((40 * 3 + 1) * 16);
  for (let y = 0; y < 16; y += 1) {
    const rowStart = y * (40 * 3 + 1);
    for (let x = 0; x < 40; x += 1) {
      const color = colors[Math.floor(y / 8) * 5 + Math.floor(x / 8)];
      const pixelOffset = rowStart + 1 + x * 3;
      scanlines[pixelOffset] = color[0];
      scanlines[pixelOffset + 1] = color[1];
      scanlines[pixelOffset + 2] = color[2];
    }
  }
  const header = new Uint8Array(13);
  writeUint32(header, 0, 40);
  writeUint32(header, 4, 16);
  header[8] = 8;
  header[9] = 2;
  return concatenate([
    new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    createChunk("IHDR", header),
    createChunk("IDAT", zlibSync(scanlines)),
    createChunk("IEND", new Uint8Array()),
  ]);
}

function createChunk(type: string, data: Uint8Array): Uint8Array {
  const payload = new Uint8Array(4 + data.length);
  for (let index = 0; index < 4; index += 1) payload[index] = type.charCodeAt(index);
  payload.set(data, 4);
  const chunk = new Uint8Array(data.length + 12);
  writeUint32(chunk, 0, data.length);
  chunk.set(payload, 4);
  writeUint32(chunk, chunk.length - 4, crc32(payload));
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
  bytes[offset + 3] = value >>> 0;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
