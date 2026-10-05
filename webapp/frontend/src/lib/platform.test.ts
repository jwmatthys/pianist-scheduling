import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("@tauri-apps/api/core", () => ({
  invoke: vi.fn(),
  isTauri: () => false,
}));
vi.mock("@tauri-apps/plugin-dialog", () => ({ open: vi.fn(), save: vi.fn() }));
vi.mock("@tauri-apps/plugin-fs", () => ({ readFile: vi.fn(), writeFile: vi.fn() }));

import { saveBinaryFile, saveTextFile } from "./platform";

describe("local Save As helpers", () => {
  afterEach(() => {
    Reflect.deleteProperty(window, "showSaveFilePicker");
  });

  it("requests a Markdown save location with a valid bare MIME type", async () => {
    const write = vi.fn();
    const close = vi.fn();
    const picker = vi.fn().mockResolvedValue({
      createWritable: vi.fn().mockResolvedValue({ write, close }),
    });
    Object.defineProperty(window, "showSaveFilePicker", { configurable: true, value: picker });

    const saved = await saveTextFile("# Synthetic report", "lesson_pianists.md", "Markdown report", "md");

    expect(saved).toBe(true);
    expect(picker).toHaveBeenCalledWith({
      suggestedName: "lesson_pianists.md",
      types: [{ description: "Markdown report", accept: { "text/markdown": [".md"] } }],
    });
    expect(write).toHaveBeenCalledOnce();
    expect(close).toHaveBeenCalledOnce();
  });

  it("requests a PDF save location and writes the binary report", async () => {
    const write = vi.fn();
    const close = vi.fn();
    const picker = vi.fn().mockResolvedValue({
      createWritable: vi.fn().mockResolvedValue({ write, close }),
    });
    Object.defineProperty(window, "showSaveFilePicker", { configurable: true, value: picker });
    const pdf = new Blob(["%PDF-synthetic"], { type: "application/pdf" });

    const saved = await saveBinaryFile(pdf, "lesson_pianists.pdf", "PDF report", "pdf");

    expect(saved).toBe(true);
    expect(picker).toHaveBeenCalledWith({
      suggestedName: "lesson_pianists.pdf",
      types: [{ description: "PDF report", accept: { "application/pdf": [".pdf"] } }],
    });
    expect(write).toHaveBeenCalledWith(pdf);
    expect(close).toHaveBeenCalledOnce();
  });
});
