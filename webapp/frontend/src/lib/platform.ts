import { invoke, isTauri } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import { readFile, writeFile } from "@tauri-apps/plugin-fs";

let apiBasePromise: Promise<string> | undefined;

export function getApiBase(): Promise<string> {
  if (!apiBasePromise) {
    apiBasePromise = isTauri()
      ? invoke<string>("get_api_base")
      : Promise.resolve(
          new URLSearchParams(window.location.search).get("apiBase")
            ?? import.meta.env.VITE_API_BASE
            ?? "http://localhost:8123"
        );
  }
  return apiBasePromise;
}

export async function chooseLocalFile(
  description: string,
  extensions: string[],
): Promise<File | null> {
  if (isTauri()) {
    const selected = await open({
      multiple: false,
      filters: [{ name: description, extensions }],
    });
    if (!selected || Array.isArray(selected)) return null;
    const contents = await readFile(selected);
    const filename = selected.split(/[\\/]/).pop() ?? `availability.${extensions[0]}`;
    const bytes = contents.buffer.slice(contents.byteOffset, contents.byteOffset + contents.byteLength) as ArrayBuffer;
    return new File([bytes], filename);
  }

  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = extensions.map((extension) => `.${extension}`).join(",");
    input.addEventListener("change", () => resolve(input.files?.[0] ?? null), { once: true });
    input.addEventListener("cancel", () => resolve(null), { once: true });
    input.click();
  });
}

export async function chooseSessionArchive(): Promise<Uint8Array | null> {
  if (isTauri()) {
    const selected = await open({
      multiple: false,
      filters: [{ name: "Music Program Scheduler Session", extensions: ["mpsession"] }],
    });
    if (!selected || Array.isArray(selected)) return null;
    return readFile(selected);
  }

  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".mpsession,application/zip";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (!file) return resolve(null);
      void file.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)));
    }, { once: true });
    input.addEventListener("cancel", () => resolve(null), { once: true });
    input.click();
  });
}

export async function saveSessionArchive(contents: Uint8Array): Promise<void> {
  if (isTauri()) {
    const destination = await save({
      defaultPath: "music-program-session.mpsession",
      filters: [{ name: "Music Program Scheduler Session", extensions: ["mpsession"] }],
    });
    if (destination) await writeFile(destination, contents);
    return;
  }

  const browserWindow = window as Window & {
    showSaveFilePicker?: (options: { suggestedName: string; types: { description: string; accept: Record<string, string[]> }[] }) => Promise<{
      createWritable: () => Promise<{ write: (data: Uint8Array) => Promise<void>; close: () => Promise<void> }>;
    }>;
  };
  if (browserWindow.showSaveFilePicker) {
    const handle = await browserWindow.showSaveFilePicker({
      suggestedName: "music-program-session.mpsession",
      types: [{ description: "Session archive", accept: { "application/zip": [".mpsession"] } }],
    });
    const writable = await handle.createWritable();
    await writable.write(contents);
    await writable.close();
    return;
  }

  const url = URL.createObjectURL(new Blob([contents], { type: "application/zip" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = "music-program-session.mpsession";
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export async function saveBinaryFile(
  contents: Blob,
  suggestedName: string,
  description: string,
  extension: string,
): Promise<boolean> {
  if (isTauri()) {
    const destination = await save({
      defaultPath: suggestedName,
      filters: [{ name: description, extensions: [extension] }],
    });
    if (!destination) return false;
    await writeFile(destination, new Uint8Array(await contents.arrayBuffer()));
    return true;
  }

  const browserWindow = window as Window & {
    showSaveFilePicker?: (options: {
      suggestedName: string;
      types: { description: string; accept: Record<string, string[]> }[];
    }) => Promise<{
      createWritable: () => Promise<{
        write: (data: Blob) => Promise<void>;
        close: () => Promise<void>;
      }>;
    }>;
  };
  if (browserWindow.showSaveFilePicker) {
    try {
      const mimeType = contents.type.split(";")[0] || "application/octet-stream";
      const handle = await browserWindow.showSaveFilePicker({
        suggestedName,
        types: [{
          description,
          accept: { [mimeType]: [`.${extension}`] },
        }],
      });
      const writable = await handle.createWritable();
      await writable.write(contents);
      await writable.close();
      return true;
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") return false;
      throw error;
    }
  }

  const url = URL.createObjectURL(contents);
  const link = document.createElement("a");
  link.href = url;
  link.download = suggestedName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  return true;
}

export async function saveTextFile(
  contents: string,
  suggestedName: string,
  description: string,
  extension: string,
): Promise<boolean> {
  const mimeType = extension.toLowerCase() === "md" ? "text/markdown;charset=utf-8" : "text/csv;charset=utf-8";
  return saveBinaryFile(new Blob([contents], { type: mimeType }), suggestedName, description, extension);
}
