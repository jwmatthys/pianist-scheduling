import { unzlibSync } from "fflate";
import type { ThemePalette } from "./ThemePalette";
import { validateThemePalette } from "./themeValidation";

export const THEME_PNG_WIDTH = 40;
export const THEME_PNG_HEIGHT = 16;
export const THEME_PNG_CELL_SIZE = 8;

const PNG_SIGNATURE = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
// Each tuple is [palette key, zero-based cell column, zero-based cell row].
// Sample centers are primary(4,4), accent(12,4), background(20,4), surface(28,4), text(36,4),
// textMuted(4,12), border(12,12), success(20,12), warning(28,12), and error(36,12).
const PALETTE_CELLS = [
  ["primary", 0, 0],
  ["accent", 1, 0],
  ["background", 2, 0],
  ["surface", 3, 0],
  ["text", 4, 0],
  ["textMuted", 0, 1],
  ["border", 1, 1],
  ["success", 2, 1],
  ["warning", 3, 1],
  ["error", 4, 1],
] as const satisfies readonly (readonly [keyof ThemePalette, number, number])[];

export type ThemePngErrorCode =
  | "PNG_INVALID"
  | "PNG_DIMENSIONS_INVALID"
  | "PNG_FORMAT_UNSUPPORTED"
  | "PNG_IMAGE_DATA_MISSING"
  | "PNG_IMAGE_DATA_INVALID"
  | "PALETTE_INVALID";

export class ThemePngError extends Error {
  constructor(message: string, readonly code: ThemePngErrorCode = "PNG_INVALID") {
    super(message);
    this.name = "ThemePngError";
  }
}

interface PngImage {
  width: number;
  height: number;
  bitDepth: number;
  colorType: number;
  interlace: number;
  imageData: Uint8Array;
}

export function parseThemePalettePng(input: ArrayBuffer | Uint8Array): ThemePalette {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const image = decodePng(bytes);
  if (image.width !== THEME_PNG_WIDTH || image.height !== THEME_PNG_HEIGHT) {
    throw new ThemePngError(
      `Theme PNG must be exactly ${THEME_PNG_WIDTH}x${THEME_PNG_HEIGHT} pixels.`,
      "PNG_DIMENSIONS_INVALID",
    );
  }
  if (image.bitDepth !== 8 || (image.colorType !== 2 && image.colorType !== 6) || image.interlace !== 0) {
    throw new ThemePngError("Theme PNG must use non-interlaced 8-bit RGB or RGBA color.", "PNG_FORMAT_UNSUPPORTED");
  }

  const channels = image.colorType === 2 ? 3 : 4;
  const rowLength = image.width * channels;
  const decoded = unfilterImageData(image.imageData, image.width, image.height, channels);
  const palette: Record<string, string> = {};

  for (const [name, column, row] of PALETTE_CELLS) {
    const x = column * THEME_PNG_CELL_SIZE + Math.floor(THEME_PNG_CELL_SIZE / 2);
    const y = row * THEME_PNG_CELL_SIZE + Math.floor(THEME_PNG_CELL_SIZE / 2);
    const offset = y * rowLength + x * channels;
    if (channels === 4 && decoded[offset + 3] !== 255) {
      throw new ThemePngError(`Theme PNG palette cell "${name}" must be fully opaque.`, "PALETTE_INVALID");
    }
    palette[name] = `#${[0, 1, 2].map((channel) => decoded[offset + channel].toString(16).padStart(2, "0")).join("")}`;
  }

  try {
    return validateThemePalette(palette);
  } catch (error) {
    throw new ThemePngError(
      error instanceof Error ? error.message : "Theme PNG palette is invalid.",
      "PALETTE_INVALID",
    );
  }
}

function decodePng(bytes: Uint8Array): PngImage {
  if (bytes.length < PNG_SIGNATURE.length || !PNG_SIGNATURE.every((byte, index) => bytes[index] === byte)) {
    throw new ThemePngError("Theme file is not a valid PNG.");
  }

  let offset = PNG_SIGNATURE.length;
  let header: Omit<PngImage, "imageData"> | null = null;
  let imageDataLength = 0;
  const imageDataChunks: Uint8Array[] = [];
  let sawImageData = false;
  let imageDataEnded = false;
  let sawPalette = false;
  let sawEnd = false;

  while (offset + 12 <= bytes.length) {
    const length = readUint32(bytes, offset);
    const chunkStart = offset + 8;
    const chunkEnd = chunkStart + length;
    const crcEnd = chunkEnd + 4;
    if (chunkEnd < chunkStart || crcEnd > bytes.length) {
      throw new ThemePngError("Theme PNG contains a truncated chunk.");
    }

    const typeBytes = bytes.subarray(offset + 4, offset + 8);
    const type = String.fromCharCode(...typeBytes);
    if (readUint32(bytes, chunkEnd) !== crc32(bytes.subarray(offset + 4, chunkEnd))) {
      throw new ThemePngError(`Theme PNG contains a corrupt ${type} chunk.`);
    }
    const data = bytes.subarray(chunkStart, chunkEnd);

    if (!header && type !== "IHDR") {
      throw new ThemePngError("Theme PNG must begin with an IHDR chunk.");
    }
    if (type === "IHDR") {
      if (header || length !== 13) throw new ThemePngError("Theme PNG has an invalid IHDR chunk.");
      header = {
        width: readUint32(data, 0),
        height: readUint32(data, 4),
        bitDepth: data[8],
        colorType: data[9],
        interlace: data[12],
      };
      if (header.width !== THEME_PNG_WIDTH || header.height !== THEME_PNG_HEIGHT) {
        throw new ThemePngError(
          `Theme PNG must be exactly ${THEME_PNG_WIDTH}x${THEME_PNG_HEIGHT} pixels.`,
          "PNG_DIMENSIONS_INVALID",
        );
      }
      if (header.bitDepth !== 8 || (header.colorType !== 2 && header.colorType !== 6) || header.interlace !== 0) {
        throw new ThemePngError(
          "Theme PNG must use non-interlaced 8-bit RGB or RGBA color.",
          "PNG_FORMAT_UNSUPPORTED",
        );
      }
      if (data[10] !== 0 || data[11] !== 0) {
        throw new ThemePngError("Theme PNG uses unsupported compression or filtering.");
      }
    } else if (type === "IDAT") {
      if (imageDataEnded) throw new ThemePngError("Theme PNG IDAT chunks must be consecutive.");
      sawImageData = true;
      imageDataChunks.push(data);
      imageDataLength += data.length;
    } else if (type === "PLTE") {
      if (sawImageData || sawPalette || length === 0 || length > 768 || length % 3 !== 0) {
        throw new ThemePngError("Theme PNG has an invalid PLTE chunk.");
      }
      sawPalette = true;
    } else if (type === "tRNS") {
      throw new ThemePngError("Theme PNG transparency-key chunks are not supported.");
    } else if (type === "IEND") {
      if (!sawImageData) {
        throw new ThemePngError("Theme PNG is missing image data.", "PNG_IMAGE_DATA_MISSING");
      }
      if (length !== 0) throw new ThemePngError("Theme PNG has an invalid IEND chunk.");
      sawEnd = true;
      offset = crcEnd;
      break;
    } else {
      if (sawImageData) imageDataEnded = true;
      if ((typeBytes[0] & 0x20) === 0) {
        throw new ThemePngError(`Theme PNG contains an unsupported critical ${type} chunk.`);
      }
    }

    offset = crcEnd;
  }

  if (!sawEnd || !header || offset !== bytes.length) {
    throw new ThemePngError("Theme PNG is incomplete or contains trailing data.", "PNG_IMAGE_DATA_MISSING");
  }

  const compressed = new Uint8Array(imageDataLength);
  let chunkOffset = 0;
  for (const chunk of imageDataChunks) {
    compressed.set(chunk, chunkOffset);
    chunkOffset += chunk.length;
  }

  let imageData: Uint8Array;
  try {
    imageData = unzlibSync(compressed);
  } catch {
    throw new ThemePngError("Theme PNG image data is corrupt.", "PNG_IMAGE_DATA_INVALID");
  }

  return { ...header, imageData };
}

function unfilterImageData(data: Uint8Array, width: number, height: number, channels: number): Uint8Array {
  const rowLength = width * channels;
  const expectedLength = (rowLength + 1) * height;
  if (data.length !== expectedLength) {
    throw new ThemePngError("Theme PNG image data has an invalid length.", "PNG_IMAGE_DATA_INVALID");
  }

  const pixels = new Uint8Array(rowLength * height);
  let sourceOffset = 0;
  for (let y = 0; y < height; y += 1) {
    const filter = data[sourceOffset];
    sourceOffset += 1;
    if (filter > 4) throw new ThemePngError("Theme PNG uses an invalid row filter.");
    const rowOffset = y * rowLength;

    for (let x = 0; x < rowLength; x += 1) {
      const raw = data[sourceOffset + x];
      const left = x >= channels ? pixels[rowOffset + x - channels] : 0;
      const above = y > 0 ? pixels[rowOffset + x - rowLength] : 0;
      const upperLeft = y > 0 && x >= channels ? pixels[rowOffset + x - rowLength - channels] : 0;
      let predictor = 0;

      if (filter === 1) predictor = left;
      else if (filter === 2) predictor = above;
      else if (filter === 3) predictor = Math.floor((left + above) / 2);
      else if (filter === 4) predictor = paethPredictor(left, above, upperLeft);

      pixels[rowOffset + x] = (raw + predictor) & 0xff;
    }
    sourceOffset += rowLength;
  }
  return pixels;
}

function paethPredictor(left: number, above: number, upperLeft: number): number {
  const estimate = left + above - upperLeft;
  const leftDistance = Math.abs(estimate - left);
  const aboveDistance = Math.abs(estimate - above);
  const upperLeftDistance = Math.abs(estimate - upperLeft);
  if (leftDistance <= aboveDistance && leftDistance <= upperLeftDistance) return left;
  return aboveDistance <= upperLeftDistance ? above : upperLeft;
}

function readUint32(bytes: Uint8Array, offset: number): number {
  return ((bytes[offset] * 0x1000000) + (bytes[offset + 1] << 16) + (bytes[offset + 2] << 8) + bytes[offset + 3]) >>> 0;
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
