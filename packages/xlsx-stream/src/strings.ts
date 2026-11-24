// Shared-strings table kept off the JS heap.
//
// Strings are appended as UTF-8 into ArrayBuffer pages; three typed arrays (page, offset, byte
// length: 10 bytes per string) index them. A string never straddles two pages. `get` decodes on
// demand, so memory is the packed bytes plus the index, not a JS string per entry.
import { parseSax } from "./excelInternals";
import { ImportError } from "./errors";
import type { ZipArchive } from "./zip";

/** Default cap for packed bytes + index (a bigger table aborts the import cleanly). */
export const STRING_BUDGET_BYTES = 256 * 1024 * 1024;
const DEFAULT_PAGE_BYTES = 8 * 1024 * 1024;
const FIRST_PAGE_BYTES = 64 * 1024;
const INDEX_BYTES_PER_STRING = 10;
const MAX_PAGES = 0xffff;

export interface PackedStringsOptions {
  /** Largest page. Default 8 MiB. */
  pageBytes?: number;
  /** Budget for packed bytes plus index. Default STRING_BUDGET_BYTES. */
  budgetBytes?: number;
}

interface Page {
  data: Uint8Array;
  used: number;
}

export class PackedStrings {
  private readonly encoder = new TextEncoder();
  private readonly decoder = new TextDecoder();
  private readonly pageBytes: number;
  private readonly budget: number;
  private pages: Page[] = [];
  private allocated = 0;
  private count = 0;
  private pageOf = new Uint16Array(1024);
  private offsetOf = new Uint32Array(1024);
  private lengthOf = new Uint32Array(1024);

  constructor(options: PackedStringsOptions = {}) {
    this.pageBytes = Math.max(16, options.pageBytes ?? DEFAULT_PAGE_BYTES);
    this.budget = options.budgetBytes ?? STRING_BUDGET_BYTES;
  }

  get length(): number {
    return this.count;
  }

  /** Packed bytes (used part of the pages) plus the index. */
  get bytes(): number {
    return this.pages.reduce((sum, page) => sum + page.used, 0) + this.count * INDEX_BYTES_PER_STRING;
  }

  add(value: string): number {
    const encoded = this.encoder.encode(value);
    let page = this.pages[this.pages.length - 1];
    if (!page || page.used + encoded.length > page.data.length) {
      const grown = Math.min(this.pageBytes, Math.max(FIRST_PAGE_BYTES, (page?.data.length ?? FIRST_PAGE_BYTES / 2) * 2));
      const size = Math.max(grown, encoded.length);
      if (this.allocated + size + (this.count + 1) * INDEX_BYTES_PER_STRING > this.budget || this.pages.length >= MAX_PAGES) {
        throw this.tooLarge();
      }
      this.allocated += size;
      page = { data: new Uint8Array(size), used: 0 };
      this.pages.push(page);
    }
    const index = this.count;
    if (index >= this.pageOf.length) this.growIndex();
    this.pageOf[index] = this.pages.length - 1;
    this.offsetOf[index] = page.used;
    this.lengthOf[index] = encoded.length;
    page.data.set(encoded, page.used);
    page.used += encoded.length;
    this.count++;
    return index;
  }

  get(index: number): string {
    if (!(index >= 0 && index < this.count)) return "";
    const length = this.lengthOf[index];
    if (length === 0) return "";
    const offset = this.offsetOf[index];
    return this.decoder.decode(this.pages[this.pageOf[index]].data.subarray(offset, offset + length));
  }

  private growIndex(): void {
    const size = this.pageOf.length * 2;
    const page = new Uint16Array(size);
    page.set(this.pageOf);
    this.pageOf = page;
    const offset = new Uint32Array(size);
    offset.set(this.offsetOf);
    this.offsetOf = offset;
    const length = new Uint32Array(size);
    length.set(this.lengthOf);
    this.lengthOf = length;
  }

  private tooLarge(): ImportError {
    const size = this.budget >= 1024 * 1024 ? `${Math.round(this.budget / 1024 / 1024)} MB` : `${Math.round(this.budget / 1024)} KB`;
    return new ImportError(
      "STRING_TABLE_TOO_LARGE",
      `This file has too many distinct text values to import in the browser (more than ${size} of text).`,
    );
  }
}

/** Excel writes control characters as _xHHHH_ (a literal "_x" is written _x005F_x). */
export function decodeEscapes(value: string): string {
  return value.includes("_x") ? value.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, hex: string) => String.fromCharCode(parseInt(hex, 16))) : value;
}

/**
 * Reads xl/sharedStrings.xml into a PackedStrings. Rich text runs are joined into plain text and
 * phonetic runs (<rPh>) are skipped. Returns null when the file has no shared strings.
 */
export async function loadSharedStrings(
  zip: ZipArchive,
  options: PackedStringsOptions & { signal?: AbortSignal } = {},
): Promise<PackedStrings | null> {
  if (!zip.has("xl/sharedStrings.xml")) return null;
  const strings = new PackedStrings(options);
  let current: string[] | null = null;
  let inText = false;
  let phonetic = 0;
  for await (const events of parseSax(zip.text("xl/sharedStrings.xml", { signal: options.signal }))) {
    for (const { eventType, value } of events) {
      if (eventType === "opentag") {
        if (value.name === "si") current = [];
        else if (value.name === "rPh") phonetic++;
        else if (value.name === "t" && current && phonetic === 0) inText = true;
      } else if (eventType === "text") {
        if (inText && current) current.push(value);
      } else if (value.name === "t") {
        inText = false;
      } else if (value.name === "rPh") {
        phonetic--;
      } else if (value.name === "si" && current) {
        strings.add(decodeEscapes(current.join("")));
        current = null;
      }
    }
  }
  return strings;
}
