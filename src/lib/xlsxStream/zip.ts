// Random-access zip reader over a Blob/File.
//
// Reads the central directory (sizes and offsets come from there, never from local headers, so
// entries written with a data descriptor work), then streams one entry at a time: slices of the
// blob are inflated with fflate and decoded to text with one streaming TextDecoder per entry, so
// multi-byte UTF-8 split across slice boundaries stays intact.
// Only stored (0) and deflated (8) entries are supported. ZIP64 is supported.
// CRC-32 is not verified: it would cost a pass over the whole inflated data, and damage that
// inflates cleanly still has to be valid XML to be read.
import { Inflate } from "fflate";

export type ZipErrorCode = "NOT_ZIP" | "UNSUPPORTED_ZIP" | "CORRUPT_ZIP" | "MISSING_ENTRY" | "CANCELLED";

export class ZipError extends Error {
  readonly code: ZipErrorCode;
  constructor(code: ZipErrorCode, message: string) {
    super(message);
    this.name = "ZipError";
    this.code = code;
  }
}

export interface ZipEntry {
  name: string;
  method: number;
  /** Compressed and uncompressed sizes from the central directory. */
  csize: number;
  usize: number;
  /** Offset of the local header. */
  offset: number;
  flags: number;
}

export interface ZipReadOptions {
  /** Compressed bytes read from the blob per step. Default 256 KiB. */
  readSize?: number;
  signal?: AbortSignal;
  /** Called after each slice with the compressed bytes of this entry read so far. */
  onRead?: (bytesRead: number) => void;
}

export interface ZipArchive {
  entries: Map<string, ZipEntry>;
  has(name: string): boolean;
  /** The entry decoded as UTF-8 text, in pieces of at most ~512 K characters. */
  text(name: string, options?: ZipReadOptions): AsyncGenerator<string>;
  /** The entry as raw (inflated) bytes. */
  bytes(name: string, options?: ZipReadOptions): AsyncGenerator<Uint8Array>;
}

const EOCD_SIG = 0x06054b50;
const EOCD64_LOCATOR_SIG = 0x07064b50;
const EOCD64_SIG = 0x06064b50;
const CENTRAL_SIG = 0x02014b50;
const LOCAL_SIG = 0x04034b50;
const MAX_COMMENT = 0xffff;
const MAX_CENTRAL_DIR_BYTES = 256 * 1024 * 1024;
/** Hard ceiling for one inflated entry, whatever the zip claims. */
const MAX_ENTRY_BYTES = 8 * 1024 * 1024 * 1024;
const DEFAULT_READ = 256 * 1024;
const TEXT_PIECE = 512 * 1024;

const normalise = (name: string) => (name.startsWith("/") ? name.slice(1) : name);

function safeNumber(value: bigint, what: string): number {
  if (value > BigInt(Number.MAX_SAFE_INTEGER)) throw new ZipError("CORRUPT_ZIP", `Zip ${what} is out of range`);
  return Number(value);
}

async function readRange(blob: Blob, start: number, end: number): Promise<Uint8Array> {
  if (start < 0 || end > blob.size || end < start) throw new ZipError("CORRUPT_ZIP", "Zip data runs past the end of the file");
  return new Uint8Array(await blob.slice(start, end).arrayBuffer());
}

export async function openZip(blob: Blob, defaults: ZipReadOptions = {}): Promise<ZipArchive> {
  const size = blob.size;
  if (size < 22) throw new ZipError("NOT_ZIP", "The file is not a zip archive");
  const tailLen = Math.min(size, MAX_COMMENT + 22);
  const tail = await readRange(blob, size - tailLen, size);
  const t = new DataView(tail.buffer, tail.byteOffset, tail.byteLength);

  let eocd = -1;
  for (let i = tailLen - 22; i >= 0; i--) {
    // The record must end exactly where the file ends (its comment length says how long that is);
    // a signature inside a comment does not.
    if (t.getUint32(i, true) === EOCD_SIG && i + 22 + t.getUint16(i + 20, true) === tailLen) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new ZipError("NOT_ZIP", "The file is not a zip archive");

  let count = t.getUint16(eocd + 10, true);
  let cdSize = t.getUint32(eocd + 12, true);
  let cdOffset = t.getUint32(eocd + 16, true);
  if (count === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
    const locator = eocd - 20;
    if (locator < 0 || t.getUint32(locator, true) !== EOCD64_LOCATOR_SIG) {
      throw new ZipError("CORRUPT_ZIP", "ZIP64 end-of-directory locator is missing");
    }
    const eocd64At = safeNumber(t.getBigUint64(locator + 8, true), "directory offset");
    const head = await readRange(blob, eocd64At, eocd64At + 56);
    const h = new DataView(head.buffer, head.byteOffset, head.byteLength);
    if (h.getUint32(0, true) !== EOCD64_SIG) throw new ZipError("CORRUPT_ZIP", "ZIP64 end-of-directory record is missing");
    count = safeNumber(h.getBigUint64(32, true), "entry count");
    cdSize = safeNumber(h.getBigUint64(40, true), "directory size");
    cdOffset = safeNumber(h.getBigUint64(48, true), "directory offset");
  }
  if (cdSize > MAX_CENTRAL_DIR_BYTES) throw new ZipError("CORRUPT_ZIP", "Zip directory is implausibly large");

  const cdBytes = await readRange(blob, cdOffset, cdOffset + cdSize);
  const cd = new DataView(cdBytes.buffer, cdBytes.byteOffset, cdBytes.byteLength);
  const decoder = new TextDecoder();
  const entries = new Map<string, ZipEntry>();
  let p = 0;
  for (let i = 0; i < count; i++) {
    if (p + 46 > cdSize || cd.getUint32(p, true) !== CENTRAL_SIG) throw new ZipError("CORRUPT_ZIP", "Zip directory is damaged");
    const flags = cd.getUint16(p + 8, true);
    const method = cd.getUint16(p + 10, true);
    let csize = cd.getUint32(p + 20, true);
    let usize = cd.getUint32(p + 24, true);
    const nameLen = cd.getUint16(p + 28, true);
    const extraLen = cd.getUint16(p + 30, true);
    const commentLen = cd.getUint16(p + 32, true);
    let offset = cd.getUint32(p + 42, true);
    if (p + 46 + nameLen + extraLen + commentLen > cdSize) throw new ZipError("CORRUPT_ZIP", "Zip directory is damaged");
    const name = decoder.decode(cdBytes.subarray(p + 46, p + 46 + nameLen));
    if (usize === 0xffffffff || csize === 0xffffffff || offset === 0xffffffff) {
      let q = p + 46 + nameLen;
      const end = q + extraLen;
      while (q + 4 <= end) {
        const id = cd.getUint16(q, true);
        const len = cd.getUint16(q + 2, true);
        if (id === 1) {
          let r = q + 4;
          if (usize === 0xffffffff) {
            usize = safeNumber(cd.getBigUint64(r, true), "entry size");
            r += 8;
          }
          if (csize === 0xffffffff) {
            csize = safeNumber(cd.getBigUint64(r, true), "entry size");
            r += 8;
          }
          if (offset === 0xffffffff) offset = safeNumber(cd.getBigUint64(r, true), "entry offset");
        }
        q += 4 + len;
      }
    }
    entries.set(normalise(name), { name, method, csize, usize, offset, flags });
    p += 46 + nameLen + extraLen + commentLen;
  }

  function lookup(name: string): ZipEntry {
    const entry = entries.get(normalise(name));
    if (!entry) throw new ZipError("MISSING_ENTRY", `The file has no "${name}" part`);
    if (entry.flags & 1) throw new ZipError("UNSUPPORTED_ZIP", "The file is encrypted");
    if (entry.method !== 0 && entry.method !== 8) {
      throw new ZipError("UNSUPPORTED_ZIP", `Unsupported zip compression method ${entry.method}`);
    }
    return entry;
  }

  async function dataStart(entry: ZipEntry): Promise<number> {
    const head = await readRange(blob, entry.offset, entry.offset + 30);
    const h = new DataView(head.buffer, head.byteOffset, head.byteLength);
    if (h.getUint32(0, true) !== LOCAL_SIG) throw new ZipError("CORRUPT_ZIP", "Zip entry header is damaged");
    const start = entry.offset + 30 + h.getUint16(26, true) + h.getUint16(28, true);
    if (start + entry.csize > size) throw new ZipError("CORRUPT_ZIP", "Zip entry runs past the end of the file");
    return start;
  }

  async function* bytes(name: string, options: ZipReadOptions = {}): AsyncGenerator<Uint8Array> {
    const entry = lookup(name);
    const readSize = Math.max(1, options.readSize ?? defaults.readSize ?? DEFAULT_READ);
    const signal = options.signal ?? defaults.signal;
    const onRead = options.onRead ?? defaults.onRead;
    const start = await dataStart(entry);
    // Slack over the declared size; inflating much more than that is treated as a zip bomb.
    const limit = Math.min(MAX_ENTRY_BYTES, entry.usize * 1.1 + 1024 * 1024);
    let produced = 0;
    const out: Uint8Array[] = [];
    const inflate = entry.method === 8 ? new Inflate((chunk) => out.push(chunk)) : null;
    if (entry.csize === 0) return;
    for (let o = 0; o < entry.csize; o += readSize) {
      if (signal?.aborted) throw new ZipError("CANCELLED", "The import was cancelled");
      const last = o + readSize >= entry.csize;
      const piece = await readRange(blob, start + o, start + Math.min(entry.csize, o + readSize));
      onRead?.(o + piece.length);
      if (!inflate) {
        yield piece;
        continue;
      }
      try {
        inflate.push(piece, last);
      } catch (error) {
        throw new ZipError("CORRUPT_ZIP", `Zip entry "${name}" is damaged: ${(error as Error).message}`);
      }
      while (out.length > 0) {
        const chunk = out.shift() as Uint8Array;
        produced += chunk.length;
        if (produced > limit) throw new ZipError("CORRUPT_ZIP", `Zip entry "${name}" is larger than it declares`);
        if (chunk.length > 0) yield chunk;
      }
    }
  }

  async function* text(name: string, options: ZipReadOptions = {}): AsyncGenerator<string> {
    const decode = new TextDecoder("utf-8");
    for await (const chunk of bytes(name, options)) {
      yield* pieces(decode.decode(chunk, { stream: true }));
    }
    yield* pieces(decode.decode());
  }

  return { entries, has: (name) => entries.has(normalise(name)), text, bytes };
}

/** Pieces of at most TEXT_PIECE characters, never cut inside a surrogate pair. */
function* pieces(value: string): Generator<string> {
  if (value.length === 0) return;
  let from = 0;
  while (value.length - from > TEXT_PIECE) {
    let to = from + TEXT_PIECE;
    const code = value.charCodeAt(to - 1);
    if (code >= 0xd800 && code <= 0xdbff) to--;
    yield value.slice(from, to);
    from = to;
  }
  yield value.slice(from);
}
