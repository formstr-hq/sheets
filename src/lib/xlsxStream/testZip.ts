// Test helper: hand-built zip files, so tests can produce layouts fflate's zipSync does not
// (data descriptors, ZIP64 records, encrypted flag, odd compression methods).
import { deflateSync } from "fflate";

export interface TestEntry {
  name: string;
  data: string | Uint8Array;
  /** 0 stored, 8 deflate (default), anything else is written as-is with that method id. */
  method?: number;
  /** Sizes live in a trailing data descriptor (flag bit 3) and are zero in the local header. */
  descriptor?: boolean;
  /** Sets general-purpose flag bit 0. */
  encrypted?: boolean;
}

export interface TestZipOptions {
  zip64?: boolean;
  comment?: string;
  /** Appended after the end-of-directory record, no effect on the structure. */
  junkAfterComment?: never;
}

class Out {
  parts: Uint8Array[] = [];
  length = 0;
  push(bytes: Uint8Array) {
    this.parts.push(bytes);
    this.length += bytes.length;
  }
  u16(v: number) {
    this.push(new Uint8Array([v & 255, (v >> 8) & 255]));
  }
  u32(v: number) {
    this.push(new Uint8Array([v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]));
  }
  u64(v: number) {
    this.u32(v % 2 ** 32);
    this.u32(Math.floor(v / 2 ** 32));
  }
  done(): Uint8Array {
    const all = new Uint8Array(this.length);
    let at = 0;
    for (const part of this.parts) {
      all.set(part, at);
      at += part.length;
    }
    return all;
  }
}

const utf8 = (value: string) => new TextEncoder().encode(value);

export function buildZip(entries: TestEntry[], options: TestZipOptions = {}): Uint8Array {
  const out = new Out();
  const central: Array<{ e: TestEntry; nameBytes: Uint8Array; csize: number; usize: number; offset: number; method: number }> = [];
  for (const e of entries) {
    const raw = typeof e.data === "string" ? utf8(e.data) : e.data;
    const method = e.method ?? 8;
    const body = method === 8 ? deflateSync(raw) : raw;
    const nameBytes = utf8(e.name);
    const offset = out.length;
    const flags = (e.descriptor ? 8 : 0) | (e.encrypted ? 1 : 0);
    out.u32(0x04034b50);
    out.u16(20);
    out.u16(flags);
    out.u16(method);
    out.u16(0);
    out.u16(0);
    out.u32(0); // crc (not checked by the reader)
    out.u32(e.descriptor ? 0 : body.length);
    out.u32(e.descriptor ? 0 : raw.length);
    out.u16(nameBytes.length);
    out.u16(0);
    out.push(nameBytes);
    out.push(body);
    if (e.descriptor) {
      out.u32(0x08074b50);
      out.u32(0);
      out.u32(body.length);
      out.u32(raw.length);
    }
    central.push({ e, nameBytes, csize: body.length, usize: raw.length, offset, method });
  }
  const cdStart = out.length;
  for (const c of central) {
    const flags = (c.e.descriptor ? 8 : 0) | (c.e.encrypted ? 1 : 0);
    const z64 = options.zip64 === true;
    out.u32(0x02014b50);
    out.u16(45);
    out.u16(z64 ? 45 : 20);
    out.u16(flags);
    out.u16(c.method);
    out.u16(0);
    out.u16(0);
    out.u32(0);
    out.u32(z64 ? 0xffffffff : c.csize);
    out.u32(z64 ? 0xffffffff : c.usize);
    out.u16(c.nameBytes.length);
    out.u16(z64 ? 28 : 0);
    out.u16(0);
    out.u16(0);
    out.u16(0);
    out.u32(0);
    out.u32(z64 ? 0xffffffff : c.offset);
    out.push(c.nameBytes);
    if (z64) {
      out.u16(1);
      out.u16(24);
      out.u64(c.usize);
      out.u64(c.csize);
      out.u64(c.offset);
    }
  }
  const cdSize = out.length - cdStart;
  const comment = utf8(options.comment ?? "");
  if (options.zip64) {
    const eocd64At = out.length;
    out.u32(0x06064b50);
    out.u64(44);
    out.u16(45);
    out.u16(45);
    out.u32(0);
    out.u32(0);
    out.u64(entries.length);
    out.u64(entries.length);
    out.u64(cdSize);
    out.u64(cdStart);
    out.u32(0x07064b50);
    out.u32(0);
    out.u64(eocd64At);
    out.u32(1);
  }
  out.u32(0x06054b50);
  out.u16(0);
  out.u16(0);
  out.u16(options.zip64 ? 0xffff : entries.length);
  out.u16(options.zip64 ? 0xffff : entries.length);
  out.u32(options.zip64 ? 0xffffffff : cdSize);
  out.u32(options.zip64 ? 0xffffffff : cdStart);
  out.u16(comment.length);
  out.push(comment);
  return out.done();
}

export const zipBlob = (entries: TestEntry[], options?: TestZipOptions) => new Blob([buildZip(entries, options) as BlobPart]);
