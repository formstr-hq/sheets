// Sheets v2 root manifest (kind 33460).
//
// The root is a small encrypted event. Each worksheet (tab) keeps its cells and
// layout in a compressed FSB2 payload that is stored one of two ways:
//
//   inline  - the payload is embedded in the root (small sheets, no server needed)
//   blossom - the payload is an encrypted Blossom blob; the root holds a BlobRef
//
// This file is pure: it never touches the network. Fetching blobs is passed in.

import { WORKBOOK_SCHEMA, type WorkbookDocument, type WorkbookSheet } from "../workbook";
import { base64ToBytes, bytesToBase64, frame, unframe } from "./codec";
import { assembleChunks, colsForSheet } from "./chunks";

export const ROOT_SCHEMA = "formstr.sheets.root";
export const ROOT_VERSION = 2;

/**
 * Largest root plaintext we will ever encrypt. The binding limit is the relay's
 * event size, not the cipher: relay.formstr.app caps event content at 102,400
 * bytes and nos.lol caps messages at 131,072 (both from their NIP-11 documents).
 * NIP-44 ciphertext is base64, so ~4/3 of the plaintext. 48 KB encrypts to about
 * 65 KB, which fits those relays and also stays inside the 65,535-byte plaintext
 * limit of the standard NIP-44 spec, so other NIP-44 clients can still decrypt it
 * (nostr-tools' extended format for larger payloads is not interoperable).
 */
export const ROOT_MAX_BYTES = 48 * 1024;

/** A root at or under this size stays inline; anything bigger moves to Blossom. */
export const INLINE_MAX_BYTES = 32 * 1024;

/** Pointer to one encrypted Blossom blob. All of it is secret: it sits inside the encrypted root. */
export interface BlobRef {
  /** sha256 of the ciphertext, which is the blob's address on Blossom. */
  x: string;
  /** Per-blob secret, from which the AES key is derived. */
  k: string;
  /** Ciphertext size in bytes. */
  n: number;
  /** Plaintext size in bytes (the FSB2 frame). */
  p: number;
  /** Encryption segment size in bytes. */
  c: number;
  /** sha256 of the plaintext frame; also lets an unchanged tab skip re-upload. */
  h: string;
}

/**
 * One chunk of a worksheet: a contiguous run of rows stored as exactly one blob. Rows
 * [rowStart, rowStart + rowCount) belong to it. The spec stores only rowCount (a chunk's start is the
 * sum of the counts before it); writers here also store rowStart, readers recompute it when absent.
 */
export interface ChunkEntry {
  id: string;
  ref: BlobRef;
  rowStart: number;
  rowCount: number;
  /** Cells in the chunk, when the writer counted them (streamed imports do). Optional; lets a reader size a tab without downloading it. */
  cells?: number;
}

type TabMeta = Pick<WorkbookSheet, "id" | "name" | "order" | "hidden" | "color" | "showGridLines">;
type TabBody = Omit<WorkbookSheet, keyof TabMeta>;

export interface RootTab extends TabMeta {
  /** Column ids in order. Chunked tabs only; an inline tab carries them in its payload. */
  cols?: string[];
  /** base64 of an FSB2 payload holding the rest of the worksheet. Inline storage. */
  inline?: string;
  /** Blossom blob holding that payload (the whole tab in one blob). Older roots only; still readable. */
  blob?: BlobRef;
  /**
   * The tab as chunks, in row order. When the list would push the root over its size limit it
   * lives in a tab-index blob (frame type 0x02) and this holds the blob's ref instead.
   */
  chunks?: ChunkEntry[] | { index: BlobRef };
}

export interface SheetRootV2 {
  schema: typeof ROOT_SCHEMA;
  v: typeof ROOT_VERSION;
  title: string;
  activeSheetId: string;
  createdAt: number;
  updatedAt: number;
  epoch: number;
  /** Blossom servers that hold this sheet's blobs. Empty for inline sheets. */
  servers: string[];
  tabs: RootTab[];
}

export type SheetStorage = "inline" | "blossom";

export class SheetTooLargeError extends Error {
  readonly bytes: number;
  constructor(bytes: number) {
    super(
      `This spreadsheet is too large to store in a single event (${Math.ceil(bytes / 1024)} KB compressed). ` +
        "Add a Blossom server to your Nostr profile (for example in Formstr Drive) to save larger sheets.",
    );
    this.name = "SheetTooLargeError";
    this.bytes = bytes;
  }
}

export class BlobUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "BlobUnavailableError";
  }
}

/**
 * Largest sheet the editor is allowed to open. The editor (FortuneSheet) holds every cell as JS
 * objects plus a dense row x column grid, so memory grows fast: a 675k-cell sheet opens, a 4M-cell
 * one needed ~3.7 GB. Not measured in a browser between those yet; 1M is a guess on the safe side.
 */
export const OPEN_MAX_CELLS = 1_000_000;
/** Same gate for chunks that carry no cell count: compressed (plaintext frame) bytes, ~10 bytes per cell. */
export const OPEN_MAX_BYTES = 12 * 1024 * 1024;

export interface SheetSize {
  /** Cells counted by the chunks that carry a count; null when no chunk does. */
  cells: number | null;
  /** Compressed bytes of all chunks. */
  bytes: number;
  tabs: { name: string; chunks: number; cells: number | null; bytes: number }[];
}

/** Raised before anything is downloaded when a sheet is too big for the editor. */
export class SheetTooLargeToOpenError extends Error {
  readonly size: SheetSize;
  /** The decrypted root, so the big-sheet view can read it chunk by chunk. */
  readonly root?: SheetRootV2;
  constructor(size: SheetSize, root?: SheetRootV2) {
    super("This spreadsheet is too large to open in the editor yet.");
    this.name = "SheetTooLargeToOpenError";
    this.size = size;
    this.root = root;
  }
}

/**
 * Size of a root's sheet from its manifest alone (no cell chunk is downloaded). A tab whose chunk
 * list lives in an index blob has hundreds of chunks, far past any open limit, so it is not fetched.
 */
export function measureRoot(root: SheetRootV2): SheetSize & { indexed: boolean } {
  let indexed = false;
  const tabs = root.tabs.map((tab) => {
    if (tab.chunks && !Array.isArray(tab.chunks)) {
      indexed = true;
      return { name: tab.name, chunks: 0, cells: null, bytes: tab.chunks.index.p };
    }
    const entries = tab.chunks ?? [];
    const counted = entries.filter((entry) => typeof entry.cells === "number");
    return {
      name: tab.name,
      chunks: entries.length,
      cells: counted.length ? counted.reduce((sum, entry) => sum + (entry.cells as number), 0) : null,
      bytes: entries.reduce((sum, entry) => sum + entry.ref.p, 0) + (tab.blob?.p ?? 0),
    };
  });
  const counted = tabs.filter((tab) => tab.cells !== null);
  return {
    indexed,
    cells: counted.length ? counted.reduce((sum, tab) => sum + (tab.cells as number), 0) : null,
    bytes: tabs.reduce((sum, tab) => sum + tab.bytes, 0),
    tabs,
  };
}

/** Throws SheetTooLargeToOpenError when the root's sheet is over the open limits. */
export function assertOpenable(root: SheetRootV2, limits = { maxCells: OPEN_MAX_CELLS, maxBytes: OPEN_MAX_BYTES }): void {
  const { indexed, ...size } = measureRoot(root);
  if (indexed || (size.cells ?? 0) > limits.maxCells || size.bytes > limits.maxBytes) {
    throw new SheetTooLargeToOpenError(size, root);
  }
}

export function isSheetRootV2(value: unknown): value is SheetRootV2 {
  return (
    !!value &&
    typeof value === "object" &&
    (value as { schema?: unknown }).schema === ROOT_SCHEMA &&
    (value as { v?: unknown }).v === ROOT_VERSION &&
    Array.isArray((value as { tabs?: unknown }).tabs)
  );
}

export function splitTab(sheet: WorkbookSheet): { meta: TabMeta; body: TabBody } {
  const { id, name, order, hidden, color, showGridLines, ...body } = sheet;
  return { meta: { id, name, order, hidden, color, showGridLines }, body };
}

function rootShell(doc: WorkbookDocument, servers: string[], tabs: RootTab[]): SheetRootV2 {
  return {
    schema: ROOT_SCHEMA,
    v: ROOT_VERSION,
    title: doc.title,
    activeSheetId: doc.activeSheetId,
    createdAt: doc.createdAt,
    updatedAt: doc.updatedAt,
    epoch: 1,
    servers,
    tabs,
  };
}

/** Root with every tab's payload embedded. */
export async function buildInlineRoot(doc: WorkbookDocument): Promise<SheetRootV2> {
  const tabs: RootTab[] = [];
  for (const sheet of doc.sheets) {
    const { meta, body } = splitTab(sheet);
    tabs.push({ ...meta, inline: bytesToBase64(await frame(body)) });
  }
  return rootShell(doc, [], tabs);
}

/** Root whose tabs point at Blossom chunks; `chunks` is parallel to `doc.sheets`. */
export function buildBlossomRoot(
  doc: WorkbookDocument,
  chunks: (ChunkEntry[] | { index: BlobRef })[],
  servers: string[],
): SheetRootV2 {
  const tabs = doc.sheets.map((sheet, index) => ({ ...splitTab(sheet).meta, cols: colsForSheet(sheet), chunks: chunks[index] }));
  return rootShell(doc, servers, tabs);
}

export function rootSize(root: SheetRootV2): number {
  return new TextEncoder().encode(JSON.stringify(root)).length;
}

export function storageOf(root: SheetRootV2): SheetStorage {
  return root.tabs.some((tab) => tab.blob || tab.chunks) ? "blossom" : "inline";
}

/** Serialises a root, refusing roots that would not fit in one encrypted event. */
export function serializeRoot(root: SheetRootV2): string {
  const json = JSON.stringify(root);
  const bytes = new TextEncoder().encode(json).length;
  if (bytes > ROOT_MAX_BYTES) throw new SheetTooLargeError(bytes);
  return json;
}

/** Frame type of a tab-index blob: `{ chunks: ChunkEntry[] }`. */
export const FRAME_TYPE_TAB_INDEX = 0x02;

/** Fetches and unframes one blob. `type` is the expected FSB2 frame type (default: a chunk/tab body). */
export type FetchTabBody = (ref: BlobRef, servers: string[], type?: number) => Promise<unknown>;

/** How many chunks of one tab are downloaded at once. */
const CHUNK_DOWNLOADS = 3;

/** The chunk list of a tab, reading its index blob when the list was moved out of the root. */
export async function resolveChunks(
  tab: RootTab,
  servers: string[],
  fetchTabBody: FetchTabBody,
): Promise<ChunkEntry[] | undefined> {
  if (!tab.chunks) return undefined;
  if (Array.isArray(tab.chunks)) return withRowStarts(tab.chunks);
  const index = (await fetchTabBody(tab.chunks.index, servers, FRAME_TYPE_TAB_INDEX)) as { chunks?: ChunkEntry[] };
  if (!Array.isArray(index.chunks)) throw new Error(`Worksheet "${tab.name}" has a damaged chunk index.`);
  return withRowStarts(index.chunks);
}

/** Fills in each chunk's first row from the row counts before it, for entries that do not store it. */
function withRowStarts(entries: ChunkEntry[]): ChunkEntry[] {
  if (entries.every((entry) => typeof entry.rowStart === "number")) return entries;
  let at = 0;
  return entries.map((entry) => {
    const filled = { ...entry, rowStart: at };
    at += entry.rowCount;
    return filled;
  });
}

async function mapLimited<T, R>(items: T[], limit: number, work: (item: T, index: number) => Promise<R>): Promise<R[]> {
  const out = new Array<R>(items.length);
  let next = 0;
  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const index = next++;
      out[index] = await work(items[index], index);
    }
  });
  await Promise.all(runners);
  return out;
}

export async function parseRootV2(
  root: SheetRootV2,
  fetchTabBody?: FetchTabBody,
  /** Called with each chunked tab's chunk list, so a later save can reuse the blobs. */
  onChunks?: (tabId: string, entries: ChunkEntry[]) => void,
): Promise<WorkbookDocument> {
  const sheets: WorkbookSheet[] = [];
  for (const tab of root.tabs) {
    let body: TabBody;
    if (tab.inline !== undefined) {
      body = (await unframe(base64ToBytes(tab.inline))) as TabBody;
    } else if (tab.chunks) {
      if (!fetchTabBody) throw new BlobUnavailableError("This sheet is stored on Blossom, but no server is available to load it from.");
      const entries = (await resolveChunks(tab, root.servers, fetchTabBody)) ?? [];
      const bodies = await mapLimited(entries, CHUNK_DOWNLOADS, (entry) =>
        fetchTabBody(entry.ref, root.servers) as Promise<Parameters<typeof assembleChunks>[1][number]>,
      );
      onChunks?.(tab.id, entries);
      body = assembleChunks(entries, bodies, tab.cols) as TabBody;
    } else if (tab.blob) {
      if (!fetchTabBody) throw new BlobUnavailableError("This sheet is stored on Blossom, but no server is available to load it from.");
      body = (await fetchTabBody(tab.blob, root.servers)) as TabBody;
    } else {
      throw new Error(`Worksheet "${tab.name}" has no data.`);
    }
    sheets.push({
      ...body,
      id: tab.id,
      name: tab.name,
      order: tab.order,
      hidden: tab.hidden,
      color: tab.color,
      showGridLines: tab.showGridLines,
    });
  }
  return {
    schema: WORKBOOK_SCHEMA,
    version: 1,
    title: root.title,
    activeSheetId: root.activeSheetId,
    createdAt: root.createdAt,
    updatedAt: root.updatedAt,
    sheets,
  };
}
