// Chunking: cutting one worksheet into several encrypted blobs, and putting it back together.
//
// A chunk is a contiguous run of rows of ONE worksheet, stored as exactly one Blossom blob.
// All sizes below are COMPRESSED bytes (the size of the FSB2 frame), because that is what is
// uploaded and what a server's upload limit applies to.
//
//   max chunk          8 MB, or the smallest upload limit among the sheet's servers if lower
//   fill when cutting  75% of that max (6 MB for the default 8 MB; 3.75 MB for a 5 MB server)
//   split              a chunk that grew above its max becomes two of about half
//   merge              two neighbours that together are under half the max become one
//   small sheet        a worksheet that fits one chunk stays exactly one chunk
//
// ROWS AND COLUMNS HAVE STABLE IDS (see workbook/ids.ts). A chunk body is
//   { rows: [[tag, start, length], ...], cells: { "<index in this chunk>:<colId>": cell }, ...layout }
// `rows` lists the ids of the chunk's rows in order, so a chunk's bytes depend on WHICH rows it
// holds, not on where it sits: inserting a row near the top changes only the chunk it lands in.
// (That needs the editor to report inserts and deletes; until it does, row ids follow positions and
// the chunks after an insert still change.)
//
// Chunks written before ids existed have no `rows` and use "row:col" keys relative to the chunk's
// rowStart. They are still read (their rows get legacy ids, see ids.ts) and become id chunks the
// next time they are rewritten. In memory a worksheet stays positional ("row:col" keys); ids sit
// beside it in `rowIds` / `colIds`, and this file translates at the storage boundary.
//
// Layout (widths, merges, frozen panes, borders, ... everything that is not a cell) is stored in
// the FIRST chunk's body next to its cells. A separate layout blob would add a second code path
// and one more upload per tab for something that is almost always small; the cost is that a
// layout edit re-uploads chunk 0 (one chunk, never the whole tab).

import type { WorkbookCell, WorkbookSheet } from "../workbook";
import {
  colIndexOf,
  colLookup,
  concatRuns,
  legacyColId,
  legacyRows,
  padCols,
  padRuns,
  sliceRuns,
  type RowRun,
} from "../workbook/ids";
import { randomId } from "../utils";
import { frame } from "./codec";
import { sha256Hex } from "./hash";
import type { BlobRef, ChunkEntry } from "./root";
import { splitTab } from "./root";

/** Largest chunk (one upload), compressed bytes. */
export const MAX_CHUNK_BYTES = 8 * 1024 * 1024;
/** Chunks are filled to this fraction of the max when a tab is cut. */
export const FILL_FRACTION = 0.75;
/** A server limit below this is not worth chunking around. */
export const MIN_CHUNK_BYTES = 64 * 1024;

export function fillBytes(max: number): number {
  return Math.floor(max * FILL_FRACTION);
}

/** A chunk already on Blossom, remembered between saves and loads. */
export interface ChunkMemory extends ChunkEntry {
  /** Servers known to hold the blob. */
  servers: string[];
}

export interface PlannedChunk {
  id: string;
  rowStart: number;
  rowCount: number;
  /** Compressed frame. Null once the chunk is known to be reusable and the bytes were dropped. */
  frame: Uint8Array | null;
  size: number;
  hash: string;
  /** Set when an identical chunk is already stored, so it need not be uploaded. */
  reuse?: ChunkMemory;
}

export interface RowGroup {
  row: number;
  cells: { col: number; cell: WorkbookCell }[];
  /** Rough uncompressed JSON size, used only to choose cut points. */
  bytes: number;
}

function groupRows(cells: Record<string, WorkbookCell>): RowGroup[] {
  const byRow = new Map<number, RowGroup>();
  for (const key of Object.keys(cells)) {
    const sep = key.indexOf(":");
    const row = Number(key.slice(0, sep));
    const col = Number(key.slice(sep + 1));
    if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) continue;
    let group = byRow.get(row);
    if (!group) {
      group = { row, cells: [], bytes: 0 };
      byRow.set(row, group);
    }
    group.cells.push({ col, cell: cells[key] });
    group.bytes += key.length + JSON.stringify(cells[key]).length + 4;
  }
  return [...byRow.values()].sort((a, b) => a.row - b.row);
}

function lowerBound(rows: RowGroup[], row: number): number {
  let lo = 0;
  let hi = rows.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (rows[mid].row < row) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

export interface Encoded {
  i: number;
  j: number;
  frame: Uint8Array;
  hash: string;
}

/** The ids a tab is encoded with. */
export interface TabIds {
  /** Row ids in order; may be shorter than the tab (the rest is padded with legacy ids). */
  rows: readonly RowRun[];
  /** Column ids in order; positions past the end use legacy ids. */
  cols?: readonly string[];
  /** One past the last row the chunks cover (the last chunk runs to here). */
  end: number;
}

/**
 * Encodes rows [i,j) as one chunk frame. The chunk covers rows [start, next) where `next` is the
 * first row of the following piece (rows[j]) or `ids.end`. Cells are keyed
 * "<row - start>:<colId>"; `layout` (everything but cells) is stored in the chunk only when given.
 */
export async function encodeRows(
  rows: RowGroup[],
  i: number,
  j: number,
  start: number,
  layout: Record<string, unknown> | undefined,
  ids: TabIds,
): Promise<Encoded> {
  const next = j < rows.length ? rows[j].row : ids.end;
  const part: Record<string, WorkbookCell> = {};
  for (let k = i; k < j; k++) {
    const group = rows[k];
    for (const { col, cell } of group.cells) part[`${group.row - start}:${ids.cols?.[col] ?? legacyColId(col)}`] = cell;
  }
  const runs = sliceRuns(padRuns(ids.rows, next), start, next);
  const bytes = await frame(layout ? { ...layout, rows: runs, cells: part } : { rows: runs, cells: part });
  return { i, j, frame: bytes, hash: await sha256Hex(bytes) };
}

/** A tab's columns as stored: its ids, padded with legacy ids to cover `count` columns. */
export function colsForSheet(sheet: Pick<WorkbookSheet, "colIds" | "colCount" | "cells">): string[] {
  let used = sheet.colCount ?? 0;
  for (const key of Object.keys(sheet.cells ?? {})) {
    const col = Number(key.slice(key.indexOf(":") + 1));
    if (col + 1 > used) used = col + 1;
  }
  return padCols(sheet.colIds, used);
}

/**
 * Encodes rows [i,j) with `encode`; if the frame is over `max`, cuts the rows at the byte midpoint
 * (by the `prefix` sums of RowGroup.bytes) and recurses. A single row over the max is kept whole.
 */
export async function fitRows(
  rows: RowGroup[],
  prefix: Float64Array,
  i: number,
  j: number,
  max: number,
  encode: (i: number, j: number) => Promise<Encoded>,
  out: Encoded[],
): Promise<void> {
  const encoded = await encode(i, j);
  if (encoded.frame.length <= max) {
    out.push(encoded);
    return;
  }
  if (j - i < 2) {
    out.push(encoded); // one row bigger than the max: nothing left to cut.
    return;
  }
  const half = prefix[i] + (prefix[j] - prefix[i]) / 2;
  let mid = lowerBound2(prefix, half, i + 1, j - 1);
  mid = Math.min(j - 1, Math.max(i + 1, mid));
  await fitRows(rows, prefix, i, mid, max, encode, out);
  await fitRows(rows, prefix, mid, j, max, encode, out);
}

/**
 * Cuts one worksheet into chunks.
 *
 * `previous` is how this tab was cut last time (from a save or a load). Using the same cut points
 * means an edit only changes the chunks it touches; the split and merge rules then adjust them.
 * Without it the tab is cut fresh, filling each chunk to 75% of `max`.
 *
 * `canReuse` says whether a previous chunk with identical content is still good to point at (for
 * example it is held by every server we now upload to). Those chunks keep no frame bytes.
 */
export async function planTab(
  sheet: WorkbookSheet,
  options: { max: number; previous?: ChunkMemory[]; canReuse?: (chunk: ChunkMemory) => boolean },
): Promise<PlannedChunk[]> {
  const max = Math.max(MIN_CHUNK_BYTES, Math.floor(options.max));
  const fill = fillBytes(max);
  const { body } = splitTab(sheet);
  const { cells, rowIds: _rowIds, colIds: _colIds, ...layout } = body as { cells: Record<string, WorkbookCell> } & Record<string, unknown>;
  void _rowIds;
  void _colIds;
  const rows = groupRows(cells ?? {});
  const n = rows.length;
  const lastEnd = Math.max(n === 0 ? 0 : rows[n - 1].row + 1, sheet.rowCount ?? 0);
  const ids: TabIds = { rows: sheet.rowIds ?? [], cols: colsForSheet(sheet), end: lastEnd };
  // prefix[k] = rough raw bytes of rows[0..k)
  const prefix = new Float64Array(n + 1);
  for (let k = 0; k < n; k++) prefix[k + 1] = prefix[k] + rows[k].bytes;
  const rawOf = (i: number, j: number) => prefix[j] - prefix[i];

  const rowStartOf = (i: number) => (i === 0 ? 0 : rows[i].row);

  const encode = (i: number, j: number) => encodeRows(rows, i, j, rowStartOf(i), i === 0 ? layout : undefined, ids);
  const fit = (i: number, j: number, out: Encoded[]) => fitRows(rows, prefix, i, j, max, encode, out);

  let encoded: Encoded[] = [];
  const previous = options.previous ?? [];
  if (n === 0) {
    encoded.push(await encode(0, 0));
  } else if (previous.length > 0) {
    const cuts = [0];
    for (const chunk of previous) {
      const at = lowerBound(rows, chunk.rowStart);
      if (at > cuts[cuts.length - 1] && at < n) cuts.push(at);
    }
    cuts.push(n);
    for (let c = 0; c + 1 < cuts.length; c++) await fit(cuts[c], cuts[c + 1], encoded);
  } else if (rawOf(0, n) / 3 <= max) {
    await fit(0, n, encoded);
  } else {
    // Fresh cut. Start by assuming raw JSON compresses about 4:1, then learn the real ratio.
    let ratio = 4;
    let i = 0;
    while (i < n) {
      const target = fill * ratio;
      let j = i + 1;
      while (j < n && rawOf(i, j + 1) <= target) j++;
      const before = encoded.length;
      await fit(i, j, encoded);
      const made = encoded.slice(before);
      const compressed = made.reduce((sum, item) => sum + item.frame.length, 0);
      if (compressed > 0) ratio = Math.min(40, Math.max(1.2, rawOf(i, j) / compressed));
      i = j;
    }
  }

  // Merge neighbours that together are under half the max.
  for (let k = 0; k + 1 < encoded.length; ) {
    const a = encoded[k];
    const b = encoded[k + 1];
    if (a.frame.length + b.frame.length < max / 2) {
      encoded.splice(k, 2, await encode(a.i, b.j));
    } else {
      k++;
    }
  }
  if (encoded.length === 0) encoded = [await encode(0, 0)];

  const byHash = new Map<string, ChunkMemory>();
  for (const chunk of options.previous ?? []) byHash.set(chunk.ref.h, chunk);

  return encoded.map((item, k) => {
    const rowStart = rowStartOf(item.i);
    const end = k + 1 < encoded.length ? rowStartOf(encoded[k + 1].i) : lastEnd;
    const memory = byHash.get(item.hash);
    const reusable = memory && (options.canReuse ? options.canReuse(memory) : true) ? memory : undefined;
    return {
      id: memory?.id ?? randomId(8),
      rowStart,
      rowCount: Math.max(0, end - rowStart),
      frame: reusable ? null : item.frame,
      size: item.frame.length,
      hash: item.hash,
      reuse: reusable,
    };
  });
}

/** First index in [lo, hi] whose prefix sum is >= target (prefix is non-decreasing). */
export function lowerBound2(prefix: Float64Array, target: number, lo: number, hi: number): number {
  let a = lo;
  let b = hi;
  while (a < b) {
    const mid = (a + b) >> 1;
    if (prefix[mid] < target) a = mid + 1;
    else b = mid;
  }
  return a;
}

/** The planned chunk as a manifest entry, once its blob exists. */
export function chunkEntry(chunk: PlannedChunk, ref: BlobRef): ChunkEntry {
  return { id: chunk.id, ref, rowStart: chunk.rowStart, rowCount: chunk.rowCount };
}

/** A chunk body as stored: id chunks have `rows`, older ones do not. */
export interface ChunkBody {
  rows?: RowRun[];
  cells?: Record<string, WorkbookCell>;
  [layout: string]: unknown;
}

/** A chunk body with positional cells, as the editor and big-sheet view use them. */
export interface PositionalChunk {
  /** Ids of the chunk's rows. Legacy ids for a chunk written before ids. */
  rows: RowRun[];
  /** Keys are "<row - chunk start>:<column index>". */
  cells: Record<string, WorkbookCell>;
  layout: Record<string, unknown>;
  /** True when the stored body carried ids. */
  hasIds: boolean;
}

/**
 * Reads one stored chunk body. `start` and `count` are the chunk's rows from the manifest; `cols` is
 * the tab's column ids. Cells whose column id is unknown (a deleted column) are dropped.
 */
export function readChunk(body: ChunkBody, start: number, count: number, cols?: readonly string[]): PositionalChunk {
  const { rows, cells = {}, ...layout } = body;
  if (!rows) return { rows: legacyRows(start, count), cells, layout, hasIds: false };
  const lookup = cols ? colLookup(cols) : undefined;
  const out: Record<string, WorkbookCell> = {};
  for (const [key, cell] of Object.entries(cells)) {
    const sep = key.indexOf(":");
    const id = key.slice(sep + 1);
    let col = colIndexOf(cols, id, lookup);
    // A legacy id ("c3") that the column list does not hold, but whose position is inside the list, belongs to a column
    // that was deleted: its cells are gone (it must not be read as whatever column moved into that position).
    if (cols && col >= 0 && col < cols.length && cols[col] !== id) col = -1;
    if (col >= 0) out[`${key.slice(0, sep)}:${col}`] = cell;
  }
  return { rows, cells: out, layout, hasIds: true };
}

/** Turns a positional chunk back into a stored body (the inverse of readChunk). */
export function writeChunk(chunk: PositionalChunk, cols?: readonly string[]): ChunkBody {
  const cells: Record<string, WorkbookCell> = {};
  for (const [key, cell] of Object.entries(chunk.cells)) {
    const sep = key.indexOf(":");
    const col = Number(key.slice(sep + 1));
    cells[`${key.slice(0, sep)}:${cols?.[col] ?? legacyColId(col)}`] = cell;
  }
  return { ...chunk.layout, rows: chunk.rows, cells };
}

/** Rebuilds a worksheet body from its chunk bodies (in manifest order). */
export function assembleChunks(
  entries: { rowStart: number; rowCount?: number }[],
  bodies: ChunkBody[],
  cols?: readonly string[],
): Omit<WorkbookSheet, "id" | "name" | "order" | "hidden" | "color" | "showGridLines"> {
  const cells: Record<string, WorkbookCell> = {};
  let layout: Record<string, unknown> = {};
  const runs: RowRun[][] = [];
  let anyIds = false;
  bodies.forEach((body, k) => {
    const entry = entries[k];
    const chunk = readChunk(body, entry.rowStart, entry.rowCount ?? 0, cols);
    if (k === 0) layout = chunk.layout;
    anyIds ||= chunk.hasIds;
    runs.push(chunk.rows);
    for (const [key, cell] of Object.entries(chunk.cells)) {
      const sep = key.indexOf(":");
      cells[`${Number(key.slice(0, sep)) + entry.rowStart}:${key.slice(sep + 1)}`] = cell;
    }
  });
  const sheet = { ...layout, cells } as ReturnType<typeof assembleChunks>;
  if (anyIds) sheet.rowIds = concatRuns(...runs);
  if (cols) sheet.colIds = [...cols];
  return sheet;
}

