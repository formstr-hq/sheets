// Streaming counterpart of planTab (chunks.ts): rows arrive one at a time, chunks leave as soon as
// they are full, and memory stays bounded by `rawCap` no matter how big the sheet is.
//
// Same chunk format as planTab: a chunk is a run of rows of one tab, cell keys are relative to the
// chunk's first row, bodies are canonical JSON in an FSB2 frame. Rows and columns get legacy ids (a
// freshly imported sheet has no other ids), so a chunk written here has the bytes planTab would write
// for the same rows, and a later planTab re-save reuses it.
//
// The layout (widths, heights, merges, sizes) lives in the body of the first chunk but is only
// complete when the sheet ends, so the first chunk is the HEAD: a small run of rows held back and
// emitted LAST by finish(layout). A sheet that ends before any other chunk was written is encoded
// as one chunk (or split by the size limit), like planTab does.
import type { WorkbookCell } from "../workbook";
import { randomId } from "../utils";
import {
  encodeRows,
  fillBytes,
  fitRows,
  MIN_CHUNK_BYTES,
  type Encoded,
  type RowGroup,
  type TabIds,
} from "./chunks";

/** Raw (estimated JSON) size of the rows held back for the head chunk. */
export const HEAD_RAW_BYTES = 2 * 1024 * 1024;
/** Most raw (estimated JSON) bytes of rows kept in memory before a chunk is cut. */
export const RAW_CAP_BYTES = 24 * 1024 * 1024;

export interface RowInput {
  /** 0-based row index; rows must be pushed in ascending order. */
  row: number;
  /** [0-based column, cell]. Rows without cells are ignored. */
  cells: [col: number, cell: WorkbookCell][];
}

export interface StreamedChunk {
  id: string;
  /** Emission order within the tab (the head chunk is last). */
  seq: number;
  isHead: boolean;
  rowStart: number;
  rowCount: number;
  cells: number;
  /** FSB2 frame. */
  frame: Uint8Array;
  hash: string;
}

/** A StreamedChunk without its bytes, for manifests. */
export type StreamedChunkInfo = Omit<StreamedChunk, "frame"> & { size: number };

export interface ChunkStreamOptions {
  /** Largest chunk, compressed bytes (or a function read before every cut, for learned server limits). */
  max: number | (() => number);
  headRawBytes?: number;
  rawCap?: number;
  /** Receives each chunk; the writer waits for it (back-pressure). */
  onChunk(chunk: StreamedChunk): Promise<void>;
}

/** Cheap estimate of a row's JSON size; only used to choose cut points and bound memory. */
function estimateRow(input: RowInput): RowGroup {
  const rowDigits = String(input.row).length;
  let bytes = 0;
  const cells: RowGroup["cells"] = [];
  for (const [col, cell] of input.cells) {
    cells.push({ col, cell });
    let size = rowDigits + String(col).length + 8; // "r:c": {}, plus separators
    if (cell.value !== undefined) size += 9 + (typeof cell.value === "string" ? cell.value.length + 2 : 8);
    if (cell.display !== undefined) size += 13 + cell.display.length;
    if (cell.formula !== undefined) size += 13 + cell.formula.length;
    if (cell.type !== undefined) size += 12;
    if (cell.style !== undefined) size += 30 + 16 * Object.keys(cell.style).length;
    if (cell.link || cell.person || cell.tag || cell.validation) size += JSON.stringify(cell).length;
    bytes += size;
  }
  return { row: input.row, cells, bytes };
}

/** Imported rows carry legacy ids (their position), and columns too; only the end of the tab varies. */
const idsEnding = (end: number): TabIds => ({ rows: [], end });

const clampRatio = (ratio: number) => Math.min(40, Math.max(1.2, ratio));

export class ChunkStreamWriter {
  private readonly options: ChunkStreamOptions;
  private readonly headRaw: number;
  private readonly rawCap: number;
  private head: RowGroup[] = [];
  private headBytes = 0;
  private headFrozen = false;
  private pending: RowGroup[] = [];
  private pendingBytes = 0;
  private ratio = 4;
  private seq = 0;
  private lastRow = -1;
  /** First row after the head, once there is one. */
  private tailStart: number | null = null;
  private emittedTail = false;
  private infos: StreamedChunkInfo[] = [];
  private finished = false;

  constructor(options: ChunkStreamOptions) {
    this.options = options;
    this.headRaw = options.headRawBytes ?? HEAD_RAW_BYTES;
    this.rawCap = options.rawCap ?? RAW_CAP_BYTES;
  }

  /** Estimated raw bytes currently held in memory (head + pending rows). */
  get heldBytes(): number {
    return this.headBytes + this.pendingBytes;
  }

  private max(): number {
    const value = typeof this.options.max === "function" ? this.options.max() : this.options.max;
    return Math.max(MIN_CHUNK_BYTES, Math.floor(value));
  }

  private cutTarget(): number {
    return Math.min(this.rawCap, fillBytes(this.max()) * this.ratio);
  }

  async push(input: RowInput): Promise<void> {
    if (this.finished) throw new Error("ChunkStreamWriter: push after finish");
    if (input.row <= this.lastRow) throw new Error(`ChunkStreamWriter: row ${input.row} is not after row ${this.lastRow}`);
    if (input.cells.length === 0) return;
    const group = estimateRow(input);
    this.lastRow = input.row;
    if (!this.headFrozen) {
      if (this.head.length === 0 || this.headBytes + group.bytes <= this.headRaw) {
        this.head.push(group);
        this.headBytes += group.bytes;
        return;
      }
      await this.freezeHead();
    }
    if (this.pending.length > 0 && this.pendingBytes + group.bytes > this.cutTarget()) await this.flush(group.row);
    if (this.tailStart === null) this.tailStart = group.row;
    this.pending.push(group);
    this.pendingBytes += group.bytes;
  }

  /** Writes out everything and returns the manifest entries (no frame bytes) in rowStart order. */
  async finish(layout: Record<string, unknown>): Promise<StreamedChunkInfo[]> {
    if (this.finished) throw new Error("ChunkStreamWriter: finish called twice");
    this.finished = true;
    const end = Math.max(this.lastRow + 1, Number(layout.rowCount) || 0);
    if (this.emittedTail) {
      if (this.pending.length > 0) await this.flush(end);
      await this.emitHead(this.head, this.tailStart ?? end, layout);
    } else {
      // Nothing but the head was ever cut: encode head and remaining rows together.
      const rows = this.head.concat(this.pending);
      await this.emitHead(rows, end, layout);
    }
    this.head = [];
    this.pending = [];
    this.headBytes = this.pendingBytes = 0;
    return this.infos.sort((a, b) => a.rowStart - b.rowStart || Number(b.isHead) - Number(a.isHead));
  }

  /** Learns the compression ratio from the head rows, so the first real cut is already close. */
  private async freezeHead(): Promise<void> {
    this.headFrozen = true;
    const rows = this.head;
    const raw = rows.reduce((sum, group) => sum + group.bytes, 0);
    const probe = await encodeRows(rows, 0, rows.length, 0, undefined, idsEnding(rows[rows.length - 1].row + 1));
    if (probe.frame.length > 0) this.ratio = clampRatio(raw / probe.frame.length);
  }

  private async flush(nextRowStart: number): Promise<void> {
    const rows = this.pending;
    const raw = this.pendingBytes;
    this.pending = [];
    this.pendingBytes = 0;
    const pieces = await this.encode(rows, (i) => rows[i].row, nextRowStart);
    const compressed = pieces.reduce((sum, piece) => sum + piece.frame.length, 0);
    if (compressed > 0) this.ratio = clampRatio(raw / compressed);
    await this.emit(rows, pieces, (i) => rows[i].row, nextRowStart, false);
    this.emittedTail = true;
  }

  private async emitHead(rows: RowGroup[], end: number, layout: Record<string, unknown>): Promise<void> {
    if (rows.length === 0) {
      const empty = await encodeRows(rows, 0, 0, 0, layout, idsEnding(end));
      await this.emit(rows, [empty], () => 0, 0, true);
      return;
    }
    const startOf = (i: number) => (i === 0 ? 0 : rows[i].row);
    const pieces = await this.encode(rows, startOf, end, layout);
    await this.emit(rows, pieces, startOf, end, true);
  }

  /** Encodes rows, cutting them further when a frame is over the max. */
  private async encode(
    rows: RowGroup[],
    startOf: (i: number) => number,
    end: number,
    layout?: Record<string, unknown>,
  ): Promise<Encoded[]> {
    const prefix = new Float64Array(rows.length + 1);
    for (let k = 0; k < rows.length; k++) prefix[k + 1] = prefix[k] + rows[k].bytes;
    const encode = (i: number, j: number) => encodeRows(rows, i, j, startOf(i), i === 0 ? layout : undefined, idsEnding(end));
    const out: Encoded[] = [];
    await fitRows(rows, prefix, 0, rows.length, this.max(), encode, out);
    return out;
  }

  private async emit(
    rows: RowGroup[],
    pieces: Encoded[],
    startOf: (i: number) => number,
    end: number,
    isHead: boolean,
  ): Promise<void> {
    for (let k = 0; k < pieces.length; k++) {
      const piece = pieces[k];
      const rowStart = startOf(piece.i);
      const next = k + 1 < pieces.length ? startOf(pieces[k + 1].i) : end;
      let cells = 0;
      for (let r = piece.i; r < piece.j; r++) cells += rows[r].cells.length;
      const chunk: StreamedChunk = {
        id: randomId(8),
        seq: this.seq++,
        isHead: isHead && k === 0,
        rowStart,
        rowCount: Math.max(0, next - rowStart),
        cells,
        frame: piece.frame,
        hash: piece.hash,
      };
      const { frame, ...info } = chunk;
      this.infos.push({ ...info, size: frame.length });
      await this.options.onChunk(chunk);
    }
  }
}
