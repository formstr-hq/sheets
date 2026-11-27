// A big sheet, read (and optionally edited) a chunk at a time, never whole.
//
// The root lists every tab's chunks with the rows they hold (rowStart/rowCount), so finding the
// chunk for a row is a binary search and showing rows 5,000,000-5,000,040 downloads one chunk.
// Decoded chunks sit in a small LRU cache; edited ("dirty") chunks are pinned until saved.
//
// Saving re-encrypts only the dirty chunks (plus a tab's index blob when its chunk list lives in
// one) and uploads them under ONE signed authorization, then returns the updated root. Formulas
// are not recalculated: an edit stores the typed value.
import type { WorkbookCell } from "../workbook";
import { encryptFrame, uploadStaged, type BlossomEnv, type StagedBlob, type UploadOptions } from "./blossom";
import { concatRuns, legacyRows, padCols, rowCountOf, type RowRun } from "../workbook/ids";
import { readChunk, writeChunk, type ChunkBody as StoredChunkBody } from "./chunks";
import { frame } from "./codec";
import { downloadTabBody } from "./blossom";
import { FRAME_TYPE_TAB_INDEX, resolveChunks, type ChunkEntry, type SheetRootV2 } from "./root";

/** A decoded chunk: cells are positional ("<row - chunk start>:<column index>"), rows are its row ids. */
interface ChunkBody {
  rows?: RowRun[];
  cells?: Record<string, WorkbookCell>;
  rowCount?: number;
  colCount?: number;
  [layout: string]: unknown;
}

/** Decoded chunks kept in memory (each is ~6 MB compressed, a few tens of MB as objects). */
const CACHE_CHUNKS = 6;
/** Largest selection one toolbar action may change (it may have to download several chunks). */
export const MAX_EDIT_CELLS = 200_000;
const UNDO_STEPS = 100;

interface Change {
  tab: number;
  row: number;
  col: number;
  before?: WorkbookCell;
  after?: WorkbookCell;
}

export interface BigTab {
  name: string;
  chunks: ChunkEntry[];
  /** Chunk list stored in an index blob (must be re-uploaded when a chunk changes). */
  indexed: boolean;
  rows: number;
  cols: number;
  /** Column ids in order (legacy ids for columns that predate ids). Written back on save. */
  colIds: string[];
}

export class BigSheetStore {
  readonly root: SheetRootV2;
  tabs: BigTab[] = [];
  private readonly servers: string[];
  private readonly env: BlossomEnv;
  private cache = new Map<string, ChunkBody>();
  private loading = new Map<string, Promise<ChunkBody>>();
  private dirty = new Set<string>();
  private undoStack: Change[][] = [];
  private redoStack: Change[][] = [];
  private listeners = new Set<() => void>();

  constructor(root: SheetRootV2, servers: string[], env: BlossomEnv = {}) {
    this.root = root;
    this.servers = [...root.servers, ...servers];
    this.env = env;
  }

  private fetchBody = (ref: ChunkEntry["ref"], servers: string[], type?: number) =>
    downloadTabBody(ref, [...servers, ...this.servers], this.env, type);

  /** Reads every tab's chunk list and each tab's first chunk (it carries the size and layout). */
  async open(): Promise<void> {
    this.tabs = [];
    for (const tab of this.root.tabs) {
      const chunks = (await resolveChunks(tab, this.root.servers, this.fetchBody)) ?? [];
      const last = chunks[chunks.length - 1];
      this.tabs.push({
        name: tab.name,
        chunks,
        indexed: !!tab.chunks && !Array.isArray(tab.chunks),
        rows: last ? last.rowStart + last.rowCount : 0,
        cols: 26,
        colIds: [],
      });
    }
    await Promise.all(
      this.tabs.map(async (tab, index) => {
        if (tab.chunks.length === 0) return;
        const head = await this.load(index, 0);
        tab.rows = Math.max(tab.rows, head.rowCount ?? 0, 100);
        tab.cols = Math.max(tab.cols, head.colCount ?? 0, this.root.tabs[index].cols?.length ?? 0);
        tab.colIds = padCols(this.root.tabs[index].cols, tab.cols);
      }),
    );
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    for (const listener of this.listeners) listener();
  }

  /** Index of the chunk holding `row` (the last chunk takes rows past the end). */
  chunkFor(tab: number, row: number): number {
    const chunks = this.tabs[tab].chunks;
    let lo = 0;
    let hi = chunks.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (chunks[mid].rowStart <= row) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  private key = (tab: number, chunk: number) => `${tab}:${chunk}`;

  /** Downloads and decodes a chunk (once), evicting the least recently used clean chunks. */
  load(tab: number, chunk: number): Promise<ChunkBody> {
    const key = this.key(tab, chunk);
    const hit = this.cache.get(key);
    if (hit) {
      this.cache.delete(key);
      this.cache.set(key, hit);
      return Promise.resolve(hit);
    }
    let pending = this.loading.get(key);
    if (!pending) {
      const entry = this.tabs[tab].chunks[chunk];
      pending = (this.fetchBody(entry.ref, this.root.servers) as Promise<StoredChunkBody>).then((stored) => {
        const read = readChunk(stored, entry.rowStart, entry.rowCount, this.root.tabs[tab].cols);
        const body: ChunkBody = { ...read.layout, rows: read.rows, cells: read.cells };
        this.cache.set(key, body);
        this.loading.delete(key);
        for (const old of this.cache.keys()) {
          if (this.cache.size <= CACHE_CHUNKS) break;
          if (!this.dirty.has(old) && old !== key) this.cache.delete(old);
        }
        this.emit();
        return body;
      });
      pending.catch(() => this.loading.delete(key));
      this.loading.set(key, pending);
    }
    return pending;
  }

  /** The cell if its chunk is in memory; `undefined` when not loaded yet, `null` when empty. */
  cell(tab: number, row: number, col: number): WorkbookCell | null | undefined {
    const chunk = this.chunkFor(tab, row);
    const body = this.cache.get(this.key(tab, chunk));
    if (!body) return undefined;
    const start = this.tabs[tab].chunks[chunk].rowStart;
    return body.cells?.[`${row - start}:${col}`] ?? null;
  }

  isLoaded(tab: number, row: number): boolean {
    return this.cache.has(this.key(tab, this.chunkFor(tab, row)));
  }

  get dirtyCount(): number {
    return this.dirty.size;
  }

  /** Parses what the user typed: a number when it reads as one, else text; empty clears the value. */
  static parseInput(text: string, previous?: WorkbookCell): WorkbookCell | undefined {
    const trimmed = text.trim();
    const style = previous?.style;
    if (trimmed === "") return style ? { style } : undefined;
    const isNumber = /^[-+]?[\d,]*\.?\d+(e[-+]?\d+)?$/i.test(trimmed);
    const value = isNumber ? Number(trimmed.replace(/,/g, "")) : text;
    return { value, ...(style ? { style } : {}) };
  }

  async setCell(tab: number, row: number, col: number, text: string): Promise<void> {
    await this.edit(tab, row, col, row, col, (cell) => BigSheetStore.parseInput(text, cell));
  }

  /**
   * Changes every cell in rows r1..r2, columns c1..c2 (inclusive) with `change`, which gets a copy
   * of the cell (or undefined) and returns the new cell (undefined clears it). One undo step.
   */
  async edit(
    tab: number,
    r1: number,
    c1: number,
    r2: number,
    c2: number,
    change: (cell: WorkbookCell | undefined, row: number, col: number) => WorkbookCell | undefined,
  ): Promise<void> {
    const count = (r2 - r1 + 1) * (c2 - c1 + 1);
    if (count > MAX_EDIT_CELLS) {
      throw new Error(`Select at most ${MAX_EDIT_CELLS.toLocaleString()} cells at once in large sheet mode.`);
    }
    const from = this.chunkFor(tab, r1);
    const to = this.chunkFor(tab, r2);
    for (let c = from; c <= to; c++) await this.load(tab, c);
    const changes: Change[] = [];
    for (let row = r1; row <= r2; row++) {
      for (let col = c1; col <= c2; col++) {
        const before = this.rawCell(tab, row, col);
        const after = change(before ? structuredClone(before) : undefined, row, col);
        if (JSON.stringify(before) === JSON.stringify(after)) continue;
        changes.push({ tab, row, col, before, after });
      }
    }
    if (changes.length === 0) return;
    this.apply(changes, "after");
    this.undoStack.push(changes);
    if (this.undoStack.length > UNDO_STEPS) this.undoStack.shift();
    this.redoStack = [];
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  undo(): void {
    const step = this.undoStack.pop();
    if (!step) return;
    this.apply(step, "before");
    this.redoStack.push(step);
  }

  redo(): void {
    const step = this.redoStack.pop();
    if (!step) return;
    this.apply(step, "after");
    this.undoStack.push(step);
  }

  private rawCell(tab: number, row: number, col: number): WorkbookCell | undefined {
    const chunk = this.chunkFor(tab, row);
    const body = this.cache.get(this.key(tab, chunk));
    return body?.cells?.[`${row - this.tabs[tab].chunks[chunk].rowStart}:${col}`];
  }

  /** Writes changes into their (pinned) chunks and marks them dirty. */
  private apply(changes: Change[], side: "before" | "after"): void {
    for (const item of changes) {
      const chunk = this.chunkFor(item.tab, item.row);
      const key = this.key(item.tab, chunk);
      const body = this.cache.get(key);
      if (!body) continue; // edited chunks are pinned, so this does not happen
      const cells = (body.cells ??= {});
      const cellKey = `${item.row - this.tabs[item.tab].chunks[chunk].rowStart}:${item.col}`;
      const value = item[side];
      if (value) cells[cellKey] = structuredClone(value);
      else delete cells[cellKey];
      const tabInfo = this.tabs[item.tab];
      tabInfo.rows = Math.max(tabInfo.rows, item.row + 1);
      tabInfo.cols = Math.max(tabInfo.cols, item.col + 1);
      if (tabInfo.colIds.length < tabInfo.cols) tabInfo.colIds = padCols(tabInfo.colIds, tabInfo.cols);
      this.dirty.add(key);
    }
    this.emit();
  }

  /**
   * Uploads the edited chunks under one signature and returns the root that points at them.
   * The caller signs and publishes the root event.
   */
  async save(options: UploadOptions): Promise<SheetRootV2> {
    if (this.dirty.size === 0) return this.root;
    const keys = [...this.dirty];
    const staged: StagedBlob[] = [];
    const targets: { tab: number; chunk: number }[] = [];
    for (const key of keys) {
      const [tab, chunk] = key.split(":").map(Number);
      const body = this.cache.get(key) as ChunkBody;
      const entry = this.tabs[tab].chunks[chunk];
      const next = this.tabs[tab].chunks[chunk + 1];
      if (!next) entry.rowCount = Math.max(entry.rowCount, this.tabs[tab].rows - entry.rowStart);
      if (chunk === 0) {
        body.rowCount = Math.max(body.rowCount ?? 0, this.tabs[tab].rows);
        body.colCount = Math.max(body.colCount ?? 0, this.tabs[tab].cols);
      }
      const { rows, cells, ...layout } = body;
      const have = rowCountOf(rows ?? []);
      const ids = concatRuns(rows ?? [], legacyRows(entry.rowStart + have, entry.rowCount - have));
      const stored = writeChunk({ rows: ids, cells: cells ?? {}, layout, hasIds: true }, this.tabs[tab].colIds);
      staged.push(await encryptFrame(await frame(stored)));
      targets.push({ tab, chunk });
    }
    // New refs are known before upload, so a tab's index blob joins the same batch (same signature).
    const before = this.tabs.map((tab) => [...tab.chunks]);
    targets.forEach(({ tab, chunk }, i) => {
      const entry = this.tabs[tab].chunks[chunk];
      const body = this.cache.get(this.key(tab, chunk)) as ChunkBody;
      this.tabs[tab].chunks[chunk] = { ...entry, ref: staged[i].ref, cells: Object.keys(body.cells ?? {}).length };
    });
    const touched = [...new Set(targets.map((t) => t.tab))];
    const indexOf = new Map<number, number>();
    for (const tab of touched) {
      if (!this.tabs[tab].indexed) continue;
      indexOf.set(tab, staged.length);
      staged.push(await encryptFrame(await frame({ chunks: this.tabs[tab].chunks }, FRAME_TYPE_TAB_INDEX)));
    }
    let uploaded;
    try {
      uploaded = await uploadStaged(staged, options);
    } catch (error) {
      before.forEach((chunks, tab) => (this.tabs[tab].chunks = chunks)); // edits stay dirty; try again
      throw error;
    }
    const holders = new Set(this.root.servers);
    uploaded.forEach((blob) => blob.servers.forEach((server) => holders.add(server)));
    const root: SheetRootV2 = {
      ...this.root,
      updatedAt: Date.now(),
      servers: [...holders],
      tabs: this.root.tabs.map((tab, index) => {
        if (!touched.includes(index)) return tab;
        const at = indexOf.get(index);
        return { ...tab, cols: [...this.tabs[index].colIds], chunks: at === undefined ? [...this.tabs[index].chunks] : { index: uploaded[at].ref } };
      }),
    };
    Object.assign(this.root, root);
    this.dirty.clear();
    this.emit();
    return root;
  }
}
