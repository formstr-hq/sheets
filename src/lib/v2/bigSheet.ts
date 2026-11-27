// A big sheet, read (and optionally edited) a chunk at a time, never whole.
//
// The root lists every tab's chunks with the rows they hold (rowStart/rowCount), so finding the
// chunk for a row is a binary search and showing rows 5,000,000-5,000,040 downloads one chunk.
// Decoded chunks sit in a small LRU cache; edited ("dirty") chunks are pinned until saved.
//
// Saving re-encrypts only the dirty chunks (plus a tab's index blob when its chunk list lives in
// one) and uploads them under ONE signed authorization, then returns the updated root. Formulas
// are not recalculated: an edit stores the typed value.
import { WORKBOOK_SCHEMA, type WorkbookCell, type WorkbookConditionalRule, type WorkbookDocument, type WorkbookSheet } from "../workbook";
import { encryptFrame, uploadStaged, type BlossomEnv, type StagedBlob, type UploadOptions } from "./blossom";
import { concatRuns, deleteRuns, insertRuns, legacyRows, mintRows, newDeviceTag, padCols, padRuns, rowCountOf, type RowRun } from "../workbook/ids";
import { readChunk, writeChunk, type ChunkBody as StoredChunkBody } from "./chunks";
import { frame } from "./codec";
import { downloadTabBody } from "./blossom";
import { editLayout, type StructEdit } from "./layoutEdit";
import type { ColumnFilter } from "./bigFilter";
import { ArrayAxis, RunAxis, dropSheetRefs, formulaFromIds, formulaToIds, hasIdRefs, tombstonesFor, type RefContext, type Tombstones } from "./idRefs";
import { FRAME_TYPE_TAB_INDEX, ROOT_SCHEMA, ROOT_VERSION, resolveChunks, type ChunkEntry, type RootTab, type SheetRootV2 } from "./root";
import { randomId } from "../utils";

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
const MAX_COLUMNS = 16_384;
const MAX_INSERT_ROWS = 100_000;
const UNDO_STEPS = 100;
/** A tab's chunk list bigger than this (as JSON) is stored in an index blob instead of in the root. */
const INDEX_AT_BYTES = 16 * 1024;

export interface CellAddress {
  tab: number;
  row: number;
  col: number;
}

/** An undo step: the cells a change touched, or a function that puts back a whole earlier state (insert/delete, tabs). */
type UndoStep = Change[] | { restore: () => void };

/** Structural steps keep a copy of the state before them, so only the latest few can be undone. */
const MAX_STRUCTURAL_UNDO = 5;

interface Snapshot {
  tabs: { tab: BigTab; name: string; chunks: ChunkEntry[]; indexed: boolean; rows: number; cols: number; colIds: string[]; layout: TabLayout }[];
  tabRoots: RootTab[];
  bodies: { key: string; body: ChunkBody; cells: Record<string, WorkbookCell>; rows?: RowRun[]; layout: Record<string, unknown> }[];
  dirty: Set<string>;
  metaDirty: boolean;
}

interface Change {
  tab: number;
  row: number;
  col: number;
  before?: WorkbookCell;
  after?: WorkbookCell;
}

/** What a tab remembers about its look besides cell contents (stored with the tab's first chunk). */
export interface TabLayout {
  columnWidths?: Record<number, number>;
  rowHeights?: Record<number, number>;
  hiddenRows?: number[];
  hiddenColumns?: number[];
  frozen?: { rows: number; columns: number };
  /** 1 = 100%. */
  zoom?: number;
  conditionalFormats?: WorkbookConditionalRule[];
  /** A filter by the values of one column: its `hidden` rows are hidden on top of rows hidden by hand. */
  filter?: ColumnFilter;
  merges?: { r: number; c: number; rs: number; cs: number }[];
  /** Rows and columns that were deleted (the neighbours they had), so formulas stored by id can still resolve. */
  tombs?: { rows: Tombstones; cols: Tombstones };
  /** Pictures floating over the cells, each anchored to a cell. */
  images?: { id: string; src: string; row: number; col: number; dx: number; dy: number; width: number; height: number }[];
  borders?: { r1: number; c1: number; r2: number; c2: number; type: "all" | "outside" | "inside" | "top" | "bottom" | "left" | "right" | "none"; color?: string; style?: string }[];
}

const LAYOUT_KEYS = ["columnWidths", "rowHeights", "hiddenRows", "hiddenColumns", "frozen", "zoom", "conditionalFormats", "filter", "merges", "borders", "images", "tombs"] as const;

export interface BigTab {
  layout: TabLayout;
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
  /** The root's tab entries as edited here (names, colours, order); they replace the root's on save. */
  private tabRoots: RootTab[] = [];
  private metaDirty = false;
  private readonly servers: string[];
  private readonly env: BlossomEnv;
  private cache = new Map<string, ChunkBody>();
  private loading = new Map<string, Promise<ChunkBody>>();
  private dirty = new Set<string>();
  /** Chunks a structural edit has loaded and is still working on: the cache must not evict them (they are not dirty yet). */
  private held = new Set<string>();
  private undoStack: UndoStep[] = [];
  /** Was the last undo() a structural one (rows, columns, tabs)? Then the caller has to rebuild what it shows. */
  structuralUndo = false;
  private redoStack: Change[][] = [];
  private listeners = new Set<() => void>();
  /** Called each time a chunk finishes loading (the formula index learns its formulas here). */
  onChunkLoaded?: (tab: number, chunk: number) => void;
  /** Ids for rows and columns this device adds. */
  private readonly idTag = newDeviceTag();
  private idCounter = 0;
  /** Bumped by every insert or delete of rows or columns; tells loaded chunks whose formulas still use the old positions. */
  private epoch = 0;
  private bodyEpoch = new Map<string, number>();
  /** Tabs whose chunk list must be written on the next save even though none of their chunks changed. */
  private forceTabs = new Set<number>();
  /** The workbook's "formulas by id" flag as this store has it. It reaches the root only when saved: the root object may be shared with other stores. */
  private byId: boolean;
  /** True while `open` reads the first chunks: they are not converted yet (see open). */
  private opening = false;

  constructor(root: SheetRootV2, servers: string[], env: BlossomEnv = {}) {
    this.root = root;
    this.byId = root.refs === "id";
    this.servers = [...root.servers, ...servers];
    this.env = env;
  }

  /** Chunk bodies of a store that lives in memory (a small sheet kept inline); then nothing is downloaded. */
  private memory?: Map<string, StoredChunkBody>;

  private fetchBody = (ref: ChunkEntry["ref"], servers: string[], type?: number) =>
    this.memory ? Promise.resolve(this.memory.get(ref.x) as StoredChunkBody) : downloadTabBody(ref, [...servers, ...this.servers], this.env, type);

  /** True when this store holds a whole small sheet in memory (saved inline as a document, not as chunks). */
  get inMemory(): boolean {
    return !!this.memory;
  }

  /**
   * A store over a sheet that is held whole in memory: one chunk per tab, nothing downloaded, saved by
   * turning it back into a document (`toDocument`). It lets small, inline sheets use the big-sheet grid.
   */
  static fromDocument(document: WorkbookDocument): BigSheetStore {
    const memory = new Map<string, StoredChunkBody>();
    const tabs: RootTab[] = [];
    const entries: ChunkEntry[][] = [];
    document.sheets.forEach((sheet, index) => {
      const rowCount = Math.max(sheet.rowCount, 1);
      const colCount = Math.max(sheet.colCount, 1);
      const colIds = padCols(sheet.colIds, colCount);
      const layout: Record<string, unknown> = { rowCount, colCount };
      for (const key of LAYOUT_KEYS) {
        const value = (sheet as unknown as Record<string, unknown>)[key];
        if (value !== undefined) layout[key] = structuredClone(value);
      }
      const rows = sheet.rowIds && rowCountOf(sheet.rowIds) === rowCount ? sheet.rowIds : legacyRows(0, rowCount);
      const ref = { x: `mem-${index}`, k: "", n: 0, p: 0, c: 0, h: "" };
      memory.set(ref.x, writeChunk({ rows, cells: structuredClone(sheet.cells), layout, hasIds: true }, colIds));
      const { id, name, order, hidden, color, showGridLines } = sheet;
      tabs.push({ id, name, order, hidden, color, showGridLines, cols: colIds });
      entries.push([{ id: `mem-${index}`, ref, rowStart: 0, rowCount }]);
    });
    const root: SheetRootV2 = {
      schema: ROOT_SCHEMA,
      v: ROOT_VERSION,
      title: document.title,
      activeSheetId: document.activeSheetId,
      createdAt: document.createdAt,
      updatedAt: document.updatedAt,
      epoch: 0,
      servers: [],
      tabs: tabs.map((tab, index) => ({ ...tab, chunks: entries[index] })),
    };
    const store = new BigSheetStore(root, []);
    store.memory = memory;
    return store;
  }

  /** The whole workbook as a document (every tab, every loaded or edited chunk). For stores made by `fromDocument`. */
  async toDocument(): Promise<WorkbookDocument> {
    const sheets: WorkbookSheet[] = [];
    for (let t = 0; t < this.tabs.length; t++) {
      const tab = this.tabs[t];
      const meta = this.tabRoots[t];
      const cells: Record<string, WorkbookCell> = {};
      const runs: RowRun[][] = [];
      for (let c = 0; c < tab.chunks.length; c++) {
        const body = await this.load(t, c);
        const start = tab.chunks[c].rowStart;
        runs.push(body.rows ?? legacyRows(start, tab.chunks[c].rowCount));
        for (const [key, cell] of Object.entries(body.cells ?? {})) {
          const at = key.indexOf(":");
          cells[`${Number(key.slice(0, at)) + start}:${key.slice(at + 1)}`] = cell;
        }
      }
      const sheet: WorkbookSheet = {
        id: meta.id,
        name: meta.name,
        order: meta.order,
        rowCount: tab.rows,
        colCount: tab.cols,
        cells,
        rowIds: concatRuns(...runs),
        colIds: [...tab.colIds],
      };
      if (meta.hidden) sheet.hidden = true;
      if (meta.color) sheet.color = meta.color;
      if (meta.showGridLines !== undefined) sheet.showGridLines = meta.showGridLines;
      for (const key of LAYOUT_KEYS) if (tab.layout[key] !== undefined) (sheet as unknown as Record<string, unknown>)[key] = structuredClone(tab.layout[key]);
      // A document always holds A1 formulas and has no use for the record of deleted rows (see parseRootV2).
      delete (sheet as { tombs?: unknown }).tombs;
      sheets.push(sheet);
    }
    return {
      schema: WORKBOOK_SCHEMA,
      version: 1,
      title: this.root.title,
      activeSheetId: this.root.activeSheetId,
      createdAt: this.root.createdAt,
      updatedAt: Date.now(),
      sheets,
    };
  }

  /** Reads every tab's chunk list and each tab's first chunk (it carries the size and layout). */
  async open(): Promise<void> {
    this.tabs = [];
    this.tabRoots = this.root.tabs.map((tab) => ({ ...tab }));
    for (const tab of this.root.tabs) {
      // Our own copies: the store changes chunk entries as rows move, and the root must stay as it was saved.
      const chunks = ((await resolveChunks(tab, this.root.servers, this.fetchBody)) ?? []).map((entry) => ({ ...entry }));
      const last = chunks[chunks.length - 1];
      this.tabs.push({
        name: tab.name,
        chunks,
        indexed: !!tab.chunks && !Array.isArray(tab.chunks),
        rows: last ? last.rowStart + last.rowCount : 0,
        cols: 26,
        colIds: [],
        layout: {},
      });
    }
    // The first chunks are read before anything is known about the columns and deleted rows of the tabs
    // (they sit in these very chunks), so their formulas are converted afterwards.
    this.opening = true;
    try {
      await Promise.all(
        this.tabs.map(async (tab, index) => {
          if (tab.chunks.length === 0) return;
          const head = await this.load(index, 0);
          tab.rows = Math.max(tab.rows, head.rowCount ?? 0, 100);
          tab.cols = Math.max(tab.cols, head.colCount ?? 0, this.tabRoots[index].cols?.length ?? 0);
          tab.colIds = padCols(this.tabRoots[index].cols, tab.cols);
          for (const key of LAYOUT_KEYS) if (head[key] !== undefined) (tab.layout as Record<string, unknown>)[key] = head[key];
        }),
      );
    } finally {
      this.opening = false;
    }
    if (this.refsById) {
      this.tabs.forEach((tab, index) => {
        const head = this.cache.get(this.key(index, 0));
        if (head && tab.chunks.length > 0) head.cells = this.fromStoredCells(index, head.cells ?? {});
      });
    }
  }

  // --- formulas stored by id (spec 4.1) ------------------------------------------------------
  //
  // In memory a formula is always A1 text, so the grid and the calculation engine do not change. When the
  // workbook stores formulas by id (root.refs), they are converted when a chunk is read and when it is
  // saved. An insert or delete then needs no rewrite of the chunks that are not open.

  /** Set to false to keep a workbook on positional formulas (rewritten in every chunk on each insert or delete). */
  allowRefsById = true;

  /** Does this workbook store formulas by id? */
  get refsById(): boolean {
    return this.byId;
  }

  /** The row ids of a tab in order, from the chunk list (a chunk without recorded ids has the legacy ones). */
  private runsOf(tab: number): RowRun[] {
    return concatRuns(...this.tabs[tab].chunks.map((entry) => entry.runs ?? legacyRows(entry.rowStart, entry.rowCount)));
  }

  /** What is needed to convert formulas between A1 and ids, as of now. `current` is the tab the formula lives on. */
  refContext(current: number): RefContext {
    return {
      current: this.tabRoots[current].id,
      tabs: this.tabs.map((info, i) => ({
        id: this.tabRoots[i].id,
        name: info.name,
        rows: new RunAxis(this.runsOf(i), info.layout.tombs?.rows),
        cols: new ArrayAxis(padCols(info.colIds, info.cols), info.layout.tombs?.cols),
      })),
    };
  }

  private fromStoredCells(tab: number, cells: Record<string, WorkbookCell>): Record<string, WorkbookCell> {
    let ctx: RefContext | undefined;
    let out = cells;
    for (const [key, cell] of Object.entries(cells)) {
      if (!cell.formula || !hasIdRefs(cell.formula)) continue;
      ctx ??= this.refContext(tab);
      if (out === cells) out = { ...cells };
      out[key] = { ...cell, formula: formulaFromIds(cell.formula, ctx) };
    }
    return out;
  }

  private toStoredCells(tab: number, cells: Record<string, WorkbookCell>): Record<string, WorkbookCell> {
    let ctx: RefContext | undefined;
    let out = cells;
    for (const [key, cell] of Object.entries(cells)) {
      if (!cell.formula) continue;
      ctx ??= this.refContext(tab);
      if (out === cells) out = { ...cells };
      out[key] = { ...cell, formula: formulaToIds(cell.formula, ctx) };
    }
    return out;
  }

  /**
   * Switches the workbook to formulas stored by id. It needs every chunk's row ids (so every chunk is read
   * once) and the chunks that hold formulas are rewritten on the next save. Returns false, changing nothing,
   * when more than `maxFormulaChunks` chunks hold formulas (they would all have to stay in memory until saved).
   */
  async enableRefsById(maxFormulaChunks: number, onProgress?: (done: number, total: number) => void): Promise<boolean> {
    if (this.refsById) return true;
    const total = this.tabs.reduce((sum, tab) => sum + tab.chunks.length, 0);
    const runs: { tab: number; chunk: number; runs: RowRun[] }[] = [];
    const withFormulas = new Map<string, ChunkBody>();
    let done = 0;
    for (let t = 0; t < this.tabs.length; t++) {
      for (let c = 0; c < this.tabs[t].chunks.length; c++) {
        const entry = this.tabs[t].chunks[c];
        const body = await this.load(t, c);
        runs.push({ tab: t, chunk: c, runs: entry.rowCount > 0 ? padRuns(body.rows ?? [], entry.rowCount) : [] });
        if (Object.values(body.cells ?? {}).some((cell) => cell.formula)) {
          withFormulas.set(this.key(t, c), body);
          if (withFormulas.size > maxFormulaChunks) return false;
        }
        onProgress?.(++done, total);
      }
    }
    // Ids must be unique within a tab, or a row id would name two rows.
    for (let t = 0; t < this.tabs.length; t++) {
      if (hasDuplicateIds(concatRuns(...runs.filter((item) => item.tab === t).map((item) => item.runs)))) return false;
    }
    for (const item of runs) {
      const entry = this.tabs[item.tab].chunks[item.chunk];
      this.tabs[item.tab].chunks[item.chunk] = { ...entry, runs: item.runs };
    }
    this.byId = true;
    // The chunks with formulas must be written again (with ids); they stay in memory until then.
    for (const [key, body] of withFormulas) {
      this.cache.set(key, body);
      this.bodyEpoch.set(key, this.epoch);
      this.dirty.add(key);
    }
    for (let t = 0; t < this.tabs.length; t++) this.forceTabs.add(t);
    this.metaDirty = true;
    this.emit();
    return true;
  }

  /**
   * After an insert or delete: rewrites the formulas of every chunk that was open before it (their text still
   * names the old positions) with `rewrite`. Chunks read after it were converted with the new positions.
   * Does not mark anything as changed: the ids they are saved with did not change.
   */
  rewriteStaleFormulas(rewrite: (tab: number, formula: string) => string): void {
    for (const [key, body] of this.cache) {
      if ((this.bodyEpoch.get(key) ?? -1) >= this.epoch) continue;
      const tab = Number(key.slice(0, key.indexOf(":")));
      // New cell objects, not changes in place: an undo snapshot still holds the old ones.
      const cells = body.cells ?? {};
      for (const [cellKey, cell] of Object.entries(cells)) {
        if (cell.formula) cells[cellKey] = { ...cell, formula: rewrite(tab, cell.formula) };
      }
      this.bodyEpoch.set(key, this.epoch);
    }
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
    if (!pending && chunk > 0 && this.tabs[tab].chunks[chunk]?.rowCount === 0) {
      // Every row of this chunk was deleted: its blob still holds them, but they are gone.
      const empty: ChunkBody = { rows: [], cells: {} };
      this.cache.set(key, empty);
      this.bodyEpoch.set(key, this.epoch);
      return Promise.resolve(empty);
    }
    if (!pending) {
      const entry = this.tabs[tab].chunks[chunk];
      pending = (this.fetchBody(entry.ref, this.root.servers) as Promise<StoredChunkBody>).then((stored) => {
        // Tabs may have been deleted or reordered while this downloaded: then it no longer belongs at this key.
        const current = this.tabs[tab]?.chunks[chunk];
        if (!current || current.ref !== entry.ref) {
          this.loading.delete(key);
          return { rows: [], cells: {} } as ChunkBody;
        }
        // The current column ids (they change when columns are inserted or deleted).
        const colIds = this.tabs[tab].colIds.length > 0 ? this.tabs[tab].colIds : this.tabRoots[tab].cols;
        const read = readChunk(stored, entry.rowStart, entry.rowCount, colIds);
        const body: ChunkBody = { ...read.layout, rows: read.rows, cells: this.refsById && !this.opening ? this.fromStoredCells(tab, read.cells) : read.cells };
        this.cache.set(key, body);
        this.bodyEpoch.set(key, this.epoch);
        this.loading.delete(key);
        for (const old of this.cache.keys()) {
          if (this.cache.size <= CACHE_CHUNKS) break;
          if (!this.dirty.has(old) && !this.held.has(old) && old !== key) this.cache.delete(old);
        }
        this.onChunkLoaded?.(tab, chunk);
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

  /** Is the chunk in memory right now? */
  hasChunk(tab: number, chunk: number): boolean {
    return this.cache.has(this.key(tab, chunk));
  }

  /** Every cell of a loaded chunk as [row, column, cell] (empty when the chunk is not in memory). */
  *cellsOfChunk(tab: number, chunk: number): Generator<[number, number, WorkbookCell]> {
    const body = this.cache.get(this.key(tab, chunk));
    if (!body?.cells) return;
    const start = this.tabs[tab].chunks[chunk].rowStart;
    for (const [key, cell] of Object.entries(body.cells)) {
      const at = key.indexOf(":");
      yield [start + Number(key.slice(0, at)), Number(key.slice(at + 1)), cell];
    }
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
    const changes: Change[] = [];
    try {
      // One chunk at a time: a range can span more chunks than the cache holds, so each chunk's changes
      // are written (and pinned as edited) before the next one is loaded.
      for (let c = from; c <= to; c++) {
        await this.load(tab, c);
        const entry = this.tabs[tab].chunks[c];
        const lo = c === from ? r1 : entry.rowStart;
        const hi = c === to ? r2 : entry.rowStart + entry.rowCount - 1;
        const local: Change[] = [];
        for (let row = lo; row <= hi; row++) {
          for (let col = c1; col <= c2; col++) {
            const before = this.rawCell(tab, row, col);
            const after = change(before ? structuredClone(before) : undefined, row, col);
            if (JSON.stringify(before) === JSON.stringify(after)) continue;
            local.push({ tab, row, col, before, after });
          }
        }
        if (local.length > 0) {
          this.apply(local, "after");
          changes.push(...local);
        }
      }
    } finally {
      // Even when a download failed half way, what was already written can be undone.
      if (changes.length > 0) {
        this.undoStack.push(changes);
        if (this.undoStack.length > UNDO_STEPS) this.undoStack.shift();
        this.redoStack = [];
      }
    }
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /** Undoes the last edit; returns the cells it changed (so formulas that read them can be refreshed). */
  undo(): CellAddress[] {
    this.structuralUndo = false;
    const step = this.undoStack.pop();
    if (!step) return [];
    if (!Array.isArray(step)) {
      // Insert, delete or a tab change: put back the state from before it. (It cannot be redone.)
      step.restore();
      this.redoStack = [];
      this.structuralUndo = true;
      this.emit();
      return [];
    }
    this.apply(step, "before");
    this.redoStack.push(step);
    return step.map(({ tab, row, col }) => ({ tab, row, col }));
  }

  redo(): CellAddress[] {
    const step = this.redoStack.pop();
    if (!step) return [];
    this.apply(step, "after");
    this.undoStack.push(step);
    return step.map(({ tab, row, col }) => ({ tab, row, col }));
  }

  /**
   * Stores a calculated value (a formula's new result). Not an undo step: it is recomputed from the
   * cells it reads, so undoing the edit that caused it recalculates it again.
   * `transient` keeps the value in memory only (the chunk is not marked as changed), for results that
   * are always recomputed when the sheet is opened, like NOW().
   */
  async setDerived(tab: number, row: number, col: number, cell: WorkbookCell, transient = false): Promise<void> {
    await this.load(tab, this.chunkFor(tab, row));
    const before = this.rawCell(tab, row, col);
    if (JSON.stringify(before) === JSON.stringify(cell)) return;
    this.apply([{ tab, row, col, before, after: cell }], "after", !transient);
  }

  /**
   * Changes what a tab remembers about its look (sizes, hidden rows, frozen panes). It is stored in
   * the tab's first chunk, which is then saved with the next save. Not an undo step.
   */
  async setLayout(tab: number, patch: Partial<TabLayout>): Promise<void> {
    const body = await this.load(tab, 0);
    const info = this.tabs[tab];
    for (const key of LAYOUT_KEYS) {
      if (!(key in patch)) continue;
      const value = patch[key];
      if (value === undefined) {
        delete body[key];
        delete info.layout[key];
      } else {
        body[key] = value;
        (info.layout as Record<string, unknown>)[key] = value;
      }
    }
    this.dirty.add(this.key(tab, 0));
    this.emit();
  }

  // --- undo for structural changes --------------------------------------------------------

  private takeSnapshot(): Snapshot {
    const pick = (body: ChunkBody) => {
      const layout: Record<string, unknown> = {};
      for (const key of LAYOUT_KEYS) if (body[key] !== undefined) layout[key] = structuredClone(body[key]);
      return layout;
    };
    return {
      tabs: this.tabs.map((tab) => ({
        tab,
        name: tab.name,
        chunks: [...tab.chunks],
        indexed: tab.indexed,
        rows: tab.rows,
        cols: tab.cols,
        colIds: [...tab.colIds],
        layout: structuredClone(tab.layout),
      })),
      tabRoots: this.tabRoots.map((root) => ({ ...root })),
      // Cells are copied one level deep: later edits replace entries in the live object, not in the copy.
      bodies: [...this.cache].map(([key, body]) => ({ key, body, cells: { ...body.cells }, rows: body.rows ? [...body.rows] : undefined, layout: pick(body) })),
      dirty: new Set(this.dirty),
      metaDirty: this.metaDirty,
    };
  }

  private restoreSnapshot(snap: Snapshot): void {
    this.tabs = snap.tabs.map(({ tab, name, chunks, indexed, rows, cols, colIds, layout }) => {
      Object.assign(tab, { name, chunks, indexed, rows, cols, colIds, layout });
      return tab;
    });
    this.tabRoots = snap.tabRoots;
    // Chunks read since are dropped (they are read again as they were saved); the ones that were open come back as they were.
    const cache = new Map<string, ChunkBody>();
    for (const { key, body, cells, rows, layout } of snap.bodies) {
      body.cells = cells;
      body.rows = rows;
      for (const name of LAYOUT_KEYS) delete body[name];
      Object.assign(body, layout);
      cache.set(key, body);
    }
    this.cache = cache;
    this.dirty = snap.dirty;
    this.metaDirty = snap.metaDirty;
    this.loading.clear();
  }

  /** Takes the state now; the returned function (call it once the change worked) makes the change undoable. */
  private beginUndoableChange(): () => void {
    const snap = this.takeSnapshot();
    return () => {
      this.undoStack.push({ restore: () => this.restoreSnapshot(snap) });
      this.redoStack = [];
      // Older steps are only valid together with the structural steps above them: dropping the oldest drops what is below it.
      const at = () => this.undoStack.map((step, i) => (Array.isArray(step) ? -1 : i)).filter((i) => i >= 0);
      let steps = at();
      while (steps.length > MAX_STRUCTURAL_UNDO) {
        this.undoStack.splice(0, steps[0] + 1);
        steps = at();
      }
    };
  }

  // --- inserting and deleting rows and columns ------------------------------------------------

  /**
   * Inserts or deletes rows or columns: moves the cells, the ids and the tab's saved look (sizes,
   * merges, borders). Formulas are NOT rewritten here (see bigStructure.ts). Undoable (the last few
   * such changes), but not redoable.
   */
  async applyEdit(tab: number, edit: StructEdit): Promise<void> {
    if (edit.count <= 0) return;
    const info = this.tabs[tab];
    const total = edit.axis === "row" ? info.rows : info.cols;
    if (edit.kind === "delete" && edit.index + edit.count > total) edit = { ...edit, count: total - edit.index };
    if (edit.count <= 0) return;
    if (edit.kind === "delete" && edit.count >= total) {
      throw new Error(`A sheet needs at least one ${edit.axis === "row" ? "row" : "column"}.`);
    }
    if (edit.kind === "insert" && edit.axis === "col" && info.cols + edit.count > MAX_COLUMNS) {
      throw new Error(`A sheet has at most ${MAX_COLUMNS.toLocaleString()} columns.`);
    }
    if (edit.kind === "insert" && edit.axis === "row" && edit.count > MAX_INSERT_ROWS) {
      throw new Error(`Insert at most ${MAX_INSERT_ROWS.toLocaleString()} rows at once.`);
    }
    // The first chunk carries the size and look, so it is always rewritten. It stays in memory until the edit is done.
    const headKey = this.key(tab, 0);
    const heldBefore = this.held.has(headKey);
    this.held.add(headKey);
    try {
      await this.load(tab, 0);
      await this.applyEditHeld(tab, edit, info);
    } finally {
      if (!heldBefore) this.held.delete(headKey);
    }
  }

  private async applyEditHeld(tab: number, edit: StructEdit, info: BigTab): Promise<void> {
    const finish = this.beginUndoableChange();
    // Rows and columns about to be deleted leave tombstones, so formulas that name them can still resolve.
    let tombs: Tombstones = {};
    if (this.refsById && edit.kind === "delete") {
      tombs = tombstonesFor(edit.axis === "row" ? new RunAxis(this.runsOf(tab)) : new ArrayAxis(padCols(info.colIds, info.cols)), edit.index, edit.count);
    }
    if (edit.axis === "row") await (edit.kind === "insert" ? this.insertRows(tab, edit) : this.deleteRows(tab, edit));
    else if (edit.kind === "insert") this.insertCols(tab, edit);
    else this.deleteCols(tab, edit);
    this.epoch++;

    const head = this.cache.get(this.key(tab, 0)) as ChunkBody;
    info.layout = editLayout(info.layout, edit);
    if (Object.keys(tombs).length > 0) {
      const have = info.layout.tombs ?? { rows: {}, cols: {} };
      const axisKey = edit.axis === "row" ? "rows" : "cols";
      info.layout = { ...info.layout, tombs: { ...have, [axisKey]: { ...have[axisKey], ...tombs } } };
    }
    for (const key of LAYOUT_KEYS) {
      const value = info.layout[key];
      if (value === undefined) delete head[key];
      else head[key] = value;
    }
    this.dirty.add(this.key(tab, 0));
    finish();
    this.emit();
  }

  private async insertRows(tab: number, edit: StructEdit): Promise<void> {
    const info = this.tabs[tab];
    const at = Math.min(edit.index, info.rows);
    const last = info.chunks[info.chunks.length - 1];
    const coverage = last.rowStart + last.rowCount;
    if (at < coverage) {
      const k = this.chunkFor(tab, at);
      const body = await this.load(tab, k);
      const entry = info.chunks[k];
      const local = at - entry.rowStart;
      body.cells = moveCells(body.cells, (row, col) => [row >= local ? row + edit.count : row, col]);
      body.rows = insertRuns(padRuns(body.rows ?? [], entry.rowCount), local, mintRows(this.idTag, this.idCounter, edit.count));
      this.idCounter += edit.count;
      info.chunks[k] = { ...entry, rowCount: entry.rowCount + edit.count, ...(this.refsById ? { runs: body.rows } : {}) };
      for (let j = k + 1; j < info.chunks.length; j++) info.chunks[j] = { ...info.chunks[j], rowStart: info.chunks[j].rowStart + edit.count };
      this.dirty.add(this.key(tab, k));
    }
    info.rows += edit.count;
  }

  private async deleteRows(tab: number, edit: StructEdit): Promise<void> {
    const info = this.tabs[tab];
    const last = info.chunks[info.chunks.length - 1];
    const coverage = last.rowStart + last.rowCount;
    const from = edit.index;
    const to = Math.min(edit.index + edit.count, coverage); // rows [from, to) live in chunks
    let removed = 0;
    if (from < to) {
      const first = this.chunkFor(tab, from);
      const lastChunk = this.chunkFor(tab, to - 1);
      // Load what must be edited before changing anything, so a failed download leaves the sheet alone.
      const partial: number[] = [];
      for (let k = first; k <= lastChunk; k++) {
        const entry = info.chunks[k];
        const covered = from <= entry.rowStart && to >= entry.rowStart + entry.rowCount;
        if (!covered || k === 0) partial.push(k);
      }
      const heldKeys = partial.map((k) => this.key(tab, k)).filter((key) => !this.held.has(key));
      for (const key of heldKeys) this.held.add(key);
      try {
        for (const k of partial) await this.load(tab, k);
      } finally {
        for (const key of heldKeys) this.held.delete(key);
      }
      const chunks = [...info.chunks];
      for (let k = 0; k < chunks.length; k++) {
        const entry = chunks[k];
        const start = entry.rowStart;
        const end = start + entry.rowCount;
        const lo = Math.max(from, start);
        const hi = Math.min(to, end);
        if (hi > lo && entry.rowCount > 0) {
          const gone = hi - lo;
          if (partial.includes(k)) {
            const body = this.cache.get(this.key(tab, k)) as ChunkBody;
            const a = lo - start;
            const b = hi - start;
            body.cells = moveCells(body.cells, (row, col) => (row >= a && row < b ? null : [row >= b ? row - gone : row, col]));
            body.rows = deleteRuns(padRuns(body.rows ?? [], entry.rowCount), a, gone);
            this.dirty.add(this.key(tab, k));
          } else {
            // Every row of the chunk is deleted: it is simply emptied (nothing to download).
            this.cache.delete(this.key(tab, k));
            this.dirty.delete(this.key(tab, k));
          }
          chunks[k] = {
            ...entry,
            rowStart: start - removed,
            rowCount: entry.rowCount - gone,
            cells: partial.includes(k) ? entry.cells : 0,
            ...(this.refsById ? { runs: partial.includes(k) ? (this.cache.get(this.key(tab, k)) as ChunkBody).rows : [] } : {}),
          };
          removed += gone;
        } else {
          chunks[k] = { ...entry, rowStart: start - removed };
        }
      }
      info.chunks = chunks;
    }
    info.rows -= edit.count;
  }

  private insertCols(tab: number, edit: StructEdit): void {
    const info = this.tabs[tab];
    const at = Math.min(edit.index, info.cols);
    const ids = padCols(info.colIds, info.cols);
    const fresh = Array.from({ length: edit.count }, () => `x${this.idTag}${this.idCounter++}`);
    info.colIds = [...ids.slice(0, at), ...fresh, ...ids.slice(at)];
    info.cols += edit.count;
    this.mapLoadedCells(tab, (row, col) => [row, col >= at ? col + edit.count : col]);
  }

  private deleteCols(tab: number, edit: StructEdit): void {
    const info = this.tabs[tab];
    const ids = padCols(info.colIds, info.cols);
    const to = edit.index + edit.count;
    info.colIds = [...ids.slice(0, edit.index), ...ids.slice(to)];
    info.cols -= edit.count;
    this.mapLoadedCells(tab, (row, col) => (col >= edit.index && col < to ? null : [row, col >= to ? col - edit.count : col]));
    // Chunks written before may still hold the deleted columns; they are ignored when read (their ids are gone).
  }

  /** Moves the cells of every loaded chunk of a tab (positions are chunk-relative rows and columns). */
  private mapLoadedCells(tab: number, move: (row: number, col: number) => [number, number] | null): void {
    for (const [key, body] of this.cache) {
      if (!key.startsWith(`${tab}:`)) continue;
      body.cells = moveCells(body.cells, move);
    }
  }

  // --- tabs ----------------------------------------------------------------------------------

  /** Are there tab changes (names, colours, order, added or deleted tabs) not saved yet? */
  get hasTabChanges(): boolean {
    return this.metaDirty;
  }

  /** The tab entry as it will be saved (colour, hidden). */
  tabMeta(tab: number): Pick<RootTab, "id" | "color" | "hidden"> {
    const { id, color, hidden } = this.tabRoots[tab];
    return { id, color, hidden };
  }

  private changeTab(tab: number, patch: Partial<Pick<RootTab, "name" | "color" | "hidden">>): void {
    const finish = this.beginUndoableChange();
    const next = { ...this.tabRoots[tab], ...patch };
    for (const key of Object.keys(next) as (keyof typeof next)[]) if (next[key] === undefined) delete next[key];
    this.tabRoots[tab] = next;
    if (patch.name !== undefined) this.tabs[tab].name = patch.name;
    this.metaDirty = true;
    finish();
    this.emit();
  }

  renameTab(tab: number, name: string): void {
    this.changeTab(tab, { name });
  }

  setTabColor(tab: number, color: string | undefined): void {
    this.changeTab(tab, { color });
  }

  setTabHidden(tab: number, hidden: boolean): void {
    if (hidden && this.tabRoots.filter((t) => !t.hidden).length <= 1) throw new Error("At least one tab has to stay visible.");
    this.changeTab(tab, { hidden: hidden || undefined });
  }

  /** Adds an empty tab at the end and returns its index. It exists in memory until saved. */
  addTab(name: string): number {
    const finish = this.beginUndoableChange();
    const index = this.tabs.length;
    const rows = 200;
    const cols = 26;
    const entry: ChunkEntry = {
      id: randomId(8),
      ref: { x: "", k: "", n: 0, p: 0, c: 0, h: "" },
      rowStart: 0,
      rowCount: rows,
    };
    this.tabRoots.push({ id: randomId(8), name, order: index });
    this.tabs.push({ name, chunks: [entry], indexed: false, rows, cols, colIds: padCols([], cols), layout: {} });
    const body: ChunkBody = { rows: mintRows(this.idTag, this.idCounter, rows), cells: {} };
    this.idCounter += rows;
    if (this.refsById) this.tabs[index].chunks[0] = { ...entry, runs: body.rows };
    this.cache.set(this.key(index, 0), body);
    this.bodyEpoch.set(this.key(index, 0), this.epoch);
    this.dirty.add(this.key(index, 0));
    this.metaDirty = true;
    finish();
    this.emit();
    return index;
  }

  /**
   * Adds a copy of a tab right after it and returns its index. Nothing is downloaded: the copy points at
   * the same (unchangeable) blobs, and a chunk gets its own blob only when it is edited. The source tab
   * must be saved first, so that what is copied is what the sheet holds.
   */
  duplicateTab(tab: number, name: string): number {
    const info = this.tabs[tab];
    if ([...this.dirty].some((key) => key.startsWith(`${tab}:`))) throw new Error("Save your changes first, then copy the tab.");
    const finish = this.beginUndoableChange();
    const source = this.tabRoots[tab];
    const at = tab + 1;
    const copy: BigTab = {
      name,
      chunks: info.chunks.map((entry) => ({ ...entry, id: randomId(8) })),
      indexed: info.indexed,
      rows: info.rows,
      cols: info.cols,
      colIds: [...info.colIds],
      layout: structuredClone(info.layout),
    };
    const entry: RootTab = {
      id: randomId(8),
      name,
      order: at,
      cols: [...info.colIds],
      chunks: Array.isArray(source.chunks) ? copy.chunks.map((chunk) => ({ ...chunk })) : source.chunks,
      ...(source.color ? { color: source.color } : {}),
      ...(source.showGridLines !== undefined ? { showGridLines: source.showGridLines } : {}),
    };
    // Insert after the source by renumbering the tabs that follow it.
    const order = this.tabs.map((_, i) => i);
    this.tabs.push(copy);
    this.tabRoots.push(entry);
    order.splice(at, 0, this.tabs.length - 1);
    this.remapTabsRaw(order);
    finish();
    return at;
  }

  /**
   * Reorders or drops tabs: `order` lists the old index of each tab that stays, in its new place.
   * Loaded chunks follow their tab. Undoable (the last few structural changes), not redoable.
   */
  remapTabs(order: number[]): void {
    if (order.length === 0) throw new Error("A workbook needs at least one tab.");
    const finish = this.beginUndoableChange();
    this.remapTabsRaw(order);
    finish();
  }

  private remapTabsRaw(order: number[]): void {
    const where = new Map(order.map((old, now) => [old, now]));
    const move = (map: Map<string, ChunkBody>) => {
      const next = new Map<string, ChunkBody>();
      for (const [key, body] of map) {
        const [tab, chunk] = key.split(":").map(Number);
        const now = where.get(tab);
        if (now !== undefined) next.set(this.key(now, chunk), body);
      }
      return next;
    };
    this.cache = move(this.cache);
    const epochs = new Map<string, number>();
    for (const [key, value] of this.bodyEpoch) {
      const [tab, chunk] = key.split(":").map(Number);
      const now = where.get(tab);
      if (now !== undefined) epochs.set(this.key(now, chunk), value);
    }
    this.bodyEpoch = epochs;
    this.forceTabs = new Set([...this.forceTabs].flatMap((tab) => (where.has(tab) ? [where.get(tab) as number] : [])));
    const dirty = new Set<string>();
    for (const key of this.dirty) {
      const [tab, chunk] = key.split(":").map(Number);
      const now = where.get(tab);
      if (now !== undefined) dirty.add(this.key(now, chunk));
    }
    this.dirty = dirty;
    this.loading.clear();
    // Formulas that name a tab that is now gone read #REF! whichever chunk they are in (the ids of unopened chunks do already).
    const kept = new Set(order);
    const survivors = new Set(order.map((old) => this.tabs[old].name.toLowerCase()));
    const dropped = this.tabs.filter((_, old) => !kept.has(old)).map((tab) => tab.name).filter((name) => !survivors.has(name.toLowerCase()));
    if (this.refsById && dropped.length > 0) {
      for (const body of this.cache.values()) {
        const cells = body.cells ?? {};
        for (const [cellKey, cell] of Object.entries(cells)) {
          if (cell.formula) {
            const next = dropSheetRefs(cell.formula, dropped);
            if (next !== cell.formula) cells[cellKey] = { ...cell, formula: next };
          }
        }
      }
    }
    this.tabs = order.map((old) => this.tabs[old]);
    this.tabRoots = order.map((old, now) => ({ ...this.tabRoots[old], order: now }));
    this.metaDirty = true;
    this.emit();
  }

  /** Chunks currently in memory, as [tab, chunk]. */
  loadedChunks(): [number, number][] {
    return [...this.cache.keys()].map((key) => key.split(":").map(Number) as [number, number]);
  }

  private rawCell(tab: number, row: number, col: number): WorkbookCell | undefined {
    const chunk = this.chunkFor(tab, row);
    const body = this.cache.get(this.key(tab, chunk));
    return body?.cells?.[`${row - this.tabs[tab].chunks[chunk].rowStart}:${col}`];
  }

  /** Writes changes into their (pinned) chunks and marks them dirty. */
  private apply(changes: Change[], side: "before" | "after", markDirty = true): void {
    for (const item of changes) {
      const chunk = this.chunkFor(item.tab, item.row);
      const key = this.key(item.tab, chunk);
      const body = this.cache.get(key);
      if (!body) {
        // Not pinned any more (it was saved, then left the cache): fetch it, then write the change.
        void this.load(item.tab, chunk).then(() => this.apply([item], side, markDirty), () => undefined);
        continue;
      }
      const cells = (body.cells ??= {});
      const cellKey = `${item.row - this.tabs[item.tab].chunks[chunk].rowStart}:${item.col}`;
      const value = item[side];
      if (value) cells[cellKey] = structuredClone(value);
      else delete cells[cellKey];
      const tabInfo = this.tabs[item.tab];
      tabInfo.rows = Math.max(tabInfo.rows, item.row + 1);
      tabInfo.cols = Math.max(tabInfo.cols, item.col + 1);
      if (tabInfo.colIds.length < tabInfo.cols) tabInfo.colIds = padCols(tabInfo.colIds, tabInfo.cols);
      if (markDirty) this.dirty.add(key);
    }
    this.emit();
  }

  /**
   * Uploads the edited chunks under one signature and returns the root that points at them.
   * The caller signs and publishes the root event.
   */
  async save(options: UploadOptions): Promise<SheetRootV2> {
    if (this.dirty.size === 0 && !this.metaDirty) return this.root;
    const keys = [...this.dirty];
    const staged: StagedBlob[] = [];
    const targets: { tab: number; chunk: number }[] = [];
    // First every changed chunk gets its row ids (rows added past the end get new ids when formulas are stored by id),
    // so that all ids are known before any formula is converted.
    const idsOf = new Map<string, RowRun[]>();
    for (const key of keys) {
      const [tab, chunk] = key.split(":").map(Number);
      const body = this.cache.get(key) as ChunkBody;
      const entry = this.tabs[tab].chunks[chunk];
      const next = this.tabs[tab].chunks[chunk + 1];
      if (!next) entry.rowCount = Math.max(entry.rowCount, this.tabs[tab].rows - entry.rowStart);
      const have = rowCountOf(body.rows ?? []);
      const missing = entry.rowCount - have;
      let ids: RowRun[];
      if (this.refsById) {
        ids = concatRuns(body.rows ?? [], mintRows(this.idTag, this.idCounter, Math.max(0, missing)));
        this.idCounter += Math.max(0, missing);
        body.rows = ids;
        this.tabs[tab].chunks[chunk] = { ...entry, runs: ids };
      } else {
        ids = concatRuns(body.rows ?? [], legacyRows(entry.rowStart + have, missing));
      }
      idsOf.set(key, ids);
    }
    for (const key of keys) {
      const [tab, chunk] = key.split(":").map(Number);
      const body = this.cache.get(key) as ChunkBody;
      if (chunk === 0) {
        body.rowCount = this.tabs[tab].rows;
        body.colCount = this.tabs[tab].cols;
      }
      const { rows: _rows, cells, ...layout } = body;
      void _rows;
      const stored = writeChunk(
        { rows: idsOf.get(key) as RowRun[], cells: this.refsById ? this.toStoredCells(tab, cells ?? {}) : (cells ?? {}), layout, hasIds: true },
        this.tabs[tab].colIds,
      );
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
    const touched = [...new Set([...targets.map((t) => t.tab), ...this.forceTabs])].filter((tab) => tab < this.tabs.length);
    // The root is one small event: a chunk list that has grown (each entry now also holds its row ids) moves into an index blob.
    for (const tab of touched) {
      if (!this.tabs[tab].indexed && JSON.stringify(this.tabs[tab].chunks).length > INDEX_AT_BYTES) this.tabs[tab].indexed = true;
    }
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
    const tabRoots = this.tabRoots.map((tab, index) => {
      if (!touched.includes(index)) return tab;
      const at = indexOf.get(index);
      return { ...tab, cols: [...this.tabs[index].colIds], chunks: at === undefined ? [...this.tabs[index].chunks] : { index: uploaded[at].ref } };
    });
    const { refs: _inherited, ...rest } = this.root;
    void _inherited;
    const root: SheetRootV2 = {
      ...rest,
      ...(this.byId ? { refs: "id" as const } : {}),
      updatedAt: Date.now(),
      servers: [...holders],
      // The tab that was open may have been deleted since.
      activeSheetId: tabRoots.some((tab) => tab.id === this.root.activeSheetId) ? this.root.activeSheetId : (tabRoots.find((tab) => !tab.hidden) ?? tabRoots[0]).id,
      tabs: tabRoots,
    };
    Object.assign(this.root, root);
    if (!this.byId) delete this.root.refs; // Object.assign cannot remove a key another store set on the shared object
    this.tabRoots = root.tabs.map((tab) => ({ ...tab }));
    this.markSaved();
    return root;
  }

  /** Index of the tab the sheet was left on (the first visible tab when that one is gone or hidden). */
  activeTab(): number {
    const at = this.tabRoots.findIndex((tab) => tab.id === this.root.activeSheetId && !tab.hidden);
    if (at >= 0) return at;
    const visible = this.tabRoots.findIndex((tab) => !tab.hidden);
    return visible >= 0 ? visible : 0;
  }

  /** Remembers which tab is open (it is saved with the sheet; not a change that needs saving). */
  setActiveTab(index: number): void {
    const id = this.tabRoots[index]?.id;
    if (id) this.root.activeSheetId = id;
  }

  /** Everything now counts as saved (after `save`, or after an in-memory store was written out as a document). */
  markSaved(): void {
    this.metaDirty = false;
    this.forceTabs.clear();
    this.dirty.clear();
    // Structural steps describe the state from before this save; once it is saved they (and the steps below them) cannot be undone.
    let last = -1;
    this.undoStack.forEach((step, i) => {
      if (!Array.isArray(step)) last = i;
    });
    if (last >= 0) this.undoStack.splice(0, last + 1);
    this.emit();
  }
}

/** Does any row id appear twice in these runs? */
function hasDuplicateIds(runs: RowRun[]): boolean {
  const byTag = new Map<string, [number, number][]>();
  for (const [tag, start, length] of runs) {
    const list = byTag.get(tag) ?? [];
    list.push([start, start + length]);
    byTag.set(tag, list);
  }
  for (const list of byTag.values()) {
    list.sort((a, b) => a[0] - b[0]);
    for (let i = 1; i < list.length; i++) if (list[i][0] < list[i - 1][1]) return true;
  }
  return false;
}

/** The cells with each position moved by `move` (null drops the cell). */
function moveCells(
  cells: Record<string, WorkbookCell> | undefined,
  move: (row: number, col: number) => [number, number] | null,
): Record<string, WorkbookCell> {
  const out: Record<string, WorkbookCell> = {};
  for (const [key, cell] of Object.entries(cells ?? {})) {
    const at = key.indexOf(":");
    const next = move(Number(key.slice(0, at)), Number(key.slice(at + 1)));
    if (next) out[`${next[0]}:${next[1]}`] = cell;
  }
  return out;
}
