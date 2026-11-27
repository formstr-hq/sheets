// Formulas typed into a big sheet.
//
// The sheet is never loaded whole. To calculate one formula we look at what it names (`=SUM(A:A)`
// names column A), download only the chunks that hold those rows, copy the needed values out
// (the chunk cache may drop the chunks right after) and run the formula engine on that copy.
//
// Not done here: recalculating other formulas when a cell they read changes. That needs an index
// of which cells hold formulas, which is the next step.
import type { WorkbookCell } from "../workbook";
import type { BigSheetStore, CellAddress } from "./bigSheet";

export interface FormulaResult {
  value: string | number | boolean | null;
  /** Set when the formula evaluates to an error (value then holds the code, e.g. "#DIV/0!"). */
  error: boolean;
}

/** Most cells one formula may read (matches the engine's own limit). */
const MAX_READ_CELLS = 5_000_000;
const KEY = 16384;
/** A read of this many cells or fewer is looked up cell by cell. */
const SMALL_READ = 512;

interface Rect {
  tab: number;
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

/** Calculates `formula` (without the leading "=") as if typed into the cell. */
export async function calculateTyped(
  store: BigSheetStore,
  tab: number,
  row: number,
  col: number,
  formula: string,
  onProgress?: (done: number, total: number) => void,
): Promise<FormulaResult> {
  const engine = await import("@formstr/sheets-formula");
  const { FormulaError } = engine;
  const home = store.tabs[tab];
  const tabIndex = (name: string | undefined) => (name === undefined ? tab : store.tabs.findIndex((t) => t.name === name));

  let node;
  try {
    node = engine.parse(formula);
  } catch {
    return { value: "#NAME?", error: true };
  }

  // What the formula reads, as rectangles per tab.
  const rects: Rect[] = [];
  for (const ref of engine.references(node)) {
    const at = tabIndex(ref.sheet);
    if (at < 0) return { value: "#REF!", error: true };
    const info = store.tabs[at];
    if (ref.kind === "cell" && ref.cell) {
      rects.push({ tab: at, r1: ref.cell.row, r2: ref.cell.row, c1: ref.cell.col, c2: ref.cell.col });
    } else if (ref.range) {
      const { r1, c1, r2, c2 } = ref.range;
      rects.push({ tab: at, r1, c1, r2: r2 ?? info.rows - 1, c2: c2 ?? info.cols - 1 });
    }
  }
  let cells = 0;
  for (const rect of rects) cells += Math.max(0, rect.r2 - rect.r1 + 1) * Math.max(0, rect.c2 - rect.c1 + 1);
  if (cells > MAX_READ_CELLS) return { value: "#VALUE!", error: true };

  // Copy the values out, one chunk at a time.
  const values = new Map<number, Map<number, WorkbookCell>>();
  const grab = (tabNo: number) => {
    let map = values.get(tabNo);
    if (!map) values.set(tabNo, (map = new Map()));
    return map;
  };
  const jobs: { rect: Rect; chunk: number }[] = [];
  for (const rect of rects) {
    const info = store.tabs[rect.tab];
    if (info.chunks.length === 0 || rect.r2 < rect.r1) continue;
    const from = store.chunkFor(rect.tab, Math.max(0, rect.r1));
    const to = store.chunkFor(rect.tab, Math.max(0, rect.r2));
    for (let chunk = from; chunk <= to; chunk++) jobs.push({ rect, chunk });
  }
  let done = 0;
  for (const { rect, chunk } of jobs) {
    await store.load(rect.tab, chunk);
    const target = grab(rect.tab);
    const entry = store.tabs[rect.tab].chunks[chunk];
    const lastRow = chunk === store.tabs[rect.tab].chunks.length - 1 ? Number.MAX_SAFE_INTEGER : entry.rowStart + entry.rowCount - 1;
    const r1 = Math.max(rect.r1, entry.rowStart);
    const r2 = Math.min(rect.r2, lastRow);
    if ((r2 - r1 + 1) * (rect.c2 - rect.c1 + 1) <= SMALL_READ) {
      // A few cells (the usual `=A1*2`): look each one up instead of walking the whole chunk.
      for (let r = r1; r <= r2; r++) {
        for (let c = rect.c1; c <= rect.c2; c++) {
          const cell = store.cell(rect.tab, r, c);
          if (cell) target.set(r * KEY + c, cell);
        }
      }
    } else {
      for (const [r, c, cell] of store.cellsOfChunk(rect.tab, chunk)) {
        if (r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2) target.set(r * KEY + c, cell);
      }
    }
    onProgress?.(++done, jobs.length);
  }
  // Cells edited since (not yet saved) are in the chunks, so they are read here too.

  const result = engine.calculate(formula, {
    sheet: home.name,
    row,
    col,
    getCell(sheetName, r, c) {
      const at = tabIndex(sheetName);
      const cell = values.get(at)?.get(r * KEY + c);
      if (!cell || cell.value === undefined || cell.value === null) return undefined;
      if (cell.type === "error" && typeof cell.value === "string") return new FormulaError(cell.value as never);
      return cell.value;
    },
    extent: (sheetName) => {
      const info = store.tabs[tabIndex(sheetName)] ?? home;
      return { rows: info.rows, cols: info.cols };
    },
    sheetNames: () => store.tabs.map((t) => t.name),
  });
  if (result instanceof FormulaError) return { value: result.code, error: true };
  // Excel keeps 15 significant digits, so 0.1+0.2 shows as 0.3 and a long sum has no float noise.
  return { value: typeof result === "number" && Number.isFinite(result) ? Number(result.toPrecision(15)) : result, error: false };
}

// ---- keeping dependent formulas up to date -----------------------------------------------------

/** What a formula reads in one tab; `null` ends mean "to the last row/column" (A:A, 1:1). */
export interface Footprint {
  tab: number;
  r1: number;
  c1: number;
  r2: number | null;
  c2: number | null;
}

export interface FormulaEntry {
  tab: number;
  row: number;
  col: number;
  formula: string;
  reads: Footprint[];
}

/** A rectangle of changed cells (inclusive). */
export interface ChangedRect {
  tab: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** One cell as a rectangle. */
export const pointRect = (cell: CellAddress): ChangedRect => ({ tab: cell.tab, r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });

const addressKey = (tab: number, row: number, col: number) => `${tab}:${row}:${col}`;
const bare = (formula: string) => formula.replace(/^=/, "");

/** Tells the caller how far a long job is; `phase` says whether it is still looking for formulas or recalculating. */
export type RefreshProgress = (done: number, total: number, phase: "scan" | "calc") => void;

/** Lets the page paint and respond now and then during a long job. */
function breather(): () => Promise<void> {
  let last = Date.now();
  return async () => {
    if (Date.now() - last < 40) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    last = Date.now();
  };
}

/** A changed area this small is checked cell by cell against the formulas that read single cells. */
const POINT_LOOKUP_MAX = 4096;

/**
 * Which cells hold formulas and what each reads. It learns formulas as chunks are opened and as
 * they are typed; `ensureScanned` opens every other chunk in the background so that nothing is
 * missed before the first recalculation.
 */
export class FormulaIndex {
  private readonly entries = new Map<string, FormulaEntry>();
  /** Formulas that read exactly one cell, by that cell. */
  private readonly byPoint = new Map<string, Set<FormulaEntry>>();
  /** Formulas that read a range (or an open-ended one): checked against every change. */
  private readonly wide = new Set<FormulaEntry>();
  private readonly scanned = new Set<string>();

  get size(): number {
    return this.entries.size;
  }

  /** Every formula the index knows. */
  all(): FormulaEntry[] {
    return [...this.entries.values()];
  }

  /** Forgets everything, including which chunks were scanned (tabs were deleted or reordered, so tab numbers changed). */
  forget(): void {
    this.reset();
    this.scanned.clear();
  }

  /** Forgets every formula (the chunks stay marked as scanned). Used after rows/columns moved, before re-adding them. */
  reset(): void {
    this.entries.clear();
    this.byPoint.clear();
    this.wide.clear();
  }

  has(tab: number, row: number, col: number): boolean {
    return this.entries.has(addressKey(tab, row, col));
  }

  remove(tab: number, row: number, col: number): void {
    const key = addressKey(tab, row, col);
    const old = this.entries.get(key);
    if (old) this.unregister(old);
    this.entries.delete(key);
  }

  private register(entry: FormulaEntry): void {
    let isWide = false;
    for (const f of entry.reads) {
      if (f.r2 === f.r1 && f.c2 === f.c1) {
        const key = addressKey(f.tab, f.r1, f.c1);
        let set = this.byPoint.get(key);
        if (!set) this.byPoint.set(key, (set = new Set()));
        set.add(entry);
      } else isWide = true;
    }
    if (isWide) this.wide.add(entry);
  }

  private unregister(entry: FormulaEntry): void {
    for (const f of entry.reads) {
      const key = addressKey(f.tab, f.r1, f.c1);
      const set = this.byPoint.get(key);
      if (set) {
        set.delete(entry);
        if (set.size === 0) this.byPoint.delete(key);
      }
    }
    this.wide.delete(entry);
  }

  /** Records (or replaces) the formula in a cell. Text that does not parse is recorded as reading nothing. */
  async set(store: BigSheetStore, tab: number, row: number, col: number, formula: string): Promise<void> {
    const engine = await import("@formstr/sheets-formula");
    const text = bare(formula);
    const reads: Footprint[] = [];
    try {
      for (const ref of engine.references(engine.parse(text))) {
        const at = ref.sheet === undefined ? tab : store.tabs.findIndex((t) => t.name === ref.sheet);
        if (at < 0) continue;
        if (ref.kind === "cell" && ref.cell) reads.push({ tab: at, r1: ref.cell.row, c1: ref.cell.col, r2: ref.cell.row, c2: ref.cell.col });
        else if (ref.range) reads.push({ tab: at, r1: ref.range.r1, c1: ref.range.c1, r2: ref.range.r2, c2: ref.range.c2 });
      }
    } catch {
      // not a formula the engine understands: it reads nothing, so nothing refreshes it
    }
    this.remove(tab, row, col);
    const entry: FormulaEntry = { tab, row, col, formula: text, reads };
    this.entries.set(addressKey(tab, row, col), entry);
    this.register(entry);
  }

  /** Learns the formulas in a loaded chunk. Returns false when the chunk was not in memory. */
  async scanChunk(store: BigSheetStore, tab: number, chunk: number): Promise<boolean> {
    if (!store.hasChunk(tab, chunk)) return false;
    const found: [number, number, string][] = [];
    for (const [row, col, cell] of store.cellsOfChunk(tab, chunk)) {
      if (cell.formula && !this.has(tab, row, col)) found.push([row, col, cell.formula]);
    }
    for (const [row, col, formula] of found) await this.set(store, tab, row, col, formula);
    this.scanned.add(`${tab}:${chunk}`);
    return true;
  }

  /** Opens every chunk not looked at yet (a few at a time stay in memory) so every formula is known. */
  async ensureScanned(store: BigSheetStore, onProgress?: RefreshProgress): Promise<void> {
    const todo: [number, number][] = [];
    store.tabs.forEach((info, tab) =>
      info.chunks.forEach((_, chunk) => {
        if (!this.scanned.has(`${tab}:${chunk}`)) todo.push([tab, chunk]);
      }),
    );
    if (todo.length === 0) return;
    const breathe = breather();
    let done = 0;
    for (const [tab, chunk] of todo) {
      // Another load may push the chunk out before it is read; then it is simply loaded again.
      for (let attempt = 0; attempt < 3; attempt++) {
        await store.load(tab, chunk);
        if (await this.scanChunk(store, tab, chunk)) break;
      }
      onProgress?.(++done, todo.length, "scan");
      await breathe();
    }
  }

  /** Brings the entry for a cell in line with what the sheet holds there now. */
  async sync(store: BigSheetStore, tab: number, row: number, col: number): Promise<void> {
    // An unloaded chunk reads as "no cell": load it first so a formula there is not forgotten.
    await store.load(tab, store.chunkFor(tab, row));
    const cell = store.cell(tab, row, col);
    if (cell?.formula) await this.set(store, tab, row, col, cell.formula);
    else this.remove(tab, row, col);
  }

  /** Brings the entries inside a rectangle in line with the sheet (a clear or an undo can remove formulas). */
  async syncArea(store: BigSheetStore, area: ChangedRect, discover = false): Promise<FormulaEntry[]> {
    const found: FormulaEntry[] = [];
    if (area.r1 === area.r2 && area.c1 === area.c2) {
      await store.load(area.tab, store.chunkFor(area.tab, area.r1));
      await this.sync(store, area.tab, area.r1, area.c1);
      const entry = this.entries.get(addressKey(area.tab, area.r1, area.c1));
      if (entry) found.push(entry);
      return found;
    }
    for (const entry of [...this.entries.values()]) {
      if (entry.tab !== area.tab || entry.row < area.r1 || entry.row > area.r2 || entry.col < area.c1 || entry.col > area.c2) continue;
      await this.sync(store, entry.tab, entry.row, entry.col);
    }
    if (discover) {
      // Cells pasted or filled into the area may hold formulas the index has not seen.
      for (let chunk = store.chunkFor(area.tab, area.r1); chunk <= store.chunkFor(area.tab, area.r2); chunk++) {
        await store.load(area.tab, chunk);
        const fresh: [number, number, string][] = [];
        for (const [row, col, cell] of store.cellsOfChunk(area.tab, chunk)) {
          if (cell.formula && row >= area.r1 && row <= area.r2 && col >= area.c1 && col <= area.c2) fresh.push([row, col, cell.formula]);
        }
        for (const [row, col, formula] of fresh) await this.set(store, area.tab, row, col, formula);
      }
    }
    for (const entry of this.entries.values()) {
      if (entry.tab === area.tab && entry.row >= area.r1 && entry.row <= area.r2 && entry.col >= area.c1 && entry.col <= area.c2) found.push(entry);
    }
    return found;
  }

  /** Does the formula read any cell of the rectangle? */
  private reads(entry: FormulaEntry, store: BigSheetStore, area: ChangedRect): boolean {
    for (const f of entry.reads) {
      if (f.tab !== area.tab) continue;
      const info = store.tabs[area.tab];
      const r2 = f.r2 ?? (info ? info.rows - 1 : Number.MAX_SAFE_INTEGER);
      const c2 = f.c2 ?? (info ? info.cols - 1 : Number.MAX_SAFE_INTEGER);
      if (f.r1 <= area.r2 && r2 >= area.r1 && f.c1 <= area.c2 && c2 >= area.c1) return true;
    }
    return false;
  }

  /** The formulas that read any cell of `area`. */
  private readersOf(store: BigSheetStore, area: ChangedRect): Set<FormulaEntry> {
    const out = new Set<FormulaEntry>();
    const cells = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1);
    if (cells <= POINT_LOOKUP_MAX) {
      for (let r = area.r1; r <= area.r2; r++) {
        for (let c = area.c1; c <= area.c2; c++) this.byPoint.get(addressKey(area.tab, r, c))?.forEach((e) => out.add(e));
      }
    } else {
      for (const [key, set] of this.byPoint) {
        const [tab, row, col] = key.split(":").map(Number);
        if (tab === area.tab && row >= area.r1 && row <= area.r2 && col >= area.c1 && col <= area.c2) set.forEach((e) => out.add(e));
      }
    }
    for (const entry of this.wide) if (this.reads(entry, store, area)) out.add(entry);
    return out;
  }

  /** The formulas that must be recalculated after `changed` cells changed, in an order that works (inputs first). */
  affected(store: BigSheetStore, changed: ChangedRect[], include: FormulaEntry[] = []): { order: FormulaEntry[]; cyclic: FormulaEntry[] } {
    const members = new Set<FormulaEntry>();
    const readers = new Map<FormulaEntry, FormulaEntry[]>();
    const waiting = new Map<FormulaEntry, number>();
    const queue: FormulaEntry[] = [];
    // Formulas that were just pasted or filled hold a result copied from elsewhere: they need their own recalculation.
    for (const entry of include) {
      members.add(entry);
      waiting.set(entry, 0);
      queue.push(entry);
    }
    const reach = (from: ChangedRect, source: FormulaEntry | null) => {
      for (const entry of this.readersOf(store, from)) {
        if (source) {
          // `entry` must wait for `source`
          const list = readers.get(source) ?? [];
          list.push(entry);
          readers.set(source, list);
          waiting.set(entry, (waiting.get(entry) ?? 0) + 1);
        }
        if (!members.has(entry)) {
          members.add(entry);
          waiting.set(entry, waiting.get(entry) ?? 0);
          queue.push(entry);
        }
      }
    };
    for (const area of changed) reach(area, null);
    for (let at = 0; at < queue.length; at++) reach(pointRect(queue[at]), queue[at]);

    const order: FormulaEntry[] = [];
    const ready = [...members].filter((e) => waiting.get(e) === 0);
    for (let at = 0; at < ready.length; at++) {
      const entry = ready[at];
      order.push(entry);
      for (const reader of readers.get(entry) ?? []) {
        const left = (waiting.get(reader) ?? 1) - 1;
        waiting.set(reader, left);
        if (left === 0) ready.push(reader);
      }
    }
    const placed = new Set(order);
    return { order, cyclic: [...members].filter((e) => !placed.has(e)) };
  }
}

/**
 * After cells changed, recalculates every formula that reads them (and the formulas that read
 * those), writing the new results into the sheet. First makes sure every chunk has been looked at
 * for formulas, in the background, so none is skipped. Returns how many formulas were recalculated.
 */
export async function refreshDependents(
  store: BigSheetStore,
  index: FormulaIndex,
  changed: ChangedRect[],
  onProgress?: RefreshProgress,
  /** Set after a paste or fill: the formulas inside the changed areas are recalculated too. */
  recalcInside = false,
  options: {
    /** Results stay in memory only (see BigSheetStore.setDerived). */
    transient?: boolean;
    /** Recalculate every formula in the workbook, not only what depends on `changed`. */
    all?: boolean;
  } = {},
): Promise<number> {
  const inside: FormulaEntry[] = [];
  for (const area of changed) inside.push(...(await index.syncArea(store, area, recalcInside)));
  await index.ensureScanned(store, onProgress);
  const { order, cyclic } = index.affected(store, changed, options.all ? index.all() : recalcInside ? inside : []);
  const write = async (entry: { tab: number; row: number; col: number }, result: FormulaResult) => {
    await store.load(entry.tab, store.chunkFor(entry.tab, entry.row));
    const current = store.cell(entry.tab, entry.row, entry.col) ?? {};
    const { type: _type, display: _display, ...rest } = current;
    void _type;
    void _display;
    await store.setDerived(
      entry.tab,
      entry.row,
      entry.col,
      {
        ...rest,
        value: result.value,
        ...(result.error ? { type: "error" as const } : {}),
      },
      options.transient,
    );
  };
  const breathe = breather();
  for (const entry of cyclic) await write(entry, { value: "#CIRC!", error: true });
  let done = 0;
  for (const entry of order) {
    await write(entry, await calculateTyped(store, entry.tab, entry.row, entry.col, entry.formula));
    onProgress?.(++done, order.length, "calc");
    await breathe();
  }
  return order.length + cyclic.length;
}

/** Functions whose result changes every time (Excel recalculates them whenever a workbook is opened). */
const VOLATILE = /\b(?:NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY)\s*\(/i;
export const isVolatile = (formula: string): boolean => VOLATILE.test(formula);

/**
 * Gives the volatile formulas (NOW, TODAY, RAND...) of a loaded chunk fresh results, kept in memory only
 * so that just opening a sheet does not make it "changed". Formulas that read them are refreshed the
 * next time something is edited. Returns how many were recalculated.
 */
export async function recalcVolatileInChunk(store: BigSheetStore, tab: number, chunk: number, limit = 5000): Promise<number> {
  const found: [number, number, string][] = [];
  for (const [row, col, cell] of store.cellsOfChunk(tab, chunk)) {
    if (cell.formula && isVolatile(cell.formula)) found.push([row, col, cell.formula]);
    if (found.length >= limit) break;
  }
  for (const [row, col, formula] of found) {
    const result = await calculateTyped(store, tab, row, col, formula);
    const current = store.cell(tab, row, col) ?? {};
    const { type: _type, display: _display, ...rest } = current;
    void _type;
    void _display;
    await store.setDerived(tab, row, col, { ...rest, value: result.value, ...(result.error ? { type: "error" as const } : {}) }, true);
  }
  return found.length;
}
