// Copy, paste and fill for big sheets.
//
// Copy gives two things: tab-separated text for the system clipboard (so it pastes into other
// programs) and a "block" kept in memory with the real cells (formulas, formats). Pasting text that
// matches the last copy uses the block, so formulas keep working and their references move with the
// cell, as in Excel. Anything else is read as plain text.
import type { WorkbookCell } from "../workbook";
import { cellDisplay } from "../workbook/format";
import { BigSheetStore, MAX_EDIT_CELLS } from "./bigSheet";

const MAX_COLUMNS = 16384;

export interface Area {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Cells copied from a sheet, remembered with where they came from. */
export interface Block {
  /** Top-left cell it was copied from. */
  row: number;
  col: number;
  rows: number;
  cols: number;
  cells: (WorkbookCell | undefined)[][];
  /** The same cells as tab-separated text (what went to the system clipboard). */
  text: string;
}

// ---- text format ------------------------------------------------------------------------------

const quote = (text: string) => (/[\t\n\r"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text);

/** Rows of cells as tab-separated text; cells holding tabs, line breaks or quotes are quoted like Excel does. */
export function toTsv(rows: string[][]): string {
  return rows.map((row) => row.map(quote).join("\t")).join("\n");
}

/** Reads tab-separated text (with Excel-style quoting) into rows of cells. */
export function fromTsv(text: string): string[][] {
  const source = text.replace(/\r\n?/g, "\n");
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < source.length; i++) {
    const ch = source[i];
    if (quoted) {
      if (ch === '"') {
        if (source[i + 1] === '"') {
          cell += '"';
          i++;
        } else quoted = false;
      } else cell += ch;
    } else if (ch === '"' && cell === "") quoted = true;
    else if (ch === "\t") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else cell += ch;
  }
  if (cell !== "" || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

// ---- copy -------------------------------------------------------------------------------------

const size = (area: Area) => (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1);

function tooBig(area: Area): Error | null {
  return size(area) > MAX_EDIT_CELLS ? new Error(`Select at most ${MAX_EDIT_CELLS.toLocaleString()} cells at once in large sheet mode.`) : null;
}

/** Are the chunks holding these rows in memory? */
export function isLoadedArea(store: BigSheetStore, tab: number, area: Area): boolean {
  for (let chunk = store.chunkFor(tab, area.r1); chunk <= store.chunkFor(tab, area.r2); chunk++) {
    if (!store.hasChunk(tab, chunk)) return false;
  }
  return true;
}

/** A block made of copies of `cells` (rows of cells of `area`). */
function blockFrom(area: Area, cells: (WorkbookCell | undefined)[][]): Block {
  const text = cells.map((row) => row.map((cell) => (cell ? cellDisplay(cell) : "")));
  return { row: area.r1, col: area.c1, rows: cells.length, cols: cells[0]?.length ?? 0, cells, text: toTsv(text) };
}

/** The cells of an area whose chunks are already in memory. */
export function copyBlockNow(store: BigSheetStore, tab: number, area: Area): Block {
  const error = tooBig(area);
  if (error) throw error;
  const cells: (WorkbookCell | undefined)[][] = [];
  for (let r = area.r1; r <= area.r2; r++) {
    const row: (WorkbookCell | undefined)[] = [];
    for (let c = area.c1; c <= area.c2; c++) {
      const cell = store.cell(tab, r, c) ?? undefined;
      row.push(cell ? structuredClone(cell) : undefined);
    }
    cells.push(row);
  }
  return blockFrom(area, cells);
}

/**
 * Like `copyBlockNow`, downloading the chunks first. They are read one at a time, so an area that spans
 * more chunks than the cache holds is still copied whole.
 */
export async function copyBlock(store: BigSheetStore, tab: number, area: Area): Promise<Block> {
  const error = tooBig(area);
  if (error) throw error;
  const cells: (WorkbookCell | undefined)[][] = Array.from({ length: area.r2 - area.r1 + 1 }, () => []);
  const first = store.chunkFor(tab, area.r1);
  const last = store.chunkFor(tab, area.r2);
  for (let chunk = first; chunk <= last; chunk++) {
    await store.load(tab, chunk);
    const entry = store.tabs[tab].chunks[chunk];
    const lo = chunk === first ? area.r1 : entry.rowStart;
    const hi = chunk === last ? area.r2 : entry.rowStart + entry.rowCount - 1;
    for (let r = lo; r <= hi; r++) {
      const row = cells[r - area.r1];
      for (let c = area.c1; c <= area.c2; c++) {
        const cell = store.cell(tab, r, c) ?? undefined;
        row.push(cell ? structuredClone(cell) : undefined);
      }
    }
  }
  return blockFrom(area, cells);
}

// ---- paste and fill ---------------------------------------------------------------------------

/** The formula as if copied by (dr, dc): relative references move, $absolute ones stay. */
async function shiftFormula(formula: string, dr: number, dc: number): Promise<string> {
  if (dr === 0 && dc === 0) return formula;
  const engine = await import("@formstr/sheets-formula");
  try {
    // Formulas are stored without the leading "=", which `print` adds.
    return engine.print(engine.shiftReferences(engine.parse(formula.replace(/^=/, "")), dr, dc)).replace(/^=/, "");
  } catch {
    return formula;
  }
}

/** Where a paste of `rows` x `cols` lands: at the cell, or tiled over the selection when one value is pasted. */
function pasteArea(rows: number, cols: number, row: number, col: number, selection?: Area): Area {
  if (rows === 1 && cols === 1 && selection && (selection.r2 > selection.r1 || selection.c2 > selection.c1)) return selection;
  return { r1: row, c1: col, r2: row + rows - 1, c2: col + cols - 1 };
}

function checkFits(area: Area): void {
  if (area.c2 >= MAX_COLUMNS) throw new Error(`That doesn't fit: a sheet has at most ${MAX_COLUMNS.toLocaleString()} columns.`);
  const error = tooBig(area);
  if (error) throw error;
}

/** Writes `block` over `target` (tiled when the target is bigger), formulas shifted. One undo step. */
async function writeBlock(store: BigSheetStore, tab: number, block: Block, target: Area): Promise<Area> {
  checkFits(target);
  const shifted = new Map<string, string>();
  for (let r = target.r1; r <= target.r2; r++) {
    for (let c = target.c1; c <= target.c2; c++) {
      const sr = (r - target.r1) % block.rows;
      const sc = (c - target.c1) % block.cols;
      const source = block.cells[sr][sc];
      if (source?.formula) shifted.set(`${r}:${c}`, await shiftFormula(source.formula, r - (block.row + sr), c - (block.col + sc)));
    }
  }
  await store.edit(tab, target.r1, target.c1, target.r2, target.c2, (_before, r, c) => {
    const source = block.cells[(r - target.r1) % block.rows][(c - target.c1) % block.cols];
    if (!source) return undefined;
    const cell = structuredClone(source);
    const formula = shifted.get(`${r}:${c}`);
    if (formula !== undefined) cell.formula = formula;
    return cell;
  });
  return target;
}

/** Pastes copied cells with their formats and formulas. Returns the area written. */
export async function pasteBlock(store: BigSheetStore, tab: number, row: number, col: number, block: Block, selection?: Area): Promise<Area> {
  if (block.rows === 0 || block.cols === 0) return { r1: row, c1: col, r2: row, c2: col };
  return writeBlock(store, tab, block, pasteArea(block.rows, block.cols, row, col, selection));
}

/** What typing `text` into a cell makes, keeping the cell's format. */
function cellFromText(text: string, before: WorkbookCell | undefined): WorkbookCell | undefined {
  const trimmed = text.trim();
  if (trimmed.startsWith("=") && trimmed.length > 1) return { ...(before?.style ? { style: before.style } : {}), formula: trimmed.slice(1) };
  return BigSheetStore.parseInput(text, before);
}

/** Pastes plain text (rows of cells). Text starting with "=" becomes a formula. Returns the area written. */
export async function pasteText(store: BigSheetStore, tab: number, row: number, col: number, rows: string[][], selection?: Area): Promise<Area> {
  const width = Math.max(1, ...rows.map((r) => r.length));
  if (rows.length === 0) return { r1: row, c1: col, r2: row, c2: col };
  const target = pasteArea(rows.length, width, row, col, selection);
  checkFits(target);
  await store.edit(tab, target.r1, target.c1, target.r2, target.c2, (before, r, c) =>
    cellFromText(rows[(r - target.r1) % rows.length][(c - target.c1) % width] ?? "", before),
  );
  return target;
}

/**
 * Ctrl+D / Ctrl+R: copies the first row (or column) of the selection over the rest, formulas
 * shifted. With one row (column) selected it copies the row above (column to the left).
 * Returns the area written, or null when there is nothing to copy from.
 */
export async function fillArea(store: BigSheetStore, tab: number, area: Area, direction: "down" | "right"): Promise<Area | null> {
  const down = direction === "down";
  const span = down ? area.r2 - area.r1 : area.c2 - area.c1;
  let source: Area;
  let target: Area;
  if (span > 0) {
    source = down ? { ...area, r2: area.r1 } : { ...area, c2: area.c1 };
    target = down ? { ...area, r1: area.r1 + 1 } : { ...area, c1: area.c1 + 1 };
  } else {
    if ((down ? area.r1 : area.c1) === 0) return null;
    source = down ? { ...area, r1: area.r1 - 1, r2: area.r1 - 1 } : { ...area, c1: area.c1 - 1, c2: area.c1 - 1 };
    target = area;
  }
  checkFits(target);
  const block = await copyBlock(store, tab, source);
  return writeBlock(store, tab, block, target);
}

// ---- fill handle ------------------------------------------------------------------------------

export type SeriesValue = number | string | undefined;

/**
 * The next `count` values of a series the cells start, or null when they are not one (then the fill
 * just repeats them). Numbers continue an arithmetic series (2 cells or more with the same step), and
 * text ending in a number ("Item 1", "Q2") counts up, keeping its zero padding ("001" -> "002").
 */
export function extendSeries(source: unknown[], count: number): SeriesValue[] | null {
  const n = source.length;
  if (n === 0 || source.some((v) => v === undefined || v === "")) return null;
  const same = (a: number, b: number) => Math.abs(a - b) < 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  if (source.every((v) => typeof v === "number")) {
    const nums = source as number[];
    if (n < 2) return null;
    const step = nums[1] - nums[0];
    for (let i = 2; i < n; i++) if (!same(nums[i] - nums[i - 1], step)) return null;
    return Array.from({ length: count }, (_, i) => Number((nums[n - 1] + step * (i + 1)).toPrecision(12)));
  }
  if (source.every((v) => typeof v === "string")) {
    const parts = (source as string[]).map((text) => /^(.*?)(\d+)$/.exec(text));
    if (parts.some((p) => !p)) return null;
    const found = parts as RegExpExecArray[];
    const prefix = found[0][1];
    if (found.some((p) => p[1] !== prefix)) return null;
    const nums = found.map((p) => Number(p[2]));
    const step = n === 1 ? 1 : nums[1] - nums[0];
    for (let i = 2; i < n; i++) if (nums[i] - nums[i - 1] !== step) return null;
    const width = found[n - 1][2].length;
    const padded = found[n - 1][2].startsWith("0") && width > 1;
    return Array.from({ length: count }, (_, i) => {
      const next = nums[n - 1] + step * (i + 1);
      if (next < 0) return prefix + String(next);
      return prefix + (padded ? String(next).padStart(width, "0") : String(next));
    });
  }
  return null;
}

/**
 * Fills out from `source` to `extended` (which holds `source` and goes further down or to the right),
 * as dragging the fill handle does: numbers and numbered text continue their series, formulas are
 * copied with their references moved, everything else is repeated. One undo step.
 * Returns the area written.
 */
export async function dragFill(store: BigSheetStore, tab: number, source: Area, extended: Area): Promise<Area | null> {
  const down = extended.r2 > source.r2;
  const right = extended.c2 > source.c2;
  if (!down && !right) return null;
  const target: Area = down ? { ...extended, r1: source.r2 + 1, c1: source.c1, c2: source.c2 } : { ...extended, c1: source.c2 + 1, r1: source.r1, r2: source.r2 };
  checkFits(target);
  const block = await copyBlock(store, tab, source);
  const along = down ? block.rows : block.cols;
  const lines = down ? block.cols : block.rows;
  const count = down ? target.r2 - target.r1 + 1 : target.c2 - target.c1 + 1;

  // For each line (column when filling down): the series, if the cells make one.
  const series: (SeriesValue[] | null)[] = [];
  for (let line = 0; line < lines; line++) {
    const cells = Array.from({ length: along }, (_, k) => (down ? block.cells[k][line] : block.cells[line][k]));
    if (cells.some((cell) => cell?.formula)) {
      series.push(null);
      continue;
    }
    series.push(extendSeries(cells.map((cell) => cell?.value), count));
  }
  const shifted = new Map<string, string>();
  for (let r = target.r1; r <= target.r2; r++) {
    for (let c = target.c1; c <= target.c2; c++) {
      const line = down ? c - source.c1 : r - source.r1;
      const offset = down ? r - source.r1 : c - source.c1;
      const k = offset % along;
      const from = down ? block.cells[k][line] : block.cells[line][k];
      if (!series[line] && from?.formula) {
        const dr = down ? r - (source.r1 + k) : 0;
        const dc = down ? 0 : c - (source.c1 + k);
        shifted.set(`${r}:${c}`, await shiftFormula(from.formula, dr, dc));
      }
    }
  }
  await store.edit(tab, target.r1, target.c1, target.r2, target.c2, (_before, r, c) => {
    const line = down ? c - source.c1 : r - source.r1;
    const offset = down ? r - source.r1 : c - source.c1;
    const k = offset % along;
    const from = down ? block.cells[k][line] : block.cells[line][k];
    const values = series[line];
    if (values) {
      const last = down ? block.cells[along - 1][line] : block.cells[line][along - 1];
      const { formula: _formula, display: _display, ...rest } = structuredClone(last ?? {});
      void _formula;
      void _display;
      return { ...rest, value: values[(down ? r - target.r1 : c - target.c1)] as WorkbookCell["value"] };
    }
    if (!from) return undefined;
    const cell = structuredClone(from);
    const formula = shifted.get(`${r}:${c}`);
    if (formula !== undefined) cell.formula = formula;
    return cell;
  });
  return target;
}

/**
 * Moves the cells of `source` so that its top-left lands on `target`'s top-left. The cells keep their
 * formulas as typed (references are not changed) and whatever was under them is replaced.
 * One undo step when the two areas fit one edit, two otherwise. Returns the area they now occupy.
 */
export async function moveCells(store: BigSheetStore, tab: number, source: Area, target: Area): Promise<Area | null> {
  const dr = target.r1 - source.r1;
  const dc = target.c1 - source.c1;
  if (dr === 0 && dc === 0) return null;
  const moved: Area = { r1: source.r1 + dr, c1: source.c1 + dc, r2: source.r2 + dr, c2: source.c2 + dc };
  checkFits(moved);
  const block = await copyBlock(store, tab, source);
  const inside = (area: Area, r: number, c: number) => r >= area.r1 && r <= area.r2 && c >= area.c1 && c <= area.c2;
  const union: Area = {
    r1: Math.min(source.r1, moved.r1),
    c1: Math.min(source.c1, moved.c1),
    r2: Math.max(source.r2, moved.r2),
    c2: Math.max(source.c2, moved.c2),
  };
  if ((union.r2 - union.r1 + 1) * (union.c2 - union.c1 + 1) <= MAX_EDIT_CELLS) {
    // Rows between the two areas are loaded too, which for a far move can be many parts: ask first.
    await store.edit(tab, union.r1, union.c1, union.r2, union.c2, (before, r, c) => {
      if (inside(moved, r, c)) {
        const cell = block.cells[r - moved.r1][c - moved.c1];
        return cell ? structuredClone(cell) : undefined;
      }
      return inside(source, r, c) ? undefined : before;
    });
  } else {
    await store.edit(tab, source.r1, source.c1, source.r2, source.c2, () => undefined);
    await store.edit(tab, moved.r1, moved.c1, moved.r2, moved.c2, (_before, r, c) => {
      const cell = block.cells[r - moved.r1][c - moved.c1];
      return cell ? structuredClone(cell) : undefined;
    });
  }
  return moved;
}
