// Find and replace across a whole big sheet. The chunks are opened one after another, so only a few are
// in memory at a time; each one is searched and then left to the cache.
import type { BigSheetStore } from "./bigSheet";
import { cellDisplay } from "../workbook/format";

export interface FindOptions {
  matchCase?: boolean;
  /** The whole cell must equal the search text. */
  wholeCell?: boolean;
}

export interface Match {
  row: number;
  col: number;
}

export interface FindResult {
  matches: Match[];
  /** True when more cells matched than MAX_MATCHES (the rest are not listed). */
  truncated: boolean;
}

export const MAX_MATCHES = 10_000;

function matcher(query: string, options: FindOptions): (text: string) => boolean {
  const needle = options.matchCase ? query : query.toLowerCase();
  return (text) => {
    const hay = options.matchCase ? text : text.toLowerCase();
    return options.wholeCell ? hay === needle : hay.includes(needle);
  };
}

/** Every cell of the tab (in reading order) whose shown text matches. `stop()` ends the search early. */
export async function findCells(
  store: BigSheetStore,
  tab: number,
  query: string,
  options: FindOptions = {},
  onProgress?: (done: number, total: number) => void,
  stop?: () => boolean,
): Promise<FindResult> {
  const result: FindResult = { matches: [], truncated: false };
  if (query === "") return result;
  const test = matcher(query, options);
  const info = store.tabs[tab];
  if (!info) return result;
  for (let chunk = 0; chunk < info.chunks.length; chunk++) {
    if (stop?.()) break;
    await store.load(tab, chunk);
    // The chunk may be pushed out by another load before it is read; then try again.
    let cells: Iterable<[number, number, import("../workbook").WorkbookCell]> | undefined;
    for (let attempt = 0; attempt < 3 && !cells; attempt++) {
      if (store.hasChunk(tab, chunk)) cells = store.cellsOfChunk(tab, chunk);
      else await store.load(tab, chunk);
    }
    const found: Match[] = [];
    for (const [row, col, cell] of cells ?? []) {
      const text = cellDisplay(cell);
      if (text !== "" && test(text)) found.push({ row, col });
    }
    found.sort((a, b) => a.row - b.row || a.col - b.col);
    for (const match of found) {
      if (result.matches.length >= MAX_MATCHES) {
        result.truncated = true;
        return result;
      }
      result.matches.push(match);
    }
    onProgress?.(chunk + 1, info.chunks.length);
    // Let the page breathe between chunks.
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  return result;
}

/** `text` with every occurrence of `query` replaced (a whole-cell match replaces the whole text). */
export function replaceIn(text: string, query: string, replacement: string, options: FindOptions = {}): string {
  if (options.wholeCell) return text;
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return text.replace(new RegExp(escaped, options.matchCase ? "g" : "gi"), () => replacement);
}

/**
 * Replaces in the given cells. Cells holding a formula are left alone (their text is the result, not
 * what was typed). Returns the cells that changed.
 */
export async function replaceCells(
  store: BigSheetStore,
  tab: number,
  matches: readonly Match[],
  query: string,
  replacement: string,
  options: FindOptions = {},
): Promise<Match[]> {
  const changed: Match[] = [];
  for (const { row, col } of matches) {
    await store.load(tab, store.chunkFor(tab, row));
    const cell = store.cell(tab, row, col);
    if (!cell || cell.formula) continue;
    const text = cellDisplay(cell);
    const next = options.wholeCell ? replacement : replaceIn(text, query, replacement, options);
    if (next === text) continue;
    await store.setCell(tab, row, col, next);
    changed.push({ row, col });
  }
  return changed;
}
