// Sorting a selected block of a big sheet by one column. The selection has to fit in memory
// (MAX_EDIT_CELLS) and may not hold formulas or merged cells, because moving those would need their
// references rewritten (a later step).
import type { WorkbookCell } from "../workbook";
import { cellDisplay } from "../workbook/format";
import { BigSheetStore, MAX_EDIT_CELLS } from "./bigSheet";
import { copyBlock } from "./bigClipboard";

export interface SortArea {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Numbers first (smallest first), then text (ignoring case), empty cells always last. */
function compare(a: WorkbookCell | undefined, b: WorkbookCell | undefined, descending: boolean): number {
  const empty = (cell: WorkbookCell | undefined) => cell === undefined || cell.value === undefined || cell.value === "";
  if (empty(a) || empty(b)) return empty(a) === empty(b) ? 0 : empty(a) ? 1 : -1; // empty stays last either way
  const num = (cell: WorkbookCell | undefined) => (typeof cell?.value === "number" ? cell.value : null);
  const x = num(a);
  const y = num(b);
  let order: number;
  if (x !== null && y !== null) order = x - y;
  else if (x !== null) order = -1;
  else if (y !== null) order = 1;
  else order = cellDisplay(a!).localeCompare(cellDisplay(b!), undefined, { sensitivity: "base", numeric: true });
  return descending ? -order : order;
}

/** Sorts the rows of `area` by column `keyCol`. One undo step. Returns false when nothing needed to move. */
export async function sortArea(
  store: BigSheetStore,
  tab: number,
  area: SortArea,
  keyCol: number,
  descending: boolean,
  merges: readonly { r: number; c: number; rs: number; cs: number }[] = [],
): Promise<boolean> {
  const rows = area.r2 - area.r1 + 1;
  const cols = area.c2 - area.c1 + 1;
  if (rows < 2) return false;
  if (rows * cols > MAX_EDIT_CELLS) {
    throw new Error(`Select at most ${MAX_EDIT_CELLS.toLocaleString()} cells to sort in large sheet mode.`);
  }
  if (keyCol < area.c1 || keyCol > area.c2) throw new Error("Pick a cell inside the selection to sort by.");
  if (merges.some((m) => m.r <= area.r2 && m.r + m.rs - 1 >= area.r1 && m.c <= area.c2 && m.c + m.cs - 1 >= area.c1)) {
    throw new Error("Cells in the selection are merged. Unmerge them to sort.");
  }
  // Read chunk by chunk (a big selection can span more chunks than the cache keeps).
  const copied = await copyBlock(store, tab, area);
  const block = copied.cells;
  if (block.some((row) => row.some((cell) => cell?.formula))) throw new Error("The selection has formulas. Sorting formulas isn't supported yet.");
  const order = block.map((_, index) => index);
  const key = keyCol - area.c1;
  order.sort((a, b) => compare(block[a][key], block[b][key], descending) || a - b);
  if (order.every((from, to) => from === to)) return false;
  await store.edit(tab, area.r1, area.c1, area.r2, area.c2, (_before, row, col) => {
    const source = block[order[row - area.r1]][col - area.c1];
    return source ? structuredClone(source) : undefined;
  });
  return true;
}
