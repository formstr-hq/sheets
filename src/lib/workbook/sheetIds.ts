// Keeps each worksheet's row and column ids in step with what the editor does.
//
// FortuneSheet reports every row/column insert and delete as an operation. The tracker applies the
// same operation to the id lists, so a row keeps its id when rows above it are inserted or deleted.
// On save the ids are attached to the sheets that are written out.
import {
  deleteRuns,
  insertRuns,
  mintRows,
  newDeviceTag,
  nextCounter,
  padCols,
  padRuns,
  rowCountOf,
  type RowRun,
} from "./ids";
import type { WorkbookSheet } from "./types";

/** This device's id tag and a counter that never goes backwards, even across page loads. */
export interface DeviceIds {
  tag: string;
  /** Reserves `count` numbers and returns the first; at least `atLeast`. */
  reserve(count: number, atLeast: number): number;
}

const STORAGE_KEY = "formstr.sheets.deviceIds";

/** The tag and counter kept in localStorage (memory only when storage is unavailable). */
export function localDeviceIds(storage: Pick<Storage, "getItem" | "setItem"> | undefined = safeStorage()): DeviceIds {
  let state: { tag: string; next: number } | undefined;
  try {
    const raw = storage?.getItem(STORAGE_KEY);
    const parsed = raw ? (JSON.parse(raw) as { tag?: unknown; next?: unknown }) : undefined;
    if (parsed && typeof parsed.tag === "string" && /^[0-9a-z]{4}$/.test(parsed.tag) && typeof parsed.next === "number") {
      state = { tag: parsed.tag, next: parsed.next };
    }
  } catch {
    // fall through to a fresh tag
  }
  const current = state ?? { tag: newDeviceTag(), next: 0 };
  const save = () => {
    try {
      storage?.setItem(STORAGE_KEY, JSON.stringify(current));
    } catch {
      // keep going in memory
    }
  };
  save();
  return {
    tag: current.tag,
    reserve(count, atLeast) {
      const first = Math.max(current.next, atLeast);
      current.next = first + count;
      save();
      return first;
    },
  };
}

function safeStorage(): Storage | undefined {
  try {
    return typeof localStorage === "undefined" ? undefined : localStorage;
  } catch {
    return undefined;
  }
}

/** The two operations of FortuneSheet's `onOp` that change rows or columns. */
export type RowColOp =
  | {
      op: "insertRowCol";
      id?: string;
      value: { id?: string; type: "row" | "column"; index: number; count: number; direction?: "lefttop" | "rightbottom" };
    }
  | { op: "deleteRowCol"; id?: string; value: { id?: string; type: "row" | "column"; start: number; end: number } };

/** The runs, extended to at least `total` rows (padRuns alone would also cut a longer list short). */
const reach = (runs: RowRun[], total: number): RowRun[] => (rowCountOf(runs) >= total ? runs : padRuns(runs, total));

interface SheetIds {
  rows: RowRun[];
  cols: string[];
}

export class SheetIdTracker {
  private readonly sheets = new Map<string, SheetIds>();

  constructor(private readonly device: DeviceIds) {}

  /**
   * Starts from the ids a loaded document carries. A sheet that predates ids gets legacy ids for
   * every row and column it has now, so an id is never handed to a different row later.
   */
  seed(sheets: readonly Pick<WorkbookSheet, "id" | "rowIds" | "colIds" | "rowCount" | "colCount">[]): void {
    this.sheets.clear();
    for (const sheet of sheets) {
      this.sheets.set(sheet.id, {
        rows: padRuns(sheet.rowIds ?? [], sheet.rowCount),
        cols: padCols(sheet.colIds, sheet.colCount),
      });
    }
  }

  /** Applies one editor operation; anything that is not a row/column insert or delete is ignored. */
  apply(operation: { op: string; id?: string; value?: unknown }): void {
    if (operation.op !== "insertRowCol" && operation.op !== "deleteRowCol") return;
    const op = operation as RowColOp;
    const sheetId = op.value.id ?? op.id;
    if (!sheetId) return;
    const ids = this.sheets.get(sheetId) ?? { rows: [], cols: [] };
    if (op.op === "insertRowCol") {
      const { type, index, count, direction } = op.value;
      if (count <= 0) return;
      const at = direction === "rightbottom" ? index + 1 : index;
      if (type === "row") {
        const start = this.device.reserve(count, nextCounter(ids.rows, this.device.tag));
        ids.rows = insertRuns(reach(ids.rows, at), at, mintRows(this.device.tag, start, count));
      } else {
        const start = this.device.reserve(count, 0);
        const minted = Array.from({ length: count }, (_, k) => `${this.device.tag}.${start + k}`);
        const cols = padCols(ids.cols, at);
        cols.splice(at, 0, ...minted);
        ids.cols = cols;
      }
    } else {
      const { type, start, end } = op.value;
      const count = end - start + 1;
      if (count <= 0) return;
      if (type === "row") {
        ids.rows = deleteRuns(reach(ids.rows, end + 1), start, count);
      } else {
        const cols = padCols(ids.cols, end + 1);
        cols.splice(start, count);
        ids.cols = cols;
      }
    }
    this.sheets.set(sheetId, ids);
  }

  /** The sheets with the tracked ids attached (sheets the tracker knows nothing about are untouched). */
  attach<T extends WorkbookSheet>(sheets: readonly T[]): T[] {
    return sheets.map((sheet) => {
      const ids = this.sheets.get(sheet.id);
      if (!ids) return sheet;
      return { ...sheet, rowIds: padRuns(ids.rows, sheet.rowCount), colIds: padCols(ids.cols, sheet.colCount).slice(0, sheet.colCount) };
    });
  }
}
