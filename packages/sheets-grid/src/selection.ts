// Selection and keyboard movement, with no DOM in it.
import type { SizeIndex } from "./sizes";

export interface Cell {
  row: number;
  col: number;
}

export interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** One or more ranges; `active` is the cell typing goes to, `anchor` the fixed corner of the last range. */
export interface Selection {
  ranges: Range[];
  active: Cell;
  anchor: Cell;
}

export interface Bounds {
  rows: SizeIndex;
  cols: SizeIndex;
}

/** True when the cell holds something (used by Ctrl+arrow to find the edge of a block of data). */
export type HasData = (row: number, col: number) => boolean;

export const normalize = (a: Cell, b: Cell): Range => ({
  r1: Math.min(a.row, b.row),
  r2: Math.max(a.row, b.row),
  c1: Math.min(a.col, b.col),
  c2: Math.max(a.col, b.col),
});

export const single = (cell: Cell): Selection => ({ ranges: [normalize(cell, cell)], active: cell, anchor: cell });

export const contains = (range: Range, row: number, col: number): boolean =>
  row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2;

export const isSelected = (selection: Selection, row: number, col: number): boolean =>
  selection.ranges.some((range) => contains(range, row, col));

/** Click: a new single cell, an extension of the last range (Shift), or an extra range (Ctrl/Cmd). */
export function click(selection: Selection | null, cell: Cell, mode: "replace" | "extend" | "add"): Selection {
  if (!selection || mode === "replace") return single(cell);
  if (mode === "extend") {
    const last = normalize(selection.anchor, cell);
    return { ranges: [...selection.ranges.slice(0, -1), last], active: selection.active, anchor: selection.anchor };
  }
  return { ranges: [...selection.ranges, normalize(cell, cell)], active: cell, anchor: cell };
}

/** Whole rows (header click) or columns; the other axis spans everything. */
export function wholeRow(bounds: Bounds, from: number, to = from): Selection {
  const range = { r1: Math.min(from, to), r2: Math.max(from, to), c1: 0, c2: bounds.cols.length - 1 };
  return { ranges: [range], active: { row: range.r1, col: 0 }, anchor: { row: from, col: 0 } };
}

export function wholeCol(bounds: Bounds, from: number, to = from): Selection {
  const range = { c1: Math.min(from, to), c2: Math.max(from, to), r1: 0, r2: bounds.rows.length - 1 };
  return { ranges: [range], active: { row: 0, col: range.c1 }, anchor: { row: 0, col: from } };
}

export function selectAll(bounds: Bounds): Selection {
  return {
    ranges: [{ r1: 0, c1: 0, r2: bounds.rows.length - 1, c2: bounds.cols.length - 1 }],
    active: { row: 0, col: 0 },
    anchor: { row: 0, col: 0 },
  };
}

export type Direction = "up" | "down" | "left" | "right";

const STEP: Record<Direction, [number, number]> = { up: [-1, 0], down: [1, 0], left: [0, -1], right: [0, 1] };

/** Moves one cell, skipping hidden rows/columns and stopping at the sheet's edge. */
function step(bounds: Bounds, from: Cell, dir: Direction): Cell {
  const [dr, dc] = STEP[dir];
  let { row, col } = from;
  if (dr) {
    for (let r = row + dr; r >= 0 && r < bounds.rows.length; r += dr) {
      if (!bounds.rows.isHidden(r)) {
        row = r;
        break;
      }
    }
  } else {
    for (let c = col + dc; c >= 0 && c < bounds.cols.length; c += dc) {
      if (!bounds.cols.isHidden(c)) {
        col = c;
        break;
      }
    }
  }
  return { row, col };
}

/** Longest run Ctrl+arrow scans before giving up (a sheet can have millions of empty rows). */
const JUMP_LIMIT = 2_000_000;

/**
 * Ctrl+arrow, as in Excel: from a cell with data whose neighbour also has data, go to the last
 * cell of that block; otherwise go to the next cell that has data; if none, the sheet's edge.
 */
export function jump(bounds: Bounds, from: Cell, dir: Direction, hasData: HasData): Cell {
  const first = step(bounds, from, dir);
  if (first.row === from.row && first.col === from.col) return from;
  const dataHere = hasData(from.row, from.col);
  const dataNext = hasData(first.row, first.col);
  let at = first;
  let scanned = 0;
  if (dataHere && dataNext) {
    // Walk to the end of the block.
    for (;;) {
      const next = step(bounds, at, dir);
      if ((next.row === at.row && next.col === at.col) || !hasData(next.row, next.col)) return at;
      at = next;
      if (++scanned > JUMP_LIMIT) return at;
    }
  }
  // Walk to the next cell that has data.
  while (!hasData(at.row, at.col)) {
    const next = step(bounds, at, dir);
    if (next.row === at.row && next.col === at.col) return at;
    at = next;
    if (++scanned > JUMP_LIMIT) return at;
  }
  return at;
}

export interface MoveOptions {
  /** Shift: extend the last range instead of moving. */
  extend?: boolean;
  /** Ctrl/Cmd: jump to the edge of the data. */
  jump?: boolean;
  hasData?: HasData;
}

/** Arrow-key movement. With `extend` the active cell stays and the range grows from the anchor. */
export function move(bounds: Bounds, selection: Selection, dir: Direction, options: MoveOptions = {}): Selection {
  const hasData = options.hasData ?? (() => false);
  if (options.extend) {
    const last = selection.ranges[selection.ranges.length - 1];
    // The moving corner is the one opposite the anchor.
    const corner: Cell = {
      row: selection.anchor.row === last.r1 ? last.r2 : last.r1,
      col: selection.anchor.col === last.c1 ? last.c2 : last.c1,
    };
    const next = options.jump ? jump(bounds, corner, dir, hasData) : step(bounds, corner, dir);
    return { ...selection, ranges: [...selection.ranges.slice(0, -1), normalize(selection.anchor, next)] };
  }
  const next = options.jump ? jump(bounds, selection.active, dir, hasData) : step(bounds, selection.active, dir);
  return single(next);
}

/** Tab / Enter: move within the selected range when it is more than one cell, wrapping at its edges. */
export function advance(bounds: Bounds, selection: Selection, dir: "next" | "previous" | "down" | "up"): Selection {
  const range = selection.ranges[selection.ranges.length - 1];
  const multi = range.r1 !== range.r2 || range.c1 !== range.c2;
  const horizontal = dir === "next" || dir === "previous";
  const sign = dir === "next" || dir === "down" ? 1 : -1;
  if (!multi) return move(bounds, selection, horizontal ? (sign > 0 ? "right" : "left") : sign > 0 ? "down" : "up");
  let { row, col } = selection.active;
  const visibleIn = (axis: SizeIndex, from: number, to: number, at: number, delta: number) => {
    for (let i = at + delta; i >= from && i <= to; i += delta) if (!axis.isHidden(i)) return i;
    return null;
  };
  if (horizontal) {
    const c = visibleIn(bounds.cols, range.c1, range.c2, col, sign);
    if (c !== null) col = c;
    else {
      const r = visibleIn(bounds.rows, range.r1, range.r2, row, sign);
      row = r ?? (sign > 0 ? range.r1 : range.r2);
      col = sign > 0 ? bounds.cols.nextVisible(range.c1, 1) : bounds.cols.nextVisible(range.c2, -1);
    }
  } else {
    const r = visibleIn(bounds.rows, range.r1, range.r2, row, sign);
    if (r !== null) row = r;
    else {
      const c = visibleIn(bounds.cols, range.c1, range.c2, col, sign);
      col = c ?? (sign > 0 ? range.c1 : range.c2);
      row = sign > 0 ? bounds.rows.nextVisible(range.r1, 1) : bounds.rows.nextVisible(range.r2, -1);
    }
  }
  return { ...selection, active: { row, col } };
}
