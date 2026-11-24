// Cell borders. They are kept as the list of "put this border on that range" requests, in order,
// and replayed for the part of the sheet on screen. No DOM in here.
import type { Range } from "./selection";

export type BorderType = "all" | "outside" | "inside" | "top" | "bottom" | "left" | "right" | "none";

export interface Border extends Range {
  type: BorderType;
  color?: string;
  /** "thin" (default), "medium", "thick", "dashed" or "dotted". */
  style?: string;
}

export interface EdgeStyle {
  color: string;
  style: string;
}

/**
 * The lines to draw in a window of cells. `h(row, col)` is the line above that cell, `v(row, col)` the
 * line to its left; a row or column one past the window's end gives the line below / right of it.
 */
export interface Edges {
  h: Map<number, EdgeStyle>;
  v: Map<number, EdgeStyle>;
}

/** Key for an edge: unique for rows below 2^22 * 2^15 which is far past any sheet. */
export const edgeKey = (row: number, col: number): number => row * 32768 + col;

const DEFAULT_COLOR = "#000000";

/** Replays the borders over cells `r1..r2` x `c1..c2` and returns the lines in and around them. */
export function edgesIn(borders: readonly Border[], r1: number, c1: number, r2: number, c2: number): Edges {
  const edges: Edges = { h: new Map(), v: new Map() };
  for (const border of borders) {
    // Lines touching the window are the ones on cells inside it or one past its end.
    if (border.r1 > r2 + 1 || border.r2 < r1 - 1 || border.c1 > c2 + 1 || border.c2 < c1 - 1) continue;
    const erase = border.type === "none";
    const edge: EdgeStyle = { color: border.color ?? DEFAULT_COLOR, style: border.style ?? "thin" };
    const put = (map: Map<number, EdgeStyle>, row: number, col: number) => {
      if (row < r1 || row > r2 + 1 || col < c1 || col > c2 + 1) return;
      if (erase) map.delete(edgeKey(row, col));
      else map.set(edgeKey(row, col), edge);
    };
    const withTop = border.type === "all" || border.type === "outside" || border.type === "top" || erase;
    const withBottom = border.type === "all" || border.type === "outside" || border.type === "bottom" || erase;
    const withLeft = border.type === "all" || border.type === "outside" || border.type === "left" || erase;
    const withRight = border.type === "all" || border.type === "outside" || border.type === "right" || erase;
    const inside = border.type === "all" || border.type === "inside" || erase;

    const fromC = Math.max(border.c1, c1);
    const toC = Math.min(border.c2, c2);
    const fromR = Math.max(border.r1, r1);
    const toR = Math.min(border.r2, r2);
    for (let c = fromC; c <= toC; c++) {
      if (withTop) put(edges.h, border.r1, c);
      if (withBottom) put(edges.h, border.r2 + 1, c);
    }
    for (let r = fromR; r <= toR; r++) {
      if (withLeft) put(edges.v, r, border.c1);
      if (withRight) put(edges.v, r, border.c2 + 1);
    }
    if (inside) {
      for (let r = Math.max(border.r1 + 1, r1); r <= Math.min(border.r2, r2 + 1); r++) {
        for (let c = fromC; c <= toC; c++) put(edges.h, r, c);
      }
      for (let c = Math.max(border.c1 + 1, c1); c <= Math.min(border.c2, c2 + 1); c++) {
        for (let r = fromR; r <= toR; r++) put(edges.v, r, c);
      }
    }
  }
  return edges;
}

/** Adds a border request to the list. A "none" removes the requests it fully covers, so the list does not grow forever. */
export function addBorder(list: readonly Border[], border: Border): Border[] {
  if (border.type !== "none") return [...list, border];
  const covered = (b: Border) => b.r1 >= border.r1 && b.r2 <= border.r2 && b.c1 >= border.c1 && b.c2 <= border.c2;
  const touches = (b: Border) => b.r1 <= border.r2 + 1 && b.r2 >= border.r1 - 1 && b.c1 <= border.c2 + 1 && b.c2 >= border.c1 - 1;
  const kept = list.filter((b) => !covered(b));
  return kept.some(touches) ? [...kept, border] : kept;
}
