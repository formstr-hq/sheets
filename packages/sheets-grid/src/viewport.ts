// Which rows and columns are on screen for a scroll position.
//
// Browsers cap how tall a scrolling element can be (about 17M px in Firefox), so a 25M-row sheet
// cannot be scrolled 1:1. Past `maxScrollPx` the scrollbar covers a shorter range and each
// scrollbar pixel moves several sheet pixels (`scale` > 1). The first and last rows stay reachable.
import type { SizeIndex } from "./sizes";

export const MAX_SCROLL_PX = 10_000_000;

export interface Axis {
  /** Size of the scrollable area in sheet pixels. */
  total: number;
  /** Height/width of the window the grid shows (without headers). */
  window: number;
}

/** Real scrollbar length for `total` sheet pixels. */
export function scrollExtent(total: number, windowPx: number, maxScrollPx = MAX_SCROLL_PX): number {
  return Math.min(total, maxScrollPx + windowPx);
}

/** Sheet pixels per scrollbar pixel (1 until the sheet is too big to scroll 1:1). */
export function scrollScale(total: number, windowPx: number, maxScrollPx = MAX_SCROLL_PX): number {
  const extent = scrollExtent(total, windowPx, maxScrollPx);
  const travel = extent - windowPx;
  const need = total - windowPx;
  return travel > 0 && need > travel ? need / travel : 1;
}

/** Sheet position (px) shown at the window's start for a scrollbar position. */
export function toSheetPx(scroll: number, total: number, windowPx: number, maxScrollPx = MAX_SCROLL_PX): number {
  return scroll * scrollScale(total, windowPx, maxScrollPx);
}

/** Scrollbar position that puts sheet position `px` at the window's start. */
export function toScrollPx(px: number, total: number, windowPx: number, maxScrollPx = MAX_SCROLL_PX): number {
  return px / scrollScale(total, windowPx, maxScrollPx);
}

export interface VisibleRange {
  first: number;
  /** One past the last visible index. */
  end: number;
  /** Pixels the first index starts above the window's start (0..size of first). */
  shift: number;
}

/** Indexes whose cells show in a window of `windowPx` that starts at sheet position `start`. */
export function visible(sizes: SizeIndex, start: number, windowPx: number): VisibleRange {
  if (sizes.length === 0) return { first: 0, end: 0, shift: 0 };
  const pos = Math.max(0, Math.min(start, Math.max(0, sizes.total - windowPx)));
  const first = sizes.indexAt(pos);
  let end = first;
  let at = sizes.offset(first);
  while (end < sizes.length && at < pos + windowPx) {
    at += sizes.size(end);
    end++;
  }
  return { first, end, shift: pos - sizes.offset(first) };
}
