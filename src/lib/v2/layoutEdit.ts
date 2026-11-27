// What inserting or deleting rows or columns does to a tab's saved look (sizes, hidden rows, merges,
// borders). Pure functions; the store applies the result.
import type { TabLayout } from "./bigSheet";

export interface StructEdit {
  kind: "insert" | "delete";
  axis: "row" | "col";
  /** First row/column affected (0-based). Inserted ones go before it. */
  index: number;
  count: number;
}

/** Where a position lands, or null when it was deleted. */
export function moveIndex(value: number, edit: StructEdit): number | null {
  if (edit.kind === "insert") return value >= edit.index ? value + edit.count : value;
  if (value < edit.index) return value;
  if (value >= edit.index + edit.count) return value - edit.count;
  return null;
}

/** A span's [first, last] after the edit; null when all of it was deleted. An insert inside a span grows it. */
export function moveSpan(first: number, last: number, edit: StructEdit): [number, number] | null {
  if (edit.kind === "insert") {
    return [first >= edit.index ? first + edit.count : first, last >= edit.index ? last + edit.count : last];
  }
  const end = edit.index + edit.count - 1;
  const newFirst = first < edit.index ? first : first > end ? first - edit.count : edit.index;
  const newLast = last < edit.index ? last : last > end ? last - edit.count : edit.index - 1;
  return newLast < newFirst ? null : [newFirst, newLast];
}

function moveKeys(record: Record<number, number> | undefined, edit: StructEdit): Record<number, number> | undefined {
  if (!record) return undefined;
  const out: Record<number, number> = {};
  for (const [key, size] of Object.entries(record)) {
    const at = moveIndex(Number(key), edit);
    if (at !== null) out[at] = size;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

function moveList(list: number[] | undefined, edit: StructEdit): number[] | undefined {
  if (!list) return undefined;
  const out = list.flatMap((value) => {
    const at = moveIndex(value, edit);
    return at === null ? [] : [at];
  });
  return out.length > 0 ? out : undefined;
}

/** The layout after the edit. Keys that end up empty are left out (so they are removed when saved). */
export function editLayout(layout: TabLayout, edit: StructEdit): TabLayout {
  const row = edit.axis === "row";
  const out: TabLayout = {};
  const widths = row ? layout.columnWidths : moveKeys(layout.columnWidths, edit);
  const heights = row ? moveKeys(layout.rowHeights, edit) : layout.rowHeights;
  const hiddenRows = row ? moveList(layout.hiddenRows, edit) : layout.hiddenRows;
  const hiddenColumns = row ? layout.hiddenColumns : moveList(layout.hiddenColumns, edit);
  if (widths) out.columnWidths = widths;
  if (heights) out.rowHeights = heights;
  if (hiddenRows?.length) out.hiddenRows = hiddenRows;
  if (hiddenColumns?.length) out.hiddenColumns = hiddenColumns;
  if (layout.frozen) out.frozen = layout.frozen;
  if (layout.zoom) out.zoom = layout.zoom;

  const merges = (layout.merges ?? []).flatMap((m) => {
    const span = row ? moveSpan(m.r, m.r + m.rs - 1, edit) : moveSpan(m.c, m.c + m.cs - 1, edit);
    if (!span) return [];
    const next = row ? { ...m, r: span[0], rs: span[1] - span[0] + 1 } : { ...m, c: span[0], cs: span[1] - span[0] + 1 };
    return next.rs * next.cs > 1 ? [next] : [];
  });
  if (merges.length > 0) out.merges = merges;

  const borders = (layout.borders ?? []).flatMap((b) => {
    const span = row ? moveSpan(b.r1, b.r2, edit) : moveSpan(b.c1, b.c2, edit);
    if (!span) return [];
    return [row ? { ...b, r1: span[0], r2: span[1] } : { ...b, c1: span[0], c2: span[1] }];
  });
  if (borders.length > 0) out.borders = borders;

  const rules = (layout.conditionalFormats ?? []).flatMap((rule) => {
    const ranges = rule.ranges.flatMap((r) => {
      const span = row ? moveSpan(r.r1, r.r2, edit) : moveSpan(r.c1, r.c2, edit);
      if (!span) return [];
      return [row ? { ...r, r1: span[0], r2: span[1] } : { ...r, c1: span[0], c2: span[1] }];
    });
    return ranges.length > 0 ? [{ ...rule, ranges }] : [];
  });
  if (rules.length > 0) out.conditionalFormats = rules;

  const images = (layout.images ?? []).flatMap((image) => {
    const at = moveIndex(row ? image.row : image.col, edit);
    return at === null ? [] : [row ? { ...image, row: at } : { ...image, col: at }];
  });
  if (images.length > 0) out.images = images;

  if (layout.filter) {
    const col = row ? layout.filter.col : moveIndex(layout.filter.col, edit);
    if (col !== null) {
      const hidden = row ? (moveList(layout.filter.hidden, edit) ?? []) : layout.filter.hidden;
      out.filter = { ...layout.filter, col, hidden };
    }
  }
  return out;
}
