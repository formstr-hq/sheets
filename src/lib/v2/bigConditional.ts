// Conditional formatting for big sheets: rules that colour a cell when its value matches.
// A rule is stored with the tab's look (see TabLayout.conditionalFormats) and checked while drawing,
// so it works on any size of sheet without touching the cells.
import type { WorkbookCell, WorkbookConditionalRule } from "../workbook";
import { cellDisplay } from "../workbook/format";
import type { BigSheetStore } from "./bigSheet";

export const CONDITIONS = [
  { value: "gt", label: "Greater than", values: 1 },
  { value: "gte", label: "Greater than or equal to", values: 1 },
  { value: "lt", label: "Less than", values: 1 },
  { value: "lte", label: "Less than or equal to", values: 1 },
  { value: "eq", label: "Equal to", values: 1 },
  { value: "neq", label: "Not equal to", values: 1 },
  { value: "between", label: "Between", values: 2 },
  { value: "contains", label: "Text contains", values: 1 },
  { value: "empty", label: "Is empty", values: 0 },
  { value: "notEmpty", label: "Is not empty", values: 0 },
  { value: "dataBar", label: "Data bar (length by value)", values: 0 },
  { value: "colorScale", label: "Colour scale (low to high)", values: 0 },
] as const;

/** Rules that colour by where the value lies between the lowest and highest of the range. */
export const SCALE_TYPES = ["dataBar", "colorScale"];

/** What a rule does to one cell: colours, and the length of a data bar (0 to 1). */
export interface Effect {
  textColor?: string;
  cellColor?: string;
  bar?: { fraction: number; color: string };
}

const hex = (color: string): [number, number, number] => {
  const m = /^#?([0-9a-f]{6})$/i.exec(color.trim());
  const n = m ? parseInt(m[1], 16) : 0;
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
};

/** The colour `t` of the way (0 to 1) from `from` to `to`, as #rrggbb. */
export function blend(from: string, to: string, t: number): string {
  const a = hex(from);
  const b = hex(to);
  const mix = a.map((v, i) => Math.round(v + (b[i] - v) * Math.min(1, Math.max(0, t))));
  return `#${mix.map((v) => v.toString(16).padStart(2, "0")).join("")}`;
}

/** Does the cell satisfy the rule's condition? */
export function matchesCondition(condition: WorkbookConditionalRule["condition"], cell: WorkbookCell | null | undefined): boolean {
  const empty = !cell || cell.value === undefined || cell.value === "";
  if (condition.type === "empty") return empty;
  if (condition.type === "notEmpty") return !empty;
  if (empty) return false;
  if (SCALE_TYPES.includes(condition.type)) return typeof cell!.value === "number" && Number.isFinite(cell!.value);
  const text = cellDisplay(cell!);
  const [first = "", second = ""] = condition.values;
  if (condition.type === "contains") return text.toLowerCase().includes(first.toLowerCase());
  const number = typeof cell!.value === "number" ? cell!.value : Number(String(cell!.value).replace(/,/g, ""));
  const a = Number(first);
  const b = Number(second);
  const numeric = Number.isFinite(number) && Number.isFinite(a) && (condition.type !== "between" || Number.isFinite(b));
  switch (condition.type) {
    case "eq":
      return numeric ? number === a : text.toLowerCase() === first.toLowerCase();
    case "neq":
      return numeric ? number !== a : text.toLowerCase() !== first.toLowerCase();
    case "gt":
      return numeric && number > a;
    case "gte":
      return numeric && number >= a;
    case "lt":
      return numeric && number < a;
    case "lte":
      return numeric && number <= a;
    case "between":
      return numeric && number >= Math.min(a, b) && number <= Math.max(a, b);
    default:
      return false;
  }
}

/** The colours the first matching rule gives a cell, or undefined. */
export function formatFor(
  rules: readonly WorkbookConditionalRule[] | undefined,
  row: number,
  col: number,
  cell: WorkbookCell | null | undefined,
): Effect | undefined {
  if (!rules) return undefined;
  for (const rule of rules) {
    if (!rule.ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2)) continue;
    if (!matchesCondition(rule.condition, cell)) continue;
    if (!SCALE_TYPES.includes(rule.condition.type)) return rule.format;
    // Scales and bars compare the value with the lowest and highest numbers of the range, found when the rule was made.
    const [min = 0, max = 0] = rule.condition.values.map(Number);
    const value = cell!.value as number;
    const t = max === min ? 1 : Math.min(1, Math.max(0, (value - min) / (max - min)));
    if (rule.condition.type === "dataBar") return { bar: { fraction: t, color: rule.format.cellColor ?? "#5b9bd5" } };
    return { cellColor: blend(rule.format.cellColor ?? "#ffffff", rule.format.toColor ?? "#63be7b", t) };
  }
  return undefined;
}

/** The lowest and highest number among the cells of an area, reading one chunk at a time. Null when it holds no numbers. */
export async function numericExtent(
  store: BigSheetStore,
  tab: number,
  area: { r1: number; c1: number; r2: number; c2: number },
): Promise<{ min: number; max: number } | null> {
  if ((area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1) > 5_000_000) throw new Error("That range is too big to colour by value.");
  let min = Infinity;
  let max = -Infinity;
  const first = store.chunkFor(tab, area.r1);
  const last = store.chunkFor(tab, area.r2);
  for (let chunk = first; chunk <= last; chunk++) {
    await store.load(tab, chunk);
    for (const [row, col, cell] of store.cellsOfChunk(tab, chunk)) {
      if (row < area.r1 || row > area.r2 || col < area.c1 || col > area.c2) continue;
      if (typeof cell.value === "number" && Number.isFinite(cell.value)) {
        if (cell.value < min) min = cell.value;
        if (cell.value > max) max = cell.value;
      }
    }
  }
  return min === Infinity ? null : { min, max };
}

/** Rules with the part of `area` taken out of their ranges (used when a rule is replaced or removed). */
export function withoutArea(
  rules: readonly WorkbookConditionalRule[],
  area: { r1: number; c1: number; r2: number; c2: number },
): WorkbookConditionalRule[] {
  const out: WorkbookConditionalRule[] = [];
  for (const rule of rules) {
    const ranges: WorkbookConditionalRule["ranges"] = [];
    for (const r of rule.ranges) {
      const overlaps = r.r1 <= area.r2 && r.r2 >= area.r1 && r.c1 <= area.c2 && r.c2 >= area.c1;
      if (!overlaps) {
        ranges.push(r);
        continue;
      }
      // Keep what is left: the strips above, below, left and right of the removed area.
      if (r.r1 < area.r1) ranges.push({ ...r, r2: area.r1 - 1 });
      if (r.r2 > area.r2) ranges.push({ ...r, r1: area.r2 + 1 });
      const r1 = Math.max(r.r1, area.r1);
      const r2 = Math.min(r.r2, area.r2);
      if (r.c1 < area.c1) ranges.push({ r1, r2, c1: r.c1, c2: area.c1 - 1 });
      if (r.c2 > area.c2) ranges.push({ r1, r2, c1: area.c2 + 1, c2: r.c2 });
    }
    if (ranges.length > 0) out.push({ ...rule, ranges });
  }
  return out;
}
