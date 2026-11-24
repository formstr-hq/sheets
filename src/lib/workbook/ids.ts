// Stable row and column ids (spec: "Rows, columns and chunks").
//
// A row id is `<tag>.<n>`: a 4-character tag chosen per device plus a counter, so two devices never
// mint the same id. A tab's row order is a list of runs `[tag, start, length]`, meaning the ids
// `tag.start`, `tag.start+1`, ... `tag.start+length-1`. A freshly imported or migrated sheet is one run.
//
// Positions change when rows are inserted or deleted; ids do not. Everything that must survive an
// edit (cells, heights, formulas) is attached to an id, so an insert only touches the small run list.

/** `[tag, first counter, how many]`. */
export type RowRun = [tag: string, start: number, length: number];

/** Tag of rows that existed before ids did: their counter is simply their position when migrated. */
export const LEGACY_TAG = "0000";

export function rowCountOf(runs: readonly RowRun[]): number {
  let total = 0;
  for (const run of runs) total += run[2];
  return total;
}

/** Appends a run, merging it into the last one when it continues it. */
function push(out: RowRun[], tag: string, start: number, length: number): void {
  if (length <= 0) return;
  const last = out[out.length - 1];
  if (last && last[0] === tag && last[1] + last[2] === start) last[2] += length;
  else out.push([tag, start, length]);
}

/** Joins run lists, merging runs that continue each other. */
export function concatRuns(...lists: readonly (readonly RowRun[])[]): RowRun[] {
  const out: RowRun[] = [];
  for (const list of lists) for (const [tag, start, length] of list) push(out, tag, start, length);
  return out;
}

/** `count` rows that predate ids, numbered by their position `start`, `start+1`, ... */
export function legacyRows(start: number, count: number): RowRun[] {
  return count > 0 ? [[LEGACY_TAG, start, count]] : [];
}

/** The runs for rows [from, to). Positions past the end of `runs` are ignored. */
export function sliceRuns(runs: readonly RowRun[], from: number, to: number): RowRun[] {
  const out: RowRun[] = [];
  let at = 0;
  for (const [tag, start, length] of runs) {
    const end = at + length;
    if (end > from && at < to) {
      const lo = Math.max(from, at);
      const hi = Math.min(to, end);
      push(out, tag, start + (lo - at), hi - lo);
    }
    at = end;
    if (at >= to) break;
  }
  return out;
}

/**
 * The runs extended to `total` rows. Rows the list does not know about get ids that continue the
 * legacy numbering, so reading an unmigrated sheet and padding a short list agree with each other.
 */
export function padRuns(runs: readonly RowRun[], total: number): RowRun[] {
  const have = rowCountOf(runs);
  if (have >= total) return sliceRuns(runs, 0, total);
  let next = 0;
  for (const [tag, start, length] of runs) if (tag === LEGACY_TAG) next = Math.max(next, start + length);
  return concatRuns(runs, legacyRows(next, total - have));
}

/** The id of the row at `index`, or undefined past the end. */
export function rowIdAt(runs: readonly RowRun[], index: number): string | undefined {
  let at = 0;
  for (const [tag, start, length] of runs) {
    if (index < at + length) return `${tag}.${start + (index - at)}`;
    at += length;
  }
  return undefined;
}

/** The position of a row id, or -1 when it is not in the list. */
export function indexOfRow(runs: readonly RowRun[], id: string): number {
  const dot = id.lastIndexOf(".");
  if (dot < 0) return -1;
  const tag = id.slice(0, dot);
  const n = Number(id.slice(dot + 1));
  let at = 0;
  for (const [runTag, start, length] of runs) {
    if (runTag === tag && n >= start && n < start + length) return at + (n - start);
    at += length;
  }
  return -1;
}

/** Inserts `inserted` so that its first row lands at position `index`. */
export function insertRuns(runs: readonly RowRun[], index: number, inserted: readonly RowRun[]): RowRun[] {
  const total = rowCountOf(runs);
  const at = Math.max(0, Math.min(index, total));
  return concatRuns(sliceRuns(runs, 0, at), inserted, sliceRuns(runs, at, total));
}

/** Removes `count` rows starting at position `index`. */
export function deleteRuns(runs: readonly RowRun[], index: number, count: number): RowRun[] {
  const total = rowCountOf(runs);
  return concatRuns(sliceRuns(runs, 0, index), sliceRuns(runs, Math.min(total, index + count), total));
}

/** `count` brand-new rows for this device, numbered from `counter`. */
export function mintRows(tag: string, counter: number, count: number): RowRun[] {
  return count > 0 ? [[tag, counter, count]] : [];
}

/** The next counter this device may use for `tag`: one past anything already in the list. */
export function nextCounter(runs: readonly RowRun[], tag: string): number {
  let next = 0;
  for (const [runTag, start, length] of runs) if (runTag === tag) next = Math.max(next, start + length);
  return next;
}

/** Four random base36 characters; never the legacy tag. */
export function newDeviceTag(random: () => number = Math.random): string {
  for (;;) {
    let tag = "";
    for (let i = 0; i < 4; i++) tag += Math.floor(random() * 36).toString(36);
    if (tag !== LEGACY_TAG) return tag;
  }
}

/** Column id of a column that predates ids: its position when migrated. */
export function legacyColId(index: number): string {
  return `c${index}`;
}

export function legacyCols(count: number): string[] {
  return Array.from({ length: Math.max(0, count) }, (_, index) => legacyColId(index));
}

/** `cols` extended with legacy ids until it has `count` entries. */
export function padCols(cols: readonly string[] | undefined, count: number): string[] {
  const out = [...(cols ?? [])];
  for (let index = out.length; index < count; index++) out.push(legacyColId(index));
  return out;
}

/** Position of a column id; legacy ids ("c12") resolve even when the list does not hold them. */
export function colIndexOf(cols: readonly string[] | undefined, id: string, lookup?: Map<string, number>): number {
  const hit = lookup ? lookup.get(id) : cols?.indexOf(id);
  if (hit !== undefined && hit >= 0) return hit;
  const legacy = /^c(\d+)$/.exec(id);
  return legacy ? Number(legacy[1]) : -1;
}

export function colLookup(cols: readonly string[]): Map<string, number> {
  const map = new Map<string, number>();
  cols.forEach((id, index) => map.set(id, index));
  return map;
}
