// Row heights / column widths for a sheet of any length.
//
// Most rows have the default size, a few have their own (resized or hidden). Finding the pixel
// offset of row 5,000,000 or the row under a pixel must not walk 5M rows, so the exceptions live
// in a sorted list with running totals and everything else is arithmetic.

export interface SizeOverride {
  index: number;
  /** Size in px; 0 hides the row/column. */
  size: number;
}

export class SizeIndex {
  private readonly defaultSize: number;
  private count: number;
  /** Sorted by index. */
  private overrides: SizeOverride[] = [];
  /** extra[i] = sum over overrides[0..i-1] of (size - defaultSize); length = overrides.length + 1. */
  private extra: number[] = [0];

  constructor(count: number, defaultSize: number, overrides: readonly SizeOverride[] = []) {
    this.count = count;
    this.defaultSize = defaultSize;
    this.setOverrides(overrides);
  }

  get length(): number {
    return this.count;
  }

  setLength(count: number): void {
    this.count = count;
    this.setOverrides(this.overrides);
  }

  private setOverrides(list: readonly SizeOverride[]): void {
    const clean = list
      .filter((o) => o.index >= 0 && o.index < this.count && o.size !== this.defaultSize)
      .sort((a, b) => a.index - b.index);
    // Last one wins when an index repeats.
    this.overrides = clean.filter((o, i) => clean[i + 1]?.index !== o.index);
    this.extra = [0];
    for (const o of this.overrides) this.extra.push(this.extra[this.extra.length - 1] + o.size - this.defaultSize);
  }

  /** Replaces every override at once (e.g. when another tab is shown). */
  replace(list: readonly SizeOverride[]): void {
    this.setOverrides(list);
  }

  /** Overrides in index order (for saving). */
  list(): SizeOverride[] {
    return this.overrides.map((o) => ({ ...o }));
  }

  set(index: number, size: number): void {
    const rest = this.overrides.filter((o) => o.index !== index);
    rest.push({ index, size });
    this.setOverrides(rest);
  }

  /** Number of overrides at positions below `index`. */
  private below(index: number): number {
    let lo = 0;
    let hi = this.overrides.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.overrides[mid].index < index) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  size(index: number): number {
    const at = this.below(index);
    const hit = this.overrides[at];
    return hit && hit.index === index ? hit.size : this.defaultSize;
  }

  isHidden(index: number): boolean {
    return this.size(index) === 0;
  }

  /** Pixel offset of the start of `index` (rows above it summed). */
  offset(index: number): number {
    const i = Math.max(0, Math.min(index, this.count));
    return i * this.defaultSize + this.extra[this.below(i)];
  }

  get total(): number {
    return this.offset(this.count);
  }

  /** The row/column that contains pixel `px` (clamped to the first/last visible one). */
  indexAt(px: number): number {
    if (this.count === 0) return 0;
    if (px <= 0) return this.nextVisible(0, 1);
    if (px >= this.total) return this.nextVisible(this.count - 1, -1);
    // offset() is non-decreasing, so binary search the largest i with offset(i) <= px.
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = Math.ceil((lo + hi) / 2);
      if (this.offset(mid) <= px) lo = mid;
      else hi = mid - 1;
    }
    // A hidden row has zero height and shares its offset with the next row; step to a visible one.
    return this.isHidden(lo) ? this.nextVisible(lo, 1) : lo;
  }

  /** `from` itself or the nearest visible index in direction `dir` (stays put when none is visible). */
  nextVisible(from: number, dir: 1 | -1): number {
    for (let i = from; i >= 0 && i < this.count; i += dir) if (!this.isHidden(i)) return i;
    for (let i = from; i >= 0 && i < this.count; i -= dir) if (!this.isHidden(i)) return i;
    return Math.max(0, Math.min(from, this.count - 1));
  }
}
