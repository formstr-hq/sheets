// Merged cells: where they are and how a selection grows to cover them. No DOM in here.
import type { Range, Selection } from "./selection";

/** A merged block: its top-left cell, and how many rows and columns it spans. */
export interface Merge {
  r: number;
  c: number;
  rs: number;
  cs: number;
}

const rangeOf = (m: Merge): Range => ({ r1: m.r, c1: m.c, r2: m.r + m.rs - 1, c2: m.c + m.cs - 1 });

export class MergeIndex {
  /** Sorted by first row. */
  private list: Merge[] = [];
  private maxRows = 1;

  get size(): number {
    return this.list.length;
  }

  set(merges: readonly Merge[]): void {
    this.list = merges.filter((m) => m.rs >= 1 && m.cs >= 1 && m.rs * m.cs > 1).sort((a, b) => a.r - b.r || a.c - b.c);
    this.maxRows = this.list.reduce((most, m) => Math.max(most, m.rs), 1);
  }

  all(): readonly Merge[] {
    return this.list;
  }

  /** The merged blocks that overlap a range of cells. */
  intersecting(r1: number, c1: number, r2: number, c2: number): Merge[] {
    if (this.list.length === 0) return [];
    // First block that could start early enough to reach r1.
    let lo = 0;
    let hi = this.list.length;
    const from = r1 - this.maxRows + 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.list[mid].r < from) lo = mid + 1;
      else hi = mid;
    }
    const found: Merge[] = [];
    for (let i = lo; i < this.list.length && this.list[i].r <= r2; i++) {
      const m = this.list[i];
      if (m.r + m.rs - 1 >= r1 && m.c <= c2 && m.c + m.cs - 1 >= c1) found.push(m);
    }
    return found;
  }

  /** The block a cell belongs to (its top-left or any covered cell). */
  at(row: number, col: number): Merge | undefined {
    return this.intersecting(row, col, row, col)[0];
  }

  /** The range grown until it cuts no merged block in half. */
  expand(range: Range): Range {
    if (this.list.length === 0) return range;
    let { r1, c1, r2, c2 } = range;
    for (let guard = 0; guard < 50; guard++) {
      let changed = false;
      for (const m of this.intersecting(r1, c1, r2, c2)) {
        const box = rangeOf(m);
        if (box.r1 < r1) (r1 = box.r1), (changed = true);
        if (box.c1 < c1) (c1 = box.c1), (changed = true);
        if (box.r2 > r2) (r2 = box.r2), (changed = true);
        if (box.c2 > c2) (c2 = box.c2), (changed = true);
      }
      if (!changed) break;
    }
    return { r1, c1, r2, c2 };
  }

  /** A selection whose ranges cover whole merged blocks and whose active cell is a block's top-left. */
  expandSelection(selection: Selection): Selection {
    if (this.list.length === 0) return selection;
    const ranges = selection.ranges.map((range) => this.expand(range));
    const block = this.at(selection.active.row, selection.active.col);
    const active = block ? { row: block.r, col: block.c } : selection.active;
    return { ranges, active, anchor: selection.anchor };
  }
}
