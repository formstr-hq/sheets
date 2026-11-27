import { describe, expect, it } from "vitest";
import { advance, click, isSelected, jump, move, selectAll, single, wholeCol, wholeRow } from "./selection";
import { SizeIndex } from "./sizes";
import { scrollExtent, scrollScale, toScrollPx, toSheetPx, visible } from "./viewport";

describe("SizeIndex", () => {
  it("is plain arithmetic without overrides, even for 25M rows", () => {
    const rows = new SizeIndex(25_000_000, 30);
    expect(rows.total).toBe(750_000_000);
    expect(rows.offset(5_000_000)).toBe(150_000_000);
    expect(rows.indexAt(150_000_029)).toBe(5_000_000);
    expect(rows.indexAt(150_000_030)).toBe(5_000_001);
  });

  it("counts resized and hidden rows", () => {
    const rows = new SizeIndex(10, 20, [
      { index: 2, size: 50 },
      { index: 4, size: 0 },
    ]);
    expect(rows.offset(2)).toBe(40);
    expect(rows.offset(3)).toBe(90);
    expect(rows.offset(5)).toBe(90 + 20 + 0);
    expect(rows.total).toBe(10 * 20 + 30 - 20);
    expect(rows.indexAt(45)).toBe(2);
    expect(rows.indexAt(90)).toBe(3);
    // The hidden row has no pixels, so a pixel just past row 3 is row 5.
    expect(rows.indexAt(110)).toBe(5);
    expect(rows.isHidden(4)).toBe(true);
  });

  it("set() replaces an override and ignores values equal to the default", () => {
    const rows = new SizeIndex(5, 10);
    rows.set(1, 30);
    rows.set(1, 10);
    expect(rows.list()).toEqual([]);
    expect(rows.total).toBe(50);
  });

  it("clamps pixels outside the sheet", () => {
    const rows = new SizeIndex(5, 10);
    expect(rows.indexAt(-5)).toBe(0);
    expect(rows.indexAt(9999)).toBe(4);
  });
});

describe("viewport", () => {
  it("lists the rows that fit and how far the first one is cut off", () => {
    const rows = new SizeIndex(1000, 30);
    expect(visible(rows, 0, 100)).toEqual({ first: 0, end: 4, shift: 0 });
    expect(visible(rows, 45, 100)).toEqual({ first: 1, end: 5, shift: 15 });
  });

  it("stops at the end of the sheet", () => {
    const rows = new SizeIndex(10, 30);
    expect(visible(rows, 10_000, 100)).toEqual({ first: 6, end: 10, shift: 20 });
  });

  it("scrolls 1:1 until the sheet is too tall, then in proportion", () => {
    expect(scrollScale(1_000_000, 500)).toBe(1);
    const total = 750_000_000;
    const scale = scrollScale(total, 600);
    expect(scale).toBeGreaterThan(70);
    expect(scrollExtent(total, 600)).toBe(10_000_600);
    // The bottom of the scrollbar shows the bottom of the sheet, and it round-trips.
    expect(toSheetPx(10_000_000, total, 600)).toBeCloseTo(total - 600, 0);
    expect(toScrollPx(toSheetPx(1234, total, 600), total, 600)).toBeCloseTo(1234, 6);
  });
});

const bounds = () => ({ rows: new SizeIndex(100, 20), cols: new SizeIndex(10, 50) });

describe("selection", () => {
  it("replaces, extends and adds ranges", () => {
    let s = single({ row: 2, col: 2 });
    s = click(s, { row: 4, col: 3 }, "extend");
    expect(s.ranges).toEqual([{ r1: 2, c1: 2, r2: 4, c2: 3 }]);
    expect(isSelected(s, 3, 3)).toBe(true);
    s = click(s, { row: 9, col: 9 }, "add");
    expect(s.ranges).toHaveLength(2);
    expect(s.active).toEqual({ row: 9, col: 9 });
    expect(click(s, { row: 0, col: 0 }, "replace").ranges).toHaveLength(1);
  });

  it("selects whole rows, columns and everything", () => {
    const b = bounds();
    expect(wholeRow(b, 3, 5).ranges[0]).toEqual({ r1: 3, r2: 5, c1: 0, c2: 9 });
    expect(wholeCol(b, 4).ranges[0]).toEqual({ c1: 4, c2: 4, r1: 0, r2: 99 });
    expect(selectAll(b).ranges[0]).toEqual({ r1: 0, c1: 0, r2: 99, c2: 9 });
  });

  it("moves one cell, stays inside the sheet and skips hidden rows", () => {
    const b = { rows: new SizeIndex(5, 20, [{ index: 2, size: 0 }]), cols: new SizeIndex(3, 50) };
    expect(move(b, single({ row: 1, col: 0 }), "down").active).toEqual({ row: 3, col: 0 });
    expect(move(b, single({ row: 3, col: 0 }), "up").active).toEqual({ row: 1, col: 0 });
    expect(move(b, single({ row: 0, col: 0 }), "left").active).toEqual({ row: 0, col: 0 });
    expect(move(b, single({ row: 4, col: 2 }), "right").active).toEqual({ row: 4, col: 2 });
  });

  it("Shift+arrow grows the range from the anchor in both directions", () => {
    const b = bounds();
    let s = single({ row: 5, col: 5 });
    s = move(b, s, "down", { extend: true });
    s = move(b, s, "down", { extend: true });
    expect(s.ranges[0]).toEqual({ r1: 5, c1: 5, r2: 7, c2: 5 });
    s = move(b, s, "up", { extend: true });
    s = move(b, s, "up", { extend: true });
    s = move(b, s, "up", { extend: true });
    expect(s.ranges[0]).toEqual({ r1: 4, c1: 5, r2: 5, c2: 5 });
    expect(s.active).toEqual({ row: 5, col: 5 });
  });

  describe("Ctrl+arrow", () => {
    const filled = new Set(["3:0", "4:0", "5:0", "9:0"]);
    const has = (r: number, c: number) => filled.has(`${r}:${c}`);

    it("goes to the end of a block of data", () => {
      expect(jump(bounds(), { row: 3, col: 0 }, "down", has)).toEqual({ row: 5, col: 0 });
    });
    it("from the end of a block goes to the next block", () => {
      expect(jump(bounds(), { row: 5, col: 0 }, "down", has)).toEqual({ row: 9, col: 0 });
    });
    it("from an empty cell goes to the next data", () => {
      expect(jump(bounds(), { row: 0, col: 0 }, "down", has)).toEqual({ row: 3, col: 0 });
      expect(jump(bounds(), { row: 8, col: 0 }, "up", has)).toEqual({ row: 5, col: 0 });
    });
    it("with nothing ahead goes to the sheet's edge", () => {
      expect(jump(bounds(), { row: 9, col: 0 }, "down", has)).toEqual({ row: 99, col: 0 });
      expect(jump(bounds(), { row: 3, col: 0 }, "left", has)).toEqual({ row: 3, col: 0 });
      expect(jump(bounds(), { row: 3, col: 0 }, "right", () => false)).toEqual({ row: 3, col: 9 });
    });
    it("extends with Shift", () => {
      const s = move(bounds(), single({ row: 3, col: 0 }), "down", { jump: true, extend: true, hasData: has });
      expect(s.ranges[0]).toEqual({ r1: 3, r2: 5, c1: 0, c2: 0 });
    });
  });

  it("Tab and Enter move inside a multi-cell selection and wrap", () => {
    const b = bounds();
    let s = click(single({ row: 1, col: 1 }), { row: 2, col: 2 }, "extend");
    s = advance(b, s, "next");
    expect(s.active).toEqual({ row: 1, col: 2 });
    s = advance(b, s, "next");
    expect(s.active).toEqual({ row: 2, col: 1 });
    s = advance(b, s, "down");
    // Already on the last row of the range, so Enter wraps to the top of the next column.
    expect(s.active).toEqual({ row: 1, col: 2 });
    s = advance(b, single({ row: 0, col: 0 }), "down");
    expect(s.active).toEqual({ row: 1, col: 0 });
  });
});
