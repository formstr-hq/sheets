import { describe, expect, it } from "vitest";
import { ArrayAxis, RunAxis, formulaFromIds, formulaToIds, tombstonesFor, type Axis, type RefContext, type TabAxes } from "./idRefs";

const mk = (rows = 20, cols = 6, id = "a", name = "S"): TabAxes => ({
  id,
  name,
  rows: new ArrayAxis(Array.from({ length: rows }, (_, i) => `${id}r${i}`)),
  cols: new ArrayAxis(Array.from({ length: cols }, (_, i) => `${id}c${i}`)),
});
const R = (t: TabAxes) => t.rows as ArrayAxis;
const C = (t: TabAxes) => t.cols as ArrayAxis;
const ctxOf = (t: TabAxes): RefContext => ({ tabs: [t], current: t.id });
const back = (f: string, t: TabAxes, edit: () => void) => {
  const c = ctxOf(t);
  const s = formulaToIds(f, c);
  edit();
  return formulaFromIds(s, c);
};

describe("limits", () => {
  // BUG: same root cause as the next tests (positions past the end are @N and never shift).
  it("XFD and row 1048576 shift with an insert before them", () => {
    const t = mk();
    expect(back("XFD1048576", t, () => R(t).insert(0, ["n"]))).toBe("XFD1048577");
  });
  // BUG: same root cause (@N never shifts).
  it("column past XFD inserted before", () => {
    const t = mk();
    expect(back("XFE5", t, () => C(t).insert(0, ["n"]))).toBe("XFF5");
  });
  // BUG: references past the end of the axis are positional ("@N") and do not move when rows/columns are
  // inserted before them. Excel (and rewriteForEdit) shift them. (idRefs.ts token() / formulaFromIds)
  it("ref past the end shifts when rows are inserted above it (like rewriteForEdit)", () => {
    const t = mk();
    expect(back("A500", t, () => R(t).insert(2, ["n"]))).toBe("A501");
  });
  // BUG: same for ranges with a far corner: SUM(B2:B1000) grows by one in Excel when a row is inserted inside.
  it("range with far corner grows when a row is inserted inside", () => {
    const t = mk();
    expect(back("SUM(B2:B1000)", t, () => R(t).insert(5, ["n"]))).toBe("SUM(B2:B1001)");
  });
  // BUG: deleting a row shrinks the far corner in Excel.
  it("range with far corner shrinks when a row inside is deleted", () => {
    const t = mk();
    expect(back("SUM(B2:B1000)", t, () => R(t).remove(5, 1))).toBe("SUM(B2:B999)");
  });
  // BUG: a ref just past the end does not move up when rows above are deleted (Excel: A18).
  it("ref just past the end moves up when rows above are deleted", () => {
    const t = mk();
    expect(back("A21", t, () => R(t).remove(0, 3))).toBe("A18");
  });
  it("row 0 / col beyond 3 letters stay text", () => {
    const t = mk();
    expect(formulaToIds("A0+AAAA1+0:3", ctxOf(t))).toBe("A0+AAAA1+0:3");
  });
  it("whole column / row tokens with empty axes", () => {
    const t = mk(0, 0);
    const c = ctxOf(t);
    expect(formulaFromIds(formulaToIds("SUM(A:B)+SUM(2:3)", c), c)).toBe("SUM(A:B)+SUM(2:3)");
    expect(formulaFromIds(formulaToIds("B3", c), c)).toBe("B3");
  });
  it("single-cell range A1:A1 and its deletion", () => {
    const t = mk();
    expect(back("SUM(A1:A1)", t, () => {})).toBe("SUM(A1:A1)");
    expect(back("SUM(A3:A3)", t, () => R(t).remove(2, 1))).toContain("#REF!");
  });
});

describe("tombstones", () => {
  it("delete same position repeatedly", () => {
    const t = mk();
    const out = back("SUM(B3:B12)", t, () => {
      for (let i = 0; i < 4; i++) R(t).remove(2, 1);
    });
    expect(out).toBe("SUM(B3:B8)");
  });
  it("delete then insert at same place: range start", () => {
    const t = mk();
    const out = back("SUM(B3:B8)", t, () => {
      R(t).remove(2, 1);
      R(t).insert(2, ["x"]);
    });
    expect(out).toBe("SUM(B4:B8)");
  });
  it("delete then insert at same place: range end", () => {
    const t = mk();
    const out = back("SUM(B3:B8)", t, () => {
      R(t).remove(7, 1);
      R(t).insert(7, ["x"]);
    });
    expect(out).toBe("SUM(B3:B7)");
  });
  it("both corners deleted, survivors between", () => {
    const t = mk();
    expect(
      back("SUM(B3:B10)", t, () => {
        R(t).remove(9, 1);
        R(t).remove(2, 1);
      }),
    ).toBe("SUM(B3:B8)");
  });
  it("both corners in the same deleted block -> #REF!", () => {
    const t = mk();
    expect(back("SUM(B3:B5)", t, () => R(t).remove(1, 6))).toBe("SUM(#REF!)");
  });
  it("range across several deleted blocks", () => {
    const t = mk();
    expect(
      back("SUM(B3:B14)", t, () => {
        R(t).remove(10, 3);
        R(t).remove(5, 2);
        R(t).remove(1, 2);
      }),
    ).toBe("SUM(B2:B7)");
  });
  it("whole-column range after column deletes", () => {
    const t = mk();
    expect(back("SUM(B:D)", t, () => C(t).remove(1, 1))).toBe("SUM(B:C)");
    const t2 = mk();
    expect(back("SUM(B:D)", t2, () => C(t2).remove(1, 3))).toContain("#REF!");
  });
  it("column tombstones: start corner deleted", () => {
    const t = mk();
    expect(back("SUM(B2:D5)", t, () => C(t).remove(1, 1))).toBe("SUM(B2:C5)");
  });
  it("tombstone cycle terminates", () => {
    const rows: Axis = {
      length: 2,
      idAt: (i) => ["x", "y"][i],
      indexOf: (id) => ["x", "y"].indexOf(id),
      tombstone: (id) => (id === "p" ? ["q", "q"] : id === "q" ? ["p", "p"] : undefined),
    };
    const t: TabAxes = { id: "a", name: "S", rows, cols: new ArrayAxis(["c0"]) };
    const f = (s: string) => formulaFromIds(s, ctxOf(t));
    expect(f("~(_:p:c0:0):~(_:q:c0:0)")).toBe("#REF!");
    expect(f("~(_:p:c0:0)")).toBe("#REF!");
  });
  it("self tombstone terminates", () => {
    const rows: Axis = { length: 1, idAt: () => "x", indexOf: (id) => (id === "x" ? 0 : -1), tombstone: (id) => (id === "p" ? ["p", "p"] : undefined) };
    const t: TabAxes = { id: "a", name: "S", rows, cols: new ArrayAxis(["c0"]) };
    expect(formulaFromIds("~(_:p:c0:0):~(_:p:c0:0)", ctxOf(t))).toBe("#REF!");
  });
  it("three hops of tombstones", () => {
    const t = mk();
    const out = back("SUM(B5:B10)", t, () => {
      R(t).remove(5, 1);
      R(t).remove(4, 1);
      R(t).remove(3, 1);
    });
    expect(out).toBe("SUM(B4:B7)");
  });
  it("ArrayAxis.remove across the end", () => {
    const t = mk(5);
    expect(back("SUM(B2:B5)", t, () => R(t).remove(3, 2))).toBe("SUM(B2:B3)");
  });
  it("range whose start moved past a live end -> #REF!", () => {
    const t = mk();
    expect(back("SUM(B3:B4)", t, () => R(t).remove(2, 2))).toBe("SUM(#REF!)");
  });
  it("range with start deleted at end of axis -> #REF!", () => {
    const t = mk(5);
    expect(back("SUM(B4:B5)", t, () => R(t).remove(3, 2))).toBe("SUM(#REF!)");
  });
  it("tombstonesFor out of range / count 0", () => {
    const a = new ArrayAxis(["a", "b"]);
    expect(tombstonesFor(a, 5, 2)).toEqual({});
    expect(tombstonesFor(a, 0, 0)).toEqual({});
    expect(tombstonesFor(a, 1, 10)).toEqual({ b: ["a", null] });
  });
  it("undo (re-insert same id) restores the reference", () => {
    const t = mk();
    expect(
      back("B3", t, () => {
        R(t).remove(2, 1);
        R(t).insert(2, ["ar2"]);
      }),
    ).toBe("B3");
  });
  it("undo then delete again", () => {
    const t = mk();
    expect(
      back("SUM(B3:B8)", t, () => {
        R(t).remove(2, 1);
        R(t).insert(2, ["ar2"]);
        R(t).remove(2, 1);
      }),
    ).toBe("SUM(B3:B7)");
  });
});

describe("RunAxis", () => {
  // BUG candidate: Number("") === 0, so the id "k3f9." counts as k3f9.0; Number accepts "1e0", "0x2", " 3".
  it("indexOf rejects malformed counters", () => {
    const a = new RunAxis([["k3f9", 0, 10]]);
    expect(a.indexOf("k3f9.")).toBe(-1);
    expect(a.indexOf("k3f9.1e0")).toBe(-1);
    expect(a.indexOf("k3f9.0x2")).toBe(-1);
    expect(a.indexOf("k3f9. 3")).toBe(-1);
    expect(a.indexOf("k3f9.3.5")).toBe(-1);
  });
  it("indexOf follows a replaced runs list", () => {
    const a = new RunAxis([["k", 0, 10]]);
    expect(a.indexOf("k.5")).toBe(5);
    a.runs = [
      ["k", 0, 3],
      ["z", 0, 5],
      ["k", 3, 7],
    ];
    expect(a.indexOf("k.5")).toBe(10); // after the 3 k rows and the 5 z rows come k.3 ... k.9
  });
  it("split runs with the same tag", () => {
    const a = new RunAxis([
      ["k", 0, 3],
      ["z", 0, 2],
      ["k", 5, 3],
    ]);
    const t: TabAxes = { id: "a", name: "S", rows: a, cols: new ArrayAxis(["c0"]) };
    const c = ctxOf(t);
    const s = formulaToIds("A6+A1", c);
    expect(formulaFromIds(s, c)).toBe("A6+A1");
  });
  it("id with a dot in the tag", () => {
    const a = new RunAxis([["a.b", 0, 3]]);
    expect(a.indexOf("a.b.2")).toBe(2);
  });
  it("zero length runs", () => {
    const a = new RunAxis([
      ["0000", 0, 0],
      ["k", 0, 2],
    ]);
    expect(a.idAt(0)).toBe("k.0");
    expect(a.indexOf("k.1")).toBe(1);
  });
});
