import { describe, expect, it } from "vitest";
import { parse, print, rewriteForEdit, type Node } from "@formstr/sheets-formula";
import { ArrayAxis, formulaFromIds, formulaToIds, hasIdRefs, RunAxis, tombstonesFor, type RefContext, type TabAxes } from "./idRefs";

function tab(id: string, name: string, rows: number, cols: number): TabAxes {
  return {
    id,
    name,
    rows: new ArrayAxis(Array.from({ length: rows }, (_, i) => `r${id}.${i}`)),
    cols: new ArrayAxis(Array.from({ length: cols }, (_, i) => `c${id}${i}`)),
  };
}

const ctxOf = (tabs: TabAxes[], current = "a"): RefContext => ({ tabs, current });
const rows = (t: TabAxes) => t.rows as ArrayAxis;
const cols = (t: TabAxes) => t.cols as ArrayAxis;

describe("formulas by id: the text conversion", () => {
  const a = tab("a", "Sheet1", 20, 8);
  const b = tab("b", "My Data", 20, 8);
  const ctx = ctxOf([a, b]);

  it("turns references into id tokens and back, leaving the rest of the text alone", () => {
    const formula = 'SUM(B2:B4) + IF(A1>"B2:B4",$C$3, 0) * \'My Data\'!D5';
    const stored = formulaToIds(formula, ctx);
    expect(hasIdRefs(stored)).toBe(true);
    expect(stored).toContain('"B2:B4"'); // inside a string: untouched
    expect(stored).not.toMatch(/\bSUM\(B2/);
    expect(formulaFromIds(stored, ctx)).toBe("SUM(B2:B4) + IF(A1>\"B2:B4\",$C$3, 0) * 'My Data'!D5");
  });

  it("keeps $ markers, whole columns and whole rows", () => {
    for (const formula of ["$A$1+B$2+$C3", "SUM(A:A)", "SUM($B:$D)", "SUM(2:4)", "SUM($2:$4)", "A1:$B$3"]) {
      expect(formulaFromIds(formulaToIds(formula, ctx), ctx)).toBe(formula);
    }
  });

  it("does not mistake function names for references", () => {
    const formula = "LOG10(100)+ATAN2(1,1)+SUM(A1)";
    const stored = formulaToIds(formula, ctx);
    expect(stored).toContain("LOG10(100)");
    expect(stored).toContain("ATAN2(1,1)");
    expect(formulaFromIds(stored, ctx)).toBe(formula);
  });

  it("writes the sheet name only when the reference is to another sheet", () => {
    const stored = formulaToIds("Sheet1!A1+'My Data'!A1", ctx);
    expect(formulaFromIds(stored, ctx)).toBe("A1+'My Data'!A1");
    // the same formula typed on the other sheet
    const onB = ctxOf([a, b], "b");
    expect(formulaFromIds(formulaToIds("Sheet1!A1+'My Data'!A1", onB), onB)).toBe("Sheet1!A1+A1");
  });

  it("a reference to the formula's own tab means whichever tab holds the formula (copied tabs share their stored formulas)", () => {
    const stored = formulaToIds("A1+B2+Sheet1!C3+'My Data'!D4", ctx);
    expect(stored).toContain("~(_:");
    // read the same stored text as a copy of the tab that has other ids for the same rows and columns
    const copy: TabAxes = { id: "c", name: "Sheet1 (2)", rows: new ArrayAxis(rows(a).ids.slice()), cols: new ArrayAxis(cols(a).ids.slice()) };
    const inCopy = ctxOf([a, b, copy], "c");
    expect(formulaFromIds(stored, inCopy)).toBe("A1+B2+C3+'My Data'!D4");
  });

  it("follows a renamed sheet because it names the tab by id", () => {
    const stored = formulaToIds("Sheet1!A1", ctxOf([a, b], "b"));
    const renamed = ctxOf([{ ...a, name: "Budget" }, b], "b");
    expect(formulaFromIds(stored, renamed)).toBe("Budget!A1");
  });

  it("leaves an unknown sheet as text and a deleted tab as #REF!", () => {
    expect(formulaToIds("Nope!A1+A1", ctx)).toMatch(/^Nope!A1\+~\(/);
    const stored = formulaToIds("'My Data'!A1+A1", ctx);
    expect(formulaFromIds(stored, ctxOf([a], "a"))).toBe("#REF!+A1");
  });

  it("a reference past the end of the sheet moves with inserts and deletes before it", () => {
    const stored = formulaToIds("A500+B2+SUM(C2:C1000)", ctx);
    expect(stored).toMatch(/~\(_:@[^:]+\+\d+:/);
    rows(a).insert(0, ["new"]);
    expect(formulaFromIds(stored, ctx)).toBe("A501+B3+SUM(C3:C1001)");
    rows(a).remove(0, 1);
    expect(formulaFromIds(stored, ctx)).toBe("A500+B2+SUM(C2:C1000)");
    rows(a).remove(3, 2);
    expect(formulaFromIds(stored, ctx)).toBe("A498+B2+SUM(C2:C998)");
  });

  it("a range typed backwards is stored the right way round", () => {
    for (const [typed, shown] of [["SUM(B5:B2)", "SUM(B2:B5)"], ["D4:A1", "A1:D4"], ["$B$5:A$1", "A$1:$B$5"], ["5:2", "2:5"], ["B:A", "A:B"]] as const) {
      expect(formulaFromIds(formulaToIds(typed, ctx), ctx)).toBe(shown);
    }
  });

  it("does not convert twice, and keeps a quoted sheet name that holds a double quote", () => {
    const once = formulaToIds("A1+B2", ctx);
    expect(formulaToIds(once, ctx)).toBe(once);
    const q = tab("q", 'Say "hi"', 20, 8);
    const quoted = ctxOf([a, q], "a");
    const stored = formulaToIds("'Say \"hi\"'!B2*2", quoted);
    expect(stored).toMatch(/^~\(q:/);
    expect(formulaFromIds(stored, quoted)).toBe("'Say \"hi\"'!B2*2");
  });
});

describe("formulas by id: inserting and deleting", () => {
  it("a formula keeps its target when rows are inserted above and columns to the left", () => {
    const a = tab("a", "S", 20, 8);
    const ctx = ctxOf([a]);
    const stored = formulaToIds("B3*2", ctx);
    rows(a).insert(0, ["x1", "x2"]);
    cols(a).insert(0, ["y1"]);
    expect(formulaFromIds(stored, ctx)).toBe("C5*2");
  });

  it("a range grows when rows are inserted inside it and not at its edges", () => {
    const a = tab("a", "S", 20, 8);
    const ctx = ctxOf([a]);
    const stored = formulaToIds("SUM(B2:B4)", ctx);
    rows(a).insert(2, ["i1"]); // between B2 and B3
    expect(formulaFromIds(stored, ctx)).toBe("SUM(B2:B5)");
    rows(a).insert(1, ["i2"]); // just above the first corner
    expect(formulaFromIds(stored, ctx)).toBe("SUM(B3:B6)");
    rows(a).insert(6, ["i3"]); // just below the last corner
    expect(formulaFromIds(stored, ctx)).toBe("SUM(B3:B6)");
  });

  it("deleting a row the formula points at gives #REF!", () => {
    const a = tab("a", "S", 20, 8);
    const ctx = ctxOf([a]);
    const stored = formulaToIds("B3*2", ctx);
    rows(a).remove(2, 1);
    expect(formulaFromIds(stored, ctx)).toBe("#REF!*2");
  });

  it("a range shrinks when its edge rows are deleted and is #REF! only when nothing is left", () => {
    const a = tab("a", "S", 20, 8);
    const ctx = ctxOf([a]);
    const stored = formulaToIds("SUM(B2:B4)", ctx);
    rows(a).remove(3, 1); // the last row, B4
    expect(formulaFromIds(stored, ctx)).toBe("SUM(B2:B3)");
    rows(a).remove(1, 1); // the first row, B2
    expect(formulaFromIds(stored, ctx)).toBe("SUM(B2:B2)");
    rows(a).remove(1, 1); // the last one left
    expect(formulaFromIds(stored, ctx)).toBe("SUM(#REF!)");
  });

  it("tombstones chain: a corner whose neighbour was deleted later still resolves", () => {
    const a = tab("a", "S", 20, 8);
    const ctx = ctxOf([a]);
    const stored = formulaToIds("SUM(B3:B8)", ctx);
    // delete the end corner's row, then the row that was its previous neighbour
    rows(a).remove(7, 1);
    rows(a).remove(6, 1);
    expect(formulaFromIds(stored, ctx)).toBe("SUM(B3:B6)");
  });

  it("a deleted start corner moves to the next surviving row", () => {
    const a = tab("a", "S", 20, 8);
    const ctx = ctxOf([a]);
    const stored = formulaToIds("SUM(B3:B8)", ctx);
    rows(a).remove(1, 3); // rows 2, 3 and 4: the start corner (row 3) is inside
    expect(formulaFromIds(stored, ctx)).toBe("SUM(B2:B5)");
  });

  it("an edit on another sheet does not disturb a formula that reads a different one", () => {
    const a = tab("a", "S", 20, 8);
    const b = tab("b", "T", 20, 8);
    const ctx = ctxOf([a, b]);
    const stored = formulaToIds("A1+T!A1", ctx);
    rows(b).insert(0, ["n1"]);
    expect(formulaFromIds(stored, ctx)).toBe("A1+T!A2");
  });

  it("tombstonesFor records the neighbours around a deleted block", () => {
    const axis = new ArrayAxis(["a", "b", "c", "d", "e"]);
    expect(tombstonesFor(axis, 1, 2)).toEqual({ b: ["a", "d"], c: ["a", "d"] });
    expect(tombstonesFor(axis, 0, 1)).toEqual({ a: [null, "b"] });
    expect(tombstonesFor(axis, 4, 1)).toEqual({ e: ["d", null] });
  });

  it("works with a run-based axis (the rows of a big sheet)", () => {
    const axis = new RunAxis([["k3f9", 0, 1_000_000]]);
    const ctx: RefContext = { tabs: [{ id: "a", name: "S", rows: axis, cols: new ArrayAxis(["c0", "c1", "c2"]) }], current: "a" };
    const stored = formulaToIds("SUM(B900000:B900002)", ctx);
    expect(stored).toContain("k3f9.899999");
    expect(formulaFromIds(stored, ctx)).toBe("SUM(B900000:B900002)");
  });
});

// A random sequence of inserts and deletes must give the same A1 formula as the position-rewriting code
// that is in use today (and is checked against Excel's behaviour): that is the proof the id scheme is
// equivalent for every case that code handles.
describe("formulas by id agree with rewriting the positions", () => {
  const FORMULAS = [
    "B2*2",
    "SUM(B2:B6)",
    "SUM(B2:D4)+A1",
    "T!B3+B4",
    "SUM(T!B2:C5)*A3",
    "SUM(B:B)",
    "SUM(2:4)",
    "$B$3+C$4+$D5",
    "IF(B2>B3,SUM(B4:B9),T!C2)",
    "SUM(B2:B3)+SUM(B3:B4)+SUM(B4:B5)",
  ];

  function rng(seed: number) {
    let s = seed >>> 0;
    return () => {
      s = (s * 1664525 + 1013904223) >>> 0;
      return s / 2 ** 32;
    };
  }

  for (let seed = 1; seed <= 200; seed++) {
    it(`scenario ${seed}`, () => {
      const rand = rng(seed);
      const pick = <T,>(list: T[]) => list[Math.floor(rand() * list.length)];
      const a = tab("a", "S", 14, 6);
      const t = tab("t", "T", 14, 6);
      const tabs = [a, t];
      const ctx = ctxOf(tabs, "a");
      // T! in the formulas is the second sheet
      const texts = FORMULAS.map((f) => f.replace(/T!/g, "T!"));
      const stored = texts.map((f) => formulaToIds(f, ctx));
      let nodes: Node[] = texts.map((f) => parse("=" + f));
      let counter = 0;
      const steps = 1 + Math.floor(rand() * 8);
      for (let step = 0; step < steps; step++) {
        const target = pick(tabs);
        const axisName = pick<"row" | "col">(["row", "col"]);
        const axis = (axisName === "row" ? target.rows : target.cols) as ArrayAxis;
        const kind = rand() < 0.5 ? "insert" : "delete";
        const count = 1 + Math.floor(rand() * 3);
        const index = Math.floor(rand() * Math.max(1, axis.length - (kind === "delete" ? count - 1 : 0)));
        if (kind === "delete" && axis.length - count < 3) continue;
        nodes = nodes.map((n) => rewriteForEdit(n, { kind, axis: axisName, sheet: target.name, index, count }, "S"));
        if (kind === "insert") axis.insert(index, Array.from({ length: count }, () => `n${counter++}`));
        else axis.remove(index, count);
      }
      stored.forEach((g, i) => {
        const fromIds = formulaFromIds(g, ctx);
        expect(print(parse("=" + fromIds)), `${texts[i]} -> ${fromIds}`).toBe(print(nodes[i]));
      });
    });
  }
});
