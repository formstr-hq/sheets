import { describe, expect, it } from "vitest";
import type { WorkbookCell } from "../workbook";
import type { BigSheetStore } from "./bigSheet";
import { FormulaIndex, pointRect, refreshDependents } from "./bigFormulas";

/** A store that holds cells in a map: one tab "S" of 100 rows (chunks of 50) and a tab "T". */
function memoryStore(rows = 100) {
  const cells = new Map<string, WorkbookCell>();
  const loaded = new Set<string>();
  const at = (tab: number, row: number, col: number) => `${tab}:${row}:${col}`;
  const store = {
    tabs: [
      {
        name: "S",
        rows,
        cols: 6,
        chunks: [
          { rowStart: 0, rowCount: 50 },
          { rowStart: 50, rowCount: 50 },
        ],
      },
      { name: "T", rows: 10, cols: 3, chunks: [{ rowStart: 0, rowCount: 10 }] },
    ],
    chunkFor: (tab: number, row: number) => (tab === 0 && row >= 50 ? 1 : 0),
    // Like the real store, a chunk is "in memory" only once it has been loaded.
    load: async (tab: number, chunk: number) => {
      loaded.add(`${tab}:${chunk}`);
    },
    hasChunk: (tab: number, chunk: number) => loaded.has(`${tab}:${chunk}`),
    cell: (tab: number, row: number, col: number) => cells.get(at(tab, row, col)) ?? null,
    async setDerived(tab: number, row: number, col: number, cell: WorkbookCell) {
      cells.set(at(tab, row, col), cell);
    },
    *cellsOfChunk(tab: number, chunk: number) {
      for (const [key, cell] of cells) {
        const [t, r, c] = key.split(":").map(Number);
        if (t === tab && store.chunkFor(t, r) === chunk) yield [r, c, cell] as [number, number, WorkbookCell];
      }
    },
  } as unknown as BigSheetStore;
  const put = (tab: number, row: number, col: number, cell: WorkbookCell) => cells.set(at(tab, row, col), cell);
  const value = (tab: number, row: number, col: number) => cells.get(at(tab, row, col))?.value;
  return { store, put, value, cells };
}

describe("refreshDependents", () => {
  it("recalculates a formula when a cell it reads changes, and chains through other formulas", async () => {
    const { store, put, value } = memoryStore();
    const index = new FormulaIndex();
    put(0, 0, 0, { value: 1 });
    put(0, 1, 0, { value: 2 });
    put(0, 0, 1, { formula: "SUM(A1:A2)", value: 3 });
    put(0, 0, 2, { formula: "B1*10", value: 30 });
    await index.set(store, 0, 0, 1, "SUM(A1:A2)");
    await index.set(store, 0, 0, 2, "B1*10");

    put(0, 1, 0, { value: 5 }); // the user changed A2
    const count = await refreshDependents(store, index, [pointRect({ tab: 0, row: 1, col: 0 })]);
    expect(count).toBe(2);
    expect(value(0, 0, 1)).toBe(6);
    expect(value(0, 0, 2)).toBe(60);
  });

  it("follows whole-column and cross-tab reads, and leaves unrelated formulas alone", async () => {
    const { store, put, value } = memoryStore();
    const index = new FormulaIndex();
    put(0, 3, 0, { value: 4 });
    put(0, 0, 3, { formula: "SUM(A:A)", value: 4 });
    put(1, 0, 0, { formula: "S!D1+1", value: 5 });
    put(0, 0, 4, { formula: "SUM(B1:B3)", value: 999 });
    await index.set(store, 0, 0, 3, "SUM(A:A)");
    await index.set(store, 1, 0, 0, "S!D1+1");
    await index.set(store, 0, 0, 4, "SUM(B1:B3)");

    put(0, 60, 0, { value: 10 }); // a far row in column A
    await refreshDependents(store, index, [pointRect({ tab: 0, row: 60, col: 0 })]);
    expect(value(0, 0, 3)).toBe(14);
    expect(value(1, 0, 0)).toBe(15);
    expect(value(0, 0, 4)).toBe(999);
  });

  it("marks a circular reference as #CIRC!", async () => {
    const { store, put, value, cells } = memoryStore();
    const index = new FormulaIndex();
    put(0, 0, 0, { formula: "B1+1", value: 1 });
    put(0, 0, 1, { formula: "A1+1", value: 2 });
    await refreshDependents(store, index, [pointRect({ tab: 0, row: 0, col: 0 }), pointRect({ tab: 0, row: 0, col: 1 })]);
    expect(value(0, 0, 0)).toBe("#CIRC!");
    expect(value(0, 0, 1)).toBe("#CIRC!");
    expect(cells.get("0:0:0")?.type).toBe("error");
  });

  it("a formula that reads itself is circular", async () => {
    const { store, put, value } = memoryStore();
    const index = new FormulaIndex();
    put(0, 5, 5, { formula: "F6+1", value: 1 });
    await refreshDependents(store, index, [pointRect({ tab: 0, row: 5, col: 5 })]);
    expect(value(0, 5, 5)).toBe("#CIRC!");
  });

  it("forgets a formula that was cleared and stops refreshing it", async () => {
    const { store, put, value, cells } = memoryStore();
    const index = new FormulaIndex();
    put(0, 0, 0, { value: 1 });
    put(0, 0, 1, { formula: "A1+1", value: 2 });
    await index.set(store, 0, 0, 1, "A1+1");
    cells.delete("0:0:1"); // the user cleared B1
    await refreshDependents(store, index, [{ tab: 0, r1: 0, c1: 0, r2: 0, c2: 5 }]);
    expect(index.has(0, 0, 1)).toBe(false);
    put(0, 0, 0, { value: 9 });
    expect(await refreshDependents(store, index, [pointRect({ tab: 0, row: 0, col: 0 })])).toBe(0);
    expect(value(0, 0, 1)).toBeUndefined();
  });

  it("learns the formulas of a chunk when it loads", async () => {
    const { store, put } = memoryStore();
    const index = new FormulaIndex();
    put(0, 70, 1, { formula: "A71*2", value: 0 });
    put(0, 1, 1, { value: 1 });
    expect(await index.scanChunk(store, 0, 1)).toBe(false); // not in memory yet
    await store.load(0, 1);
    expect(await index.scanChunk(store, 0, 1)).toBe(true);
    expect(index.has(0, 70, 1)).toBe(true);
    expect(index.size).toBe(1);
  });

  it("finds and updates formulas in chunks nobody opened, all the way down a column", async () => {
    const { store, put, value } = memoryStore(6000);
    const index = new FormulaIndex();
    for (let r = 0; r < 6000; r++) {
      put(0, r, 0, { value: r });
      put(0, r, 1, { formula: `A${r + 1}*2`, value: r * 2 });
    }
    // The index has seen none of these formulas; the user changes the whole input column.
    const seen: string[] = [];
    for (let r = 0; r < 6000; r++) put(0, r, 0, { value: r + 1 });
    const count = await refreshDependents(store, index, [{ tab: 0, r1: 0, c1: 0, r2: 5999, c2: 0 }], (_d, _t, phase) => {
      if (!seen.includes(phase)) seen.push(phase);
    });
    expect(count).toBe(6000);
    expect(seen).toEqual(["scan", "calc"]);
    expect(value(0, 0, 1)).toBe(2);
    expect(value(0, 5999, 1)).toBe(12000);
  });

  it("orders a long chain correctly (each row reads the one above)", async () => {
    const { store, put, value } = memoryStore(2000);
    const index = new FormulaIndex();
    put(0, 0, 0, { value: 1 });
    for (let r = 1; r < 2000; r++) put(0, r, 0, { formula: `A${r}+1`, value: 0 });
    put(0, 0, 0, { value: 10 });
    await refreshDependents(store, index, [pointRect({ tab: 0, row: 0, col: 0 })]);
    expect(value(0, 1999, 0)).toBe(2009);
  });
});
