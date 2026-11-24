import { describe, expect, it } from "vitest";
import type { WorkbookCell } from "../workbook";
import type { BigSheetStore } from "./bigSheet";
import { calculateTyped } from "./bigFormulas";

/** Two tabs; "Data" has 10 rows in two chunks of 5: column A = 1..10, column B = "a".."j". */
function fakeStore(): { store: BigSheetStore; loaded: number[] } {
  const loaded: number[] = [];
  const chunk = (from: number): [number, number, WorkbookCell][] => {
    const out: [number, number, WorkbookCell][] = [];
    for (let r = from; r < from + 5; r++) {
      out.push([r, 0, { value: r + 1 }], [r, 1, { value: String.fromCharCode(97 + r) }]);
    }
    return out;
  };
  const err: [number, number, WorkbookCell][] = [[0, 0, { value: "#N/A", type: "error" }]];
  const store = {
    tabs: [
      { name: "Data", rows: 10, cols: 2, chunks: [{}, {}] },
      { name: "Other", rows: 1, cols: 1, chunks: [{}] },
    ],
    chunkFor: (tab: number, row: number) => (tab === 0 ? (row >= 5 ? 1 : 0) : 0),
    load: async (tab: number, c: number) => void loaded.push(tab * 10 + c),
    *cellsOfChunk(tab: number, c: number) {
      yield* tab === 0 ? chunk(c * 5) : err;
    },
  } as unknown as BigSheetStore;
  return { store, loaded };
}

describe("calculateTyped", () => {
  it("reads only the chunks the formula names", async () => {
    const { store, loaded } = fakeStore();
    expect(await calculateTyped(store, 0, 20, 5, "SUM(A1:A3)")).toEqual({ value: 6, error: false });
    expect(loaded).toEqual([0]);
  });

  it("reads a whole column across chunks and reports progress", async () => {
    const { store, loaded } = fakeStore();
    const steps: string[] = [];
    const result = await calculateTyped(store, 0, 20, 5, "SUM(A:A)", (done, total) => steps.push(`${done}/${total}`));
    expect(result).toEqual({ value: 55, error: false });
    expect(loaded).toEqual([0, 1]);
    expect(steps).toEqual(["1/2", "2/2"]);
  });

  it("follows references to another tab and passes errors through", async () => {
    const { store } = fakeStore();
    expect(await calculateTyped(store, 0, 20, 5, "Other!A1")).toEqual({ value: "#N/A", error: true });
    expect(await calculateTyped(store, 0, 20, 5, "Nope!A1")).toEqual({ value: "#REF!", error: true });
  });

  it("rounds to 15 digits like Excel and reports errors and bad syntax", async () => {
    const { store } = fakeStore();
    expect(await calculateTyped(store, 0, 20, 5, "0.1+0.2")).toEqual({ value: 0.3, error: false });
    expect(await calculateTyped(store, 0, 20, 5, "1/0")).toEqual({ value: "#DIV/0!", error: true });
    expect(await calculateTyped(store, 0, 20, 5, "SUM(")).toEqual({ value: "#NAME?", error: true });
    expect(await calculateTyped(store, 0, 20, 5, 'IF(A5>4,"big","small")')).toEqual({ value: "big", error: false });
  });
});
