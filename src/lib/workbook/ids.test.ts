import { describe, expect, it } from "vitest";
import {
  colIndexOf,
  concatRuns,
  deleteRuns,
  indexOfRow,
  insertRuns,
  legacyRows,
  mintRows,
  newDeviceTag,
  nextCounter,
  padCols,
  padRuns,
  rowCountOf,
  rowIdAt,
  sliceRuns,
  type RowRun,
} from "./ids";

describe("row runs", () => {
  const runs: RowRun[] = [["ab12", 0, 3], ["0000", 10, 2]];

  it("counts, finds and slices rows", () => {
    expect(rowCountOf(runs)).toBe(5);
    expect(rowIdAt(runs, 0)).toBe("ab12.0");
    expect(rowIdAt(runs, 3)).toBe("0000.10");
    expect(rowIdAt(runs, 5)).toBeUndefined();
    expect(indexOfRow(runs, "ab12.2")).toBe(2);
    expect(indexOfRow(runs, "0000.11")).toBe(4);
    expect(indexOfRow(runs, "zz99.1")).toBe(-1);
    expect(sliceRuns(runs, 1, 4)).toEqual([["ab12", 1, 2], ["0000", 10, 1]]);
  });

  it("merges runs that continue each other", () => {
    expect(concatRuns([["a", 0, 2]], [["a", 2, 3]], [["b", 0, 1]])).toEqual([["a", 0, 5], ["b", 0, 1]]);
    expect(concatRuns(legacyRows(0, 4), legacyRows(4, 4))).toEqual([["0000", 0, 8]]);
  });

  it("an insert keeps every other row's id", () => {
    const next = insertRuns(runs, 1, mintRows("zz01", 7, 2));
    expect(next).toEqual([["ab12", 0, 1], ["zz01", 7, 2], ["ab12", 1, 2], ["0000", 10, 2]]);
    expect(rowCountOf(next)).toBe(7);
    // The row that was at position 2 is now at position 4, with the same id.
    expect(indexOfRow(next, "ab12.2")).toBe(4);
  });

  it("a delete removes just those rows", () => {
    const next = deleteRuns(runs, 1, 3);
    expect(next).toEqual([["ab12", 0, 1], ["0000", 11, 1]]);
    expect(indexOfRow(next, "ab12.1")).toBe(-1);
  });

  it("pads a short list with legacy ids and trims a long one", () => {
    expect(padRuns([], 3)).toEqual([["0000", 0, 3]]);
    expect(padRuns(runs, 7)).toEqual([["ab12", 0, 3], ["0000", 10, 4]]);
    expect(padRuns(runs, 2)).toEqual([["ab12", 0, 2]]);
  });

  it("mints ids that cannot collide with this device's earlier ones", () => {
    expect(nextCounter(runs, "ab12")).toBe(3);
    expect(nextCounter(runs, "nope")).toBe(0);
    for (let i = 0; i < 200; i++) expect(newDeviceTag()).toMatch(/^[0-9a-z]{4}$/);
    // A draw that spells the legacy tag is thrown away.
    const draws = [0, 0, 0, 0, 0.5, 0.5, 0.5, 0.5];
    expect(newDeviceTag(() => draws.shift() as number)).toBe("iiii");
  });
});

describe("column ids", () => {
  it("pads with legacy ids and resolves them even when the list lacks them", () => {
    expect(padCols(["x.1"], 3)).toEqual(["x.1", "c1", "c2"]);
    expect(colIndexOf(["x.1", "c0"], "c0")).toBe(1); // listed ids win over the legacy reading
    expect(colIndexOf(["x.1"], "c7")).toBe(7);
    expect(colIndexOf(["x.1"], "gone.2")).toBe(-1);
  });
});
