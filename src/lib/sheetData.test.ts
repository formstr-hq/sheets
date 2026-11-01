import { describe, expect, it } from "vitest";
import type { Sheet } from "@fortune-sheet/core";
import { activeSheetId, storedSheets, toStoredSheet } from "./sheetData";

describe("FortuneSheet adapters", () => {
  it("converts a data matrix to sparse celldata and drops the matrix", () => {
    const sheet: Sheet = {
      id: "s",
      name: "Sheet 1",
      row: 20,
      column: 10,
      status: 1,
      data: [
        [{ v: "hello" }, null],
        [null, { v: 42 }],
      ],
    };
    const stored = toStoredSheet(sheet);
    expect(stored.data).toBeUndefined();
    expect(stored.celldata).toEqual([
      { r: 0, c: 0, v: { v: "hello" } },
      { r: 1, c: 1, v: { v: 42 } },
    ]);
    expect(stored.row).toBe(200);
    expect(stored.column).toBe(26);
  });

  it("finds the active sheet", () => {
    const sheets = [
      { id: "a", name: "A", status: 0 },
      { id: "b", name: "B", status: 1 },
    ] as Sheet[];
    expect(activeSheetId(sheets)).toBe("b");
  });

  it("strips data from every stored sheet", () => {
    const sheets = [
      {
        id: "a",
        name: "A",
        status: 1,
        data: [[{ v: "x" }]],
      },
    ] as Sheet[];
    expect(storedSheets(sheets)[0].data).toBeUndefined();
  });
});
