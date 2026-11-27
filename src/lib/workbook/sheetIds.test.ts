import { describe, expect, it } from "vitest";
import { indexOfRow, rowCountOf, rowIdAt } from "./ids";
import { localDeviceIds, SheetIdTracker, type DeviceIds } from "./sheetIds";
import type { WorkbookSheet } from "./types";

function device(tag = "dev1"): DeviceIds {
  let next = 0;
  return {
    tag,
    reserve(count, atLeast) {
      const first = Math.max(next, atLeast);
      next = first + count;
      return first;
    },
  };
}

const sheet = (rows: number, cols: number): WorkbookSheet => ({ id: "s1", name: "S", order: 0, rowCount: rows, colCount: cols, cells: {} });
const insertRow = (index: number, count = 1, direction: "lefttop" | "rightbottom" = "lefttop") => ({
  op: "insertRowCol",
  id: "s1",
  value: { id: "s1", type: "row", index, count, direction },
});
const deleteRow = (start: number, end: number) => ({ op: "deleteRowCol", id: "s1", value: { id: "s1", type: "row", start, end } });

describe("SheetIdTracker", () => {
  it("gives every row of an old sheet its legacy id, then follows inserts and deletes", () => {
    const tracker = new SheetIdTracker(device());
    tracker.seed([sheet(10, 3)]);
    tracker.apply(insertRow(2));
    // The sheet now has 11 rows; the new row sits at position 2 and every other row keeps its id.
    const [after] = tracker.attach([sheet(11, 3)]);
    expect(rowCountOf(after.rowIds ?? [])).toBe(11);
    expect(rowIdAt(after.rowIds ?? [], 2)).toBe("dev1.0");
    expect(rowIdAt(after.rowIds ?? [], 3)).toBe("0000.2");
    expect(rowIdAt(after.rowIds ?? [], 10)).toBe("0000.9");

    tracker.apply(deleteRow(0, 1));
    const [again] = tracker.attach([sheet(9, 3)]);
    expect(indexOfRow(again.rowIds ?? [], "0000.0")).toBe(-1);
    expect(rowIdAt(again.rowIds ?? [], 0)).toBe("dev1.0");
    expect(rowIdAt(again.rowIds ?? [], 8)).toBe("0000.9");
  });

  it("'insert below' lands after the clicked row", () => {
    const tracker = new SheetIdTracker(device());
    tracker.seed([sheet(5, 2)]);
    tracker.apply(insertRow(1, 2, "rightbottom"));
    const [out] = tracker.attach([sheet(7, 2)]);
    expect(rowIdAt(out.rowIds ?? [], 1)).toBe("0000.1");
    expect(rowIdAt(out.rowIds ?? [], 2)).toBe("dev1.0");
    expect(rowIdAt(out.rowIds ?? [], 3)).toBe("dev1.1");
    expect(rowIdAt(out.rowIds ?? [], 4)).toBe("0000.2");
  });

  it("never hands out a number twice, even after the rows that used it are deleted", () => {
    const tracker = new SheetIdTracker(device());
    tracker.seed([sheet(4, 2)]);
    tracker.apply(insertRow(0, 2));
    tracker.apply(deleteRow(0, 1)); // the two new rows go away again
    tracker.apply(insertRow(0, 1));
    const [out] = tracker.attach([sheet(5, 2)]);
    expect(rowIdAt(out.rowIds ?? [], 0)).toBe("dev1.2");
  });

  it("an insert near the top keeps the ids of every row below it", () => {
    const tracker = new SheetIdTracker(device());
    tracker.seed([{ ...sheet(10, 2), rowIds: [["ab12", 5, 10]] }]);
    tracker.apply(insertRow(2));
    tracker.apply(deleteRow(0, 0));
    const [out] = tracker.attach([sheet(10, 2)]);
    expect(rowIdAt(out.rowIds ?? [], 0)).toBe("ab12.6");
    expect(rowIdAt(out.rowIds ?? [], 1)).toBe("dev1.0");
    expect(rowIdAt(out.rowIds ?? [], 9)).toBe("ab12.14");
  });

  it("tracks columns too", () => {
    const tracker = new SheetIdTracker(device());
    tracker.seed([sheet(3, 4)]);
    tracker.apply({ op: "insertRowCol", id: "s1", value: { id: "s1", type: "column", index: 1, count: 1, direction: "lefttop" } });
    expect(tracker.attach([sheet(3, 5)])[0].colIds).toEqual(["c0", "dev1.0", "c1", "c2", "c3"]);
    tracker.apply({ op: "deleteRowCol", id: "s1", value: { id: "s1", type: "column", start: 0, end: 0 } });
    expect(tracker.attach([sheet(3, 4)])[0].colIds).toEqual(["dev1.0", "c1", "c2", "c3"]);
  });

  it("ignores other operations and sheets it does not know", () => {
    const tracker = new SheetIdTracker(device());
    tracker.seed([sheet(3, 3)]);
    tracker.apply({ op: "replace", id: "s1", value: { v: 1 } });
    const other = { ...sheet(3, 3), id: "other" };
    expect(tracker.attach([other])[0]).toBe(other);
    expect(tracker.attach([sheet(3, 3)])[0].rowIds).toEqual([["0000", 0, 3]]);
  });

  it("keeps ids a document already carries", () => {
    const tracker = new SheetIdTracker(device());
    tracker.seed([{ ...sheet(3, 2), rowIds: [["ab12", 5, 3]], colIds: ["x.1", "x.2"] }]);
    const [out] = tracker.attach([sheet(3, 2)]);
    expect(out.rowIds).toEqual([["ab12", 5, 3]]);
    expect(out.colIds).toEqual(["x.1", "x.2"]);
  });
});

describe("localDeviceIds", () => {
  it("keeps the tag and a counter that survives a reload", () => {
    const store = new Map<string, string>();
    const storage = { getItem: (k: string) => store.get(k) ?? null, setItem: (k: string, v: string) => void store.set(k, v) };
    const first = localDeviceIds(storage);
    expect(first.reserve(3, 0)).toBe(0);
    const second = localDeviceIds(storage);
    expect(second.tag).toBe(first.tag);
    expect(second.reserve(1, 0)).toBe(3);
    expect(second.reserve(1, 100)).toBe(100);
  });

  it("works without storage", () => {
    const ids = localDeviceIds(undefined);
    expect(ids.tag).toMatch(/^[0-9a-z]{4}$/);
    expect(ids.reserve(2, 0)).toBe(0);
  });
});
