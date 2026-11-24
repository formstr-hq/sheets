import { describe, expect, it } from "vitest";
import { normalizeWorkbook, WORKBOOK_SCHEMA, type WorkbookDocument } from "../workbook";
import { BigSheetStore } from "./bigSheet";
import { FormulaIndex } from "./bigFormulas";
import { editStructure } from "./bigStructure";

function sample(): WorkbookDocument {
  return normalizeWorkbook({
    schema: WORKBOOK_SCHEMA,
    version: 1,
    title: "Small",
    activeSheetId: "s2",
    createdAt: 1,
    updatedAt: 1,
    sheets: [
      {
        id: "s1",
        name: "First",
        order: 0,
        rowCount: 50,
        colCount: 8,
        color: "#ff0000",
        cells: { "0:0": { value: "Item", type: "text" }, "1:0": { value: "Rent", type: "text" }, "1:1": { value: 1200, type: "number", style: { bold: true } } },
        columnWidths: { 1: 180 },
        hiddenRows: [4],
        frozen: { rows: 1, columns: 0 },
        merges: [{ r: 6, c: 2, rs: 2, cs: 2 }],
      },
      { id: "s2", name: "Second", order: 1, hidden: true, rowCount: 20, colCount: 5, cells: { "2:2": { value: "x", type: "text" } } },
    ],
  });
}

describe("a small sheet held in memory by the big-sheet store", () => {
  it("opens, shows its cells and layout, and turns back into the same document", async () => {
    const document = sample();
    const store = BigSheetStore.fromDocument(document);
    await store.open();
    expect(store.inMemory).toBe(true);
    expect(store.tabs.map((tab) => tab.name)).toEqual(["First", "Second"]);
    await store.load(0, 0);
    expect(store.cell(0, 1, 1)?.value).toBe(1200);
    expect(store.tabs[0].layout.merges).toEqual([{ r: 6, c: 2, rs: 2, cs: 2 }]);
    expect(store.tabs[0].rows).toBe(200);
    const back = await store.toDocument();
    expect(back.title).toBe("Small");
    expect(back.activeSheetId).toBe("s2");
    expect(back.sheets.map((s) => [s.id, s.name, s.hidden ?? false, s.color ?? null])).toEqual([
      ["s1", "First", false, "#ff0000"],
      ["s2", "Second", true, null],
    ]);
    expect(back.sheets[0].cells["1:1"]).toMatchObject({ value: 1200, style: { bold: true } });
    expect(back.sheets[0].columnWidths).toEqual({ 1: 180 });
    expect(back.sheets[0].hiddenRows).toEqual([4]);
    expect(back.sheets[0].frozen).toEqual({ rows: 1, columns: 0 });
    expect(back.sheets[0].rowCount).toBe(200);
    expect(back.sheets[1].cells["2:2"]).toMatchObject({ value: "x" });
  });

  it("keeps edits, inserted rows and a formula rewrite when it is turned back into a document", async () => {
    const store = BigSheetStore.fromDocument(sample());
    await store.open();
    await store.load(0, 0);
    const index = new FormulaIndex();
    await store.setCell(0, 5, 1, "7");
    await store.edit(0, 3, 1, 3, 1, () => ({ formula: "B2*2", value: 2400 }));
    await index.set(store, 0, 3, 1, "B2*2");
    await editStructure(store, index, 0, { kind: "insert", axis: "row", index: 1, count: 2 });
    const back = await store.toDocument();
    const sheet = back.sheets[0];
    expect(sheet.rowCount).toBe(202);
    expect(sheet.cells["3:1"]).toMatchObject({ value: 1200 });
    expect(sheet.cells["5:1"]).toMatchObject({ formula: "B4*2" });
    expect(sheet.cells["7:1"]).toMatchObject({ value: 7 });
    // The ids line up with the rows and columns.
    expect(sheet.colIds?.length).toBe(sheet.colCount);
    // Reading it again gives the same thing.
    const again = BigSheetStore.fromDocument(back);
    await again.open();
    expect((await again.toDocument()).sheets[0].cells).toEqual(sheet.cells);
  });
});

describe("pictures in a sheet", () => {
  const image = { id: "i1", src: "data:image/png;base64,AAAA", row: 5, col: 2, dx: 4, dy: 4, width: 100, height: 60 };
  it("move with inserted and deleted rows and survive turning the sheet into a document", async () => {
    const store = BigSheetStore.fromDocument(sample());
    await store.open();
    await store.load(0, 0);
    await store.setLayout(0, { images: [image] });
    const index = new FormulaIndex();
    await editStructure(store, index, 0, { kind: "insert", axis: "row", index: 2, count: 3 });
    expect(store.tabs[0].layout.images?.[0]).toMatchObject({ row: 8, col: 2 });
    await editStructure(store, index, 0, { kind: "insert", axis: "col", index: 0, count: 1 });
    expect(store.tabs[0].layout.images?.[0]).toMatchObject({ row: 8, col: 3 });
    const back = await store.toDocument();
    expect(back.sheets[0].images?.[0]).toMatchObject({ id: "i1", row: 8, col: 3, width: 100 });
    await editStructure(store, index, 0, { kind: "delete", axis: "row", index: 8, count: 1 });
    expect(store.tabs[0].layout.images).toBeUndefined();
  });
});
