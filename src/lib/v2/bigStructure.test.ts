import { describe, expect, it } from "vitest";
import { importXlsxToBlossom } from "./streamImport";
import { BigSheetStore } from "./bigSheet";
import { calculateTyped, FormulaIndex, isVolatile, recalcVolatileInChunk, refreshDependents } from "./bigFormulas";
import { editStructure, renameTabAndFormulas } from "./bigStructure";
import { sortArea } from "./bigSort";
import { BLANK, NO_COLOR, rowsToHide, rowsToHideByColor, scanColumn } from "./bigFilter";
import { copyBlock, dragFill, extendSeries, moveCells } from "./bigClipboard";
import { editLayout } from "./layoutEdit";
import { fakeBlossom, localSigner } from "./testUtils";
import { buildLargeXlsx } from "../xlsxStream/testFixtures";
import { createWorkerHandler } from "../xlsxStream/workerCore";
import type { WorkerLike } from "../xlsxStream/client";
import type { MainToWorker } from "../xlsxStream/protocol";

function inProcessWorker(): WorkerLike {
  const worker: WorkerLike = {
    onmessage: null,
    onerror: null,
    onmessageerror: null,
    postMessage(message: MainToWorker) {
      queueMicrotask(() => handle(message));
    },
    terminate() {
      handle({ t: "cancel" });
    },
  };
  const handle = createWorkerHandler((message) => setTimeout(() => worker.onmessage?.({ data: message }), 0));
  return worker;
}

const START = { maxChunkBytes: 64 * 1024, headRawBytes: 100_000, rawCap: 1_000_000 };
const COLS = 8;

async function open(label: string, tabs = 1, rows = 3000) {
  const bytes = await buildLargeXlsx({ tabs, rows, cols: COLS, sharedStrings: false });
  const fake = fakeBlossom();
  const signer = localSigner();
  const servers = [`https://${label}.test`];
  const env = { fetch: fake.fetch, sleep: async () => {} };
  const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), {
    title: label, servers, signer, env, createWorker: inProcessWorker, start: START,
  });
  const store = new BigSheetStore(root, servers, env);
  await store.open();
  const reopen = async (from: BigSheetStore = store) => {
    const next = await from.save({ servers, signer, ...env });
    const again = new BigSheetStore(next, servers, env);
    await again.open();
    return again;
  };
  return { store, reopen };
}

/** Every cell's value, row by row (opens every chunk). */
async function snapshot(store: BigSheetStore, tab = 0): Promise<(unknown[] | null)[]> {
  const rows: (unknown[] | null)[] = [];
  for (let c = 0; c < store.tabs[tab].chunks.length; c++) await store.load(tab, c);
  for (let r = 0; r < store.tabs[tab].rows; r++) {
    const values = Array.from({ length: store.tabs[tab].cols }, (_, c) => store.cell(tab, r, c)?.value ?? null);
    rows.push(values.every((v) => v === null) ? null : values);
  }
  return rows;
}

describe("layout after inserting or deleting", () => {
  it("shifts sizes, hidden rows, merges and borders", () => {
    const layout = {
      rowHeights: { 2: 40, 10: 50 },
      hiddenRows: [3, 12],
      columnWidths: { 1: 200 },
      merges: [{ r: 8, c: 0, rs: 4, cs: 2 }],
      borders: [{ r1: 5, c1: 0, r2: 6, c2: 1, type: "all" as const }],
    };
    const inserted = editLayout(layout, { kind: "insert", axis: "row", index: 4, count: 3 });
    expect(inserted.rowHeights).toEqual({ 2: 40, 13: 50 });
    expect(inserted.hiddenRows).toEqual([3, 15]);
    expect(inserted.merges).toEqual([{ r: 11, c: 0, rs: 4, cs: 2 }]);
    expect(inserted.borders?.[0]).toMatchObject({ r1: 8, r2: 9 });
    expect(inserted.columnWidths).toEqual({ 1: 200 });
    // An insert inside a merged block makes it taller.
    expect(editLayout(layout, { kind: "insert", axis: "row", index: 10, count: 2 }).merges).toEqual([{ r: 8, c: 0, rs: 6, cs: 2 }]);
    // Deleting rows of a block shrinks it; deleting all of it removes it.
    expect(editLayout(layout, { kind: "delete", axis: "row", index: 9, count: 2 }).merges).toEqual([{ r: 8, c: 0, rs: 2, cs: 2 }]);
    expect(editLayout(layout, { kind: "delete", axis: "row", index: 8, count: 4 }).merges).toBeUndefined();
    // Deleting a row drops its height and moves the ones below up.
    expect(editLayout(layout, { kind: "delete", axis: "row", index: 2, count: 1 }).rowHeights).toEqual({ 9: 50 });
    const columns = editLayout(layout, { kind: "insert", axis: "col", index: 1, count: 1 });
    expect(columns.columnWidths).toEqual({ 2: 200 });
    expect(columns.merges).toEqual([{ r: 8, c: 0, rs: 4, cs: 3 }]);
  });
});

describe("inserting and deleting in a big sheet", () => {
  it("inserts rows in the middle, keeps everything below, and survives save and reload", async () => {
    const { store, reopen } = await open("struct-insert");
    expect(store.tabs[0].chunks.length).toBeGreaterThan(3);
    const before = await snapshot(store);
    const at = store.tabs[0].chunks[1].rowStart + 7;
    await store.setLayout(0, { rowHeights: { [at + 1]: 70 } });

    await store.applyEdit(0, { kind: "insert", axis: "row", index: at, count: 5 });
    expect(store.tabs[0].rows).toBe(before.length + 5);
    const expected = [...before.slice(0, at), ...Array(5).fill(null), ...before.slice(at)];
    expect(await snapshot(store)).toEqual(expected);
    expect(store.tabs[0].layout.rowHeights).toEqual({ [at + 6]: 70 });

    const again = await reopen();
    expect(again.tabs[0].rows).toBe(before.length + 5);
    expect(await snapshot(again)).toEqual(expected);
    expect(again.tabs[0].layout.rowHeights).toEqual({ [at + 6]: 70 });
  }, 120_000);

  it("deletes rows across chunks, including whole chunks, and survives save and reload", async () => {
    const { store, reopen } = await open("struct-delete");
    const chunks = store.tabs[0].chunks;
    expect(chunks.length).toBeGreaterThan(4);
    const before = await snapshot(store);
    const from = chunks[1].rowStart + 3;
    const to = chunks[3].rowStart + 4; // exclusive: chunk 2 is deleted whole
    await store.applyEdit(0, { kind: "delete", axis: "row", index: from, count: to - from });
    const expected = [...before.slice(0, from), ...before.slice(to)];
    expect(store.tabs[0].rows).toBe(expected.length);
    expect(await snapshot(store)).toEqual(expected);

    const again = await reopen();
    expect(again.tabs[0].rows).toBe(expected.length);
    expect(await snapshot(again)).toEqual(expected);
  }, 120_000);

  it("inserts and deletes columns, in loaded and unloaded chunks, through save and reload", async () => {
    const { store, reopen } = await open("struct-cols");
    const before = await snapshot(store);
    const cols0 = store.tabs[0].cols;
    await store.applyEdit(0, { kind: "insert", axis: "col", index: 2, count: 2 });
    const widen = (row: unknown[] | null) => (row ? [...row.slice(0, 2), null, null, ...row.slice(2)] : null);
    const inserted = before.map(widen);
    expect(store.tabs[0].cols).toBe(cols0 + 2);
    expect(await snapshot(store)).toEqual(inserted);

    const again = await reopen();
    expect(await snapshot(again)).toEqual(inserted);

    await again.applyEdit(0, { kind: "delete", axis: "col", index: 0, count: 1 });
    const trimmed = inserted.map((row) => (row ? row.slice(1) : null));
    expect(await snapshot(again)).toEqual(trimmed);
    // Read fresh from the saved sheet, the deleted column does not come back.
    const fresh = await reopen(again);
    expect(await snapshot(fresh)).toEqual(trimmed);
    expect(fresh.tabs[0].cols).toBe(cols0 + 1);
  }, 120_000);

  it("rewrites formulas, moves them with their rows, and gives #REF! for deleted cells", async () => {
    const { store, reopen } = await open("struct-formulas");
    const rowOf = (chunk: number, extra: number) => store.tabs[0].chunks[chunk].rowStart + extra;
    const source = rowOf(2, 10);
    const formulaRow = rowOf(1, 4);
    await store.load(0, store.chunkFor(0, source));
    await store.load(0, store.chunkFor(0, formulaRow));
    await store.setCell(0, source, 3, "10");
    const index = new FormulaIndex();
    const type = async (row: number, col: number, formula: string) => {
      const result = await calculateTyped(store, 0, row, col, formula);
      await store.edit(0, row, col, row, col, () => ({ formula, value: result.value }));
      await index.set(store, 0, row, col, formula);
    };
    await type(formulaRow, 5, `D${source + 1}*2`); // F = D(source)*2
    await type(formulaRow, 6, `SUM(D${source - 1}:D${source + 3})`); // a range around the source
    expect(store.cell(0, formulaRow, 5)?.value).toBe(20);

    // Insert 3 rows above the source: the formulas keep pointing at the same cells.
    await editStructure(store, index, 0, { kind: "insert", axis: "row", index: source - 3, count: 3 });
    expect(store.cell(0, source + 3, 3)?.value).toBe(10);
    expect(store.cell(0, formulaRow, 5)?.formula).toBe(`D${source + 4}*2`);
    expect(store.cell(0, formulaRow, 5)?.value).toBe(20);
    expect(store.cell(0, formulaRow, 6)?.formula).toBe(`SUM(D${source + 2}:D${source + 6})`);

    // Insert above the formula: it moves down and still reads the same cell.
    await editStructure(store, index, 0, { kind: "insert", axis: "row", index: formulaRow, count: 2 });
    expect(store.cell(0, formulaRow + 2, 5)?.formula).toBe(`D${source + 6}*2`);
    expect(store.cell(0, formulaRow + 2, 5)?.value).toBe(20);

    const again = await reopen();
    await again.load(0, again.chunkFor(0, formulaRow + 2));
    expect(again.cell(0, formulaRow + 2, 5)?.formula).toBe(`D${source + 6}*2`);

    // Delete the row the first formula reads: it becomes #REF!.
    await editStructure(store, index, 0, { kind: "delete", axis: "row", index: source + 5, count: 1 });
    const broken = store.cell(0, formulaRow + 2, 5);
    expect(broken?.formula).toContain("#REF!");
    expect(broken?.type).toBe("error");
  }, 180_000);
});

describe("sorting a selection", () => {
  it("sorts rows by a column, keeps empty cells last, and undoes in one step", async () => {
    const { store } = await open("struct-sort");
    await store.load(0, 0);
    const top = 10;
    const input = [["pear", 3], ["Apple", 10], [undefined, 1], ["fig", 2], ["banana", 9]] as const;
    for (let i = 0; i < input.length; i++) {
      const [name, n] = input[i];
      await store.setCell(0, top + i, 0, name ?? "");
      await store.setCell(0, top + i, 1, String(n));
    }
    const area = { r1: top, c1: 0, r2: top + 4, c2: 1 };
    const read = () => Array.from({ length: 5 }, (_, i) => [store.cell(0, top + i, 0)?.value ?? null, store.cell(0, top + i, 1)?.value]);

    expect(await sortArea(store, 0, area, 0, false)).toBe(true);
    expect(read().map((r) => r[0])).toEqual(["Apple", "banana", "fig", "pear", null]);
    expect(read().map((r) => r[1])).toEqual([10, 9, 2, 3, 1]); // rows moved whole

    await sortArea(store, 0, area, 1, true);
    expect(read().map((r) => r[1])).toEqual([10, 9, 3, 2, 1]);

    store.undo();
    expect(read().map((r) => r[0])).toEqual(["Apple", "banana", "fig", "pear", null]);

    await store.edit(0, top, 2, top, 2, () => ({ formula: "A1", value: 1 }));
    await expect(sortArea(store, 0, { ...area, c2: 2 }, 0, false)).rejects.toThrow(/formulas/);
    await expect(sortArea(store, 0, area, 0, false, [{ r: top, c: 0, rs: 2, cs: 1 }])).rejects.toThrow(/merged/);
  }, 120_000);
});

describe("links and comments", () => {
  it("keep their data through save and reload and move with inserted rows", async () => {
    const { store, reopen } = await open("struct-notes");
    await store.load(0, 0);
    await store.edit(0, 5, 2, 5, 2, () => ({ value: "site", link: { url: "https://example.com" }, comment: "check me", validation: { type: "list", options: ["site", "other"] } }));
    await store.applyEdit(0, { kind: "insert", axis: "row", index: 2, count: 2 });
    const again = await reopen();
    await again.load(0, 0);
    expect(again.cell(0, 7, 2)).toMatchObject({ value: "site", link: { url: "https://example.com" }, comment: "check me", validation: { type: "list", options: ["site", "other"] } });
    expect(again.cell(0, 5, 2)?.comment).toBeUndefined();
  }, 120_000);
});

describe("conditional formatting rules in the layout", () => {
  it("move with inserted rows, shrink with deleted ones, and survive save and reload", async () => {
    const rule = { ranges: [{ r1: 4, c1: 0, r2: 8, c2: 1 }], condition: { type: "gt", values: ["5"] }, format: { cellColor: "#fee" } };
    expect(editLayout({ conditionalFormats: [rule] }, { kind: "insert", axis: "row", index: 2, count: 3 }).conditionalFormats?.[0].ranges).toEqual([{ r1: 7, c1: 0, r2: 11, c2: 1 }]);
    expect(editLayout({ conditionalFormats: [rule] }, { kind: "delete", axis: "row", index: 6, count: 10 }).conditionalFormats?.[0].ranges).toEqual([{ r1: 4, c1: 0, r2: 5, c2: 1 }]);
    expect(editLayout({ conditionalFormats: [rule] }, { kind: "delete", axis: "row", index: 4, count: 5 }).conditionalFormats).toBeUndefined();
    expect(editLayout({ conditionalFormats: [rule] }, { kind: "insert", axis: "col", index: 1, count: 1 }).conditionalFormats?.[0].ranges).toEqual([{ r1: 4, c1: 0, r2: 8, c2: 2 }]);

    const { store, reopen } = await open("struct-cf");
    await store.setLayout(0, { conditionalFormats: [rule] });
    const again = await reopen();
    expect(again.tabs[0].layout.conditionalFormats).toEqual([rule]);
  }, 120_000);
});

describe("tabs", () => {
  it("renames (formulas follow), recolours, hides, adds, reorders and deletes tabs through save and reload", async () => {
    const { store, reopen } = await open("struct-tabs", 2);
    expect(store.tabs).toHaveLength(2);
    const [first, second] = store.tabs.map((tab) => tab.name);
    await store.load(0, 0);
    await store.load(1, 0);
    await store.setCell(1, 3, 3, "42");
    const index = new FormulaIndex();
    const formula = `'${second}'!D4*2`;
    const result = await calculateTyped(store, 0, 1, 1, formula);
    await store.edit(0, 1, 1, 1, 1, () => ({ formula, value: result.value }));
    await index.set(store, 0, 1, 1, formula);
    expect(store.cell(0, 1, 1)?.value).toBe(84);

    await renameTabAndFormulas(store, index, 1, "Budget");
    expect(store.tabs[1].name).toBe("Budget");
    expect(store.cell(0, 1, 1)?.formula).toBe("Budget!D4*2");
    store.setTabColor(0, "#ff0000");
    expect(() => store.setTabHidden(0, true)).not.toThrow();
    expect(() => store.setTabHidden(1, true)).toThrow(/stay visible/);
    store.setTabHidden(0, false);

    const added = store.addTab("Notes");
    expect(added).toBe(2);
    await store.setCell(2, 4, 1, "hello");

    const again = await reopen();
    expect(again.tabs.map((t) => t.name)).toEqual([first, "Budget", "Notes"]);
    expect(again.tabMeta(0).color).toBe("#ff0000");
    expect(again.tabs[2].rows).toBe(200);
    await again.load(2, 0);
    expect(again.cell(2, 4, 1)?.value).toBe("hello");
    await again.load(0, 0);
    expect(again.cell(0, 1, 1)?.formula).toBe("Budget!D4*2");

    // Reorder: Notes first. Cells follow their tab, also when saved and reloaded.
    again.remapTabs([2, 0, 1]);
    expect(again.tabs.map((t) => t.name)).toEqual(["Notes", first, "Budget"]);
    expect(again.cell(0, 4, 1)?.value).toBe("hello");
    const third = await reopen(again);
    expect(third.tabs.map((t) => t.name)).toEqual(["Notes", first, "Budget"]);
    await third.load(0, 0);
    expect(third.cell(0, 4, 1)?.value).toBe("hello");
    await third.load(2, 0);
    expect(third.cell(2, 3, 3)?.value).toBe(42);

    // Delete the first tab: the others keep their data.
    third.remapTabs([1, 2]);
    const fourth = await reopen(third);
    expect(fourth.tabs.map((t) => t.name)).toEqual([first, "Budget"]);
    await fourth.load(1, 0);
    expect(fourth.cell(1, 3, 3)?.value).toBe(42);
    expect(fourth.root.tabs.map((t) => t.order)).toEqual([0, 1]);
    expect(fourth.root.tabs.some((t) => t.id === fourth.root.activeSheetId)).toBe(true);
    expect(() => fourth.remapTabs([])).toThrow();
  }, 180_000);
});

describe("filtering by a column", () => {
  it("lists the values, hides the rows that do not match, keeps the header, and moves with the layout", async () => {
    const { store, reopen } = await open("struct-filter");
    await store.load(0, 0);
    const last = store.tabs[0].rows - 1;
    await store.load(0, store.chunkFor(0, last));
    await store.setCell(0, 1, 8, "red");
    await store.setCell(0, 2, 8, "blue");
    await store.setCell(0, 3, 8, "red");
    await store.setCell(0, last, 8, "green");

    const progress: number[] = [];
    const scan = (await scanColumn(store, 0, 8, (done) => progress.push(done)))!;
    expect(progress.at(-1)).toBe(store.tabs[0].chunks.length);
    const counts = Object.fromEntries(scan.values.map((v) => [v.value, v.count]));
    expect(counts.red).toBe(2);
    expect(counts.blue).toBe(1);
    expect(counts.green).toBe(1);
    expect(counts[BLANK]).toBe(store.tabs[0].rows - 4);

    // Show only red: with a header row, row 0 is never hidden.
    const hidden = rowsToHide(scan, store.tabs[0].rows, true, (text) => text === "red");
    expect(hidden).not.toContain(0);
    expect(hidden).not.toContain(1);
    expect(hidden).toContain(2);
    expect(hidden).toContain(last);
    expect(hidden).toHaveLength(store.tabs[0].rows - 1 - 2);
    // Without a header row, row 0 can be hidden too.
    expect(rowsToHide(scan, store.tabs[0].rows, false, (text) => text === "red")).toContain(0);

    const filter = { col: 8, header: true, hidden, summary: "is red" };
    await store.setLayout(0, { filter });
    expect((await reopen()).tabs[0].layout.filter).toEqual(filter);

    // Inserting a row above moves the hidden rows; deleting the column drops the filter.
    expect(editLayout({ filter }, { kind: "insert", axis: "row", index: 2, count: 1 }).filter?.hidden[0]).toBe(hidden[0] + 1);
    expect(editLayout({ filter }, { kind: "insert", axis: "col", index: 0, count: 1 }).filter?.col).toBe(9);
    expect(editLayout({ filter }, { kind: "delete", axis: "col", index: 8, count: 1 }).filter).toBeUndefined();
  }, 120_000);
});

describe("fill handle", () => {
  it("continues numbers and numbered text, and repeats the rest", () => {
    expect(extendSeries([1, 2, 3], 3)).toEqual([4, 5, 6]);
    expect(extendSeries([10, 8], 2)).toEqual([6, 4]);
    expect(extendSeries([0.1, 0.2], 1)).toEqual([0.3]);
    expect(extendSeries([5], 2)).toBeNull(); // one number repeats
    expect(extendSeries([1, 2, 4], 2)).toBeNull(); // not a steady step
    expect(extendSeries(["Item 1"], 3)).toEqual(["Item 2", "Item 3", "Item 4"]);
    expect(extendSeries(["Q1", "Q2"], 2)).toEqual(["Q3", "Q4"]);
    expect(extendSeries(["a007"], 2)).toEqual(["a008", "a009"]);
    expect(extendSeries(["a", "b"], 2)).toBeNull();
    expect(extendSeries(["x1", "y2"], 2)).toBeNull();
    expect(extendSeries([1, undefined], 2)).toBeNull();
  });

  it("fills down and right, shifting formulas, and undoes in one step", async () => {
    const { store } = await open("struct-fill");
    await store.load(0, 0);
    const set = (r: number, c: number, text: string) => store.setCell(0, r, c, text);
    await set(10, 0, "1");
    await set(11, 0, "3"); // step 2
    await set(10, 1, "Week 1"); // text series
    await set(10, 2, "x"); // repeats
    await set(11, 2, "y");
    await store.edit(0, 10, 3, 10, 3, () => ({ formula: "A11*10", value: 10 }));

    const done = await dragFill(store, 0, { r1: 10, c1: 0, r2: 11, c2: 2 }, { r1: 10, c1: 0, r2: 15, c2: 2 });
    expect(done).toEqual({ r1: 12, c1: 0, r2: 15, c2: 2 });
    expect([12, 13, 14, 15].map((r) => store.cell(0, r, 0)?.value)).toEqual([5, 7, 9, 11]);
    // Two different texts that are not a series repeat in turn.
    expect([12, 13].map((r) => store.cell(0, r, 1)?.value)).toEqual([store.cell(0, 10, 1)?.value, store.cell(0, 11, 1)?.value]);
    expect([12, 13, 14, 15].map((r) => store.cell(0, r, 2)?.value)).toEqual(["x", "y", "x", "y"]);

    // A formula copied down moves its relative references.
    const f = await dragFill(store, 0, { r1: 10, c1: 3, r2: 10, c2: 3 }, { r1: 10, c1: 3, r2: 12, c2: 3 });
    expect(f).toEqual({ r1: 11, c1: 3, r2: 12, c2: 3 });
    expect(store.cell(0, 11, 3)?.formula).toBe("A12*10");
    expect(store.cell(0, 12, 3)?.formula).toBe("A13*10");

    // Right: "Week 1" continues across.
    await dragFill(store, 0, { r1: 10, c1: 1, r2: 10, c2: 1 }, { r1: 10, c1: 1, r2: 10, c2: 1 }); // nothing to extend
    store.undo();
    expect(store.cell(0, 12, 3)?.formula).toBeUndefined();
    expect(await dragFill(store, 0, { r1: 10, c1: 1, r2: 10, c2: 1 }, { r1: 10, c1: 1, r2: 10, c2: 1 })).toBeNull();
    await dragFill(store, 0, { r1: 20, c1: 0, r2: 20, c2: 0 }, { r1: 20, c1: 0, r2: 20, c2: 3 });
    await set(30, 1, "Week 1");
    await dragFill(store, 0, { r1: 30, c1: 1, r2: 30, c2: 1 }, { r1: 30, c1: 1, r2: 30, c2: 4 });
    expect([2, 3, 4].map((c) => store.cell(0, 30, c)?.value)).toEqual(["Week 2", "Week 3", "Week 4"]);
  }, 120_000);
});

describe("ranges that span more chunks than the cache holds", () => {
  it("edits, copies, sorts and moves every row of a range over many chunks", async () => {
    const { store } = await open("struct-many", 1, 12000);
    const chunks = store.tabs[0].chunks.length;
    expect(chunks).toBeGreaterThan(8); // the cache keeps 6
    const last = store.tabs[0].rows - 1;
    const rows = Math.min(last, 11_999);

    await store.edit(0, 0, 20, rows, 20, (_cell, row) => ({ value: rows - row }));
    for (let c = 0; c < chunks; c++) await store.load(0, c);
    const missing: number[] = [];
    for (let c = 0; c < chunks; c++) {
      for (const [row, col, cell] of store.cellsOfChunk(0, c)) void row, col, cell;
    }
    for (let r = 0; r <= rows; r += 7) if (store.cell(0, r, 20)?.value !== rows - r) missing.push(r);
    expect(missing).toEqual([]);

    // Copy reads every chunk, not just the ones still cached.
    const block = await copyBlock(store, 0, { r1: 0, c1: 20, r2: rows, c2: 20 });
    expect(block.rows).toBe(rows + 1);
    expect(block.cells.every((row, i) => row[0]?.value === rows - i)).toBe(true);

    // Sort all of it, smallest first.
    await sortArea(store, 0, { r1: 0, c1: 20, r2: rows, c2: 20 }, 20, false);
    const sorted = await copyBlock(store, 0, { r1: 0, c1: 20, r2: rows, c2: 20 });
    expect(sorted.cells.every((row, i) => row[0]?.value === i)).toBe(true);
  }, 240_000);
});

describe("moving cells", () => {
  it("moves values, formats and formulas, replaces what was there, and undoes in one step", async () => {
    const { store } = await open("struct-move");
    await store.load(0, 0);
    await store.edit(0, 10, 0, 10, 0, () => ({ value: "keep", style: { bold: true } }));
    await store.edit(0, 11, 0, 11, 0, () => ({ formula: "A11*2", value: 2 }));
    await store.setCell(0, 20, 5, "old"); // will be overwritten

    const moved = await moveCells(store, 0, { r1: 10, c1: 0, r2: 11, c2: 0 }, { r1: 20, c1: 5, r2: 21, c2: 5 });
    expect(moved).toEqual({ r1: 20, c1: 5, r2: 21, c2: 5 });
    expect(store.cell(0, 20, 5)).toMatchObject({ value: "keep", style: { bold: true } });
    expect(store.cell(0, 21, 5)?.formula).toBe("A11*2"); // references are not changed
    expect(store.cell(0, 10, 0)).toBeNull();
    expect(store.cell(0, 11, 0)).toBeNull();

    store.undo();
    expect(store.cell(0, 10, 0)?.value).toBe("keep");
    expect(store.cell(0, 20, 5)?.value).toBe("old");
    expect(await moveCells(store, 0, { r1: 10, c1: 0, r2: 10, c2: 0 }, { r1: 10, c1: 0, r2: 10, c2: 0 })).toBeNull();

    // Overlapping move: shifting down by one keeps every cell.
    await store.setCell(0, 30, 2, "a");
    await store.setCell(0, 31, 2, "b");
    await moveCells(store, 0, { r1: 30, c1: 2, r2: 31, c2: 2 }, { r1: 31, c1: 2, r2: 32, c2: 2 });
    expect([30, 31, 32].map((r) => store.cell(0, r, 2)?.value ?? null)).toEqual([null, "a", "b"]);
  }, 120_000);
});

describe("volatile formulas and full recalculation", () => {
  it("recognises volatile functions", () => {
    expect(isVolatile("NOW()")).toBe(true);
    expect(isVolatile("A1+today( )")).toBe(true);
    expect(isVolatile("IF(RAND()>0.5,1,0)")).toBe(true);
    expect(isVolatile("SUM(A1:A5)")).toBe(false);
    expect(isVolatile("NOWHERE(1)")).toBe(false);
  });

  it("refreshes NOW() when a saved sheet is read again, without marking it changed", async () => {
    const { store, reopen } = await open("struct-volatile");
    await store.load(0, 0);
    await store.edit(0, 3, 3, 3, 3, () => ({ formula: "NOW()", value: 1 })); // an old, stale result
    const again = await reopen();
    await again.load(0, 0);
    expect(again.cell(0, 3, 3)?.value).toBe(1);
    expect(again.dirtyCount).toBe(0);
    expect(await recalcVolatileInChunk(again, 0, 0)).toBe(1);
    expect(Number(again.cell(0, 3, 3)?.value)).toBeGreaterThan(40_000); // a date serial
    expect(again.dirtyCount).toBe(0); // opening alone leaves the sheet "saved"
  }, 120_000);

  it("recalculates every formula on request", async () => {
    const { store } = await open("struct-recalc-all");
    await store.load(0, 0);
    const index = new FormulaIndex();
    await store.setCell(0, 1, 0, "5");
    // Two formulas whose stored results are stale; nothing changed that would refresh them.
    await store.edit(0, 2, 0, 2, 0, () => ({ formula: "A2*2", value: 999 }));
    await store.edit(0, 3, 0, 3, 0, () => ({ formula: "A3+1", value: 999 }));
    await index.set(store, 0, 2, 0, "A2*2");
    await index.set(store, 0, 3, 0, "A3+1");
    const count = await refreshDependents(store, index, [], undefined, false, { all: true });
    expect(count).toBeGreaterThanOrEqual(2);
    expect(store.cell(0, 2, 0)?.value).toBe(10);
    expect(store.cell(0, 3, 0)?.value).toBe(11); // chained: A3 = A2*2 + 1 = 11
  }, 120_000);
});

describe("copying a tab", () => {
  it("points the copy at the same data without downloading it, and edits stay in their own tab", async () => {
    const { store, reopen } = await open("struct-copy");
    await store.load(0, 0);
    await store.setCell(0, 3, 3, "original");
    expect(() => store.duplicateTab(0, "Copy")).toThrow(/Save your changes first/);
    const saved = await reopen();

    const at = saved.duplicateTab(0, "Copy of Tab");
    expect(at).toBe(1);
    expect(saved.tabs.map((t) => t.name)).toHaveLength(2);
    expect(saved.tabs[1].chunks.map((c) => c.ref)).toEqual(saved.tabs[0].chunks.map((c) => c.ref)); // shared blobs
    expect(new Set([...saved.tabs[0].chunks, ...saved.tabs[1].chunks].map((c) => c.id)).size).toBe(saved.tabs[0].chunks.length * 2);
    await saved.load(1, 0);
    expect(saved.cell(1, 3, 3)?.value).toBe("original");

    await saved.setCell(1, 3, 3, "changed in the copy");
    await saved.load(0, 0);
    expect(saved.cell(0, 3, 3)?.value).toBe("original");

    const again = await reopen(saved);
    expect(again.tabs).toHaveLength(2);
    await again.load(0, 0);
    await again.load(1, 0);
    expect(again.cell(0, 3, 3)?.value).toBe("original");
    expect(again.cell(1, 3, 3)?.value).toBe("changed in the copy");
    expect(again.tabs[1].chunks.length).toBe(again.tabs[0].chunks.length);
  }, 180_000);
});

describe("filtering by colour", () => {
  it("keeps the rows with the chosen fill or text colour", async () => {
    const { store } = await open("struct-filter-color");
    await store.load(0, 0);
    await store.edit(0, 1, 9, 1, 9, () => ({ value: "a", style: { background: "#FF0000", color: "#0000ff" } }));
    await store.edit(0, 2, 9, 2, 9, () => ({ value: "b", style: { background: "#ff0000" } }));
    await store.edit(0, 3, 9, 3, 9, () => ({ value: "c", style: { background: "#00ff00" } }));
    const scan = (await scanColumn(store, 0, 9))!;
    const counts = Object.fromEntries(scan.colorCounts.background.map((c) => [c.value, c.count]));
    expect(counts["#ff0000"]).toBe(2); // colours are compared ignoring case
    expect(counts["#00ff00"]).toBe(1);
    expect(counts[NO_COLOR]).toBe(store.tabs[0].rows - 3);
    const rows = store.tabs[0].rows;
    const redOnly = rowsToHideByColor(scan, rows, true, "background", "#ff0000");
    expect(redOnly).not.toContain(0); // header
    expect(redOnly).not.toContain(1);
    expect(redOnly).not.toContain(2);
    expect(redOnly).toContain(3);
    expect(redOnly).toHaveLength(rows - 1 - 2);
    const blueText = rowsToHideByColor(scan, rows, false, "font", "#0000ff");
    expect(blueText).toHaveLength(rows - 1);
    expect(blueText).not.toContain(1);
    expect(rowsToHideByColor(scan, rows, true, "background", NO_COLOR)).not.toContain(4); // rows with no fill stay
  }, 120_000);
});

describe("person tags and tag pills", () => {
  it("keep their data through save and reload", async () => {
    const { store, reopen } = await open("struct-pills");
    await store.load(0, 0);
    const person = { pubkey: "a".repeat(64), npub: "npub1xyz", name: "Ada", picture: "https://x.test/a.png" };
    await store.edit(0, 2, 1, 2, 1, () => ({ value: "Ada", person }));
    await store.edit(0, 3, 1, 3, 1, () => ({ value: "Urgent", tag: "Urgent" }));
    const again = await reopen();
    await again.load(0, 0);
    expect(again.cell(0, 2, 1)).toMatchObject({ value: "Ada", person });
    expect(again.cell(0, 3, 1)).toMatchObject({ value: "Urgent", tag: "Urgent" });
  }, 120_000);
});

describe("undo for insert, delete and tab changes", () => {
  it("puts back rows after an insert and after a delete that drops whole chunks", async () => {
    const { store, reopen } = await open("struct-undo-rows");
    const before = await snapshot(store);
    const chunks = store.tabs[0].chunks;
    await store.applyEdit(0, { kind: "insert", axis: "row", index: chunks[1].rowStart + 4, count: 3 });
    expect(store.canUndo).toBe(true);
    store.undo();
    expect(store.structuralUndo).toBe(true);
    expect(store.tabs[0].rows).toBe(before.length);
    expect(await snapshot(store)).toEqual(before);

    // A delete across chunks (chunk 2 is dropped whole), then undo.
    const from = chunks[1].rowStart + 3;
    const to = chunks[3].rowStart + 4;
    await store.applyEdit(0, { kind: "delete", axis: "row", index: from, count: to - from });
    expect(store.tabs[0].rows).toBe(before.length - (to - from));
    store.undo();
    expect(store.tabs[0].rows).toBe(before.length);
    expect(await snapshot(store)).toEqual(before);

    // After undoing, saving and reading again gives the original sheet.
    const again = await reopen();
    expect(await snapshot(again)).toEqual(before);
  }, 180_000);

  it("puts back columns, formulas and layout, and older cell edits stay undoable in order", async () => {
    const { store } = await open("struct-undo-cols");
    await store.load(0, 0);
    const index = new FormulaIndex();
    await store.setCell(0, 5, 0, "7");
    await store.edit(0, 6, 1, 6, 1, () => ({ formula: "A6*2", value: 14 }));
    await index.set(store, 0, 6, 1, "A6*2");
    await store.setLayout(0, { columnWidths: { 2: 99 } });
    const before = await snapshot(store);
    const layoutBefore = structuredClone(store.tabs[0].layout);

    // A cell edit, then a structural change: undo goes back through both, in order.
    await store.setCell(0, 5, 3, "edit");
    await editStructure(store, index, 0, { kind: "insert", axis: "col", index: 0, count: 2 });
    expect(store.cell(0, 6, 3)?.formula).toBe("C6*2"); // rewritten and moved
    store.undo();
    expect(store.structuralUndo).toBe(true);
    expect(store.cell(0, 6, 1)?.formula).toBe("A6*2"); // formula text came back with its cell
    expect(store.tabs[0].layout.columnWidths).toEqual(layoutBefore.columnWidths);
    expect(store.cell(0, 5, 3)?.value).toBe("edit");
    store.undo(); // the cell edit
    expect(store.structuralUndo).toBe(false);
    expect(await snapshot(store)).toEqual(before);
  }, 180_000);

  it("undoes added, renamed, recoloured and reordered tabs", async () => {
    const { store } = await open("struct-undo-tabs", 2);
    const names = store.tabs.map((t) => t.name);
    store.renameTab(1, "Budget");
    store.undo();
    expect(store.tabs[1].name).toBe(names[1]);
    store.setTabColor(0, "#ff0000");
    store.undo();
    expect(store.tabMeta(0).color).toBeUndefined();
    store.addTab("Notes");
    expect(store.tabs).toHaveLength(3);
    store.undo();
    expect(store.tabs.map((t) => t.name)).toEqual(names);
    store.remapTabs([1, 0]);
    store.undo();
    expect(store.tabs.map((t) => t.name)).toEqual(names);
    await store.load(1, 0);
    expect(store.cell(1, 0, 0)).toBeDefined();
    store.remapTabs([0]); // delete the second tab
    expect(store.tabs).toHaveLength(1);
    store.undo();
    expect(store.tabs.map((t) => t.name)).toEqual(names);
    await store.load(1, 0);
    expect(store.cell(1, 0, 0)).toBeDefined(); // its data is still there
  }, 180_000);

  it("forgets structural steps once the sheet is saved, and keeps only the latest few", async () => {
    const { store, reopen } = await open("struct-undo-save");
    await store.applyEdit(0, { kind: "insert", axis: "row", index: 2, count: 1 });
    const saved = await reopen();
    expect(saved.canUndo).toBe(false);
    for (let i = 0; i < 7; i++) await store.applyEdit(0, { kind: "insert", axis: "row", index: 1, count: 1 });
    let undone = 0;
    while (store.canUndo) {
      store.undo();
      undone++;
    }
    expect(undone).toBe(5);
  }, 180_000);
});

