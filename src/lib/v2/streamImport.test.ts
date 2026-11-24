import { describe, expect, it } from "vitest";
import { importXlsxToBlossom } from "./streamImport";
import { BigSheetStore } from "./bigSheet";
import { calculateTyped, FormulaIndex, pointRect, refreshDependents } from "./bigFormulas";
import { findCells, replaceCells } from "./bigFind";
import { copyBlock, fillArea, fromTsv, pasteBlock, pasteText } from "./bigClipboard";
import { downloadTabBody } from "./blossom";
import { assertOpenable, parseRootV2, SheetTooLargeToOpenError, type ChunkEntry } from "./root";
import { fakeBlossom, localSigner } from "./testUtils";
import { buildLargeXlsx } from "../xlsxStream/testFixtures";
import { createWorkerHandler } from "../xlsxStream/workerCore";
import type { WorkerLike } from "../xlsxStream/client";
import type { MainToWorker } from "../xlsxStream/protocol";
import type { UploadProgress } from "./progress";

/** The worker logic in this thread, behind the Worker message interface. */
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

describe("importXlsxToBlossom", () => {
  it("streams a workbook to Blossom under one signature and builds a readable root", async () => {
    const bytes = await buildLargeXlsx({ tabs: 2, rows: 3000, cols: 8, sharedStrings: true });
    const fake = fakeBlossom();
    const base = localSigner();
    let signatures = 0;
    const signer = { ...base, signEvent: (t: Parameters<typeof base.signEvent>[0]) => (signatures++, base.signEvent(t)) };
    const phases: UploadProgress["phase"][] = [];
    const servers = ["https://stream-a.test", "https://stream-b.test"];

    const { root, summary } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), {
      title: "Big",
      servers,
      replicas: 2,
      signer,
      env: { fetch: fake.fetch, sleep: async () => {} },
      onProgress: (p) => phases.push(p.phase),
      createWorker: inProcessWorker,
      start: START,
    });

    const chunks = summary.tabs.reduce((sum, tab) => sum + tab.chunks, 0);
    expect(chunks).toBeGreaterThan(4);
    expect(signatures).toBe(1);
    // Every chunk on both servers, nothing else uploaded.
    expect(fake.uploads).toHaveLength(chunks * 2);
    expect(phases).toContain("uploading");
    expect(phases[phases.length - 1]).toBe("done");

    expect(root.title).toBe("Big");
    expect(root.tabs).toHaveLength(2);
    const lists = root.tabs.map((tab) => tab.chunks as ChunkEntry[]);
    lists.forEach((list, i) => {
      expect(list.reduce((sum, entry) => sum + (entry.cells ?? 0), 0)).toBe(summary.tabs[i].cells);
      expect(list[0].rowStart).toBe(0);
    });

    // Reads back as the same cells.
    const doc = await parseRootV2(root, (ref, rootServers, type) =>
      downloadTabBody(ref, rootServers, { fetch: fake.fetch }, type),
    );
    doc.sheets.forEach((sheet, i) => {
      expect(Object.keys(sheet.cells)).toHaveLength(summary.tabs[i].cells);
      expect(sheet.name).toBe(summary.tabs[i].name);
    });

    // The open gate reads the counts from the root alone.
    expect(() => assertOpenable(root)).not.toThrow();
    expect(() => assertOpenable(root, { maxCells: 1000, maxBytes: Infinity })).toThrow(SheetTooLargeToOpenError);
  }, 60_000);

  it("uploads nothing and asks for no signature when the file is not a workbook", async () => {
    const fake = fakeBlossom();
    const base = localSigner();
    let signatures = 0;
    const signer = { ...base, signEvent: (t: Parameters<typeof base.signEvent>[0]) => (signatures++, base.signEvent(t)) };
    await expect(
      importXlsxToBlossom(new Blob([new Uint8Array([1, 2, 3, 4])]), {
        title: "Bad",
        servers: ["https://stream-c.test"],
        signer,
        env: { fetch: fake.fetch },
        createWorker: inProcessWorker,
        start: START,
      }),
    ).rejects.toThrow();
    expect(signatures).toBe(0);
    expect(fake.uploads).toHaveLength(0);
  });

  it("edits a big sheet chunk by chunk and saves only the changed chunk under one signature", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 3000, cols: 8, sharedStrings: false });
    const fake = fakeBlossom();
    const base = localSigner();
    let signatures = 0;
    const signer = { ...base, signEvent: (t: Parameters<typeof base.signEvent>[0]) => (signatures++, base.signEvent(t)) };
    const servers = ["https://big-edit.test"];
    const env = { fetch: fake.fetch, sleep: async () => {} };
    const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), {
      title: "Edit", servers, signer, env, createWorker: inProcessWorker, start: START,
    });
    const store = new BigSheetStore(root, servers, env);
    await store.open();
    const tab = store.tabs[0];
    expect(tab.chunks.length).toBeGreaterThan(2);
    const row = tab.chunks[2].rowStart + 5;
    await store.load(0, store.chunkFor(0, row));
    await store.setCell(0, row, 3, "1234.5");
    await store.setCell(0, row, 4, "hello");
    expect(store.dirtyCount).toBe(1);

    signatures = 0;
    const before = fake.uploads.length;
    const next = await store.save({ servers, signer, ...env });
    expect(signatures).toBe(1);
    expect(fake.uploads.length - before).toBe(1);

    const again = new BigSheetStore(next, servers, env);
    await again.open();
    await again.load(0, again.chunkFor(0, row));
    expect(again.cell(0, row, 3)?.value).toBe(1234.5);
    expect(again.cell(0, row, 4)?.value).toBe("hello");
  }, 60_000);

  it("copies, pastes, fills and undoes with formulas moving with their cells", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 3000, cols: 8, sharedStrings: false });
    const fake = fakeBlossom();
    const signer = localSigner();
    const servers = ["https://big-clip.test"];
    const env = { fetch: fake.fetch, sleep: async () => {} };
    const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), {
      title: "Clip", servers, signer, env, createWorker: inProcessWorker, start: START,
    });
    const store = new BigSheetStore(root, servers, env);
    await store.open();
    const far = store.tabs[0].chunks[2].rowStart + 5; // a row in another chunk
    await store.load(0, 0);
    await store.load(0, store.chunkFor(0, far));

    // A1 = 4, B1 = A1*2 (typed with its result), a styled cell
    await store.setCell(0, 0, 0, "4");
    await store.edit(0, 0, 1, 0, 1, () => ({ formula: "A1*2", value: 8, style: { bold: true } }));

    // Copy A1:B1 and paste at A{far}: the formula now reads A{far}.
    const block = await copyBlock(store, 0, { r1: 0, c1: 0, r2: 0, c2: 1 });
    expect(block.text).toBe("4\t8");
    const index = new FormulaIndex();
    const done = await pasteBlock(store, 0, far, 0, block);
    expect(done).toEqual({ r1: far, c1: 0, r2: far, c2: 1 });
    expect(store.cell(0, far, 1)?.formula).toBe(`A${far + 1}*2`);
    expect(store.cell(0, far, 1)?.style).toEqual({ bold: true });
    await refreshDependents(store, index, [{ tab: 0, ...done }], undefined, true);
    expect(store.cell(0, far, 1)?.value).toBe(8);
    await store.setCell(0, far, 0, "10");
    await refreshDependents(store, index, [pointRect({ tab: 0, row: far, col: 0 })]);
    expect(store.cell(0, far, 1)?.value).toBe(20);

    // Absolute references stay put.
    await store.edit(0, 1, 1, 1, 1, () => ({ formula: "$A$1+1", value: 5 }));
    const abs = await copyBlock(store, 0, { r1: 1, c1: 1, r2: 1, c2: 1 });
    await pasteBlock(store, 0, 2, 3, abs);
    expect(store.cell(0, 2, 3)?.formula).toBe("$A$1+1");

    // One typed value pasted over a selection fills it all, as one undo step.
    const before = store.canUndo;
    const filled = await pasteText(store, 0, 5, 5, [["7"]], { r1: 5, c1: 5, r2: 6, c2: 6 });
    expect(filled).toEqual({ r1: 5, c1: 5, r2: 6, c2: 6 });
    expect([store.cell(0, 5, 5)?.value, store.cell(0, 6, 6)?.value]).toEqual([7, 7]);
    store.undo();
    expect(store.cell(0, 5, 5)?.value).not.toBe(7);
    expect(before).toBe(true);

    // Text from outside: rows and columns, numbers become numbers, "=" starts a formula.
    await pasteText(store, 0, 10, 0, fromTsv("1\t2\n3\t=A11+B11"));
    expect(store.cell(0, 10, 1)?.value).toBe(2);
    expect(store.cell(0, 11, 1)?.formula).toBe("A11+B11");

    // Ctrl+D: the first row of the selection goes down the rest, formulas shifted.
    await store.edit(0, 20, 0, 20, 1, (_c, _r, col) => (col === 0 ? { value: 1 } : { formula: "A21+1", value: 2 }));
    await fillArea(store, 0, { r1: 20, c1: 0, r2: 23, c2: 1 }, "down");
    expect(store.cell(0, 23, 0)?.value).toBe(1);
    expect(store.cell(0, 23, 1)?.formula).toBe("A24+1");
    // With one row selected it copies the row above.
    expect(await fillArea(store, 0, { r1: 25, c1: 0, r2: 25, c2: 1 }, "down")).not.toBeNull();
    expect(await fillArea(store, 0, { r1: 0, c1: 0, r2: 0, c2: 1 }, "down")).toBeNull();
    // Ctrl+R
    await fillArea(store, 0, { r1: 20, c1: 1, r2: 20, c2: 3 }, "right");
    expect(store.cell(0, 20, 3)?.formula).toBe("C21+1");

    await expect(pasteText(store, 0, 0, 16383, [["a", "b"]])).rejects.toThrow(/at most 16,384 columns/);
  }, 60_000);

  it("keeps column widths, row heights, hidden rows and frozen panes through save and reload", async () => {
    const bytes = await buildLargeXlsx({ tabs: 2, rows: 3000, cols: 8, sharedStrings: false });
    const fake = fakeBlossom();
    const signer = localSigner();
    const servers = ["https://big-layout.test"];
    const env = { fetch: fake.fetch, sleep: async () => {} };
    const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), {
      title: "Layout", servers, signer, env, createWorker: inProcessWorker, start: START,
    });
    const store = new BigSheetStore(root, servers, env);
    await store.open();
    await store.setLayout(1, { columnWidths: { 2: 220 }, rowHeights: { 5: 60 }, hiddenRows: [7], frozen: { rows: 1, columns: 2 }, zoom: 1.5, merges: [{ r: 3, c: 1, rs: 2, cs: 2 }], borders: [{ r1: 0, c1: 0, r2: 1, c2: 1, type: "all", color: "#ff0000" }] });
    expect(store.dirtyCount).toBe(1);

    const next = await store.save({ servers, signer, ...env });
    const again = new BigSheetStore(next, servers, env);
    await again.open();
    expect(again.tabs[1].layout).toEqual({ columnWidths: { 2: 220 }, rowHeights: { 5: 60 }, hiddenRows: [7], frozen: { rows: 1, columns: 2 }, zoom: 1.5, merges: [{ r: 3, c: 1, rs: 2, cs: 2 }], borders: [{ r1: 0, c1: 0, r2: 1, c2: 1, type: "all", color: "#ff0000" }] });
    expect(again.tabs[0].layout).toEqual({});

    // Unfreezing removes the setting.
    await again.setLayout(1, { frozen: undefined });
    const third = new BigSheetStore(await again.save({ servers, signer, ...env }), servers, env);
    await third.open();
    expect(third.tabs[1].layout.frozen).toBeUndefined();
    expect(third.tabs[1].layout.columnWidths).toEqual({ 2: 220 });
  }, 60_000);

  it("finds and replaces text across every chunk of a big sheet", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 3000, cols: 8, sharedStrings: false });
    const fake = fakeBlossom();
    const signer = localSigner();
    const servers = ["https://big-find.test"];
    const env = { fetch: fake.fetch, sleep: async () => {} };
    const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), {
      title: "Find", servers, signer, env, createWorker: inProcessWorker, start: START,
    });
    const store = new BigSheetStore(root, servers, env);
    await store.open();
    expect(store.tabs[0].chunks.length).toBeGreaterThan(2);
    const lastRow = store.tabs[0].rows - 1;
    await store.load(0, store.chunkFor(0, lastRow));
    await store.setCell(0, lastRow, 2, "needle one");
    await store.load(0, 0);
    await store.setCell(0, 3, 1, "Needle two");

    const progress: number[] = [];
    const found = await findCells(store, 0, "needle", {}, (done) => progress.push(done));
    expect(found.matches).toEqual([{ row: 3, col: 1 }, { row: lastRow, col: 2 }]);
    expect(progress.at(-1)).toBe(store.tabs[0].chunks.length);
    expect((await findCells(store, 0, "needle", { matchCase: true })).matches).toEqual([{ row: lastRow, col: 2 }]);
    expect((await findCells(store, 0, "needle one", { wholeCell: true })).matches).toHaveLength(1);
    expect((await findCells(store, 0, "needle", { wholeCell: true })).matches).toHaveLength(0);

    const changed = await replaceCells(store, 0, found.matches, "needle", "pin");
    expect(changed).toHaveLength(2);
    expect(store.cell(0, 3, 1)?.value).toBe("pin two");
    expect(store.cell(0, lastRow, 2)?.value).toBe("pin one");
    expect((await findCells(store, 0, "needle")).matches).toHaveLength(0);
  }, 60_000);

  it("keeps formulas, their results and errors through save and reload, and refreshes them after reload", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 3000, cols: 8, sharedStrings: false });
    const fake = fakeBlossom();
    const signer = localSigner();
    const servers = ["https://big-formula.test"];
    const env = { fetch: fake.fetch, sleep: async () => {} };
    const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), {
      title: "Formulas", servers, signer, env, createWorker: inProcessWorker, start: START,
    });
    const store = new BigSheetStore(root, servers, env);
    await store.open();
    const row = store.tabs[0].chunks[2].rowStart + 5;
    await store.load(0, store.chunkFor(0, row));
    await store.setCell(0, row, 3, "10");

    const index = new FormulaIndex();
    const typed = async (col: number, formula: string) => {
      const result = await calculateTyped(store, 0, row, col, formula);
      await store.edit(0, row, col, row, col, () => ({
        formula,
        value: result.value,
        ...(result.error ? { type: "error" as const } : {}),
      }));
      await index.set(store, 0, row, col, formula);
    };
    const name = (n: number) => String.fromCharCode(65 + n);
    await typed(5, `${name(3)}${row + 1}*2`); // F = D*2
    await typed(6, "1/0"); // G = #DIV/0!
    expect(store.cell(0, row, 5)?.value).toBe(20);

    const next = await store.save({ servers, signer, ...env });
    const again = new BigSheetStore(next, servers, env);
    await again.open();
    await again.load(0, again.chunkFor(0, row));
    expect(again.cell(0, row, 5)).toMatchObject({ formula: `${name(3)}${row + 1}*2`, value: 20 });
    expect(again.cell(0, row, 6)).toMatchObject({ formula: "1/0", value: "#DIV/0!", type: "error" });

    // A fresh session knows nothing about the formulas; an edit still finds and refreshes them.
    const fresh = new FormulaIndex();
    await again.setCell(0, row, 3, "50");
    await refreshDependents(again, fresh, [pointRect({ tab: 0, row, col: 3 })]);
    expect(again.cell(0, row, 5)?.value).toBe(100);
  }, 60_000);
});
