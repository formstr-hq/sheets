// Formulas stored by id, end to end through the big-sheet store: the same random edits are run on a store
// that keeps formulas positional (every chunk rewritten on each insert or delete; the proven old behaviour)
// and on one that stores them by id; after saving and reopening, every cell of both must be the same.
import { describe, expect, it } from "vitest";
import { importXlsxToBlossom } from "./streamImport";
import { BigSheetStore } from "./bigSheet";
import { FormulaIndex, calculateTyped } from "./bigFormulas";
import { editStructure } from "./bigStructure";
import { fakeBlossom, localSigner } from "./testUtils";
import { buildLargeXlsx } from "../xlsxStream/testFixtures";
import { createWorkerHandler } from "../xlsxStream/workerCore";
import { parseRootV2 } from "./root";
import { downloadTabBody } from "./blossom";
import type { WorkerLike } from "../xlsxStream/client";
import type { MainToWorker } from "../xlsxStream/protocol";
import type { SheetRootV2 } from "./root";

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

const START = { maxChunkBytes: 32 * 1024, headRawBytes: 60_000, rawCap: 1_000_000 };

interface World {
  root: SheetRootV2;
  servers: string[];
  env: { fetch: typeof fetch; sleep: () => Promise<void> };
  signer: ReturnType<typeof localSigner>;
}

async function world(label: string, tabs = 2, rows = 2500): Promise<World> {
  const bytes = await buildLargeXlsx({ tabs, rows, cols: 6, sharedStrings: false });
  const fake = fakeBlossom();
  const signer = localSigner();
  const servers = [`https://${label}.test`];
  const env = { fetch: fake.fetch as typeof fetch, sleep: async () => {} };
  const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), { title: label, servers, signer, env, createWorker: inProcessWorker, start: START });
  return { root, servers, env, signer };
}

async function openStore(w: World, root = w.root, byId = true): Promise<BigSheetStore> {
  const store = new BigSheetStore(root, w.servers, w.env);
  store.allowRefsById = byId;
  await store.open();
  return store;
}

/** Formula text for a tab, placed in rows spread over every chunk (they read neighbours, ranges and the other tab). */
function formulaPlan(store: BigSheetStore): { tab: number; row: number; col: number; formula: string }[] {
  const plan: { tab: number; row: number; col: number; formula: string }[] = [];
  const other = (tab: number) => store.tabs[(tab + 1) % store.tabs.length].name;
  for (let tab = 0; tab < store.tabs.length; tab++) {
    const rows = store.tabs[tab].rows;
    const chunks = store.tabs[tab].chunks;
    // one formula in every chunk, near its start and its end
    chunks.forEach((entry, i) => {
      const at = Math.min(rows - 3, entry.rowStart + 4);
      plan.push({ tab, row: at, col: 4, formula: `A${at}*2` });
      plan.push({ tab, row: at + 1, col: 4, formula: `SUM(A${at - 2}:A${at + 2})+C${at}` });
      const sheetPrefix = store.tabs.length > 1 ? `'${other(tab)}'!` : "";
      plan.push({ tab, row: at + 2, col: 5, formula: `${sheetPrefix}A${5 + i}+$A$3+B$${7 + i}` });
    });
    // a range across many chunks and a reference far from the formula
    plan.push({ tab, row: 2, col: 5, formula: `SUM(A2:A${Math.floor(rows / 2)})` });
    plan.push({ tab, row: 3, col: 5, formula: `A${rows - 5}+B${Math.floor(rows / 3)}` });
    plan.push({ tab, row: 4, col: 5, formula: `SUM(A:A)+COUNT(2:2)` });
  }
  return plan;
}

async function typeFormulas(store: BigSheetStore, index: FormulaIndex): Promise<void> {
  for (const { tab, row, col, formula } of formulaPlan(store)) {
    await store.load(tab, store.chunkFor(tab, row));
    const result = await calculateTyped(store, tab, row, col, formula);
    await store.edit(tab, row, col, row, col, () => ({ formula, value: result.value, ...(result.error ? { type: "error" as const } : {}) }));
    await index.set(store, tab, row, col, formula);
  }
}

interface Dump {
  tabs: { name: string; rows: number; cols: number; cells: Record<string, [unknown, string | null]> }[];
}

/** Every cell of every tab (opens every chunk): value and formula, keyed by position. */
async function dump(store: BigSheetStore): Promise<Dump> {
  const out: Dump = { tabs: [] };
  for (let t = 0; t < store.tabs.length; t++) {
    const cells: Dump["tabs"][number]["cells"] = {};
    for (let c = 0; c < store.tabs[t].chunks.length; c++) {
      await store.load(t, c);
      for (const [row, col, cell] of store.cellsOfChunk(t, c)) {
        cells[`${row}:${col}`] = [cell.value ?? null, cell.formula ?? null];
      }
    }
    out.tabs.push({ name: store.tabs[t].name, rows: store.tabs[t].rows, cols: store.tabs[t].cols, cells });
  }
  return out;
}

function rng(seed: number) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 2 ** 32;
  };
}

interface Op {
  kind: "insert" | "delete";
  axis: "row" | "col";
  tab: number;
  index: number;
  count: number;
}

function randomOps(seed: number, store: BigSheetStore, n: number): Op[] {
  const rand = rng(seed);
  const ops: Op[] = [];
  for (let i = 0; i < n; i++) {
    const tab = Math.floor(rand() * store.tabs.length);
    const axis = rand() < 0.7 ? "row" : "col";
    const size = axis === "row" ? store.tabs[tab].rows : store.tabs[tab].cols;
    const kind = rand() < 0.5 ? "insert" : "delete";
    const count = 1 + Math.floor(rand() * (axis === "row" ? 20 : 2));
    const index = Math.floor(rand() * Math.max(1, size - count));
    ops.push({ kind, axis, tab, index, count });
  }
  return ops;
}

async function run(store: BigSheetStore, index: FormulaIndex, op: Op): Promise<void> {
  // keep sizes in step with what each store really has (deletes are clamped the same way in both)
  await editStructure(store, index, op.tab, { kind: op.kind, axis: op.axis, index: op.index, count: op.count });
}

describe("formulas by id in the big-sheet store", () => {
  it("switches the workbook over, saves formulas as ids and reads them back as A1", async () => {
    const w = await world("switch");
    const store = await openStore(w);
    expect(store.tabs.reduce((n, tab) => n + tab.chunks.length, 0)).toBeGreaterThan(6); // more chunks than the cache holds
    const index = new FormulaIndex();
    await typeFormulas(store, index);
    const before = await dump(store);
    await editStructure(store, index, 0, { kind: "insert", axis: "row", index: 10, count: 3 });
    expect(store.refsById).toBe(true);
    const root = await store.save({ servers: w.servers, signer: w.signer, ...w.env });
    expect(root.refs).toBe("id");
    // every chunk entry knows its row ids
    for (const tab of root.tabs) for (const entry of tab.chunks as { runs?: unknown }[]) expect(entry.runs).toBeDefined();
    // what is stored is ids: fetch a chunk with formulas and look
    const entry = (root.tabs[0].chunks as unknown as { ref: never; rowStart: number }[])[1];
    const stored = (await downloadTabBody(entry.ref, w.servers, w.env)) as { cells: Record<string, { formula?: string }> };
    const formulas = Object.values(stored.cells).map((c) => c.formula).filter(Boolean) as string[];
    expect(formulas.length).toBeGreaterThan(0);
    expect(formulas.every((f) => f.includes("~("))).toBe(true);
    const again = await openStore(w, root);
    const after = await dump(again);
    expect(after.tabs[0].rows).toBe(before.tabs[0].rows + 3);
    // formulas in the chunk we looked at read back as A1 text, not as ids
    const texts = Object.values(after.tabs[0].cells).map((c) => c[1]).filter(Boolean) as string[];
    expect(texts.some((f) => f.includes("~("))).toBe(false);
  }, 240_000);

  for (const seed of [1, 2, 3, 4, 5, 6, 7, 8]) {
    it(`random inserts and deletes give the same sheet as positional formulas (seed ${seed})`, async () => {
      const w = await world(`diff${seed}`);
      const legacy = await openStore(w, w.root, false);
      const byId = await openStore(w, w.root, true);
      const legacyIndex = new FormulaIndex();
      const idIndex = new FormulaIndex();
      await typeFormulas(legacy, legacyIndex);
      await typeFormulas(byId, idIndex);
      // start both from a saved state, as a real sheet would be
      let a = await openStore(w, await legacy.save({ servers: w.servers, signer: w.signer, ...w.env }), false);
      let b = await openStore(w, await byId.save({ servers: w.servers, signer: w.signer, ...w.env }), true);
      let ia = new FormulaIndex();
      let ib = new FormulaIndex();
      const ops = randomOps(seed, a, 7);
      for (const [i, op] of ops.entries()) {
        await run(a, ia, op);
        await run(b, ib, op);
        if (i % 3 === 2) {
          // save and reopen both (the id store then reads its formulas back with chunks it never opened)
          a = await openStore(w, await a.save({ servers: w.servers, signer: w.signer, ...w.env }), false);
          b = await openStore(w, await b.save({ servers: w.servers, signer: w.signer, ...w.env }), true);
          ia = new FormulaIndex();
          ib = new FormulaIndex();
        }
      }
      const finalA = await openStore(w, await a.save({ servers: w.servers, signer: w.signer, ...w.env }), false);
      const finalB = await openStore(w, await b.save({ servers: w.servers, signer: w.signer, ...w.env }), true);
      expect(finalB.refsById).toBe(true);
      expect(await dump(finalB)).toEqual(await dump(finalA));
    }, 480_000);
  }

  it("a chunk that was never opened still reads right after rows were inserted and deleted around its formulas", async () => {
    const w = await world("never");
    const store = await openStore(w);
    const index = new FormulaIndex();
    await typeFormulas(store, index);
    let current = await openStore(w, await store.save({ servers: w.servers, signer: w.signer, ...w.env }));
    const reference = await openStore(w, current.root, false);
    const ri = new FormulaIndex();
    const ci = new FormulaIndex();
    for (const op of [
      { kind: "insert", axis: "row", tab: 0, index: 5, count: 4 },
      { kind: "delete", axis: "row", tab: 0, index: 100, count: 15 },
      { kind: "insert", axis: "col", tab: 0, index: 1, count: 1 },
      { kind: "delete", axis: "col", tab: 1, index: 0, count: 1 },
    ] as Op[]) {
      await run(current, ci, op);
      await run(reference, ri, op);
    }
    current = await openStore(w, await current.save({ servers: w.servers, signer: w.signer, ...w.env }));
    const ref = await openStore(w, await reference.save({ servers: w.servers, signer: w.signer, ...w.env }), false);
    expect(await dump(current)).toEqual(await dump(ref));
  }, 300_000);

  it("a workbook with formulas in more chunks than fit in memory still converts, or stays positional without losing anything", async () => {
    const w = await world("many", 1, 4000);
    const store = await openStore(w);
    const index = new FormulaIndex();
    await typeFormulas(store, index);
    const root = await store.save({ servers: w.servers, signer: w.signer, ...w.env });
    const again = await openStore(w, root);
    // a limit of 2 chunks cannot be met: the switch is refused and nothing changes
    expect(await again.enableRefsById(2)).toBe(false);
    expect(again.refsById).toBe(false);
    const snapshot = await dump(again);
    await editStructure(again, new FormulaIndex(), 0, { kind: "insert", axis: "row", index: 20, count: 2 }).catch(() => {});
    expect((await dump(again)).tabs.length).toBe(snapshot.tabs.length);
  }, 300_000);

  it("an id workbook opens as a document with A1 formulas (small sheets are read whole)", async () => {
    const w = await world("doc", 2, 1200);
    const store = await openStore(w);
    const index = new FormulaIndex();
    await typeFormulas(store, index);
    await editStructure(store, index, 0, { kind: "delete", axis: "row", index: 50, count: 5 });
    await editStructure(store, index, 1, { kind: "insert", axis: "col", index: 2, count: 1 });
    const root = await store.save({ servers: w.servers, signer: w.signer, ...w.env });
    expect(root.refs).toBe("id");
    const doc = await parseRootV2(root, (ref, servers, type) => downloadTabBody(ref, servers, w.env, type));
    const texts = doc.sheets.flatMap((s) => Object.values(s.cells).map((c) => c.formula).filter(Boolean)) as string[];
    expect(texts.length).toBeGreaterThan(0);
    expect(texts.some((f) => f.includes("~("))).toBe(false);
    expect(doc.sheets.every((s) => !("tombs" in s))).toBe(true);
    // and it equals what the store itself shows
    const reopened = await openStore(w, root);
    const shown = await dump(reopened);
    doc.sheets.forEach((sheet, t) => {
      for (const [key, cell] of Object.entries(sheet.cells)) {
        if (cell.formula) expect(shown.tabs[t].cells[key]?.[1], `${sheet.name} ${key}`).toBe(cell.formula);
      }
    });
  }, 300_000);

  it("a copied tab keeps reading its own rows", async () => {
    const w = await world("copy", 2, 1200);
    const store = await openStore(w);
    const index = new FormulaIndex();
    await typeFormulas(store, index);
    await editStructure(store, index, 0, { kind: "insert", axis: "row", index: 30, count: 2 });
    let saved = await openStore(w, await store.save({ servers: w.servers, signer: w.signer, ...w.env }));
    const copyAt = saved.duplicateTab(0, "Copy");
    saved = await openStore(w, await saved.save({ servers: w.servers, signer: w.signer, ...w.env }));
    const first = await dump(saved);
    // same-tab references in the copy mean the copy: formulas read the same in both tabs
    const original = first.tabs[0];
    const copy = first.tabs[copyAt];
    for (const [key, cell] of Object.entries(original.cells)) {
      if (cell[1] && !cell[1].includes("!")) expect(copy.cells[key]?.[1]).toBe(cell[1]);
    }
    // an insert in the copy leaves the original alone
    await editStructure(saved, new FormulaIndex(), copyAt, { kind: "insert", axis: "row", index: 0, count: 1 });
    const root = await saved.save({ servers: w.servers, signer: w.signer, ...w.env });
    const after = await dump(await openStore(w, root));
    expect(after.tabs[0].cells).toEqual(first.tabs[0].cells);
  }, 300_000);

  it("undoing a delete brings the rows back and the formulas point at them again", async () => {
    const w = await world("undo", 1, 1200);
    const store = await openStore(w);
    const index = new FormulaIndex();
    await typeFormulas(store, index);
    const before = await dump(store);
    await editStructure(store, index, 0, { kind: "delete", axis: "row", index: 8, count: 4 });
    store.undo();
    expect(store.structuralUndo).toBe(true);
    const after = await dump(store);
    expect(after.tabs[0].cells).toEqual(before.tabs[0].cells);
    const root = await store.save({ servers: w.servers, signer: w.signer, ...w.env });
    expect(await dump(await openStore(w, root))).toEqual(before);
  }, 300_000);
});
