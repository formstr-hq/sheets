// Adversarial tests for formulas stored by id: persistence, migration and lifecycle edge cases.
// Reference = the positional store (allowRefsById = false). Both are saved, reopened and dumped.
import { describe, expect, it } from "vitest";
import { importXlsxToBlossom } from "./streamImport";
import { BigSheetStore } from "./bigSheet";
import { FormulaIndex, calculateTyped } from "./bigFormulas";
import { editStructure, renameTabAndFormulas } from "./bigStructure";
import { fakeBlossom, localSigner, type FakeBlossom } from "./testUtils";
import { buildLargeXlsx } from "../xlsxStream/testFixtures";
import { createWorkerHandler } from "../xlsxStream/workerCore";
import { parseRootV2, FRAME_TYPE_TAB_INDEX } from "./root";
import { downloadTabBody, encryptFrame, uploadStaged } from "./blossom";
import { frame } from "./codec";
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
  fake: FakeBlossom;
  failPut: { on: boolean };
  env: { fetch: typeof fetch; sleep: () => Promise<void> };
  signer: ReturnType<typeof localSigner>;
}

async function world(label: string, tabs = 2, rows = 1500): Promise<World> {
  const bytes = await buildLargeXlsx({ tabs, rows, cols: 6, sharedStrings: false });
  const fake = fakeBlossom();
  const signer = localSigner();
  const servers = [`https://${label}.test`];
  const failPut = { on: false };
  const env = {
    fetch: (async (input: RequestInfo | URL, init?: RequestInit) => {
      if (failPut.on && (init?.method ?? "GET").toUpperCase() === "PUT") return new Response("nope", { status: 413 });
      return fake.fetch(input, init);
    }) as typeof fetch,
    sleep: async () => {},
  };
  const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), { title: label, servers, signer, env, createWorker: inProcessWorker, start: START });
  return { root, servers, fake, failPut, env, signer };
}

async function openStore(w: World, root: SheetRootV2, byId = true): Promise<BigSheetStore> {
  const store = new BigSheetStore(root, w.servers, w.env);
  store.allowRefsById = byId;
  await store.open();
  return store;
}

const save = (w: World, store: BigSheetStore) => store.save({ servers: w.servers, signer: w.signer, ...w.env });

function formulaPlan(store: BigSheetStore): { tab: number; row: number; col: number; formula: string }[] {
  const plan: { tab: number; row: number; col: number; formula: string }[] = [];
  const other = (tab: number) => store.tabs[(tab + 1) % store.tabs.length].name;
  for (let tab = 0; tab < store.tabs.length; tab++) {
    const rows = store.tabs[tab].rows;
    store.tabs[tab].chunks.forEach((entry, i) => {
      const at = Math.min(rows - 3, entry.rowStart + 4);
      plan.push({ tab, row: at, col: 4, formula: `A${at}*2` });
      plan.push({ tab, row: at + 1, col: 4, formula: `SUM(A${at - 2}:A${at + 2})+C${at}` });
      const sheetPrefix = store.tabs.length > 1 ? `'${other(tab)}'!` : "";
      plan.push({ tab, row: at + 2, col: 5, formula: `${sheetPrefix}A${5 + i}+$A$3+B$${7 + i}` });
    });
    plan.push({ tab, row: 2, col: 5, formula: `SUM(A2:A${Math.floor(rows / 2)})` });
    plan.push({ tab, row: 3, col: 5, formula: `A${rows - 5}+B${Math.floor(rows / 3)}` });
    plan.push({ tab, row: 4, col: 5, formula: `SUM(A:A)+COUNT(2:2)` });
  }
  return plan;
}

async function putFormula(store: BigSheetStore, index: FormulaIndex, tab: number, row: number, col: number, formula: string): Promise<void> {
  await store.load(tab, store.chunkFor(tab, row));
  const result = await calculateTyped(store, tab, row, col, formula);
  await store.edit(tab, row, col, row, col, () => ({ formula, value: result.value, ...(result.error ? { type: "error" as const } : {}) }));
  await index.set(store, tab, row, col, formula);
}

async function typeFormulas(store: BigSheetStore, index: FormulaIndex, extra: { tab: number; row: number; col: number; formula: string }[] = []): Promise<void> {
  for (const { tab, row, col, formula } of [...formulaPlan(store), ...extra]) await putFormula(store, index, tab, row, col, formula);
}

interface Dump {
  tabs: { name: string; rows: number; cols: number; cells: Record<string, [unknown, string | null]> }[];
}

async function dump(store: BigSheetStore): Promise<Dump> {
  const out: Dump = { tabs: [] };
  for (let t = 0; t < store.tabs.length; t++) {
    const cells: Dump["tabs"][number]["cells"] = {};
    for (let c = 0; c < store.tabs[t].chunks.length; c++) {
      await store.load(t, c);
      for (const [row, col, cell] of store.cellsOfChunk(t, c)) cells[`${row}:${col}`] = [cell.value ?? null, cell.formula ?? null];
    }
    out.tabs.push({ name: store.tabs[t].name, rows: store.tabs[t].rows, cols: store.tabs[t].cols, cells });
  }
  return out;
}

/** Diff two dumps into a readable list (first few differences). */
function diff(a: Dump, b: Dump): string[] {
  const out: string[] = [];
  if (a.tabs.length !== b.tabs.length) out.push(`tab count ${a.tabs.length} vs ${b.tabs.length}`);
  for (let t = 0; t < Math.min(a.tabs.length, b.tabs.length); t++) {
    const x = a.tabs[t];
    const y = b.tabs[t];
    if (x.name !== y.name) out.push(`tab ${t} name ${x.name} vs ${y.name}`);
    if (x.rows !== y.rows) out.push(`tab ${t} rows ${x.rows} vs ${y.rows}`);
    if (x.cols !== y.cols) out.push(`tab ${t} cols ${x.cols} vs ${y.cols}`);
    const keys = new Set([...Object.keys(x.cells), ...Object.keys(y.cells)]);
    for (const k of keys) {
      if (JSON.stringify(x.cells[k]) !== JSON.stringify(y.cells[k])) out.push(`tab ${t} cell ${k}: ${JSON.stringify(x.cells[k])} vs ${JSON.stringify(y.cells[k])}`);
      if (out.length > 12) return out;
    }
  }
  return out;
}

type Scenario = (ctx: { w: World; store: BigSheetStore; index: FormulaIndex; mode: "legacy" | "id" }) => Promise<BigSheetStore | void>;

interface Twin {
  w: World;
  legacy: BigSheetStore;
  byId: BigSheetStore;
  li: FormulaIndex;
  bi: FormulaIndex;
}

/** Both stores start from the same saved, positional workbook with the same formulas. */
async function twin(label: string, opts: { tabs?: number; rows?: number; extra?: Parameters<typeof typeFormulas>[2]; prep?: (w: World) => Promise<void> } = {}): Promise<Twin> {
  const w = await world(label, opts.tabs ?? 2, opts.rows ?? 1500);
  const a = await openStore(w, structuredClone(w.root), false);
  const b = await openStore(w, structuredClone(w.root), true);
  await typeFormulas(a, new FormulaIndex(), opts.extra);
  await typeFormulas(b, new FormulaIndex(), opts.extra);
  const legacy = await openStore(w, await save(w, a), false);
  const byId = await openStore(w, await save(w, b), true);
  return { w, legacy, byId, li: new FormulaIndex(), bi: new FormulaIndex() };
}

/** Runs the scenario on both, saves, reopens, and returns the differences between the two sheets. */
async function compare(t: Twin, scenario: Scenario): Promise<{ problems: string[]; id: BigSheetStore; legacy: BigSheetStore }> {
  const la = (await scenario({ w: t.w, store: t.legacy, index: t.li, mode: "legacy" })) ?? t.legacy;
  const ib = (await scenario({ w: t.w, store: t.byId, index: t.bi, mode: "id" })) ?? t.byId;
  const finalA = await openStore(t.w, structuredClone(await save(t.w, la)), false);
  const finalB = await openStore(t.w, structuredClone(await save(t.w, ib)), true);
  return { problems: diff(await dump(finalA), await dump(finalB)), id: finalB, legacy: finalA };
}

const op = (tab: number, kind: "insert" | "delete", axis: "row" | "col", index: number, count: number) => ({ tab, edit: { kind, axis, index, count } as const });
async function doOps(store: BigSheetStore, index: FormulaIndex, ops: ReturnType<typeof op>[]) {
  for (const o of ops) await editStructure(store, index, o.tab, o.edit);
}

describe("1. failed save after the switch-over", () => {
  it("a save that fails (413) then succeeds loses nothing", async () => {
    const t = await twin("fail");
    const r = await compare(t, async ({ w, store, index, mode }) => {
      await doOps(store, index, [op(0, "insert", "row", 10, 3), op(0, "delete", "row", 400, 5)]);
      if (mode === "id") {
        w.failPut.on = true;
        await expect(save(w, store)).rejects.toBeTruthy();
        w.failPut.on = false;
        expect(store.refsById).toBe(true);
        // keep editing after the failure, then retry
        await doOps(store, new FormulaIndex(), [op(0, "insert", "row", 700, 2)]);
        await save(w, store);
      } else {
        await doOps(store, index, [op(0, "insert", "row", 700, 2)]);
      }
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("failed save, then retry with no further edit (second save must still write everything)", async () => {
    const t = await twin("fail2");
    const r = await compare(t, async ({ w, store, index, mode }) => {
      await doOps(store, index, [op(1, "delete", "row", 50, 4)]);
      if (mode === "id") {
        w.failPut.on = true;
        await expect(save(w, store)).rejects.toBeTruthy();
        w.failPut.on = false;
      }
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("the failed first save does not flip root.refs on the root object", async () => {
    const t = await twin("fail3");
    const store = t.byId;
    await doOps(store, t.bi, [op(0, "insert", "row", 10, 1)]);
    t.w.failPut.on = true;
    await expect(save(t.w, store)).rejects.toBeTruthy();
    t.w.failPut.on = false;
    expect(store.root.refs).not.toBe("id");
    // and the entries of the root still have no runs (nothing half-saved)
    for (const tab of store.root.tabs) for (const e of tab.chunks as { runs?: unknown }[]) expect(e.runs).toBeUndefined();
  }, 300_000);
});

describe("2. undo around the switch-over", () => {
  it("insert then undo before saving, save, reopen equals the untouched sheet", async () => {
    const t = await twin("undo1");
    const before = await dump(t.legacy);
    await doOps(t.byId, t.bi, [op(0, "insert", "row", 10, 3)]);
    t.byId.undo();
    expect(t.byId.structuralUndo).toBe(true);
    const again = await openStore(t.w, structuredClone(await save(t.w, t.byId)), true);
    expect(diff(before, await dump(again))).toEqual([]);
    // and a later structural edit still works from that saved state
    const ref = await openStore(t.w, structuredClone(t.legacy.root), false);
    await doOps(again, new FormulaIndex(), [op(0, "delete", "row", 30, 4)]);
    await doOps(ref, new FormulaIndex(), [op(0, "delete", "row", 30, 4)]);
    const a = await openStore(t.w, structuredClone(await save(t.w, again)), true);
    const b = await openStore(t.w, structuredClone(await save(t.w, ref)), false);
    expect(diff(await dump(b), await dump(a))).toEqual([]);
  }, 300_000);

  it("delete, edit a cell, more structure, undo the last, undo the cell, then save", async () => {
    const t = await twin("undo2");
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [op(0, "delete", "row", 50, 3)]);
      await store.load(0, store.chunkFor(0, 60));
      await store.setCell(0, 60, 1, "123");
      await doOps(store, new FormulaIndex(), [op(0, "insert", "row", 700, 2)]);
      store.undo(); // the insert
      store.undo(); // the cell
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("delete rows, undo, then insert, save (tombstones of the undone delete must be gone)", async () => {
    const t = await twin("undo3");
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [op(0, "delete", "row", 8, 6)]);
      store.undo();
      await doOps(store, new FormulaIndex(), [op(0, "insert", "row", 20, 2)]);
    });
    expect(r.problems).toEqual([]);
    // tombstones of the undone delete must not survive in the saved sheet
    const head = r.id.tabs[0].layout as { tombs?: { rows: Record<string, unknown> } };
    expect(Object.keys(head.tombs?.rows ?? {})).toEqual([]);
  }, 300_000);

  it("delete rows in an id store, save, then undo is not offered across the save", async () => {
    const t = await twin("undo4");
    await doOps(t.byId, t.bi, [op(0, "delete", "row", 8, 6)]);
    const saved = await save(t.w, t.byId);
    t.byId.undo();
    // nothing to undo: the saved state stands
    const again = await openStore(t.w, structuredClone(await save(t.w, t.byId)), true);
    const base = await openStore(t.w, structuredClone(saved), true);
    expect(diff(await dump(base), await dump(again))).toEqual([]);
  }, 300_000);
});

describe("3. several tabs", () => {
  it("reorder tabs after the switch-over", async () => {
    const t = await twin("reorder");
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [op(0, "insert", "row", 12, 2), op(1, "delete", "row", 300, 4)]);
      store.remapTabs([1, 0]);
      await doOps(store, new FormulaIndex(), [op(1, "insert", "row", 5, 1)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("reorder tabs, save, reopen, then structural edit on a never-opened chunk", async () => {
    const t = await twin("reorder2");
    const r = await compare(t, async ({ w, store, index }) => {
      await doOps(store, index, [op(0, "insert", "row", 12, 2)]);
      store.remapTabs([1, 0]);
      const reopened = await openStore(w, structuredClone(await save(w, store)), store.refsById);
      reopened.allowRefsById = store.allowRefsById;
      await doOps(reopened, new FormulaIndex(), [op(0, "delete", "row", 200, 3)]);
      return reopened;
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("delete a tab: formulas that named it read the same in every chunk (opened or not)", async () => {
    const t = await twin("deltab");
    const store = t.byId;
    await doOps(store, t.bi, [op(0, "insert", "row", 12, 2)]);
    // keep one chunk open (it holds A1 text with the deleted tab's name), the rest are not
    store.remapTabs([0]);
    const again = await openStore(t.w, structuredClone(await save(t.w, store)), true);
    const d = await dump(again);
    const texts = Object.values(d.tabs[0].cells).map((c) => c[1]).filter((f): f is string => !!f && (f.includes("Tab 2") || f.includes("#REF!")));
    const named = texts.filter((f) => f.includes("Tab 2")).length;
    const refs = texts.filter((f) => f.includes("#REF!")).length;
    // BUG-candidate: some chunks keep the old name, others say #REF!
    expect({ named, refs }.named === 0 || { named, refs }.refs === 0, JSON.stringify({ named, refs })).toBe(true);
  }, 300_000);

  it("delete a tab, save, add a tab with the same name: old formulas must not silently point at it", async () => {
    const t = await twin("deltab2");
    const store = t.byId;
    await doOps(store, t.bi, [op(0, "insert", "row", 12, 2)]);
    store.remapTabs([0]);
    const saved = await openStore(t.w, structuredClone(await save(t.w, store)), true);
    saved.addTab("Tab 2");
    const again = await openStore(t.w, structuredClone(await save(t.w, saved)), true);
    const d = await dump(again);
    const pointing = Object.values(d.tabs[0].cells).map((c) => c[1]).filter((f): f is string => !!f && f.includes("Tab 2")).length;
    const bad = Object.values(d.tabs[0].cells).map((c) => c[1]).filter((f): f is string => !!f && f.includes("#REF!")).length;
    // consistent behaviour: either all of them resurrect or none
    expect(pointing === 0 || bad === 0, JSON.stringify({ pointing, bad })).toBe(true);
  }, 300_000);

  it("rename a tab with formulas that name it", async () => {
    const t = await twin("rename");
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [op(0, "delete", "row", 12, 2)]);
      await renameTabAndFormulas(store, new FormulaIndex(), 1, "Renamed it's");
      await doOps(store, new FormulaIndex(), [op(1, "insert", "row", 40, 2)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("rename, save, reopen, delete rows in the renamed tab", async () => {
    const t = await twin("rename2");
    const r = await compare(t, async ({ w, store, index }) => {
      await renameTabAndFormulas(store, index, 1, "Other");
      const reopened = await openStore(w, structuredClone(await save(w, store)), store.allowRefsById);
      await doOps(reopened, new FormulaIndex(), [op(1, "delete", "row", 100, 6), op(0, "delete", "row", 100, 6)]);
      return reopened;
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("add a tab after the switch-over, put formulas in it that read the others, and others that read it", async () => {
    const t = await twin("addtab");
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [op(0, "insert", "row", 12, 2)]);
      const at = store.addTab("Fresh");
      const idx = new FormulaIndex();
      await putFormula(store, idx, at, 3, 0, "'Tab 1'!A20*2+SUM('Tab 2'!A2:A9)");
      await putFormula(store, idx, at, 4, 0, "A4+1");
      await putFormula(store, idx, 0, 6, 3, "Fresh!A4+Fresh!A5");
      await doOps(store, new FormulaIndex(), [op(at, "insert", "row", 0, 2), op(0, "delete", "row", 3, 2)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("add a tab BEFORE the first structural edit (unsaved tab with a formula), then the switch-over", async () => {
    const t = await twin("addtab2");
    const r = await compare(t, async ({ store, index }) => {
      const at = store.addTab("Fresh");
      const idx = new FormulaIndex();
      await putFormula(store, idx, at, 3, 0, "'Tab 1'!A20*2");
      await putFormula(store, idx, at, 4, 0, "A4+1");
      await doOps(store, index, [op(0, "insert", "row", 12, 2), op(at, "delete", "row", 1, 1)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("duplicate a tab (saved first), edit the copy and the original", async () => {
    const t = await twin("dup");
    const r = await compare(t, async ({ w, store, index }) => {
      await doOps(store, index, [op(0, "insert", "row", 12, 2)]);
      const saved = await openStore(w, structuredClone(await save(w, store)), store.allowRefsById);
      const copy = saved.duplicateTab(0, "Copy");
      const s2 = await openStore(w, structuredClone(await save(w, saved)), saved.allowRefsById);
      await doOps(s2, new FormulaIndex(), [op(copy, "delete", "row", 100, 3), op(0, "insert", "row", 300, 2), op(copy, "insert", "col", 1, 1)]);
      return s2;
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("duplicate a tab in the same session as the switch-over (unsaved) is refused cleanly", async () => {
    const t = await twin("dup2");
    await doOps(t.byId, t.bi, [op(0, "insert", "row", 12, 2)]);
    expect(() => t.byId.duplicateTab(0, "Copy")).toThrow();
  }, 300_000);

  it("duplicate a tab, then delete the ORIGINAL tab: the copy keeps its own formulas", async () => {
    const t = await twin("dup3");
    const r = await compare(t, async ({ w, store, index }) => {
      await doOps(store, index, [op(0, "insert", "row", 12, 2)]);
      const saved = await openStore(w, structuredClone(await save(w, store)), store.allowRefsById);
      const copy = saved.duplicateTab(0, "Copy");
      const s2 = await openStore(w, structuredClone(await save(w, saved)), saved.allowRefsById);
      const keep = s2.tabs.map((_, i) => i).filter((i) => i !== 0);
      s2.remapTabs(keep);
      void copy;
      return s2;
    });
    // formulas of the copy that named its own tab explicitly (self refs) must still be A1 not #REF!
    const d = await dump(r.id);
    const own = Object.values(d.tabs[0].cells).filter((c) => c[1]?.includes("#REF!")).length;
    const ownLegacy = Object.values((await dump(r.legacy)).tabs[0].cells).filter((c) => c[1]?.includes("#REF!")).length;
    expect(own).toBe(ownLegacy);
  }, 300_000);
});

describe("4. chunk list in a separate index blob", () => {
  async function indexed(w: World, root: SheetRootV2): Promise<SheetRootV2> {
    // Move every tab's chunk list into an index blob, as the writer does for sheets with many chunks.
    const tabs = [] as SheetRootV2["tabs"];
    for (const tab of root.tabs) {
      const staged = await encryptFrame(await frame({ chunks: tab.chunks }, FRAME_TYPE_TAB_INDEX));
      const [up] = await uploadStaged([staged], { servers: w.servers, signer: w.signer, ...w.env });
      tabs.push({ ...tab, chunks: { index: up.ref } });
    }
    return { ...root, tabs };
  }

  it("runs survive save and reopen when the chunk list is an index blob", async () => {
    const t = await twin("index");
    const rootA = await indexed(t.w, structuredClone(t.legacy.root));
    const rootB = await indexed(t.w, structuredClone(t.byId.root));
    const a = await openStore(t.w, rootA, false);
    const b = await openStore(t.w, rootB, true);
    expect(b.tabs[0].indexed).toBe(true);
    await doOps(a, new FormulaIndex(), [op(0, "insert", "row", 12, 2), op(0, "delete", "row", 400, 5), op(1, "delete", "col", 1, 1)]);
    await doOps(b, new FormulaIndex(), [op(0, "insert", "row", 12, 2), op(0, "delete", "row", 400, 5), op(1, "delete", "col", 1, 1)]);
    const sa = await save(t.w, a);
    const sb = await save(t.w, b);
    expect(sb.refs).toBe("id");
    for (const tab of sb.tabs) expect(Array.isArray(tab.chunks)).toBe(false);
    const fa = await openStore(t.w, structuredClone(sa), false);
    const fb = await openStore(t.w, structuredClone(sb), true);
    for (const tab of fb.tabs) for (const e of tab.chunks) expect(e.runs, "runs after reopen").toBeDefined();
    expect(diff(await dump(fa), await dump(fb))).toEqual([]);
    // another round from the reopened state
    await doOps(fa, new FormulaIndex(), [op(0, "delete", "row", 700, 5)]);
    await doOps(fb, new FormulaIndex(), [op(0, "delete", "row", 700, 5)]);
    const ga = await openStore(t.w, structuredClone(await save(t.w, fa)), false);
    const gb = await openStore(t.w, structuredClone(await save(t.w, fb)), true);
    expect(diff(await dump(ga), await dump(gb))).toEqual([]);
  }, 300_000);

  it("switch-over in index mode when only the first tab is touched: the other tab's index is still written with runs", async () => {
    const t = await twin("index2");
    const rootB = await indexed(t.w, structuredClone(t.byId.root));
    const b = await openStore(t.w, rootB, true);
    await doOps(b, new FormulaIndex(), [op(0, "insert", "row", 12, 2)]);
    const sb = await save(t.w, b);
    const fb = await openStore(t.w, structuredClone(sb), true);
    for (const tab of fb.tabs) for (const e of tab.chunks) expect(e.runs).toBeDefined();
  }, 300_000);
});

describe("5. rows past the end, ends, clamping, whole chunks", () => {
  it("formula far below the last row (a reference beyond the end) in a chunk never opened, then rows inserted above", async () => {
    const t = await twin("beyond", { rows: 1500, extra: [{ tab: 0, row: 600, col: 4, formula: "A5000+B4000" }] });
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [op(0, "insert", "row", 10, 3)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  // BUG: same cause as the next test: range end B5010 is past the end, stored as "@N", and does not move on a later insert
  it("same, with a delete above and then the referenced rows becoming real (typing at row 5000)", async () => {
    const t = await twin("beyond2", { rows: 1500, extra: [{ tab: 0, row: 600, col: 4, formula: "A5000+SUM(B4990:B5010)" }] });
    const r = await compare(t, async ({ w, store, index }) => {
      await doOps(store, index, [op(0, "delete", "row", 10, 3)]);
      await store.load(0, store.chunkFor(0, 5000));
      await store.setCell(0, 5000, 0, "77");
      const s = await openStore(w, structuredClone(await save(w, store)), store.allowRefsById);
      await doOps(s, new FormulaIndex(), [op(0, "insert", "row", 5, 2)]);
      return s;
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  // BUG: a "@N" (past the end) reference stored in a chunk that is not open does not move when rows are inserted above
  it("reference past the end stored as @N, rows inserted above in a later session (chunk not open)", async () => {
    const t = await twin("beyond3", { rows: 1500, extra: [{ tab: 0, row: 600, col: 4, formula: "A5000+SUM(B4990:B5010)" }] });
    const r = await compare(t, async ({ w, store, index }) => {
      await doOps(store, index, [op(0, "insert", "row", 10, 1)]); // switch-over; saves "@4999" tokens
      const s = await openStore(w, structuredClone(await save(w, store)), store.allowRefsById);
      await doOps(s, new FormulaIndex(), [op(0, "insert", "row", 10, 3)]); // chunk with the formula is not open
      return s;
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("typing far below the end with formulas in the new rows, then insert at 0, delete at 0, insert at the end", async () => {
    const t = await twin("far");
    const r = await compare(t, async ({ w, store, index }) => {
      const rows = store.tabs[0].rows;
      const idx = new FormulaIndex();
      await putFormula(store, idx, 0, rows + 40, 0, `A${rows + 39}+A1`);
      await putFormula(store, idx, 0, rows + 41, 0, `SUM(A${rows + 30}:A${rows + 40})`);
      await doOps(store, index, [op(0, "insert", "row", 0, 2), op(0, "delete", "row", 0, 1)]);
      const s = await openStore(w, structuredClone(await save(w, store)), store.allowRefsById);
      await doOps(s, new FormulaIndex(), [op(0, "insert", "row", s.tabs[0].rows, 3), op(0, "delete", "row", s.tabs[0].rows - 2, 99)]);
      return s;
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("delete more rows than exist (clamped) and a block covering whole chunks", async () => {
    const t = await twin("chunks", { rows: 4000 });
    const r = await compare(t, async ({ store, index }) => {
      const e = store.tabs[0].chunks;
      expect(e.length).toBeGreaterThan(4);
      const from = e[1].rowStart + 3;
      const count = e[3].rowStart + e[3].rowCount - from + 2; // covers chunks 2 and 3 completely
      await doOps(store, index, [op(0, "delete", "row", from, count)]);
      await doOps(store, new FormulaIndex(), [op(1, "delete", "row", 1200, 100000)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  // BUG (also in the positional store): after deleting the trailing chunks completely, typing below reads the old blob of the emptied last chunk and the deleted cells come back
  it("delete the trailing chunks completely, then type below: nothing of the deleted rows comes back", async () => {
    const t = await twin("tail");
    const r = await compare(t, async ({ w, store, index }) => {
      const e = store.tabs[0].chunks;
      const from = e[e.length - 2].rowStart + 1;
      await doOps(store, index, [op(0, "delete", "row", from, store.tabs[0].rows - from)]);
      await store.load(0, store.chunkFor(0, from + 2));
      await store.setCell(0, from + 2, 0, "9");
      const s = await openStore(w, structuredClone(await save(w, store)), store.allowRefsById);
      return s;
    });
    expect(r.problems).toEqual([]);
    // the deleted rows' old text must not return
    const d = await dump(r.id);
    const rows = r.id.tabs[0].rows;
    for (const key of Object.keys(d.tabs[0].cells)) expect(Number(key.split(":")[0])).toBeLessThan(rows + 1);
  }, 300_000);

  it("tiny sheet: one chunk, edits at both ends", async () => {
    const t = await twin("tiny", { tabs: 1, rows: 30 });
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [op(0, "insert", "row", 0, 1), op(0, "delete", "row", 25, 50), op(0, "insert", "col", 0, 1), op(0, "delete", "col", 2, 1)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);
});

describe("6. columns", () => {
  it("insert and delete columns that formulas reference, ranges A:A, formulas that live in a deleted column", async () => {
    const t = await twin("cols");
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [
        op(0, "insert", "col", 0, 1),
        op(0, "delete", "col", 1, 2),
        op(0, "delete", "col", 4, 1), // column E, where most formulas live
        op(1, "insert", "col", 3, 2),
        op(1, "delete", "col", 0, 1),
      ]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("200+ columns: insert 250 columns, formulas around, delete some, save, reopen", async () => {
    const t = await twin("wide", { tabs: 1, rows: 600 });
    const r = await compare(t, async ({ w, store, index }) => {
      await doOps(store, index, [op(0, "insert", "col", 2, 250)]);
      const idx = new FormulaIndex();
      await putFormula(store, idx, 0, 10, 260, "A5+SUM(B3:IZ3)+$IA$4");
      await putFormula(store, idx, 0, 11, 0, "SUM(C:C)+SUM(D5:HZ5)");
      const s = await openStore(w, structuredClone(await save(w, store)), store.allowRefsById);
      await doOps(s, new FormulaIndex(), [op(0, "delete", "col", 100, 20), op(0, "insert", "col", 0, 1)]);
      return s;
    });
    expect(r.problems).toEqual([]);
  }, 300_000);
});

describe("7/9. values and formulas whose own cells go away", () => {
  it("formulas that read deleted cells are recalculated; formulas on deleted rows vanish; values match the positional store", async () => {
    const t = await twin("vals", {
      extra: [
        { tab: 0, row: 20, col: 3, formula: "A10+A11" },
        { tab: 0, row: 21, col: 3, formula: "SUM(A9:A12)" },
        { tab: 0, row: 600, col: 3, formula: "A10*2" },
        { tab: 1, row: 22, col: 3, formula: "'Tab 1'!A10+1" },
      ],
    });
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [op(0, "delete", "row", 9, 2), op(0, "delete", "row", 19, 2)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  it("a single delete that removes the chunk holding formulas that other tabs read", async () => {
    const t = await twin("vals2");
    const r = await compare(t, async ({ store, index }) => {
      const e = store.tabs[0].chunks[2];
      await doOps(store, index, [op(0, "delete", "row", e.rowStart, e.rowCount)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);
});

describe("8. two stores on one root object", () => {
  it("opening and switching over in one store does not change what another store or a new open sees before a save", async () => {
    const t = await twin("shared");
    const root = structuredClone(t.legacy.root);
    const a = await openStore(t.w, root, true);
    const b = await openStore(t.w, root, true);
    await doOps(a, new FormulaIndex(), [op(0, "insert", "row", 10, 2)]);
    expect(root.refs).not.toBe("id");
    const c = await openStore(t.w, root, true);
    expect(c.refsById).toBe(false);
    expect(b.refsById).toBe(false);
    for (const tab of root.tabs) for (const e of tab.chunks as { runs?: unknown }[]) expect(e.runs).toBeUndefined();
  }, 300_000);

  // BUG: Object.assign(this.root, root) in save() mutates the shared root; a positional store's save then keeps root.refs === "id"
  it("a store opened earlier (positional) that saves after another store switched the shared root must not claim id storage", async () => {
    const t = await twin("shared2");
    const root = structuredClone(t.legacy.root);
    const a = await openStore(t.w, root, true);
    const b = await openStore(t.w, root, false);
    await doOps(a, new FormulaIndex(), [op(0, "insert", "row", 10, 2)]);
    await save(t.w, a);
    expect(root.refs).toBe("id"); // shared root object was updated by a's save
    await doOps(b, new FormulaIndex(), [op(1, "insert", "row", 10, 1)]);
    const sb = await save(t.w, b);
    // BUG-candidate: b stores plain A1 formulas but its root says refs: "id"
    expect(sb.refs).not.toBe("id");
  }, 300_000);

  // BUG: consequence of the above: reopen as id workbook, an insert leaves A1 formulas of unopened chunks unshifted
  it("consequence: a root wrongly marked refs:id over positional chunks gives wrong formulas after the next insert", async () => {
    const t = await twin("shared3");
    const root = structuredClone(t.legacy.root);
    const a = await openStore(t.w, root, true);
    const b = await openStore(t.w, root, false);
    await doOps(a, new FormulaIndex(), [op(0, "insert", "row", 10, 2)]);
    await save(t.w, a);
    await doOps(b, new FormulaIndex(), [op(1, "insert", "row", 10, 1)]);
    const sb = await save(t.w, b);
    const c = await openStore(t.w, structuredClone(sb), true);
    const before = await dump(c);
    await doOps(c, new FormulaIndex(), [op(0, "insert", "row", 5, 3)]);
    const after = await openStore(t.w, structuredClone(await save(t.w, c)), true);
    const d2 = await dump(after);
    const key = Object.keys(before.tabs[0].cells).find((k) => before.tabs[0].cells[k][1]?.startsWith("A") && Number(k.split(":")[0]) > 700) as string;
    const row = Number(key.split(":")[0]);
    const was = before.tabs[0].cells[key][1] as string;
    const now = d2.tabs[0].cells[`${row + 3}:${key.split(":")[1]}`]?.[1];
    const num = (f: string) => Number(/^A(\d+)/.exec(f)?.[1]);
    expect(num(now ?? "A0")).toBe(num(was) + (num(was) > 5 ? 3 : 0));
  }, 300_000);
});

describe("10. root parsing and documents", () => {
  it("parseRootV2 of an id workbook (after deletes, with tombstones) equals the store's view", async () => {
    const t = await twin("parse");
    await doOps(t.byId, t.bi, [op(0, "delete", "row", 8, 4), op(1, "delete", "col", 1, 1), op(0, "insert", "row", 300, 2), op(0, "delete", "row", 600, 3)]);
    const root = await save(t.w, t.byId);
    const doc = await parseRootV2(root, (ref, servers, type) => downloadTabBody(ref, servers, t.w.env, type));
    const store = await openStore(t.w, structuredClone(root), true);
    const shown = await dump(store);
    doc.sheets.forEach((sheet, ti) => {
      expect(Object.keys(sheet.cells).length).toBe(Object.keys(shown.tabs[ti].cells).length);
      for (const [key, cell] of Object.entries(sheet.cells)) {
        expect(cell.formula ?? null, `${sheet.name} ${key}`).toBe(shown.tabs[ti].cells[key]?.[1] ?? null);
      }
    });
    expect(doc.sheets.every((s) => !("tombs" in s))).toBe(true);
  }, 300_000);

  it("an id store turned into a document keeps no tombs / id leftovers in the document", async () => {
    const t = await twin("doc");
    await doOps(t.byId, t.bi, [op(0, "delete", "row", 8, 4)]);
    const doc = await t.byId.toDocument();
    // BUG: toDocument copies layout keys including 'tombs' into the sheet (parseRootV2 deletes them; toDocument does not)
    expect(doc.sheets.every((s) => !("tombs" in s))).toBe(true);
    const texts = doc.sheets.flatMap((s) => Object.values(s.cells).map((c) => c.formula).filter(Boolean)) as string[];
    expect(texts.some((f) => f.includes("~("))).toBe(false);
  }, 300_000);

  it("fromDocument of a small sheet, switch over and save as a document again keeps formulas right", async () => {
    const w = await world("small", 1, 40);
    const a = await openStore(w, structuredClone(w.root), true);
    await typeFormulas(a, new FormulaIndex());
    const doc = await a.toDocument();
    const mem = BigSheetStore.fromDocument(doc);
    await mem.open();
    const idx = new FormulaIndex();
    await editStructure(mem, idx, 0, { kind: "delete", axis: "row", index: 3, count: 2 });
    const doc2 = await mem.toDocument();
    const ref = BigSheetStore.fromDocument(doc);
    ref.allowRefsById = false;
    await ref.open();
    await editStructure(ref, new FormulaIndex(), 0, { kind: "delete", axis: "row", index: 3, count: 2 });
    const doc3 = await ref.toDocument();
    const f = (d: typeof doc2) => d.sheets.map((s) => Object.fromEntries(Object.entries(s.cells).map(([k, c]) => [k, [c.value, c.formula]])));
    expect(f(doc2)).toEqual(f(doc3));
    // round trip through the store again
    const mem2 = BigSheetStore.fromDocument(doc2);
    await mem2.open();
    expect(diff(await dump(mem2), await dump(ref))).toEqual([]);
  }, 300_000);
});

describe("misc lifecycle", () => {
  it("switch-over with no formulas anywhere: runs and refs are still saved", async () => {
    const w = await world("noformula", 1, 1500);
    const a = await openStore(w, structuredClone(w.root), true);
    expect(await a.enableRefsById(24)).toBe(true);
    const root = await save(w, a);
    expect(root.refs).toBe("id");
    for (const tab of root.tabs) for (const e of tab.chunks as { runs?: unknown }[]) expect(e.runs).toBeDefined();
    const b = await openStore(w, structuredClone(root), true);
    expect(b.refsById).toBe(true);
    const ref = await openStore(w, structuredClone(w.root), false);
    expect(diff(await dump(ref), await dump(b))).toEqual([]);
  }, 300_000);

  it("tombstones from an earlier session still shrink ranges in chunks opened only later", async () => {
    const t = await twin("tombs", { extra: [{ tab: 0, row: 700, col: 3, formula: "SUM(A400:A410)" }, { tab: 0, row: 701, col: 3, formula: "SUM(A401:A402)+A405" }] });
    const r = await compare(t, async ({ w, store, index }) => {
      await doOps(store, index, [op(0, "delete", "row", 399, 2)]);
      let s = await openStore(w, structuredClone(await save(w, store)), store.allowRefsById);
      await doOps(s, new FormulaIndex(), [op(0, "delete", "row", 400, 1), op(0, "insert", "row", 398, 2)]);
      s = await openStore(w, structuredClone(await save(w, s)), s.allowRefsById);
      await doOps(s, new FormulaIndex(), [op(0, "delete", "row", 396, 12)]);
      return s;
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  // A tab name typed in another case comes back in the tab's own spelling when formulas go through ids (the cells are the same); the test uses the real spelling.
  it("formulas with lower-case refs, quoted tab names in other case, absolute marks and whole rows", async () => {
    const t = await twin("textforms", {
      extra: [
        { tab: 0, row: 800, col: 3, formula: "a10+'Tab 2'!b12+$A$9+SUM(5:7)+SUM($B:$C)" },
        { tab: 0, row: 801, col: 3, formula: 'IF(A10>1,"A10 stays","x")&LOG10(A12)' },
      ],
    });
    const r = await compare(t, async ({ store, index }) => {
      await doOps(store, index, [op(0, "insert", "row", 6, 2), op(1, "delete", "row", 3, 1), op(0, "delete", "col", 1, 1)]);
    });
    expect(r.problems).toEqual([]);
  }, 300_000);

  for (const seed of [11, 12, 13]) {
    it(`random mix of structure edits, renames, reorders, saved and reopened often (seed ${seed})`, async () => {
      const t = await twin(`mix${seed}`, { rows: 1200 });
      let s = (0x9e3779b9 ^ seed) >>> 0;
      const rand = () => ((s = (s * 1664525 + 1013904223) >>> 0), s / 2 ** 32);
      const script: ((store: BigSheetStore, n: number) => Promise<void>)[] = [];
      for (let i = 0; i < 9; i++) {
        const kind = Math.floor(rand() * 6);
        const tabPick = rand();
        const idx = rand();
        const cnt = 1 + Math.floor(rand() * 6);
        script.push(async (store, n) => {
          const tab = Math.floor(tabPick * 2) % store.tabs.length;
          if (kind <= 2) {
            await doOps(store, new FormulaIndex(), [op(tab, kind === 2 ? "insert" : "delete", "row", Math.floor(idx * (store.tabs[tab].rows - 10)), cnt)]);
          } else if (kind === 3) {
            await doOps(store, new FormulaIndex(), [op(tab, "insert", "col", Math.floor(idx * 5), 1)]);
          } else if (kind === 4) {
            await renameTabAndFormulas(store, new FormulaIndex(), tab, `Name ${n}`);
          } else {
            store.remapTabs(store.tabs.map((_, k) => k).reverse());
          }
        });
      }
      const r = await compare(t, async ({ w, store }) => {
        let cur = store;
        for (const [n, step] of script.entries()) {
          await step(cur, n);
          if (n % 2 === 1) cur = await openStore(w, structuredClone(await save(w, cur)), store.allowRefsById);
        }
        return cur;
      });
      expect(r.problems).toEqual([]);
    }, 480_000);
  }
});

describe("delete tab in a second session (formulas already stored by id)", () => {
  // BUG: after the tab is gone, a dirty (open, edited) chunk keeps the old tab name as text while clean chunks read it as #REF!
  it("every formula that named the deleted tab reads the same whichever chunk it is in", async () => {
    const t = await twin("deltab3");
    await doOps(t.byId, t.bi, [op(0, "insert", "row", 12, 1)]);
    const s1 = await openStore(t.w, structuredClone(await save(t.w, t.byId)), true);
    // edit one cell in chunk 2 so that chunk is dirty and re-saved; the other chunks are not touched
    const row = s1.tabs[0].chunks[2].rowStart + 50;
    await s1.load(0, 2);
    await s1.setCell(0, row, 1, "5");
    s1.remapTabs([0]);
    const s2 = await openStore(t.w, structuredClone(await save(t.w, s1)), true);
    const d = await dump(s2);
    const forms = Object.values(d.tabs[0].cells).map((c) => c[1]).filter((f): f is string => !!f && (f.includes("Tab 2") || f.includes("#REF!")));
    const named = forms.filter((f) => f.includes("Tab 2")).length;
    const refs = forms.filter((f) => f.includes("#REF!")).length;
    expect(named === 0 || refs === 0, JSON.stringify({ named, refs })).toBe(true);
  }, 300_000);

  it("a new tab with the old name does not re-attach formulas of dirty chunks only", async () => {
    const t = await twin("deltab4");
    await doOps(t.byId, t.bi, [op(0, "insert", "row", 12, 1)]);
    const s1 = await openStore(t.w, structuredClone(await save(t.w, t.byId)), true);
    const row = s1.tabs[0].chunks[2].rowStart + 50;
    await s1.load(0, 2);
    await s1.setCell(0, row, 1, "5");
    s1.remapTabs([0]);
    const s2 = await openStore(t.w, structuredClone(await save(t.w, s1)), true);
    s2.addTab("Tab 2");
    const s3 = await openStore(t.w, structuredClone(await save(t.w, s2)), true);
    const d = await dump(s3);
    const forms = Object.values(d.tabs[0].cells).map((c) => c[1]).filter((f): f is string => !!f && (f.includes("Tab 2") || f.includes("#REF!")));
    const named = forms.filter((f) => f.includes("Tab 2")).length;
    const refs = forms.filter((f) => f.includes("#REF!")).length;
    expect(named === 0 || refs === 0, JSON.stringify({ named, refs })).toBe(true);
  }, 300_000);
});
