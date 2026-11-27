// Adversarial differential fuzzing harness (helper for src/lib/v2/advFuzz*.test.ts; not a test itself).
//
// The same random operation sequence is run on two stores opened from the same saved workbook:
//   * the REFERENCE: `allowRefsById = false` (formula text rewritten positionally in every chunk; the old behaviour)
//   * the SUBJECT:   `allowRefsById = true`  (formulas stored by id after the first structural edit)
// and a full dump (every cell value+formula, sizes, layout, tab names) is compared at every save and at the end.
import { importXlsxToBlossom } from "./streamImport";
import { BigSheetStore, type TabLayout } from "./bigSheet";
import { FormulaIndex, calculateTyped, pointRect, refreshDependents } from "./bigFormulas";
import { editStructure, renameTabAndFormulas } from "./bigStructure";
import { sortArea } from "./bigSort";
import { copyBlock, dragFill, fillArea, moveCells, pasteBlock, type Area } from "./bigClipboard";
import { findCells, replaceCells } from "./bigFind";
import { fakeBlossom, localSigner } from "./testUtils";
import { buildLargeXlsx } from "../xlsxStream/testFixtures";
import { createWorkerHandler } from "../xlsxStream/workerCore";
import type { WorkerLike } from "../xlsxStream/client";
import type { MainToWorker } from "../xlsxStream/protocol";
import type { SheetRootV2 } from "./root";

// ---------------------------------------------------------------------------------------------
// world

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

/** Small chunks (about 45 rows each) so that a sheet has dozens of chunks and the 6-chunk cache never holds them all. */
export const START = { maxChunkBytes: 8192, headRawBytes: 12_000, rawCap: 24_000 };

export interface World {
  root: SheetRootV2;
  servers: string[];
  env: { fetch: typeof fetch; sleep: () => Promise<void> };
  signer: ReturnType<typeof localSigner>;
}

const worlds = new Map<string, Promise<World>>();

/** A workbook with `tabs` tabs of `rows` rows (8 data columns A..H). Built once per (tabs, rows) and shared (blobs are immutable). */
export function world(tabs: number, rows: number): Promise<World> {
  const key = `${tabs}:${rows}`;
  let w = worlds.get(key);
  if (!w) {
    w = (async () => {
      const bytes = await buildLargeXlsx({ tabs, rows, cols: 8, sharedStrings: false });
      const fake = fakeBlossom();
      const signer = localSigner();
      const servers = [`https://adv${tabs}x${rows}.test`];
      const env = { fetch: fake.fetch as typeof fetch, sleep: async () => {} };
      const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), { title: "adv", servers, signer, env, createWorker: inProcessWorker, start: START });
      return { root, servers, env, signer };
    })();
    worlds.set(key, w);
  }
  return w;
}

export interface Side {
  store: BigSheetStore;
  index: FormulaIndex;
  byId: boolean;
  /** Undo steps made since the last save (BigSheetStore.apply silently drops an undo for a chunk that is no longer in memory, which can only happen after a save; both modes). */
  undoable?: number;
  redoable?: number;
}

export async function openSide(w: World, root: SheetRootV2, byId: boolean): Promise<Side> {
  const store = new BigSheetStore(root, w.servers, w.env);
  store.allowRefsById = byId;
  await store.open();
  return { store, index: new FormulaIndex(), byId };
}

/** Saves; with `reopen` the side continues from the saved root with a fresh store and formula index. */
export async function saveSide(w: World, side: Side, reopen: boolean): Promise<Side> {
  const root = await side.store.save({ servers: w.servers, signer: w.signer, ...w.env });
  side.undoable = 0;
  side.redoable = 0;
  if (!reopen) return side;
  return openSide(w, root, side.byId);
}

// ---------------------------------------------------------------------------------------------
// random numbers

export function rng(seed: number) {
  let a = (seed * 2654435761) >>> 0 || 1;
  const next = () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  return next;
}

// ---------------------------------------------------------------------------------------------
// operations

export type Op =
  | { k: "struct"; kind: "insert" | "delete"; axis: "row" | "col"; tab: number; index: number; count: number }
  | { k: "type"; tab: number; row: number; col: number; text: string }
  | { k: "sort"; tab: number; r1: number; c1: number; r2: number; c2: number; key: number; desc: boolean }
  | { k: "drag"; tab: number; r1: number; c1: number; r2: number; c2: number; grow: number; dir: "down" | "right" }
  | { k: "fill"; tab: number; r1: number; c1: number; r2: number; c2: number; dir: "down" | "right" }
  | { k: "paste"; tab: number; r1: number; c1: number; r2: number; c2: number; toTab: number; row: number; col: number }
  | { k: "move"; tab: number; r1: number; c1: number; r2: number; c2: number; dr: number; dc: number }
  | { k: "replace"; tab: number; q: string; rep: string }
  | { k: "layout"; tab: number; patch: Partial<TabLayout> }
  | { k: "addTab"; name: string }
  | { k: "rename"; tab: number; name: string }
  | { k: "dup"; tab: number; name: string }
  | { k: "remap"; order: number[] }
  | { k: "undo" }
  | { k: "redo" }
  | { k: "save"; reopen: boolean }
  | { k: "recalc" };

export const describeOp = (op: Op): string => JSON.stringify(op);

class Skip extends Error {}

const area = (op: { r1: number; c1: number; r2: number; c2: number }): Area => ({ r1: op.r1, c1: op.c1, r2: op.r2, c2: op.c2 });

function clampArea(store: BigSheetStore, tab: number, a: Area): Area {
  const info = store.tabs[tab];
  const r1 = Math.max(0, Math.min(a.r1, a.r2, info.rows - 1));
  const r2 = Math.max(r1, Math.min(Math.max(a.r1, a.r2), info.rows - 1));
  const c1 = Math.max(0, Math.min(a.c1, a.c2, info.cols - 1));
  const c2 = Math.max(c1, Math.min(Math.max(a.c1, a.c2), info.cols - 1));
  return { r1, c1, r2, c2 };
}

async function refresh(side: Side, tab: number, a: Area): Promise<void> {
  await refreshDependents(side.store, side.index, [{ tab, r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 }], undefined, true);
}

export async function typeCell(side: Side, tab: number, row: number, col: number, text: string): Promise<void> {
  const { store, index } = side;
  row = Math.max(0, Math.min(row, store.tabs[tab].rows - 1));
  col = Math.max(0, Math.min(col, store.tabs[tab].cols - 1));
  await store.load(tab, store.chunkFor(tab, row));
  if (text.startsWith("=")) {
    const formula = text.slice(1);
    const result = await calculateTyped(store, tab, row, col, formula);
    await store.edit(tab, row, col, row, col, (before) => ({
      ...(before?.style ? { style: before.style } : {}),
      formula,
      value: result.value,
      ...(result.error ? { type: "error" as const } : {}),
    }));
    await index.set(store, tab, row, col, formula);
  } else {
    await store.setCell(tab, row, col, text);
  }
  await refreshDependents(store, index, [pointRect({ tab, row, col })]);
}

/** Applies one operation to one side. Throws `Skip` when the operation does not make sense in the current state. */
export async function apply(w: World, sideRef: { side: Side }, op: Op): Promise<void> {
  const side0 = sideRef.side;
  if (op.k === "undo" && !(side0.undoable ?? 0)) throw new Skip();
  if (op.k === "redo" && !(side0.redoable ?? 0)) throw new Skip();
  await applyInner(w, sideRef, op);
  const side1 = sideRef.side;
  if (op.k === "undo") {
    side1.undoable = (side1.undoable ?? 1) - 1;
    side1.redoable = side1.store.structuralUndo ? 0 : (side1.redoable ?? 0) + 1;
  } else if (op.k === "redo") {
    side1.redoable = (side1.redoable ?? 1) - 1;
    side1.undoable = (side1.undoable ?? 0) + 1;
  } else if (op.k !== "save" && op.k !== "recalc" && op.k !== "layout") {
    side1.undoable = (side1.undoable ?? 0) + 1;
    side1.redoable = 0;
  }
}

async function applyInner(w: World, sideRef: { side: Side }, op: Op): Promise<void> {
  let side = sideRef.side;
  const { store, index } = side;
  const nTabs = store.tabs.length;
  const tab = "tab" in op ? ((op.tab % nTabs) + nTabs) % nTabs : 0;
  switch (op.k) {
    case "struct": {
      const info = store.tabs[tab];
      const size = op.axis === "row" ? info.rows : info.cols;
      let { index: at, count } = op;
      if (op.kind === "delete") {
        at = Math.min(at, size - 1);
        count = Math.min(count, size - at);
        if (count >= size) throw new Skip();
      } else at = Math.min(at, size);
      if (count <= 0) throw new Skip();
      // pin the tab's first chunk as edited so that loads during the edit cannot evict it (applyEdit crashes otherwise when 6+ edited chunks are pinned; both modes)
      await store.setLayout(tab, {});
      await editStructure(store, index, tab, { kind: op.kind, axis: op.axis, index: at, count });
      return;
    }
    case "type":
      await typeCell(side, tab, op.row, op.col, op.text);
      return;
    case "sort": {
      const a = clampArea(store, tab, area(op));
      const key = Math.max(a.c1, Math.min(a.c2, op.key));
      const moved = await sortArea(store, tab, a, key, op.desc, store.tabs[tab].layout.merges ?? []);
      if (moved) await refresh(side, tab, a);
      return;
    }
    case "drag": {
      const source = clampArea(store, tab, area(op));
      const extended = op.dir === "down" ? { ...source, r2: source.r2 + op.grow } : { ...source, c2: source.c2 + op.grow };
      if (extended.r2 >= store.tabs[tab].rows || extended.c2 >= store.tabs[tab].cols) throw new Skip();
      const target = await dragFill(store, tab, source, extended);
      if (target) await refresh(side, tab, target);
      return;
    }
    case "fill": {
      const a = clampArea(store, tab, area(op));
      const target = await fillArea(store, tab, a, op.dir);
      if (target) await refresh(side, tab, target);
      return;
    }
    case "paste": {
      const a = clampArea(store, tab, area(op));
      const toTab = ((op.toTab % nTabs) + nTabs) % nTabs;
      const block = await copyBlock(store, tab, a);
      const row = Math.min(op.row, store.tabs[toTab].rows - block.rows);
      const col = Math.min(op.col, store.tabs[toTab].cols - block.cols);
      if (row < 0 || col < 0) throw new Skip();
      const target = await pasteBlock(store, toTab, row, col, block);
      await refresh(side, toTab, target);
      return;
    }
    case "move": {
      const a = clampArea(store, tab, area(op));
      const info = store.tabs[tab];
      const dr = Math.max(-a.r1, Math.min(op.dr, info.rows - 1 - a.r2));
      const dc = Math.max(-a.c1, Math.min(op.dc, info.cols - 1 - a.c2));
      const target: Area = { r1: a.r1 + dr, c1: a.c1 + dc, r2: a.r2 + dr, c2: a.c2 + dc };
      const moved = await moveCells(store, tab, a, target);
      if (moved) {
        await refresh(side, tab, a);
        await refresh(side, tab, moved);
      }
      return;
    }
    case "replace": {
      const found = await findCells(store, tab, op.q, {});
      const matches = found.matches.slice(0, 60);
      const changed = await replaceCells(store, tab, matches, op.q, op.rep, {});
      for (const { row, col } of changed) await refresh(side, tab, { r1: row, c1: col, r2: row, c2: col });
      return;
    }
    case "layout":
      await store.setLayout(tab, op.patch);
      return;
    case "addTab": {
      if (store.tabs.some((t) => t.name.toLowerCase() === op.name.toLowerCase())) throw new Skip();
      store.addTab(op.name);
      index.forget();
      return;
    }
    case "rename": {
      if (store.tabs.some((t, i) => i !== tab && t.name.toLowerCase() === op.name.toLowerCase())) throw new Skip();
      await renameTabAndFormulas(store, index, tab, op.name);
      return;
    }
    case "dup": {
      if (store.tabs.some((t) => t.name.toLowerCase() === op.name.toLowerCase())) throw new Skip();
      side = await saveSide(w, side, false);
      store.duplicateTab(tab, op.name);
      index.forget();
      return;
    }
    case "remap": {
      const seen = new Set<number>();
      const order = op.order.map((i) => ((i % nTabs) + nTabs) % nTabs).filter((i) => (seen.has(i) ? false : (seen.add(i), true)));
      if (order.length === 0) throw new Skip();
      store.remapTabs(order);
      index.forget();
      return;
    }
    case "undo": {
      const cells = store.undo();
      if (store.structuralUndo) index.forget();
      else for (const c of cells.slice(0, 40)) await refresh(side, c.tab, { r1: c.row, c1: c.col, r2: c.row, c2: c.col });
      return;
    }
    case "redo": {
      const cells = store.redo();
      for (const c of cells.slice(0, 40)) await refresh(side, c.tab, { r1: c.row, c1: c.col, r2: c.row, c2: c.col });
      return;
    }
    case "save": {
      sideRef.side = await saveSide(w, side, op.reopen);
      return;
    }
    case "recalc":
      await refreshDependents(store, index, [], undefined, false, { all: true });
      return;
  }
}

// ---------------------------------------------------------------------------------------------
// dumps and comparison

export interface Dump {
  tabs: { name: string; rows: number; cols: number; layout: string; cells: Record<string, string> }[];
  meta: string;
}

const LAYOUT_COMPARED = ["columnWidths", "rowHeights", "hiddenRows", "hiddenColumns", "frozen", "zoom", "conditionalFormats", "filter", "merges", "borders", "images"] as const;

/** Every cell of every tab (opens every chunk): value, formula, type, style; sizes; layout (without tombstones); tab names. */
export async function dump(store: BigSheetStore): Promise<Dump> {
  const out: Dump = { tabs: [], meta: "" };
  for (let t = 0; t < store.tabs.length; t++) {
    const cells: Record<string, string> = {};
    for (let c = 0; c < store.tabs[t].chunks.length; c++) {
      await store.load(t, c);
      for (const [row, col, cell] of store.cellsOfChunk(t, c)) {
        cells[`${row}:${col}`] = JSON.stringify([cell.value ?? null, cell.formula ?? null, cell.type ?? null, cell.style ?? null]);
      }
    }
    const layout: Record<string, unknown> = {};
    for (const key of LAYOUT_COMPARED) if (store.tabs[t].layout[key] !== undefined) layout[key] = store.tabs[t].layout[key];
    out.tabs.push({ name: store.tabs[t].name, rows: store.tabs[t].rows, cols: store.tabs[t].cols, layout: JSON.stringify(layout), cells });
    const meta = store.tabMeta(t);
    out.meta += `${store.tabs[t].name}|${meta.color ?? ""}|${meta.hidden ? 1 : 0};`;
  }
  return out;
}

/** The first difference between two dumps, or null. */
export function diffDumps(ref: Dump, sub: Dump): string | null {
  if (ref.meta !== sub.meta) return `tab names/meta differ: ref=${ref.meta} sub=${sub.meta}`;
  if (ref.tabs.length !== sub.tabs.length) return `tab count ${ref.tabs.length} vs ${sub.tabs.length}`;
  for (let t = 0; t < ref.tabs.length; t++) {
    const a = ref.tabs[t];
    const b = sub.tabs[t];
    if (a.rows !== b.rows || a.cols !== b.cols) return `tab ${a.name}: size ${a.rows}x${a.cols} vs ${b.rows}x${b.cols}`;
    if (a.layout !== b.layout) return `tab ${a.name}: layout differs ref=${a.layout} sub=${b.layout}`;
    const keys = new Set([...Object.keys(a.cells), ...Object.keys(b.cells)]);
    const bad: string[] = [];
    for (const k of keys) if (a.cells[k] !== b.cells[k]) bad.push(`${k} ref=${a.cells[k]} sub=${b.cells[k]}`);
    if (bad.length) {
      bad.sort();
      return `tab ${a.name}: ${bad.length} cells differ, e.g. ${bad.slice(0, 4).join(" || ")}`;
    }
  }
  return null;
}

// ---------------------------------------------------------------------------------------------
// the workbook everything starts from: formulas in rows spread over a handful of chunks of every tab

export interface Seeded {
  plan: { tab: number; row: number; col: number; text: string }[];
}

/** Anchor rows (a few chunks per tab, kept under the 24-chunk limit of positional rewrites). */
export function anchorsFor(store: BigSheetStore, tab: number): number[] {
  const chunks = store.tabs[tab].chunks;
  const picks = [0, 2, Math.floor(chunks.length / 3), Math.floor(chunks.length / 2), chunks.length - 8, chunks.length - 6];
  const rows = store.tabs[tab].rows;
  const out = [...new Set(picks.filter((i) => i >= 0 && i < chunks.length))].map((i) => chunks[i].rowStart + Math.min(10, Math.max(0, chunks[i].rowCount - 12)) + 3).filter((at) => at + 140 < rows);
  return out.length >= 2 ? out : [3, Math.max(4, Math.min(rows - 150, 40))];
}

export function formulaPlan(store: BigSheetStore): { tab: number; row: number; col: number; text: string }[] {
  const plan: { tab: number; row: number; col: number; text: string }[] = [];
  const nTabs = store.tabs.length;
  for (let tab = 0; tab < nTabs; tab++) {
    const rows = store.tabs[tab].rows;
    const other = (d: number) => store.tabs[(tab + d) % nTabs].name;
    const q = (name: string) => `'${name}'!`;
    anchorsFor(store, tab).forEach((at, i) => {
      // K = f(A), L = f(K, neighbours), M = cross tab + absolute, N = range over chunks, O = reads formula cells above
      const r = at + 1; // 1-based row number in A1 text
      plan.push({ tab, row: at, col: 10, text: `=A${r}*2` });
      plan.push({ tab, row: at, col: 11, text: `=K${r}+SUM(B${r - 2}:B${r + 2})` });
      plan.push({ tab, row: at + 1, col: 12, text: nTabs > 1 ? `=${q(other(1))}A${5 + i}+$A$3+B$${7 + i}` : `=$A$3+B$${7 + i}` });
      plan.push({ tab, row: at + 2, col: 13, text: `=SUM(K${r - 5}:K${r + 40})` });
      plan.push({ tab, row: at + 3, col: 14, text: `=L${r}-K${r}+C${r}` });
      plan.push({ tab, row: at + 4, col: 15, text: `=SUM(A${Math.max(1, r - 120)}:C${r + 120})` });
    });
    plan.push({ tab, row: 1, col: 16, text: `=SUM(A2:A${Math.floor(rows / 2)})` });
    plan.push({ tab, row: 2, col: 16, text: `=A${rows - 5}+B${Math.floor(rows / 3)}` });
    plan.push({ tab, row: 3, col: 16, text: `=SUM(A:A)+COUNT(2:2)` });
    plan.push({ tab, row: 4, col: 16, text: `=SUM(K:K)+MAX(${q(other(1))}A1:B${Math.floor(rows / 4)})` });
    if (nTabs > 2) plan.push({ tab, row: 5, col: 16, text: `=SUM(${q(other(2))}K${anchorsFor(store, (tab + 2) % nTabs)[0] + 1}:N${anchorsFor(store, (tab + 2) % nTabs)[1] + 1})` });
  }
  return plan;
}

// ---------------------------------------------------------------------------------------------
// random generation (needs only the sizes, taken from the reference store)

export interface Mix {
  struct: number;
  type: number;
  sort: number;
  drag: number;
  fill: number;
  paste: number;
  move: number;
  replace: number;
  layout: number;
  tabs: number;
  undo: number;
  save: number;
  recalc: number;
  /** probability that a row/column count is huge (many chunks) */
  nasty: number;
  /** allow deleting columns (the legacy column-id fallback resurrects deleted columns in chunks that are not in memory: a bug that exists in both modes) */
  colDelete?: boolean;
  /** allow deleting tabs (the subject writes #REF! where the reference keeps the old tab name) */
  tabDelete?: boolean;
}

export const DEFAULT_MIX: Mix = { struct: 30, type: 18, sort: 3, drag: 4, fill: 3, paste: 5, move: 4, replace: 2, layout: 5, tabs: 4, undo: 0, save: 8, recalc: 2, nasty: 0.12 };

const NAMES = ["Data 2", "X", "Summary", "Q1", "Tab 2", "Tab 1", "Tab 3", "Zed"];

export interface GenState {
  /** formula chunks typed outside the anchor chunks so far (kept small so the 24-chunk limit is not what is tested) */
  strayChunks: number;
  names: number;
}

export function generate(rand: () => number, store: BigSheetStore, ref: Side, state: GenState, mix: Mix): Op {
  const skipKey = (k: string) => k === "nasty" || k === "colDelete" || k === "tabDelete";
  const total = Object.entries(mix).reduce((s, [k, v]) => (skipKey(k) ? s : s + (v as number)), 0);
  let pick = rand() * total;
  let kind = "struct";
  for (const [k, v] of Object.entries(mix)) {
    if (skipKey(k)) continue;
    if ((pick -= v as number) < 0) {
      kind = k;
      break;
    }
  }
  const n = store.tabs.length;
  const tab = Math.floor(rand() * n);
  const info = store.tabs[tab];
  const int = (lo: number, hi: number) => lo + Math.floor(rand() * (hi - lo + 1));
  /** A row: uniformly, or at a chunk boundary (starts or ends of chunks are where bugs hide). */
  const row = () => {
    const x = rand();
    if (x < 0.4 && info.chunks.length > 1) {
      const c = info.chunks[int(0, info.chunks.length - 1)];
      return Math.max(0, Math.min(info.rows - 1, c.rowStart + int(-2, 2) + (rand() < 0.5 ? c.rowCount : 0)));
    }
    return int(0, Math.max(0, info.rows - 1));
  };
  const anchors = anchorsFor(store, tab);
  const nearAnchor = () => Math.max(0, Math.min(info.rows - 1, anchors[int(0, anchors.length - 1)] + int(-6, 12)));
  const anyCol = () => (rand() < 0.7 ? int(0, 7) : int(8, 16));
  const colName = (c: number) => String.fromCharCode(65 + (c % 26)); // columns up to Z are enough here
  const cellRef = (t: number) => {
    const i = store.tabs[t];
    const r = rand() < 0.5 ? nearAnchorOf(t) : int(0, i.rows - 1);
    const c = Math.min(25, anyCol(), store.tabs[t].cols - 1);
    const abs = rand();
    const text = `${abs < 0.15 ? "$" : ""}${colName(c)}${abs > 0.85 && abs < 0.93 ? "$" : ""}${r + 1}`;
    return text;
  };
  const nearAnchorOf = (t: number) => {
    const a = anchorsFor(store, t);
    return Math.max(0, Math.min(store.tabs[t].rows - 1, a[int(0, a.length - 1)] + int(-6, 12)));
  };
  const prefix = (t: number) => (t === tab ? "" : `'${store.tabs[t].name}'!`);
  const randomFormula = (): string => {
    const t = rand() < 0.65 ? tab : int(0, n - 1);
    const p = prefix(t);
    const i = store.tabs[t];
    const x = rand();
    if (x < 0.2) return `=${p}${cellRef(t)}*2+${cellRef(tab)}`;
    if (x < 0.4) {
      const r1 = int(0, i.rows - 1);
      const r2 = Math.min(i.rows - 1, r1 + (rand() < 0.5 ? int(0, 30) : int(30, 900)));
      const c1 = int(0, 7);
      const c2 = Math.min(25, c1 + int(0, 3));
      return `=SUM(${p}${colName(c1)}${r1 + 1}:${colName(c2)}${r2 + 1})`;
    }
    if (x < 0.5) {
      const c1 = int(0, 12);
      return `=SUM(${p}${colName(c1)}:${colName(Math.min(25, c1 + int(0, 3)))})`;
    }
    if (x < 0.58) {
      const a = int(1, 30);
      return `=COUNT(${p}${a}:${a + int(0, 20)})`;
    }
    if (x < 0.7) return `=IF(${cellRef(tab)}>1000,${p}${cellRef(t)},${cellRef(tab)})`;
    if (x < 0.8) return `=${cellRef(tab)}+${cellRef(tab)}`;
    if (x < 0.9) {
      const r1 = int(1, i.rows - 40);
      const c1 = int(0, 10);
      return `=MAX(${p}${rand() < 0.2 ? "$" : ""}${colName(c1)}${rand() < 0.2 ? "$" : ""}${r1}:${colName(c1 + int(0, 2))}${r1 + int(0, 30)})`;
    }
    const ca = int(0, 7);
    const ra = int(1, 5);
    return `=SUM(${p}${colName(ca)}:${colName(ca + int(0, 3))})+COUNT(${ra}:${ra + int(0, 3)})`;
  };
  const formulaRowNote = () => {
    // most typed formulas go next to the anchors; a few go anywhere (bounded)
    if (rand() < 0.85 || state.strayChunks >= 3) return nearAnchor();
    state.strayChunks++;
    return row();
  };

  switch (kind) {
    case "struct": {
      const axis = rand() < 0.72 ? "row" : "col";
      const ins = rand() < 0.5;
      let count: number;
      let index: number;
      if (axis === "row") {
        if (rand() < mix.nasty) {
          // a block covering whole chunks, or a huge insert
          if (ins && rand() < 0.4) count = 100_000;
          else count = int(100, 500);
          if (ins && count < 1000) count = int(1, 40);
        } else count = int(1, 22);
        // aim at a formula's row sometimes (delete rows that other formulas read, or that hold formulas)
        const entries = ref.index.all();
        if (!ins && entries.length && rand() < 0.35) {
          const e = entries[int(0, entries.length - 1)];
          index = Math.max(0, e.row - int(0, 2));
          return { k: "struct", kind: "delete", axis, tab: e.tab, index, count: int(1, 5) };
        }
        index = rand() < 0.5 ? row() : nearAnchor();
        if (!ins && rand() < 0.15 && info.chunks.length > 3) {
          const c = info.chunks[int(1, info.chunks.length - 2)];
          index = c.rowStart;
          count = c.rowCount + int(0, 3);
        }
      } else {
        count = int(1, 2);
        index = int(0, 16);
      }
      if (axis === "col" && !ins && !mix.colDelete) return { k: "struct", kind: "insert", axis, tab, index, count };
      return { k: "struct", kind: ins ? "insert" : "delete", axis, tab, index, count };
    }
    case "type": {
      const x = rand();
      const r = x < 0.75 ? formulaRowNote() : row();
      const c = x < 0.75 ? int(10, 16) : anyCol();
      const y = rand();
      const text = y < 0.65 ? randomFormula() : y < 0.8 ? String(int(0, 5000)) : y < 0.9 ? `item-${int(0, 99)}` : "";
      return { k: "type", tab, row: r, col: c, text };
    }
    case "sort": {
      const r1 = row();
      return { k: "sort", tab, r1, c1: int(0, 2), r2: r1 + int(2, 60), c2: int(1, 4), key: int(0, 3), desc: rand() < 0.5 };
    }
    case "drag": {
      const r = rand() < 0.5 ? nearAnchor() : row();
      const c = int(8, 15);
      return { k: "drag", tab, r1: r, c1: c, r2: r, c2: c + int(0, 2), grow: int(1, 80), dir: "down" };
    }
    case "fill": {
      const r = rand() < 0.5 ? nearAnchor() : row();
      return { k: "fill", tab, r1: r, c1: int(8, 14), r2: r + int(0, 40), c2: int(9, 16), dir: rand() < 0.7 ? "down" : "right" };
    }
    case "paste": {
      const r = rand() < 0.6 ? nearAnchor() : row();
      const c = rand() < 0.6 ? int(10, 15) : int(0, 7);
      return { k: "paste", tab, r1: r, c1: c, r2: r + int(0, 12), c2: c + int(0, 3), toTab: rand() < 0.5 ? tab : int(0, n - 1), row: row(), col: int(0, 14) };
    }
    case "move": {
      const r = rand() < 0.6 ? nearAnchor() : row();
      const c = rand() < 0.6 ? int(10, 15) : int(0, 7);
      return { k: "move", tab, r1: r, c1: c, r2: r + int(0, 8), c2: c + int(0, 2), dr: rand() < 0.5 ? int(-300, 300) : int(-5, 5), dc: int(-3, 3) };
    }
    case "replace":
      return { k: "replace", tab, q: ["item-1", "item-2", "-1", "5"][int(0, 3)], rep: ["zz", "7", "item-9"][int(0, 2)] };
    case "layout": {
      const x = rand();
      const r = row();
      const patch: Partial<TabLayout> =
        x < 0.2
          ? { hiddenRows: [...new Set([r, row(), row()])].sort((a, b) => a - b) }
          : x < 0.35
            ? { rowHeights: { [r]: int(20, 60), [row()]: int(20, 60) } }
            : x < 0.5
              ? { merges: [{ r, c: int(0, 3), rs: int(1, 4), cs: int(1, 3) }] }
              : x < 0.62
                ? { frozen: { rows: int(0, 3), columns: int(0, 2) } }
                : x < 0.75
                  ? { borders: [{ r1: r, c1: 0, r2: r + int(0, 5), c2: int(0, 4), type: "all" }] }
                  : x < 0.85
                    ? { hiddenColumns: [int(0, 12)] }
                    : x < 0.95
                      ? { filter: { col: int(0, 3), header: true, hidden: [row(), row()].sort((a, b) => a - b), summary: "adv" } }
                      : { columnWidths: { [int(0, 12)]: int(50, 200) } };
      return { k: "layout", tab, patch };
    }
    case "tabs": {
      const x = rand();
      if (x < 0.25 && n < 5) return { k: "addTab", name: `${NAMES[int(0, NAMES.length - 1)]}${state.names++ % 3 === 0 ? "" : " " + state.names}` };
      if (x < 0.5) return { k: "rename", tab, name: NAMES[int(0, NAMES.length - 1)] };
      if (x < 0.65 && n < 5) return { k: "dup", tab, name: `Copy ${state.names++}` };
      if (x < 0.85 && n > 1) {
        const order = Array.from({ length: n }, (_, i) => i).sort(() => rand() - 0.5);
        return { k: "remap", order };
      }
      if (n > 1 && mix.tabDelete) return { k: "remap", order: Array.from({ length: n }, (_, i) => i).filter((i) => i !== tab) };
      if (n > 1) return { k: "remap", order: Array.from({ length: n }, (_, i) => i).reverse() };
      return { k: "rename", tab, name: NAMES[int(0, NAMES.length - 1)] };
    }
    case "undo":
      return rand() < 0.75 ? { k: "undo" } : { k: "redo" };
    case "save":
      return { k: "save", reopen: rand() < 0.7 };
    case "recalc":
      return { k: "recalc" };
  }
  return { k: "recalc" };
}

// ---------------------------------------------------------------------------------------------
// running a sequence on both sides

export interface RunOptions {
  tabs: number;
  rows: number;
  mix?: Mix;
}

export interface Failure {
  at: number; // op index (ops.length for the final checks)
  why: string;
}

const setups = new Map<World, Promise<{ ref: SheetRootV2; sub: SheetRootV2 }>>();

/** Opens both sides from the initial workbook with its formulas typed in and saved (computed once per world, then copied). */
export async function setup(w: World): Promise<{ ref: { side: Side }; sub: { side: Side } }> {
  let roots = setups.get(w);
  if (!roots) {
    roots = (async () => {
      const out: Record<string, SheetRootV2> = {};
      for (const [name, byId] of [["ref", false], ["sub", true]] as const) {
        let side = await openSide(w, structuredClone(w.root), byId);
        for (const p of formulaPlan(side.store)) await typeCell(side, p.tab, p.row, p.col, p.text);
        side = await saveSide(w, side, true);
        out[name] = structuredClone(side.store.root);
      }
      return { ref: out.ref, sub: out.sub };
    })();
    setups.set(w, roots);
  }
  const r = await roots;
  const ref = { side: await openSide(w, structuredClone(r.ref), false) };
  const sub = { side: await openSide(w, structuredClone(r.sub), true) };
  const d = diffDumps(await dump(ref.side.store), await dump(sub.side.store));
  if (d) throw new Error(`setup already differs: ${d}`);
  return { ref, sub };
}

export async function stepBoth(w: World, ref: { side: Side }, sub: { side: Side }, op: Op): Promise<string | null> {
  // Many edited (pinned) chunks make applyEdit crash in both modes (see advFuzzE.test.ts); keep the fuzzers away from that by saving first.
  if (op.k === "struct" && ref.side.store.dirtyCount >= 4) {
    ref.side = await saveSide(w, ref.side, false);
    sub.side = await saveSide(w, sub.side, false);
  }
  let refError: unknown = null;
  try {
    await apply(w, ref, op);
  } catch (error) {
    refError = error;
  }
  if (refError instanceof Skip) return null;
  // the positional store refuses edits whose formulas are spread over too many chunks; nothing was applied then
  if (refError instanceof Error && /too many parts/.test(refError.message)) return null;
  let subError: unknown = null;
  try {
    await apply(w, sub, op);
  } catch (error) {
    subError = error;
  }
  if (subError instanceof Skip) return refError ? null : "subject skipped but reference applied";
  if (!!refError !== !!subError) {
    return `only ${refError ? "REFERENCE" : "SUBJECT"} threw: ${String((refError ?? subError) instanceof Error ? ((refError ?? subError) as Error).stack?.split("\n").slice(0, 4).join(" | ") : refError ?? subError)}`;
  }
  return null;
}

export async function checkpoint(ref: Side, sub: Side, label: string): Promise<string | null> {
  const d = diffDumps(await dump(ref.store), await dump(sub.store));
  return d ? `${label}: ${d}` : null;
}

/** Runs a fixed op list; returns the first failure. */
export async function replay(w: World, ops: Op[], opts: { finalChecks?: boolean } = {}): Promise<Failure | null> {
  const { ref, sub } = await setup(w);
  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    let bad: string | null;
    try {
      bad = await stepBoth(w, ref, sub, op);
    } catch (error) {
      return { at: i, why: `harness error: ${String(error)}` };
    }
    if (bad) return { at: i, why: bad };
    if (op.k === "save") {
      const c = await checkpoint(ref.side, sub.side, "after save/reopen");
      if (c) return { at: i, why: c };
    }
    if (op.k === "undo" || op.k === "redo" || op.k === "recalc" || op.k === "struct") {
      // cheap in-memory comparison of what is loaded would miss chunks; compare everything only now and then
    }
  }
  if (opts.finalChecks !== false) {
    const c0 = await checkpoint(ref.side, sub.side, "in memory at end");
    if (c0) return { at: ops.length, why: c0 };
    for (const s of [ref, sub]) s.side = await saveSide(w, s.side, true);
    const c1 = await checkpoint(ref.side, sub.side, "after final save and reopen");
    if (c1) return { at: ops.length, why: c1 };
    for (const s of [ref, sub]) await refreshDependents(s.side.store, s.side.index, [], undefined, false, { all: true });
    const c2 = await checkpoint(ref.side, sub.side, "after full recalculation");
    if (c2) return { at: ops.length, why: c2 };
    for (const s of [ref, sub]) s.side = await saveSide(w, s.side, true);
    const c3 = await checkpoint(ref.side, sub.side, "after recalculation saved and reopened");
    if (c3) return { at: ops.length, why: c3 };
  }
  return null;
}

/** Generates and runs `count` operations; on a failure shrinks the list (bounded effort) and returns it. */
export async function fuzz(
  seed: number,
  count: number,
  opts: RunOptions,
  log: (s: string) => void = () => {},
): Promise<{ ops: Op[]; failure: Failure; minimal: Op[]; minimalFailure: Failure } | null> {
  const w = await world(opts.tabs, opts.rows);
  const rand = rng(seed);
  const { ref, sub } = await setup(w);
  const state: GenState = { strayChunks: 0, names: 1 };
  const ops: Op[] = [];
  let failure: Failure | null = null;
  for (let i = 0; i < count && !failure; i++) {
    const op = generate(rand, ref.side.store, ref.side, state, opts.mix ?? DEFAULT_MIX);
    ops.push(op);
    let bad: string | null;
    const t0 = Date.now();
    try {
      bad = await stepBoth(w, ref, sub, op);
    } catch (error) {
      bad = `harness error: ${String(error)}`;
    }
    if (!bad && op.k === "save") bad = await checkpoint(ref.side, sub.side, "after save/reopen");
    if (process.env.ADV_VERBOSE) log(`  op ${i} (${Date.now() - t0} ms) ${describeOp(op)}`);
    if (bad) failure = { at: i, why: bad };
  }
  if (!failure) {
    const c0 = await checkpoint(ref.side, sub.side, "in memory at end");
    if (c0) failure = { at: ops.length, why: c0 };
  }
  if (!failure) {
    for (const s of [ref, sub]) s.side = await saveSide(w, s.side, true);
    failure = await finalChecks(ref, sub, ops.length);
  }
  if (!failure) return null;
  log(`seed ${seed}: failure at op ${failure.at}: ${failure.why}`);
  const minimal = await shrink(w, ops.slice(0, Math.min(ops.length, failure.at + 1)), log, 150_000, failClass(failure.why));
  const minimalFailure = (await replay(w, minimal)) ?? failure;
  return { ops, failure, minimal, minimalFailure };
}

async function finalChecks(ref: { side: Side }, sub: { side: Side }, at: number): Promise<Failure | null> {
  const c1 = await checkpoint(ref.side, sub.side, "after final save and reopen");
  if (c1) return { at, why: c1 };
  for (const s of [ref, sub]) await refreshDependents(s.side.store, s.side.index, [], undefined, false, { all: true });
  const c2 = await checkpoint(ref.side, sub.side, "after full recalculation");
  if (c2) return { at, why: c2 };
  return null;
}

/** Delta-debugging style shrink: drops runs of ops while the sequence still fails (bounded by time). */
const failClass = (why: string): string => (why.startsWith("only") ? why.slice(0, 20) : why.startsWith("harness") ? "harness" : "diff");

export async function shrink(w: World, ops: Op[], log: (s: string) => void, budgetMs = 150_000, wanted = "diff"): Promise<Op[]> {
  const start = Date.now();
  let cur = ops;
  let chunk = Math.max(1, Math.floor(cur.length / 2));
  while (chunk >= 1 && Date.now() - start < budgetMs) {
    let removedAny = false;
    for (let i = 0; i + chunk <= cur.length && Date.now() - start < budgetMs; ) {
      const trial = [...cur.slice(0, i), ...cur.slice(i + chunk)];
      const f = trial.length === 0 ? null : await replay(w, trial);
      if (f && failClass(f.why) === wanted) {
        cur = trial;
        removedAny = true;
      } else i += chunk;
    }
    if (!removedAny || chunk === 1) chunk = chunk === 1 ? 0 : Math.max(1, Math.floor(chunk / 2));
  }
  log(`shrunk to ${cur.length} ops`);
  return cur;
}

export function report(seed: number, r: NonNullable<Awaited<ReturnType<typeof fuzz>>>): string {
  return [
    `SEED ${seed} FAILED at op ${r.failure.at} of ${r.ops.length}: ${r.failure.why}`,
    `MINIMAL SEQUENCE (${r.minimal.length} ops), still fails with: ${r.minimalFailure.why}`,
    ...r.minimal.map((op, i) => `  ${i}: ${describeOp(op)}`),
  ].join("\n");
}
