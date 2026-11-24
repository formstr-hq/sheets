import { beforeEach, describe, expect, it } from "vitest";
import { buildSheetEvent, decodeSheetEvent, decryptWithSecret, generateDocKeys, type TabMemory } from "../crypto";
import { createEmptyWorkbook, type WorkbookCell, type WorkbookDocument, type WorkbookSheet } from "../workbook";
import { MAX_CHUNK_BYTES, MIN_CHUNK_BYTES, assembleChunks, fillBytes, planTab, type ChunkMemory } from "./chunks";
import { forgetServerLimits, prepareTab, uploadFrames, chunkLimitFor } from "./blossom";
import { unframe } from "./codec";
import type { UploadProgress } from "./progress";
import { isSheetRootV2, parseRootV2, splitTab, type SheetRootV2 } from "./root";
import { fakeBlossom, localSigner } from "./testUtils";

const A = "https://a.example";
const B = "https://b.example";

/** Pseudo-random text that deflate cannot shrink much, so sizes track cell counts. */
function noisySheet(rows: number, cols: number): WorkbookSheet {
  const doc = createEmptyWorkbook("Noisy");
  let seed = 11;
  const next = () => (seed = (seed * 1664525 + 1013904223) >>> 0);
  const cells: Record<string, WorkbookCell> = {};
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      cells[`${r}:${c}`] = { value: next().toString(36) + next().toString(36) + next().toString(36), type: "text" };
    }
  }
  return { ...doc.sheets[0], rowCount: rows, colCount: cols, cells, columnWidths: { 0: 200 }, merges: [{ r: 0, c: 0, rs: 1, cs: 2 }] };
}

function docOf(...sheets: WorkbookSheet[]): WorkbookDocument {
  const doc = createEmptyWorkbook("Doc");
  doc.sheets = sheets.map((sheet, i) => ({ ...sheet, id: `t${i}`, name: `Tab ${i}`, order: i }));
  doc.activeSheetId = "t0";
  return doc;
}

async function decodeChunk(frame: Uint8Array) {
  return (await unframe(frame)) as { cells: Record<string, WorkbookCell> };
}

const MAX = 120 * 1024; // small enough to chunk a test sheet, above the 64 KB floor

beforeEach(() => forgetServerLimits());

describe("chunk size rules", () => {
  it("fills to 75% of the max: 6 MB for the default 8 MB, 3.75 MB for a 5 MB server", () => {
    expect(MAX_CHUNK_BYTES).toBe(8 * 1024 * 1024);
    expect(fillBytes(MAX_CHUNK_BYTES)).toBe(6 * 1024 * 1024);
    expect(fillBytes(5 * 1024 * 1024)).toBe(3.75 * 1024 * 1024);
  });

  it("keeps a worksheet that fits one chunk as exactly one chunk", async () => {
    const sheet = noisySheet(50, 4);
    const chunks = await planTab(sheet, { max: MAX });
    expect(chunks).toHaveLength(1);
    expect(chunks[0].rowStart).toBe(0);
    expect(chunks[0].rowCount).toBe(50);
  });

  it("cuts a big worksheet into contiguous row ranges, each within the max, and reassembles it exactly", async () => {
    const sheet = noisySheet(4000, 5);
    const chunks = await planTab(sheet, { max: MAX });
    expect(chunks.length).toBeGreaterThan(2);
    let next = 0;
    for (const chunk of chunks) {
      expect(chunk.size).toBeLessThanOrEqual(MAX);
      expect(chunk.rowStart).toBe(next);
      next += chunk.rowCount;
    }
    expect(next).toBe(4000);
    // Chunks are filled toward 75% of the max, not left tiny (except the last).
    for (const chunk of chunks.slice(0, -1)) expect(chunk.size).toBeGreaterThan(MAX * 0.3);

    const bodies = await Promise.all(chunks.map((chunk) => decodeChunk(chunk.frame as Uint8Array)));
    // Cell keys inside a chunk are relative to its rowStart.
    expect(Object.keys(bodies[1].cells).every((key) => Number(key.split(":")[0]) < chunks[1].rowCount)).toBe(true);
    const rebuilt = assembleChunks(chunks, bodies);
    expect(rebuilt.cells).toEqual(sheet.cells);
    // Layout rides in the first chunk.
    expect(rebuilt.columnWidths).toEqual({ 0: 200 });
    expect(rebuilt.merges).toEqual(sheet.merges);
  });

  it("splits a chunk that grew above the max into two of about half", async () => {
    const small = noisySheet(300, 5);
    const [only] = await planTab(small, { max: MAX });
    const memory: ChunkMemory = {
      id: only.id,
      rowStart: 0,
      rowCount: only.rowCount,
      ref: { x: "x", k: "k", n: 1, p: 1, c: 1, h: only.hash },
      servers: [A],
    };
    // The same tab after a lot of rows were added: its single old chunk would now be too big.
    const grown = noisySheet(1500, 5);
    const chunks = await planTab(grown, { max: MAX, previous: [memory] });
    expect(chunks.length).toBeGreaterThanOrEqual(2);
    for (const chunk of chunks) expect(chunk.size).toBeLessThanOrEqual(MAX);
    const total = chunks.reduce((sum, chunk) => sum + chunk.size, 0);
    // Halves, not slivers: no chunk takes the whole tab.
    expect(Math.max(...chunks.map((chunk) => chunk.size))).toBeLessThan(total * 0.75);
  });

  it("merges neighbours that together are under half the max", async () => {
    const sheet = noisySheet(4000, 5);
    const cut = await planTab(sheet, { max: MAX });
    expect(cut.length).toBeGreaterThan(3);
    // Pretend the tab was cut into slivers last time; every pair is far below max/2.
    const slivers: ChunkMemory[] = [0, 10, 20, 30, 40, 50].map((rowStart, i) => ({
      id: `s${i}`,
      rowStart,
      rowCount: 10,
      ref: { x: "x", k: "k", n: 1, p: 1, c: 1, h: `nomatch${i}` },
      servers: [A],
    }));
    const tiny = noisySheet(60, 3);
    const merged = await planTab(tiny, { max: MAX, previous: slivers });
    expect(merged).toHaveLength(1);
    expect(merged[0].rowCount).toBe(60);
  });

  it("chunks by row without stable ids: the cell keys stay positional", async () => {
    const sheet = noisySheet(3000, 4);
    const chunks = await planTab(sheet, { max: MAX });
    const body = splitTab(sheet).body;
    expect(Object.keys(body.cells)[0]).toMatch(/^\d+:\d+$/);
    expect(chunks.length).toBeGreaterThan(1);
  });
});

describe("saving and loading chunked sheets", () => {
  async function save(doc: WorkbookDocument, net: ReturnType<typeof fakeBlossom>, opts: { servers?: string[]; reuse?: Map<string, TabMemory>; replicas?: number; rootMaxBytes?: number } = {}) {
    const keys = generateDocKeys();
    const reuse = opts.reuse ?? new Map<string, TabMemory>();
    const event = await buildSheetEvent({
      sheetId: "s",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: {
        servers: opts.servers ?? [A],
        replicas: opts.replicas,
        signer: localSigner(),
        reuse,
        env: { fetch: net.fetch },
        rootMaxBytes: opts.rootMaxBytes,
      },
    });
    return { event, keys, reuse };
  }

  it("round-trips a multi-chunk workbook through the fake Blossom", async () => {
    const net = fakeBlossom();
    net.limits.set(A, MAX); // forces chunking by the server's limit
    const doc = docOf(noisySheet(3000, 5), noisySheet(40, 3));
    const { event, keys } = await save(doc, net);
    const root = JSON.parse(decryptWithSecret(event.content, keys.viewKey)) as SheetRootV2;
    expect(isSheetRootV2(root)).toBe(true);
    const [big, small] = root.tabs;
    expect(Array.isArray(big.chunks) && big.chunks.length).toBeGreaterThan(1);
    expect(Array.isArray(small.chunks) && small.chunks.length).toBe(1);
    for (const store of net.stores.values()) {
      for (const bytes of store.values()) expect(bytes.byteLength).toBeLessThanOrEqual(MAX);
    }
    const loaded = await decodeSheetEvent(event, keys.viewKey, "x", { servers: [A], env: { fetch: net.fetch } });
    expect(loaded.sheets[0].cells).toEqual(doc.sheets[0].cells);
    expect(loaded.sheets[1].cells).toEqual(doc.sheets[1].cells);
    expect(loaded.sheets[0].columnWidths).toEqual({ 0: 200 });
  });

  it("re-uploads only the chunks that changed", async () => {
    const net = fakeBlossom();
    net.limits.set(A, MAX);
    const doc = docOf(noisySheet(3000, 5));
    const first = await save(doc, net);
    const initial = net.uploads.length;
    expect(initial).toBeGreaterThan(2);

    // Edit one cell in the last row: only the last chunk's bytes change.
    const edited = structuredClone(doc);
    edited.sheets[0].cells["2999:0"] = { value: "changed", type: "text" };
    await save(edited, net, { reuse: first.reuse });
    expect(net.uploads.length - initial).toBe(1);

    // Saving the identical document again uploads nothing.
    const before = net.uploads.length;
    await save(edited, net, { reuse: first.reuse });
    expect(net.uploads.length).toBe(before);
  });

  it("learns a small server's upload limit from a 413 and cuts smaller chunks", async () => {
    const net = fakeBlossom();
    net.limits.set(A, 90 * 1024); // the preflight sees the largest blob and answers 413
    const doc = docOf(noisySheet(3000, 5));
    // The limit starts unknown: the first plan uses the 8 MB default, so one chunk.
    expect(chunkLimitFor([A])).toBe(MAX_CHUNK_BYTES);
    const { event, keys } = await save(doc, net);
    expect(net.preflights.length).toBeGreaterThan(0);
    expect(chunkLimitFor([A])).toBeLessThan(MAX_CHUNK_BYTES);
    expect(chunkLimitFor([A])).toBeGreaterThanOrEqual(MIN_CHUNK_BYTES);
    for (const store of net.stores.values()) {
      for (const bytes of store.values()) expect(bytes.byteLength).toBeLessThanOrEqual(90 * 1024);
    }
    const loaded = await decodeSheetEvent(event, keys.viewKey, "x", { servers: [A], env: { fetch: net.fetch } });
    expect(loaded.sheets[0].cells).toEqual(doc.sheets[0].cells);
  });

  it("uses the smallest limit among the servers", async () => {
    const net = fakeBlossom();
    net.limits.set(A, 400 * 1024);
    net.limits.set(B, 100 * 1024);
    const doc = docOf(noisySheet(3000, 5));
    await save(doc, net, { servers: [A, B], replicas: 2 });
    for (const [origin, store] of net.stores) {
      for (const bytes of store.values()) expect(bytes.byteLength, origin).toBeLessThanOrEqual(100 * 1024);
    }
  });

  it("copies each chunk to every ticked server (replica count)", async () => {
    const net = fakeBlossom();
    net.limits.set(A, MAX);
    const doc = docOf(noisySheet(3000, 5));
    await save(doc, net, { servers: [A, B], replicas: 2 });
    const a = net.stores.get(A)!;
    const b = net.stores.get(B)!;
    expect(a.size).toBeGreaterThan(1);
    expect([...b.keys()].sort()).toEqual([...a.keys()].sort());
    expect(net.uploads).toHaveLength(a.size + b.size);
  });

  it("re-uploads a reused chunk when it is not on every ticked server yet", async () => {
    const net = fakeBlossom();
    net.limits.set(A, MAX);
    const doc = docOf(noisySheet(3000, 5));
    const first = await save(doc, net, { servers: [A] });
    const stored = net.stores.get(A)!.size;
    net.uploads.length = 0;
    await save(doc, net, { servers: [A, B], replicas: 2, reuse: first.reuse });
    // Every chunk is now missing from B, so every chunk is uploaded again (to both).
    expect(net.uploads.length).toBe(stored * 2);
  });

  it("reports upload progress: per-tab chunk counts, bytes, and a final done", async () => {
    const net = fakeBlossom();
    net.limits.set(A, MAX);
    const events: UploadProgress[] = [];
    const doc = docOf(noisySheet(3000, 5));
    const keys = generateDocKeys();
    await buildSheetEvent({
      sheetId: "s",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: [A, B], replicas: 2, signer: localSigner(), env: { fetch: net.fetch }, onProgress: (p) => events.push(p) },
    });
    const last = events[events.length - 1];
    expect(events.some((e) => e.phase === "preparing")).toBe(true);
    expect(last.phase).toBe("done");
    expect(last.blobsTotal).toBeGreaterThan(1);
    expect(last.blobsDone).toBe(last.blobsTotal);
    expect(last.bytesDone).toBe(last.bytesTotal);
    expect(last.tabs[0].stored).toBe(last.tabs[0].chunks);
    // Copies count: bytes total covers both servers.
    const stored = [...net.stores.get(A)!.values()].reduce((sum, b) => sum + b.byteLength, 0);
    expect(last.bytesTotal).toBe(stored * 2);
  });

  it("moves the chunk list into a tab-index blob when the root would be too big", async () => {
    const net = fakeBlossom();
    net.limits.set(A, MAX);
    const doc = docOf(noisySheet(3000, 5));
    const { event, keys } = await save(doc, net, { rootMaxBytes: 800 });
    const root = JSON.parse(decryptWithSecret(event.content, keys.viewKey)) as SheetRootV2;
    expect(root.tabs[0].chunks).toHaveProperty("index");
    const loaded = await decodeSheetEvent(event, keys.viewKey, "x", { servers: [A], env: { fetch: net.fetch } });
    expect(loaded.sheets[0].cells).toEqual(doc.sheets[0].cells);
  });
});

describe("reading older formats", () => {
  it("still reads a root whose tabs are one blob each", async () => {
    const net = fakeBlossom();
    const doc = docOf(noisySheet(200, 4));
    const { frameBytes } = await prepareTab(splitTab(doc.sheets[0]).body);
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch });
    const { meta } = splitTab(doc.sheets[0]);
    const root: SheetRootV2 = {
      schema: "formstr.sheets.root",
      v: 2,
      title: "Old",
      activeSheetId: meta.id,
      createdAt: 1,
      updatedAt: 1,
      epoch: 1,
      servers: [A],
      tabs: [{ ...meta, blob: uploaded.ref }],
    };
    const { downloadTabBody } = await import("./blossom");
    const loaded = await parseRootV2(root, (ref, servers, type) => downloadTabBody(ref, servers, { fetch: net.fetch }, type));
    expect(loaded.sheets[0].cells).toEqual(doc.sheets[0].cells);
  });

  it("still reads an inline root", async () => {
    const keys = generateDocKeys();
    const doc = docOf(noisySheet(5, 3));
    const event = await buildSheetEvent({ sheetId: "s", document: doc, viewKey: keys.viewKey, editKey: keys.editKey });
    const loaded = await decodeSheetEvent(event, keys.viewKey);
    expect(loaded.sheets[0].cells).toEqual(doc.sheets[0].cells);
  });
});
