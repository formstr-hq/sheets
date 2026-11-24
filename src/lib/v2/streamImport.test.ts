import { describe, expect, it } from "vitest";
import { importXlsxToBlossom } from "./streamImport";
import { BigSheetStore } from "./bigSheet";
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
});
