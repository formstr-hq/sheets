import { describe, expect, it, vi } from "vitest";
import { importXlsxStreaming, type WorkerLike } from "./client";
import { CREDIT_WINDOW, type MainToWorker, type WorkerToMain } from "./protocol";
import { buildLargeXlsx } from "./testFixtures";
import { CreditGate, createWorkerHandler } from "./workerCore";

/** Runs the worker logic in this thread, behind the same message interface. */
function inProcessWorker(spy?: { sent: WorkerToMain[]; terminated: boolean }): WorkerLike {
  const worker: WorkerLike = {
    onmessage: null,
    onerror: null,
    onmessageerror: null,
    postMessage(message: MainToWorker) {
      queueMicrotask(() => handle(message));
    },
    terminate() {
      if (spy) spy.terminated = true;
      handle({ t: "cancel" });
    },
  };
  const handle = createWorkerHandler((message) => {
    spy?.sent.push(message);
    setTimeout(() => worker.onmessage?.({ data: message }), 0);
  });
  return worker;
}

const START = { maxChunkBytes: 64 * 1024, headRawBytes: 100_000, rawCap: 1_000_000 };

describe("CreditGate", () => {
  it("hands out credits and queues waiters", async () => {
    const gate = new CreditGate(2);
    await gate.acquire();
    await gate.acquire();
    let third = false;
    const waiting = gate.acquire().then(() => (third = true));
    await Promise.resolve();
    expect(third).toBe(false);
    gate.release();
    await waiting;
    expect(third).toBe(true);
    gate.release(2);
    await gate.acquire();
    await gate.acquire();
  });
});

describe("importXlsxStreaming", () => {
  it("delivers chunks, tab summaries and the final summary", async () => {
    const bytes = await buildLargeXlsx({ tabs: 2, rows: 3000, cols: 8, sharedStrings: true });
    const chunks: number[] = [];
    let tabs = 0;
    let progress = 0;
    const handle = importXlsxStreaming(
      new Blob([bytes as BlobPart]),
      {
        onChunk: (chunk) => {
          expect(chunk.frame).toBeInstanceOf(Uint8Array);
          chunks.push(chunk.frame.length);
        },
        onTabDone: () => tabs++,
        onProgress: () => progress++,
      },
      { ...START, createWorker: () => inProcessWorker() },
    );
    const summary = await handle.promise;
    expect(summary.cells).toBe(2 * 3001 * 8);
    expect(tabs).toBe(2);
    expect(progress).toBeGreaterThan(0);
    expect(chunks.length).toBe(summary.tabs.reduce((sum, tab) => sum + tab.chunks, 0));
  });

  it("keeps at most CREDIT_WINDOW chunks in flight until the consumer finishes with them", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 8000, cols: 10, sharedStrings: true });
    const spy = { sent: [] as WorkerToMain[], terminated: false };
    const release: Array<() => void> = [];
    let received = 0;
    const handle = importXlsxStreaming(
      new Blob([bytes as BlobPart]),
      {
        onChunk: () => {
          received++;
          return new Promise<void>((resolve) => release.push(resolve));
        },
      },
      { ...START, createWorker: () => inProcessWorker(spy) },
    );
    // Wait until the window is full (slow machines need longer), then a little more to see it stays full.
    await vi.waitFor(() => expect(received).toBe(CREDIT_WINDOW), { timeout: 4000 });
    await new Promise((resolve) => setTimeout(resolve, 300));
    expect(received).toBe(CREDIT_WINDOW);
    expect(spy.sent.filter((m) => m.t === "chunk")).toHaveLength(CREDIT_WINDOW);
    // Finishing one chunk lets exactly one more through.
    release.shift()!();
    await vi.waitFor(() => expect(received).toBe(CREDIT_WINDOW + 1), { timeout: 4000 });
    await new Promise((resolve) => setTimeout(resolve, 300));
    expect(received).toBe(CREDIT_WINDOW + 1);
    // Drain everything.
    const drain = setInterval(() => release.splice(0).forEach((resolve) => resolve()), 5);
    const summary = await handle.promise;
    clearInterval(drain);
    expect(summary.cells).toBe(8001 * 10);
  });

  it("cancel() terminates the worker and rejects with CANCELLED", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 8000, cols: 10, sharedStrings: true });
    const spy = { sent: [] as WorkerToMain[], terminated: false };
    const handle = importXlsxStreaming(new Blob([bytes as BlobPart]), { onChunk: () => new Promise<void>(() => {}) }, { ...START, createWorker: () => inProcessWorker(spy) });
    await new Promise((resolve) => setTimeout(resolve, 200));
    handle.cancel();
    await expect(handle.promise).rejects.toMatchObject({ code: "CANCELLED" });
    expect(spy.terminated).toBe(true);
  });

  it("rejects with the consumer's error and stops the worker", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 3000, cols: 10, sharedStrings: true });
    const spy = { sent: [] as WorkerToMain[], terminated: false };
    const handle = importXlsxStreaming(
      new Blob([bytes as BlobPart]),
      {
        onChunk: () => {
          throw new Error("upload refused");
        },
      },
      { ...START, createWorker: () => inProcessWorker(spy) },
    );
    await expect(handle.promise).rejects.toThrow("upload refused");
    expect(spy.terminated).toBe(true);
  });

  it("passes worker errors through with their code", async () => {
    const handle = importXlsxStreaming(new Blob(["not a zip at all, not a zip at all"]), { onChunk: () => {} }, { ...START, createWorker: () => inProcessWorker() });
    await expect(handle.promise).rejects.toMatchObject({ code: "NOT_XLSX" });
  });

  it("reports a crashed worker as an unexpected stop", async () => {
    const worker: WorkerLike = { onmessage: null, onerror: null, onmessageerror: null, postMessage() {}, terminate() {} };
    const handle = importXlsxStreaming(new Blob(["x"]), { onChunk: () => {} }, { ...START, createWorker: () => worker });
    worker.onerror?.({});
    await expect(handle.promise).rejects.toMatchObject({ code: "UNKNOWN", message: expect.stringMatching(/out of memory/) });
  });

  it("forwards a learned size limit to the worker", async () => {
    const worker: WorkerLike & { posted: MainToWorker[] } = {
      posted: [],
      onmessage: null,
      onerror: null,
      onmessageerror: null,
      postMessage(message) {
        this.posted.push(message);
      },
      terminate() {},
    };
    const handle = importXlsxStreaming(new Blob(["x"]), { onChunk: () => {} }, { ...START, createWorker: () => worker });
    handle.setMaxChunkBytes(100_000);
    expect(worker.posted.map((m) => m.t)).toEqual(["start", "limit"]);
    handle.cancel();
    await expect(handle.promise).rejects.toBeDefined();
  });
});
