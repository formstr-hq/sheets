// Main-thread side of the streaming import: starts the worker, passes chunks to the caller and
// returns credits when the caller is done with them.
import { ImportError } from "./errors";
import type { MainToWorker, StartOptions, WireChunk, WorkerToMain } from "./protocol";
import type { EmittedChunk, ImportSummary, ReadProgress, TabSummary } from "./runImport";

export interface ImportHandlers {
  /**
   * Called for each chunk. The worker may send the next one only after the returned promise
   * resolves, so do the slow work (encrypting, staging) before resolving. If it rejects the
   * import is cancelled with that error. Up to CREDIT_WINDOW (4) calls can be in flight at once,
   * so do not assume they finish in the order they started.
   */
  onChunk(chunk: EmittedChunk): Promise<void> | void;
  onTabDone?(summary: TabSummary): void;
  onProgress?(progress: ReadProgress): void;
}

export interface ImportHandle {
  promise: Promise<ImportSummary>;
  /** Stops the worker immediately; `promise` rejects with CANCELLED. */
  cancel(): void;
  /** Later chunks are cut no bigger than this (a server limit was learned). */
  setMaxChunkBytes(bytes: number): void;
}

/** The slice of the Worker API the client uses (lets tests plug in an in-process worker). */
export interface WorkerLike {
  postMessage(message: MainToWorker, transfer?: Transferable[]): void;
  terminate(): void;
  onmessage: ((event: { data: WorkerToMain }) => void) | null;
  onerror: ((event: unknown) => void) | null;
  onmessageerror: ((event: unknown) => void) | null;
}

export interface ImportClientOptions extends StartOptions {
  /** Test hook. Default: the real module worker. */
  createWorker?: () => WorkerLike;
}

/** True when the browser has what the streaming import needs. */
export function isStreamImportSupported(): boolean {
  return (
    typeof Worker !== "undefined" &&
    typeof Blob !== "undefined" &&
    typeof Blob.prototype.slice === "function" &&
    typeof TextDecoder !== "undefined" &&
    typeof CompressionStream !== "undefined" &&
    typeof DecompressionStream !== "undefined" &&
    typeof crypto !== "undefined" &&
    typeof crypto.subtle !== "undefined"
  );
}

function defaultWorker(): WorkerLike {
  return new Worker(new URL("./worker.ts", import.meta.url), { type: "module" }) as unknown as WorkerLike;
}

const toChunk = (wire: WireChunk): EmittedChunk => ({ ...wire, frame: new Uint8Array(wire.frame) });

export function importXlsxStreaming(file: Blob, handlers: ImportHandlers, options: ImportClientOptions): ImportHandle {
  const worker = (options.createWorker ?? defaultWorker)();
  let settled = false;
  let reject: (error: Error) => void = () => {};
  let resolve: (summary: ImportSummary) => void = () => {};
  const promise = new Promise<ImportSummary>((ok, fail) => {
    resolve = ok;
    reject = fail;
  });

  const finish = (action: () => void) => {
    if (settled) return;
    settled = true;
    worker.onmessage = worker.onerror = worker.onmessageerror = null;
    worker.terminate();
    action();
  };

  // Chunks the caller is still working on. "done" can arrive before they finish (the worker only
  // waits for a credit, not for the work), so the import resolves once they have all settled.
  const inFlight = new Set<Promise<void>>();

  worker.onmessage = (event) => {
    const message = event.data;
    switch (message.t) {
      case "chunk": {
        const work = Promise.resolve()
          .then(() => handlers.onChunk(toChunk(message.chunk)))
          .then(
            () => {
              if (!settled) worker.postMessage({ t: "credit", n: 1 });
            },
            (error) => finish(() => reject(error instanceof Error ? error : new Error(String(error)))),
          )
          .finally(() => inFlight.delete(work));
        inFlight.add(work);
        return;
      }
      case "tabDone":
        handlers.onTabDone?.(message.summary);
        return;
      case "progress":
        handlers.onProgress?.(message.progress);
        return;
      case "done": {
        const summary = message.summary;
        void Promise.all(inFlight).then(() => finish(() => resolve(summary)));
        return;
      }
      case "error":
        finish(() => reject(new ImportError(message.code, message.message, message.detail)));
        return;
    }
  };
  const crashed = () =>
    finish(() => reject(new ImportError("UNKNOWN", "The import stopped unexpectedly (probably out of memory).")));
  worker.onerror = crashed;
  worker.onmessageerror = crashed;

  const { createWorker: _createWorker, ...start } = options;
  void _createWorker;
  worker.postMessage({ t: "start", file, options: start });

  return {
    promise,
    cancel: () => finish(() => reject(new ImportError("CANCELLED", "The import was cancelled"))),
    setMaxChunkBytes: (bytes) => {
      if (!settled) worker.postMessage({ t: "limit", maxChunkBytes: bytes });
    },
  };
}
