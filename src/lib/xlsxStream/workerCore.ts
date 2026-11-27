// The import worker's logic, separate from `self` so it can be tested in-process.
import { ImportError, toImportError } from "./errors";
import { CREDIT_WINDOW, type MainToWorker, type WorkerToMain } from "./protocol";
import { runImport, type EmittedChunk, type ImportSink } from "./runImport";

/** Counting gate: acquire() waits until a credit is free. */
export class CreditGate {
  private credits: number;
  private waiting: Array<() => void> = [];

  constructor(credits: number) {
    this.credits = credits;
  }

  /** Waits for a credit; rejects when `signal` aborts first. */
  acquire(signal?: AbortSignal): Promise<void> {
    if (signal?.aborted) return Promise.reject(new ImportError("CANCELLED", "The import was cancelled"));
    if (this.credits > 0) {
      this.credits--;
      return Promise.resolve();
    }
    return new Promise((resolve, reject) => {
      const waiter = () => {
        signal?.removeEventListener("abort", onAbort);
        resolve();
      };
      const onAbort = () => {
        const at = this.waiting.indexOf(waiter);
        if (at >= 0) this.waiting.splice(at, 1);
        reject(new ImportError("CANCELLED", "The import was cancelled"));
      };
      signal?.addEventListener("abort", onAbort, { once: true });
      this.waiting.push(waiter);
    });
  }

  release(n = 1): void {
    for (let i = 0; i < n; i++) {
      const next = this.waiting.shift();
      if (next) next();
      else this.credits++;
    }
  }
}

export type Post = (message: WorkerToMain, transfer?: Transferable[]) => void;

/** Returns the message handler for one worker. */
export function createWorkerHandler(post: Post): (message: MainToWorker) => void {
  const gate = new CreditGate(CREDIT_WINDOW);
  const controller = new AbortController();
  let started = false;
  let maxChunkBytes = 0;

  return (message) => {
    switch (message.t) {
      case "credit":
        gate.release(message.n);
        return;
      case "limit":
        maxChunkBytes = Math.min(maxChunkBytes, message.maxChunkBytes);
        return;
      case "cancel":
        controller.abort();
        return;
      case "start": {
        if (started) return;
        started = true;
        maxChunkBytes = message.options.maxChunkBytes;
        const sink: ImportSink = {
          async chunk(chunk: EmittedChunk) {
            await gate.acquire(controller.signal);
            // Copy only when the frame is a view into a bigger buffer; otherwise transfer as is.
            const { frame, ...rest } = chunk;
            const buffer =
              frame.byteOffset === 0 && frame.byteLength === frame.buffer.byteLength
                ? (frame.buffer as ArrayBuffer)
                : (frame.slice().buffer as ArrayBuffer);
            post({ t: "chunk", chunk: { ...rest, frame: buffer } }, [buffer]);
          },
          tabDone: (summary) => post({ t: "tabDone", summary }),
          progress: (progress) => post({ t: "progress", progress }),
        };
        runImport(message.file, sink, {
          maxChunkBytes: () => maxChunkBytes,
          stringBudgetBytes: message.options.stringBudgetBytes,
          headRawBytes: message.options.headRawBytes,
          rawCap: message.options.rawCap,
          signal: controller.signal,
        }).then(
          (summary) => post({ t: "done", summary }),
          (error) => {
            const failure = toImportError(error);
            post({ t: "error", code: failure.code, message: failure.message, detail: failure.detail });
          },
        );
        return;
      }
    }
  };
}
