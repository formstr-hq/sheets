// Messages between the main thread and the import worker.
//
// Back-pressure is credit based: the worker starts with CREDIT_WINDOW credits, every chunk message
// spends one, and the main thread returns a credit once it has finished with the chunk (encrypted
// and staged). With no credit the worker stops reading the file.
import type { ImportErrorCode } from "./errors";
import type { EmittedChunk, ImportSummary, ReadProgress, TabSummary } from "./runImport";

/** Chunks the worker may have in flight before it must wait for the main thread. */
export const CREDIT_WINDOW = 4;

export interface StartOptions {
  maxChunkBytes: number;
  stringBudgetBytes?: number;
  headRawBytes?: number;
  rawCap?: number;
}

export type MainToWorker =
  | { t: "start"; file: Blob; options: StartOptions }
  | { t: "credit"; n: number }
  /** A smaller upload limit was learned; later cuts use it. */
  | { t: "limit"; maxChunkBytes: number }
  | { t: "cancel" };

/** A chunk as it crosses the thread boundary: the frame is a transferred ArrayBuffer. */
export type WireChunk = Omit<EmittedChunk, "frame"> & { frame: ArrayBuffer };

export type WorkerToMain =
  | { t: "chunk"; chunk: WireChunk }
  | { t: "tabDone"; summary: TabSummary }
  | { t: "progress"; progress: ReadProgress }
  | { t: "done"; summary: ImportSummary }
  | { t: "error"; code: ImportErrorCode; message: string; detail?: string };
