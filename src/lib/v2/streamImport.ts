// Big xlsx import straight to Blossom, without ever holding the workbook in memory.
//
//   1. The worker (xlsxStream) reads the file and cuts each worksheet into chunks.
//   2. Each chunk is encrypted here as it arrives and its ciphertext staged in a Blob; the
//      plaintext is dropped. The worker waits (credits) while we are busy, so memory stays flat.
//   3. When the whole file is read every hash is known, so the upload is signed ONCE (BUD-11
//      multi-x; split only if a server cannot take a header that big) and the blobs are sent.
//   4. The root is built from the chunk lists, exactly as an editor save would build it.
//
// Nothing is uploaded until the file has been read completely: a file that fails to parse costs
// no signer prompt and leaves nothing behind on the servers.
import type { FileSigner } from "@formstr/drive-sdk";
import { finishBlossomRoot, type TabMemory } from "../crypto";
import { WORKBOOK_SCHEMA, type WorkbookDocument, type WorkbookSheet } from "../workbook";
import { randomId } from "../utils";
import { importXlsxStreaming, type ImportClientOptions } from "../xlsxStream/client";
import type { StartOptions } from "../xlsxStream/protocol";
import type { ImportSummary } from "../xlsxStream/runImport";
import {
  MAX_REPLICAS,
  PayloadTooLargeError,
  chunkLimitFor,
  encryptFrame,
  learnServerLimit,
  preflightChunkSize,
  sanitizeServers,
  uploadStaged,
  type BlossomEnv,
  type StagedBlob,
  type UploadedBlob,
} from "./blossom";
import { MAX_CHUNK_BYTES } from "./chunks";
import type { UploadProgress, UploadTabProgress } from "./progress";
import { SheetTooLargeError, type SheetRootV2 } from "./root";

/** xlsx files above this size use the streaming import; smaller ones keep the full-fidelity in-memory importer. */
export const STREAM_IMPORT_MIN_BYTES = 8 * 1024 * 1024;

export interface StreamImportOptions {
  title: string;
  servers: string[];
  replicas?: number;
  signer?: FileSigner;
  env?: BlossomEnv;
  onProgress?: (progress: UploadProgress) => void;
  signal?: AbortSignal;
  /** Test hook passed to the worker client. */
  createWorker?: ImportClientOptions["createWorker"];
  /** Test hook: chunk writer knobs (a smaller max gives several chunks from a small file). */
  start?: Partial<StartOptions>;
}

export interface StreamImportResult {
  root: SheetRootV2;
  summary: ImportSummary;
}

const mb = (bytes: number) => `${(bytes / 1024 / 1024).toFixed(bytes < 10 * 1024 * 1024 ? 1 : 0)} MB`;

export async function importXlsxToBlossom(file: Blob, options: StreamImportOptions): Promise<StreamImportResult> {
  const servers = sanitizeServers(options.servers);
  if (!options.signer || servers.length === 0) throw new SheetTooLargeError(file.size);
  const signer = options.signer;
  const replicas = Math.max(1, options.replicas ?? MAX_REPLICAS);

  const progress: UploadProgress = {
    phase: "preparing",
    message: "Checking your storage servers…",
    tabs: [],
    bytesDone: 0,
    bytesTotal: 0,
    blobsDone: 0,
    blobsTotal: 0,
    servers,
  };
  const publish = (patch: Partial<UploadProgress>) => options.onProgress?.({ ...Object.assign(progress, patch) });
  const fail = (error: unknown): never => {
    publish({ phase: "error", message: "The import did not finish. Nothing was added to your library.", error: error instanceof Error ? error.message : String(error) });
    throw error;
  };

  try {
    options.signal?.throwIfAborted();
    publish({});
    // Learn upload limits before reading anything, so chunks are cut to fit the first time.
    await preflightChunkSize(servers, MAX_CHUNK_BYTES, options.env);

    // 1-2. Read and stage.
    const staged = new Map<string, StagedBlob>();
    const tabs: UploadTabProgress[] = [];
    const handle = importXlsxStreaming(
      file,
      {
        onChunk: async (chunk) => {
          staged.set(`${chunk.tab}:${chunk.id}`, await encryptFrame(chunk.frame));
        },
        onTabDone: (summary) => {
          tabs[summary.tab] = { name: summary.name, chunks: summary.chunks, stored: 0 };
          publish({ tabs: [...tabs] });
        },
        onProgress: (read) =>
          publish({
            message: `Reading '${read.tabName}' (${read.tab + 1} of ${read.tabCount}): ${mb(read.bytesRead)} of ${mb(read.bytesTotal)}, ${read.cells.toLocaleString()} cells`,
          }),
      },
      {
        ...options.start,
        maxChunkBytes: Math.min(options.start?.maxChunkBytes ?? Infinity, chunkLimitFor(servers)),
        createWorker: options.createWorker,
      },
    );
    const abort = () => handle.cancel();
    options.signal?.addEventListener("abort", abort, { once: true });
    let summary: ImportSummary;
    try {
      summary = await handle.promise;
    } finally {
      options.signal?.removeEventListener("abort", abort);
    }
    options.signal?.throwIfAborted();

    // 3. One authorization over every hash, then upload in manifest order.
    const order = summary.tabs.flatMap((tab) => tab.entries.map((entry) => ({ tab: tab.tab, entry })));
    const items = order.map(({ tab, entry }) => {
      const item = staged.get(`${tab}:${entry.id}`);
      if (!item) throw new Error(`Chunk ${entry.id} of tab ${tab + 1} was never received.`);
      return item;
    });
    staged.clear();
    const tabNames = summary.tabs.map((tab) => ({ name: tab.name, chunks: tab.chunks, stored: 0 }));
    publish({
      message: `Sending ${items.length} encrypted file${items.length === 1 ? "" : "s"} (${summary.cells.toLocaleString()} cells)…`,
      tabs: tabNames,
      blobsTotal: items.length,
      blobsDone: 0,
    });
    const tabOfIndex = order.map(({ tab }) => tab);
    const upload = (list: StagedBlob[], trackTabs: boolean) =>
      uploadStaged(list, {
        servers,
        signer,
        replicas,
        ...options.env,
        onProgress: (event) => {
          const patch: Partial<UploadProgress> = { phase: "uploading", bytesDone: event.bytesDone, bytesTotal: event.bytesTotal };
          if (trackTabs) {
            const stored = tabNames.map(() => 0);
            tabOfIndex.slice(0, event.blobsDone).forEach((tab) => (stored[tab] += 1));
            patch.blobsDone = event.blobsDone;
            patch.tabs = tabNames.map((tab, i) => ({ ...tab, stored: stored[i] }));
          }
          publish(patch);
        },
      });
    let uploaded: UploadedBlob[];
    try {
      uploaded = await upload(items, true);
    } catch (error) {
      if (error instanceof PayloadTooLargeError) {
        learnServerLimit(error.server, error.size);
        throw new Error(
          `${error.server} refused a ${mb(error.size)} piece of this file. Its limit is now known: import the file again and the pieces will be cut smaller.`,
        );
      }
      throw error;
    }

    // 4. Root.
    const now = Date.now();
    const sheets = summary.tabs.map(
      (tab, order): WorkbookSheet => ({
        id: randomId(8),
        name: tab.name,
        order,
        hidden: tab.hidden || undefined,
        rowCount: tab.rows,
        colCount: tab.cols,
        cells: {},
      }),
    );
    const doc: WorkbookDocument = {
      schema: WORKBOOK_SCHEMA,
      version: 1,
      title: options.title,
      activeSheetId: (sheets.find((sheet) => !sheet.hidden) ?? sheets[0]).id,
      createdAt: now,
      updatedAt: now,
      sheets,
    };
    let next = 0;
    const memory: TabMemory[] = summary.tabs.map((tab) =>
      tab.entries.map((entry) => {
        const blob = uploaded[next++];
        return { id: entry.id, ref: blob.ref, rowStart: entry.rowStart, rowCount: entry.rowCount, cells: entry.cells, servers: blob.servers };
      }),
    );
    const root = await finishBlossomRoot(doc, memory, (frames) =>
      Promise.all(frames.map(encryptFrame)).then((list) => upload(list, false)),
    );
    publish({ phase: "done", message: `Imported ${summary.cells.toLocaleString()} cells. All files are stored and encrypted.`, blobsDone: items.length });
    return { root, summary };
  } catch (error) {
    return fail(error);
  }
}
