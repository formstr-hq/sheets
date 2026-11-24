// The streaming xlsx import, independent of where it runs: reads a Blob/File, cuts every sheet into
// encoded chunks, and hands each chunk to a sink. worker.ts adapts the sink to postMessage; tests
// and tools use it directly in Node. Memory stays bounded: the shared strings table (off-heap),
// one sheet's parser state, and at most one chunk's worth of rows per tab.
import { ChunkStreamWriter, type StreamedChunkInfo } from "../v2/chunkStream";
import { ImportError, toImportError } from "./errors";
import { loadSharedStrings, STRING_BUDGET_BYTES } from "./strings";
import { readSheet } from "./sheetStream";
import { readWorkbookMeta } from "./workbookMeta";
import { openZip } from "./zip";

export interface EmittedChunk {
  /** Index of the tab in workbook order. */
  tab: number;
  /** Emission order within the tab; the head chunk (the one carrying the layout) is last. */
  seq: number;
  isHead: boolean;
  id: string;
  rowStart: number;
  rowCount: number;
  cells: number;
  /** FSB2 frame (compressed canonical JSON). */
  frame: Uint8Array;
  /** sha256 hex of the frame. */
  hash: string;
}

export interface TabSummary {
  tab: number;
  name: string;
  hidden: boolean;
  /** Rows and columns the editor should size the tab to (never below the editor defaults). */
  rows: number;
  cols: number;
  /** Rows that hold at least one cell. */
  filledRows: number;
  cells: number;
  chunks: number;
  /** Manifest entries in rowStart order (no frame bytes). */
  entries: StreamedChunkInfo[];
}

export interface ReadProgress {
  tab: number;
  tabName: string;
  tabCount: number;
  /** Compressed bytes of worksheet parts read so far, and in total. */
  bytesRead: number;
  bytesTotal: number;
  /** Cells read so far across all tabs. */
  cells: number;
}

export interface ImportSink {
  /** Resolves when the consumer can take more; the import waits for it (back-pressure). */
  chunk(chunk: EmittedChunk): Promise<void>;
  tabDone(summary: TabSummary): void;
  progress(progress: ReadProgress): void;
}

export interface ImportOptions {
  /** Largest chunk in compressed bytes, or a function read before every cut. */
  maxChunkBytes: number | (() => number);
  stringBudgetBytes?: number;
  signal?: AbortSignal;
  /** Raw-size knobs of the chunk writer (tests use small values). */
  headRawBytes?: number;
  rawCap?: number;
  /** Minimum ms between progress calls. Default 250. */
  progressIntervalMs?: number;
}

export interface ImportSummary {
  title: string;
  tabs: TabSummary[];
  cells: number;
  /** What was not carried over, in words for the user. */
  dropped: string[];
}

export async function runImport(source: Blob, sink: ImportSink, options: ImportOptions): Promise<ImportSummary> {
  try {
    return await importAll(source, sink, options);
  } catch (error) {
    throw options.signal?.aborted ? new ImportError("CANCELLED", "The import was cancelled") : toImportError(error);
  }
}

async function importAll(source: Blob, sink: ImportSink, options: ImportOptions): Promise<ImportSummary> {
  const { signal } = options;
  const zip = await openZip(source, { signal });
  const meta = await readWorkbookMeta(zip, signal);
  if (meta.sheets.length === 0) throw new ImportError("NOT_XLSX", "This workbook has no worksheets.");
  const strings = await loadSharedStrings(zip, { budgetBytes: options.stringBudgetBytes ?? STRING_BUDGET_BYTES, signal });

  const bytesTotal = meta.sheets.reduce((sum, ref) => sum + (zip.entries.get(ref.path.replace(/^\//, ""))?.csize ?? 0), 0);
  const interval = options.progressIntervalMs ?? 250;
  let bytesDone = 0;
  let cellsTotal = 0;
  let lastProgress = 0;

  const tabs: TabSummary[] = [];
  const dropped: string[] = [...meta.notes];
  for (let tab = 0; tab < meta.sheets.length; tab++) {
    const ref = meta.sheets[tab];
    const entryBytes = zip.entries.get(ref.path.replace(/^\//, ""))?.csize ?? 0;
    const report = (partial: number, force = false) => {
      const now = Date.now();
      if (!force && now - lastProgress < interval) return;
      lastProgress = now;
      sink.progress({ tab, tabName: ref.name, tabCount: meta.sheets.length, bytesRead: bytesDone + partial, bytesTotal, cells: cellsTotal });
    };
    let entryRead = 0;
    const stream = readSheet(zip, ref, {
      strings,
      styles: meta.styles,
      date1904: meta.date1904,
      signal,
      onRead: (read) => (entryRead = read),
    });
    const writer = new ChunkStreamWriter({
      max: options.maxChunkBytes,
      headRawBytes: options.headRawBytes,
      rawCap: options.rawCap,
      onChunk: async (chunk) => {
        const { frame, ...rest } = chunk;
        await sink.chunk({ tab, ...rest, frame });
      },
    });
    let cells = 0;
    let filledRows = 0;
    report(0, true);
    for await (const row of stream.rows()) {
      cells += row.cells.length;
      cellsTotal += row.cells.length;
      filledRows++;
      await writer.push({ row: row.row, cells: row.cells });
      report(entryRead);
    }
    const scan = stream.scan();
    const entries = await writer.finish(scan.layout);
    bytesDone += entryBytes;
    dropped.push(...scan.notes);
    const summary: TabSummary = {
      tab,
      name: ref.name,
      hidden: ref.hidden,
      rows: scan.layout.rowCount,
      cols: scan.layout.colCount,
      filledRows,
      cells,
      chunks: entries.length,
      entries,
    };
    tabs.push(summary);
    sink.tabDone(summary);
    report(0, true);
  }
  const title = typeof File !== "undefined" && source instanceof File ? source.name.replace(/\.[^.]+$/, "") : "";
  return { title, tabs, cells: cellsTotal, dropped };
}
