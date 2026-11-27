// Streams one worksheet: rows through ExcelJS's WorksheetReader, plus a side scan of the text
// before <sheetData> (column widths) and after </sheetData> (merged cells). The reader itself
// ignores both: its column support is broken in 4.4.0 (it calls Column.fromModel with the wrong
// arguments) and it never reads mergeCells.
import { DEFAULT_COLS, DEFAULT_ROWS } from "../constants";
import type { WorkbookMerge } from "../workbook";
import { ImportError } from "./errors";
import { createSheetReader, parseSax, type StyleManager } from "./excelInternals";
import { mapRow, type MappedRow } from "./rowMap";
import type { PackedStrings } from "./strings";
import type { SheetRef } from "./workbookMeta";
import type { ZipArchive } from "./zip";

/** Text kept before <sheetData> (<cols> lives there). */
const HEAD_CAP_CHARS = 1024 * 1024;
/** Text kept after </sheetData> (merges, validations, conditional formats). */
const TAIL_CAP_CHARS = 32 * 1024 * 1024;
/** Per-sheet limit on stored custom row heights, to keep the layout block small. */
const MAX_ROW_HEIGHTS = 200_000;
const MAX_COLUMNS = 16384;
const MAX_SHEET_ROWS = 1_048_576;
const SHEET_DATA_CLOSE = "</sheetData>";
const SHEET_DATA_OPEN = /<sheetData\b[^>]*>/;

export interface SheetContext {
  strings: PackedStrings | null;
  styles: StyleManager;
  date1904: boolean;
  signal?: AbortSignal;
  /** Compressed bytes of this sheet's part read so far. */
  onRead?: (bytesRead: number) => void;
}

/** What is known about a sheet once all its rows have been read. */
export interface SheetScan {
  /** Data validations (dropdowns) found after the rows; they are not imported. */
  validations: number;
  /** Shared-formula dependents whose formula was dropped. */
  sharedFormulaValues: number;
  /** Layout fields for the tab body, in the app's units (same conversions as the xlsx importer). */
  layout: {
    rowCount: number;
    colCount: number;
    defaultRowHeight: number;
    defaultColWidth: number;
    columnWidths?: Record<number, number>;
    rowHeights?: Record<number, number>;
    merges?: WorkbookMerge[];
  };
  /** Things that were not carried over, in words for the user. */
  notes: string[];
}

export interface SheetStream {
  /** Rows with at least one cell, ascending. */
  rows(): AsyncGenerator<MappedRow>;
  /** Only valid after `rows()` has finished. */
  scan(): SheetScan;
}

/** Passes text through while capturing what precedes <sheetData> and what follows it. */
class SheetScanner {
  head = "";
  readonly tail: string[] = [];
  tailTruncated = false;
  private tailChars = 0;
  private phase: "head" | "data" | "tail" = "head";
  private headBuffer = "";
  private carry = "";

  async *pass(source: AsyncIterable<string>): AsyncGenerator<string> {
    for await (const chunk of source) {
      this.feed(chunk);
      yield chunk;
    }
  }

  private feed(chunk: string): void {
    if (this.phase === "tail") return this.addTail(chunk);
    if (this.phase === "data") return this.scanData(chunk);
    this.headBuffer += chunk;
    const match = SHEET_DATA_OPEN.exec(this.headBuffer);
    if (match) {
      this.head = this.headBuffer.slice(0, match.index);
      const rest = this.headBuffer.slice(match.index + match[0].length);
      this.headBuffer = "";
      if (match[0].endsWith("/>")) {
        this.phase = "tail";
        this.addTail(rest);
      } else {
        this.phase = "data";
        this.scanData(rest);
      }
    } else if (this.headBuffer.length > HEAD_CAP_CHARS) {
      this.headBuffer = "";
      this.phase = "data";
    }
  }

  private scanData(text: string): void {
    const joined = this.carry + text;
    const at = joined.indexOf(SHEET_DATA_CLOSE);
    if (at >= 0) {
      this.phase = "tail";
      this.carry = "";
      this.addTail(joined.slice(at + SHEET_DATA_CLOSE.length));
    } else {
      this.carry = joined.slice(-(SHEET_DATA_CLOSE.length - 1));
    }
  }

  private addTail(text: string): void {
    if (text.length === 0 || this.tailTruncated) return;
    if (this.tailChars + text.length > TAIL_CAP_CHARS) {
      this.tailTruncated = true;
      return;
    }
    this.tailChars += text.length;
    this.tail.push(text);
  }
}

async function* chunksOf(pieces: string[]): AsyncGenerator<string> {
  for (const piece of pieces) yield piece;
}

function columnIndex(letters: string): number {
  let index = 0;
  for (const char of letters.toUpperCase()) index = index * 26 + (char.charCodeAt(0) - 64);
  return index - 1;
}

function decodeAddress(address: string): { row: number; column: number } | null {
  const match = address.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  return match ? { column: columnIndex(match[1]), row: Number(match[2]) - 1 } : null;
}

/** <cols> as the xlsx importer sees it: width in pixels per 0-based column index. */
async function parseColumnWidths(head: string): Promise<Record<number, number>> {
  const cols: { min: number; max: number; width: number }[] = [];
  const open = head.indexOf("<cols");
  if (open < 0) return {};
  const close = head.indexOf("</cols>", open);
  const block = close < 0 ? head.slice(open) + "</cols>" : head.slice(open, close + 7);
  for await (const events of parseSax(chunksOf([block]))) {
    for (const { eventType, value } of events) {
      if (eventType !== "opentag" || value.name !== "col") continue;
      const min = parseInt(value.attributes.min, 10);
      const max = parseInt(value.attributes.max, 10);
      // ExcelJS's workbook loader reads a missing width as 9.
      const width = value.attributes.width === undefined ? 9 : parseFloat(value.attributes.width);
      if (Number.isFinite(min) && Number.isFinite(max)) cols.push({ min, max, width });
    }
  }
  cols.sort((a, b) => a.min - b.min);
  const widths: Record<number, number> = {};
  for (const col of cols) {
    if (!col.width) continue;
    const pixels = Math.max(64, Math.round(col.width * 7.5));
    for (let n = Math.max(1, col.min); n <= Math.min(col.max, MAX_COLUMNS); n++) widths[n - 1] = pixels;
  }
  return widths;
}

async function parseTail(tail: string[]): Promise<{ merges: WorkbookMerge[]; validations: number }> {
  const merges: WorkbookMerge[] = [];
  let validations = 0;
  // The tail starts after </sheetData>, so give the parser the root element it closes.
  for await (const events of parseSax(chunksOf(["<worksheet>", ...tail]))) {
    for (const { eventType, value } of events) {
      if (eventType !== "opentag") continue;
      if (value.name === "dataValidation") validations++;
      if (value.name !== "mergeCell") continue;
      const [start, end] = String(value.attributes.ref ?? "").split(":");
      if (!start || !end) continue;
      const from = decodeAddress(start);
      const to = decodeAddress(end);
      if (!from || !to) continue;
      merges.push({ r: from.row, c: from.column, rs: to.row - from.row + 1, cs: to.column - from.column + 1 });
    }
  }
  return { merges, validations };
}

export function readSheet(zip: ZipArchive, ref: SheetRef, context: SheetContext): SheetStream {
  const scanner = new SheetScanner();
  const rowHeights: Record<number, number> = {};
  let heightCount = 0;
  let heightsDropped = false;
  let maxRow = 0;
  let maxCols = 0;
  let sharedFormulaValues = 0;
  let scan: SheetScan | null = null;

  async function* rows(): AsyncGenerator<MappedRow> {
    const reader = createSheetReader({
      id: ref.path,
      source: scanner.pass(zip.text(ref.path, { signal: context.signal, onRead: context.onRead })),
      context: { styles: context.styles, properties: { model: { date1904: context.date1904 } } },
    });
    let last = -1;
    for await (const row of reader) {
      if (!Number.isFinite(row.number) || row.number < 1) continue;
      if (row.number - 1 <= last) {
        throw new ImportError(
          "CORRUPT",
          `Sheet "${ref.name}" has rows that are not in order (row ${row.number}), which this import cannot handle. Open and re-save the file in a spreadsheet app, then try again.`,
        );
      }
      last = row.number - 1;
      const mapped = mapRow(row, context.strings);
      maxRow = Math.max(maxRow, row.number);
      maxCols = Math.max(maxCols, mapped.cellCount);
      sharedFormulaValues += mapped.sharedFormulaValues;
      // Like the in-memory importer, heights are kept only for rows that hold values.
      if (mapped.height && mapped.hasValues) {
        if (heightCount < MAX_ROW_HEIGHTS) {
          rowHeights[mapped.row] = Math.round(mapped.height * 1.333);
          heightCount++;
        } else {
          heightsDropped = true;
        }
      }
      if (mapped.cells.length > 0) yield mapped;
    }
    const notes: string[] = [];
    let columnWidths: Record<number, number> = {};
    let merges: WorkbookMerge[] = [];
    let validations = 0;
    try {
      columnWidths = await parseColumnWidths(scanner.head);
    } catch {
      notes.push(`Column widths could not be read for "${ref.name}".`);
    }
    try {
      ({ merges, validations } = await parseTail(scanner.tail));
    } catch {
      notes.push(`Merged cells could not be read for "${ref.name}".`);
    }
    if (scanner.tailTruncated) notes.push(`Merged cells after row data were not read for "${ref.name}" (the section is very large).`);
    if (heightsDropped) notes.push(`Some custom row heights were not kept in "${ref.name}".`);
    if (maxRow > MAX_SHEET_ROWS || maxCols > MAX_COLUMNS) notes.push(`"${ref.name}" is larger than a spreadsheet can be; its size was limited.`);
    if (validations > 0) notes.push(`${validations} data validation(s) (dropdown lists) in "${ref.name}" were not imported.`);
    if (sharedFormulaValues > 0) {
      notes.push(`${sharedFormulaValues} cell(s) in "${ref.name}" use a formula shared with other cells; they keep their values but not the formula.`);
    }
    scan = {
      validations,
      sharedFormulaValues,
      layout: {
        rowCount: Math.min(Math.max(maxRow, DEFAULT_ROWS), MAX_SHEET_ROWS),
        colCount: Math.min(Math.max(maxCols, DEFAULT_COLS), MAX_COLUMNS),
        defaultRowHeight: 30,
        defaultColWidth: 116,
        columnWidths: Object.keys(columnWidths).length > 0 ? columnWidths : undefined,
        rowHeights: heightCount > 0 ? rowHeights : undefined,
        merges: merges.length > 0 ? merges : undefined,
      },
      notes,
    };
  }

  return {
    rows,
    scan() {
      if (!scan) throw new Error("scan() called before the sheet was fully read");
      return scan;
    },
  };
}
