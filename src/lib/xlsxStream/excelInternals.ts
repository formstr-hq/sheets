// The ONLY file that deep-imports ExcelJS internals (private API, exceljs is pinned to 4.4.0).
// Everything else in xlsxStream talks to this typed facade. deepimports.test.ts guards upgrades.
import type * as ExcelJS from "exceljs";
import WorksheetReaderCjs from "exceljs/lib/stream/xlsx/worksheet-reader";
import RowCjs from "exceljs/lib/doc/row";
import utilsCjs from "exceljs/lib/utils/utils";
import StylesXformCjs from "exceljs/lib/xlsx/xform/style/styles-xform";
import WorkbookXformCjs from "exceljs/lib/xlsx/xform/book/workbook-xform";
import RelationshipsXformCjs from "exceljs/lib/xlsx/xform/core/relationships-xform";
import parseSaxCjs, { type SaxEvent } from "exceljs/lib/utils/parse-sax";

export type { SaxEvent };
export type TextSource = AsyncIterable<string | Uint8Array>;

/** An ExcelJS row as produced by the stream reader. */
export type ExcelRow = ExcelJS.Row;
export type ExcelCell = ExcelJS.Cell;

export interface WorksheetColumn {
  number?: number;
  width?: number;
}

export interface SheetReader extends AsyncIterable<ExcelRow> {
  /** Column definitions from <cols>; available once the first row has been produced. */
  readonly columns: WorksheetColumn[] | null;
}

export interface StyleManager {
  parseStream(stream: TextSource): Promise<unknown>;
  getStyleModel(id: number): unknown;
}

export interface ReaderContext {
  /** Undefined: cells hold `{ sharedString: index }` and the caller resolves them. */
  sharedStrings?: undefined;
  styles: StyleManager;
  properties: { model: { date1904?: boolean } };
}

export interface SheetRefModel {
  name: string;
  id: number;
  state?: string;
  rId?: string;
}

export interface WorkbookModel {
  sheets: SheetRefModel[];
  properties: { date1904?: boolean };
}

export interface RelationshipModel {
  Id: string;
  Type: string;
  Target: string;
  TargetMode?: string;
}

export function parseSax(source: TextSource): AsyncGenerator<SaxEvent[]> {
  return parseSaxCjs(source);
}

/** 1-based column number from the letters of a cell reference ("AB12" -> 28); 0 when absent. */
function columnOfRef(ref: string | undefined): number {
  if (!ref) return 0;
  let col = 0;
  for (let i = 0; i < ref.length; i++) {
    const code = ref.charCodeAt(i);
    if (code < 65 || code > 90) break;
    col = col * 26 + code - 64;
  }
  return col;
}

interface PendingCell {
  ref?: string;
  s: number;
  t?: string;
  f?: { text: string };
  v?: { text: string };
}

/**
 * ExcelJS WorksheetReader.parse with fixes for what its stream reader gets wrong (4.4.0):
 *  - formula results with t="e" (#N/A ...) were parseFloat'ed to NaN, losing the error;
 *  - inline strings kept only the last <t> (rich text runs) and included phonetic <rPh> text;
 *  - empty <v/> produced NaN numbers;
 *  - cells without r= and rows without r= crashed;
 *  - date-formatted formula results ignored the 1904 date system.
 * Everything else follows the original (hyperlinks, merges and <cols> are not read here).
 *
 * Derived from ExcelJS (https://github.com/exceljs/exceljs), MIT License,
 * Copyright (c) 2014-2019 Guyon Roche. Permission notice: see node_modules/exceljs/LICENSE.
 */
class FixedSheetReader extends WorksheetReaderCjs {
  async *parse(): AsyncGenerator<{ eventType: string; value: unknown }[]> {
    const { styles, properties } = this.workbook as ReaderContext;
    const date1904 = properties.model.date1904 === true;
    let inRows = false;
    let inInline = false;
    let phonetic = 0;
    let row: any = null;
    let cell: PendingCell | null = null;
    let current: { text: string } | null = null;
    let lastRow = 0;
    let lastCol = 0;
    for await (const events of parseSaxCjs(this.iterator)) {
      const out: { eventType: string; value: unknown }[] = [];
      for (const { eventType, value } of events) {
        if (eventType === "opentag") {
          const attrs = value.attributes;
          switch (value.name) {
            case "sheetData":
              inRows = true;
              break;
            case "row":
              if (inRows) {
                const r = attrs.r === undefined ? lastRow + 1 : parseInt(attrs.r, 10);
                row = new RowCjs(this, r);
                lastRow = r;
                lastCol = 0;
                if (attrs.ht) row.height = parseFloat(attrs.ht);
                if (attrs.s) {
                  const model = styles.getStyleModel(parseInt(attrs.s, 10));
                  if (model) row.style = model;
                }
              }
              break;
            case "c":
              if (row) cell = { ref: attrs.r, s: parseInt(attrs.s, 10), t: attrs.t };
              break;
            case "f":
              if (cell) current = cell.f = { text: "" };
              break;
            case "v":
              if (cell) current = cell.v = { text: "" };
              break;
            case "is":
              if (cell) {
                inInline = true;
                cell.v = { text: "" };
              }
              break;
            case "rPh":
              if (inInline) phonetic++;
              break;
            case "t":
              if (cell && inInline && phonetic === 0) current = cell.v ?? null;
              break;
          }
        } else if (eventType === "text") {
          if (current) current.text += value;
        } else {
          switch (value.name) {
            case "f":
            case "v":
            case "t":
              current = null;
              break;
            case "rPh":
              if (phonetic > 0) phonetic--;
              break;
            case "is":
              inInline = false;
              break;
            case "row":
              out.push({ eventType: "row", value: row });
              row = null;
              break;
            case "c":
              if (row && cell) {
                lastCol = columnOfRef(cell.ref) || lastCol + 1;
                fillCell(row.getCell(lastCol), cell, styles, date1904);
                cell = null;
              }
              break;
          }
        }
      }
      if (out.length > 0) yield out;
    }
  }
}

function fillCell(target: any, c: PendingCell, styles: StyleManager, date1904: boolean): void {
  if (c.s) {
    const model = styles.getStyleModel(c.s);
    if (model) target.style = model;
  }
  const text = c.v?.text;
  const number = text === undefined ? NaN : parseFloat(text);
  if (c.f) {
    const result: { formula: string; result?: unknown } = { formula: c.f.text };
    if (c.v) {
      if (c.t === "str") result.result = utilsCjs.xmlDecode(c.v.text);
      else if (c.t === "e") result.result = { error: c.v.text };
      else if (!Number.isNaN(number)) {
        result.result = date1904 && utilsCjs.isDateFmt(target.numFmt) ? utilsCjs.excelToDate(number, true) : number;
      }
    }
    target.value = result;
    return;
  }
  if (!c.v) return;
  switch (c.t) {
    case "s":
      if (!Number.isNaN(number)) target.value = { sharedString: number };
      break;
    case "inlineStr":
    case "str":
      target.value = utilsCjs.xmlDecode(c.v.text);
      break;
    case "e":
      target.value = { error: c.v.text };
      break;
    case "b":
      if (!Number.isNaN(number)) target.value = number !== 0;
      break;
    default:
      if (Number.isNaN(number)) break;
      target.value = utilsCjs.isDateFmt(target.numFmt) ? utilsCjs.excelToDate(number, date1904) : number;
  }
}

export function createSheetReader(args: { id: string; source: TextSource; context: ReaderContext }): SheetReader {
  return new FixedSheetReader({
    workbook: args.context,
    id: args.id,
    iterator: args.source,
    options: { worksheets: "emit", hyperlinks: "ignore" },
  }) as unknown as SheetReader;
}

/**
 * A style manager. Until parseStream() has loaded xl/styles.xml it knows no styles. ExcelJS throws
 * for a missing or partial styles.xml (for example no cellXfs), so failed lookups return null.
 */
export function createStyleManager(): StyleManager {
  const inner = new StylesXformCjs();
  inner.init();
  return {
    parseStream: (stream) => inner.parseStream(stream),
    getStyleModel(id) {
      try {
        return inner.getStyleModel(id);
      } catch {
        return null;
      }
    },
  };
}

export async function parseWorkbookXml(source: TextSource): Promise<WorkbookModel> {
  return (await new WorkbookXformCjs().parseStream(source)) as WorkbookModel;
}

export async function parseRelationships(source: TextSource): Promise<RelationshipModel[]> {
  return ((await new RelationshipsXformCjs().parseStream(source)) ?? []) as RelationshipModel[];
}
