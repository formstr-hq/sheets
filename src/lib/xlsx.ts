import type * as ExcelJS from "exceljs";
import type { Cell, CellWithRowAndCol, Sheet } from "./legacy/fortuneFormat";
import { DEFAULT_COLS, DEFAULT_ROWS } from "./constants";
import { fileSafeName, randomId } from "./utils";
import { valueFromExcelCell } from "./xlsxCells";
import {
  fromDateSerial as fromSerial,
  workbookToFortuneSheets,
  type WorkbookDocument,
} from "./workbook";

type ExcelJSModule = typeof import("exceljs");

async function loadExcelJS(): Promise<ExcelJSModule> {
  const imported = await import("exceljs");
  return ((imported as unknown as { default?: ExcelJSModule }).default ?? imported) as ExcelJSModule;
}

function hexToArgb(value?: string): string | undefined {
  if (!value) return undefined;
  const clean = value.replace("#", "");
  if (clean.length === 6) return `FF${clean.toUpperCase()}`;
  if (clean.length === 8) return clean.toUpperCase();
  return undefined;
}

function decodeAddress(address: string): { row: number; column: number } | null {
  const match = address.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) return null;
  return { column: columnIndexFromName(match[1]), row: Number(match[2]) - 1 };
}

function parseMerges(worksheet: ExcelJS.Worksheet): Array<{ r: number; c: number; rs: number; cs: number }> {
  const merges = (worksheet as unknown as { model?: { merges?: unknown[] } }).model?.merges ?? [];
  const out: Array<{ r: number; c: number; rs: number; cs: number }> = [];
  for (const merge of merges) {
    if (typeof merge === "string") {
      const [start, end] = merge.split(":");
      if (!start || !end) continue;
      const startCell = decodeAddress(start);
      const endCell = decodeAddress(end);
      if (!startCell || !endCell) continue;
      out.push({
        r: startCell.row,
        c: startCell.column,
        rs: endCell.row - startCell.row + 1,
        cs: endCell.column - startCell.column + 1,
      });
    }
  }
  return out;
}

function parseDataValidations(worksheet: ExcelJS.Worksheet): Record<string, unknown> {
  const model = (worksheet as unknown as { dataValidations?: { model?: Record<string, unknown> } })
    .dataValidations?.model;
  if (!model) return {};
  const result: Record<string, unknown> = {};
  for (const [address, raw] of Object.entries(model)) {
    const validation = raw as {
      type?: string;
      formulae?: unknown[];
      allowBlank?: boolean;
    };
    if (validation.type !== "list") continue;
    const formula = validation.formulae?.[0];
    if (typeof formula !== "string") continue;
    const options = formula
      .replace(/^"|"$/g, "")
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean);
    if (options.length === 0) continue;
    const rowCol = address.match(/^([A-Z]+)(\d+)$/i);
    if (!rowCol) continue;
    const column = columnIndexFromName(rowCol[1]);
    const row = Number(rowCol[2]) - 1;
    result[`${row}_${column}`] = {
      type: "dropdown",
      type2: "",
      rangeTxt: address,
      value1: options.join(","),
      value2: "",
      validity: "",
      remote: false,
      prohibitInput: false,
      hintShow: false,
      hintValue: "",
      checked: false,
    };
  }
  return result;
}

function columnIndexFromName(name: string): number {
  let index = 0;
  for (const char of name.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index - 1;
}

export interface ImportedWorkbook {
  title: string;
  sheets: Sheet[];
}

function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let value = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (quoted) {
      if (char === '"') {
        if (text[i + 1] === '"') {
          value += '"';
          i++;
        } else {
          quoted = false;
        }
      } else {
        value += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ",") {
      row.push(value);
      value = "";
    } else if (char === "\n") {
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
    } else if (char === "\r") {
      if (text[i + 1] === "\n") i++;
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
    } else {
      value += char;
    }
  }
  if (value.length > 0 || row.length > 0) {
    row.push(value);
    rows.push(row);
  }
  return rows;
}

function csvToSheet(text: string, name: string): Sheet {
  const rows = parseCsv(text);
  const values: CellWithRowAndCol[] = [];
  rows.forEach((row, r) => {
    row.forEach((raw, c) => {
      if (raw === "") return;
      values.push({
        r,
        c,
        v: { v: raw, m: raw, ct: { fa: "General", t: "s" } },
      });
    });
  });
  const maxCols = rows.reduce((max, row) => Math.max(max, row.length), 0);
  return {
    id: randomId(10),
    name,
    row: Math.max(rows.length, DEFAULT_ROWS),
    column: Math.max(maxCols, DEFAULT_COLS),
    status: 1,
    order: 0,
    celldata: values,
    config: { rowlen: {}, columnlen: {}, merge: {}, borderInfo: [] },
    showGridLines: 1,
    defaultRowHeight: 30,
    defaultColWidth: 116,
  };
}

export async function parseSpreadsheetFile(file: File, fallbackTitle: string): Promise<ImportedWorkbook> {
  const title = file.name.replace(/\.[^.]+$/, "") || fallbackTitle;
  if (/\.csv$/i.test(file.name)) {
    const text = await file.text();
    return { title, sheets: [csvToSheet(text, "Sheet 1")] };
  }
  const buffer = await file.arrayBuffer();
  const ExcelJS = await loadExcelJS();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as ArrayBuffer);
  const sheets: Sheet[] = workbook.worksheets.map((worksheet, index) => {
    const values: CellWithRowAndCol[] = [];
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const parsed = valueFromExcelCell(cell);
        if (parsed) values.push({ r: rowNumber - 1, c: colNumber - 1, v: parsed });
      });
    });
    const rowHeights: Record<string, number> = {};
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (row.height) rowHeights[rowNumber - 1] = Math.round(row.height * 1.333);
    });
    const columnWidths: Record<string, number> = {};
    worksheet.columns?.forEach((column, columnIndex) => {
      if (column.width) columnWidths[columnIndex] = Math.max(64, Math.round(column.width * 7.5));
    });
    const merges = parseMerges(worksheet);
    return {
      id: randomId(10),
      name: worksheet.name,
      row: Math.max(worksheet.rowCount || 0, DEFAULT_ROWS),
      column: Math.max(worksheet.columnCount || 0, DEFAULT_COLS),
      status: index === 0 ? 1 : 0,
      order: index,
      celldata: values,
      config: {
        rowlen: rowHeights,
        columnlen: columnWidths,
        merge: Object.fromEntries(
          merges.map((merge) => [`${merge.r}_${merge.c}`, merge]),
        ),
        borderInfo: [],
      },
      dataVerification: parseDataValidations(worksheet),
      showGridLines: 1,
      defaultRowHeight: 30,
      defaultColWidth: 116,
    } satisfies Sheet;
  });
  return {
    title,
    sheets: sheets.length > 0 ? sheets : [],
  };
}

function applyCellToWorksheet(sheet: ExcelJS.Worksheet, cell: Cell, row: number, column: number): void {
  const excelCell = sheet.getCell(row, column);
  const value = cell.v;
  const format = cell.ct?.fa;

  if (cell.f) {
    const formula = cell.f.startsWith("=") ? cell.f.slice(1) : cell.f;
    let result: unknown = value;
    if (cell.ct?.t === "d" && typeof value === "number") result = fromSerial(value);
    excelCell.value = {
      formula,
      result: result as number | string | boolean | Date | undefined,
    };
  } else if (cell.ct?.t === "d" && typeof value === "number") {
    excelCell.value = fromSerial(value);
  } else if (typeof value === "number" || typeof value === "string" || typeof value === "boolean") {
    excelCell.value = value;
  } else {
    return;
  }

  if (cell.bl) excelCell.font = { ...excelCell.font, bold: true };
  if (cell.it) excelCell.font = { ...excelCell.font, italic: true };
  if (cell.un) excelCell.font = { ...excelCell.font, underline: true };
  if (cell.cl) excelCell.font = { ...excelCell.font, strike: true };
  if (cell.fs) excelCell.font = { ...excelCell.font, size: cell.fs };
  if (cell.ff) excelCell.font = { ...excelCell.font, name: String(cell.ff) };
  if (cell.fc) excelCell.font = { ...excelCell.font, color: { argb: hexToArgb(cell.fc) } };
  if (cell.bg) {
    excelCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: hexToArgb(cell.bg) },
    };
  }
  const horizontal = cell.ht === 0 ? "center" : cell.ht === 1 ? "left" : cell.ht === 2 ? "right" : undefined;
  const vertical = cell.vt === 0 ? "middle" : cell.vt === 1 ? "top" : cell.vt === 2 ? "bottom" : undefined;
  excelCell.alignment = {
    horizontal,
    vertical,
    wrapText: cell.tb === "2",
  };
  if (format && format !== "General" && format !== "@") {
    excelCell.numFmt = format;
  }
}

export async function exportSpreadsheet(document: WorkbookDocument): Promise<void> {
  const ExcelJS = await loadExcelJS();
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Formstr Sheets";
  workbook.created = new Date(document.createdAt);
  workbook.modified = new Date();

  const sheets = workbookToFortuneSheets(document);
  for (const sheet of sheets) {
    const worksheet = workbook.addWorksheet(sheet.name.slice(0, 31) || "Sheet");
    const celldata = sheet.celldata ?? [];
    for (const item of celldata) {
      if (!item.v) continue;
      applyCellToWorksheet(worksheet, item.v as Cell, item.r + 1, item.c + 1);
    }

    const widths = sheet.config?.columnlen ?? {};
    for (const [column, width] of Object.entries(widths)) {
      worksheet.getColumn(Number(column) + 1).width = Number(width) / 7.5;
    }
    const heights = sheet.config?.rowlen ?? {};
    for (const [row, height] of Object.entries(heights)) {
      worksheet.getRow(Number(row) + 1).height = Number(height) * 0.75;
    }

    const merges = Object.values(sheet.config?.merge ?? {});
    for (const merge of merges) {
      try {
        worksheet.mergeCells(
          merge.r + 1,
          merge.c + 1,
          merge.r + merge.rs,
          merge.c + merge.cs,
        );
      } catch {
        // ignore overlapping merges from imported files
      }
    }

    const validations = sheet.dataVerification ?? {};
    for (const [key, raw] of Object.entries(validations)) {
      const validation = raw as { type?: string; value1?: string };
      if (validation.type !== "dropdown" || !validation.value1) continue;
      const [rowText, colText] = key.split("_");
      const row = Number(rowText);
      const col = Number(colText);
      if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
      const options = validation.value1.replace(/"/g, '""');
      try {
        worksheet.getCell(row + 1, col + 1).dataValidation = {
          type: "list",
          allowBlank: true,
          formulae: [`"${options}"`],
        };
      } catch {
        // ignore invalid validation ranges
      }
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  downloadBlob(`${fileSafeName(document.title)}.xlsx`, blob);
}

export function exportActiveSheetCsv(document: WorkbookDocument, sheetId?: string): void {
  const sheets = workbookToFortuneSheets(document);
  const targetId = sheetId ?? document.activeSheetId;
  const sheet =
    sheets.find((item) => item.id === targetId) ??
    sheets.find((item) => item.status === 1) ??
    sheets[0];
  if (!sheet) return;
  const celldata = sheet.celldata ?? [];
  let maxRow = 0;
  let maxCol = 0;
  for (const item of celldata) {
    maxRow = Math.max(maxRow, item.r);
    maxCol = Math.max(maxCol, item.c);
  }
  // Index the cells once: a lookup per grid position would be quadratic on a big sheet.
  const byPosition = new Map<number, string>();
  for (const item of celldata) {
    const cell = item.v;
    if (cell) byPosition.set(item.r * (maxCol + 1) + item.c, String(cell.m ?? cell.v ?? ""));
  }
  const rows: string[][] = [];
  for (let r = 0; r <= maxRow; r++) {
    const row: string[] = [];
    for (let c = 0; c <= maxCol; c++) row.push(byPosition.get(r * (maxCol + 1) + c) ?? "");
    rows.push(row);
  }
  const csv = rows
    .map((row) =>
      row
        .map((value) => {
          const escaped = value.replace(/"/g, '""');
          return /[",\n\r]/.test(value) ? `"${escaped}"` : escaped;
        })
        .join(","),
    )
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  downloadBlob(`${fileSafeName(document.title)}-${fileSafeName(sheet.name)}.csv`, blob);
}

export function downloadBlob(filename: string, blob: Blob): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
