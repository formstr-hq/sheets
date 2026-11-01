import type { Cell, CellWithRowAndCol, Sheet } from "@fortune-sheet/core";
import { DEFAULT_COLS, DEFAULT_ROWS } from "./constants";
import { randomId } from "./utils";

export interface FormstrCellMeta {
  fs: "user" | "tag";
  pubkey?: string;
  npub?: string;
  name?: string;
  picture?: string;
}

export interface FormstrCell extends Cell {
  ct?: {
    fa?: string;
    t?: string;
    s?: FormstrCellMeta;
  };
}

export function sheetName(index: number): string {
  return `Sheet ${index + 1}`;
}

export function createEmptySheet(name = "Sheet 1", rows = DEFAULT_ROWS, cols = DEFAULT_COLS): Sheet {
  return {
    id: randomId(10),
    name,
    row: rows,
    column: cols,
    status: 1,
    order: 0,
    celldata: [],
    config: {
      columnlen: {},
      rowlen: {},
      merge: {},
      borderInfo: [],
    },
  };
}

export function matrixToCelldata(data: (Cell | null)[][] | undefined): CellWithRowAndCol[] {
  if (!data) return [];
  const out: CellWithRowAndCol[] = [];
  data.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (cell !== null && cell !== undefined) out.push({ r, c, v: cell });
    });
  });
  return out;
}

export function celldataToMatrix(
  celldata: CellWithRowAndCol[] | undefined,
  rows: number,
  cols: number,
): (Cell | null)[][] {
  const maxRow = celldata?.reduce((max, item) => Math.max(max, item.r + 1), 0) ?? 0;
  const maxCol = celldata?.reduce((max, item) => Math.max(max, item.c + 1), 0) ?? 0;
  const rowCount = Math.max(rows || 0, maxRow, 1);
  const colCount = Math.max(cols || 0, maxCol, 1);
  const matrix = Array.from({ length: rowCount }, () =>
    Array<Cell | null>(colCount).fill(null),
  );
  celldata?.forEach((item) => {
    if (item.r >= 0 && item.c >= 0 && item.r < rowCount && item.c < colCount) {
      matrix[item.r][item.c] = item.v;
    }
  });
  return matrix;
}

export function toStoredSheet(sheet: Sheet): Sheet {
  const stored: Sheet = { ...sheet };
  const celldata =
    sheet.data && sheet.data.length > 0
      ? matrixToCelldata(sheet.data)
      : sheet.celldata ?? [];
  delete stored.data;
  stored.celldata = celldata;
  const maxRow = celldata.reduce((max, item) => Math.max(max, item.r + 1), 0);
  const maxCol = celldata.reduce((max, item) => Math.max(max, item.c + 1), 0);
  stored.row = Math.max(sheet.row ?? 0, maxRow, DEFAULT_ROWS);
  stored.column = Math.max(sheet.column ?? 0, maxCol, DEFAULT_COLS);
  return stored;
}

export function storedSheets(sheets: Sheet[]): Sheet[] {
  return sheets.map(toStoredSheet);
}

export function sheetsForUpdate(sheets: Sheet[]): Sheet[] {
  return sheets.map((sheet) => {
    const stored = { ...sheet };
    if (!stored.data) {
      stored.data = celldataToMatrix(
        stored.celldata,
        stored.row ?? DEFAULT_ROWS,
        stored.column ?? DEFAULT_COLS,
      );
    }
    delete stored.celldata;
    return stored;
  });
}

export function activeSheetId(sheets: Sheet[]): string | undefined {
  return sheets.find((sheet) => sheet.status === 1 && sheet.hide !== 1)?.id ?? sheets[0]?.id;
}

export function withActiveSheet(sheets: Sheet[], id: string): Sheet[] {
  return sheets.map((sheet) => ({ ...sheet, status: sheet.id === id ? 1 : 0 }));
}

export function cellMeta(cell: Cell | null | undefined): FormstrCellMeta | null {
  const meta = (cell as FormstrCell | null | undefined)?.ct?.s;
  return meta?.fs ? meta : null;
}

export function cellDisplayText(cell: Cell | null | undefined): string {
  if (!cell) return "";
  if (typeof cell.m === "string" || typeof cell.m === "number") return String(cell.m);
  if (typeof cell.v === "string" || typeof cell.v === "number") return String(cell.v);
  return "";
}
