// Turns one ExcelJS row (from the stream reader) into app cells, using the same converter as the
// in-memory importer so both produce identical cells.
import { cellFromFortune, type WorkbookCell } from "../workbook";
import { valueFromExcelCell } from "../xlsxCells";
import type { ExcelRow } from "./excelInternals";
import type { PackedStrings } from "./strings";

export interface MappedRow {
  /** 0-based row index. */
  row: number;
  /** [0-based column, cell], left to right; empty cells are left out. */
  cells: [number, WorkbookCell][];
  /** Highest column number (1-based) of any <c> element, empty or not, like ExcelJS's columnCount. */
  cellCount: number;
  /** True when the row holds at least one non-empty value (ExcelJS's hasValues). */
  hasValues: boolean;
  /** Row height in points from <row ht>, if present. */
  height?: number;
  /** Cells that were shared-formula dependents: their formula was dropped, the value kept. */
  sharedFormulaValues: number;
}

interface SharedStringValue {
  sharedString: number;
}

interface FormulaValue {
  formula?: string;
  sharedFormula?: string;
  result?: unknown;
}

/**
 * The stream reader leaves shared strings unresolved (`{ sharedString: i }`) and, for shared
 * formula dependents, a formula with empty text. Both are fixed here before conversion.
 * Returns "skip" when the cell has nothing to keep, "shared" for a shared-formula dependent.
 */
function prepare(cell: { value: unknown }, strings: PackedStrings | null): "ok" | "skip" | "shared" {
  const raw = cell.value;
  if (raw === null || raw === undefined || typeof raw !== "object" || raw instanceof Date) return "ok";
  if ("sharedString" in raw) {
    cell.value = strings ? strings.get((raw as SharedStringValue).sharedString) : "";
    return "ok";
  }
  if ("formula" in raw) {
    const value = raw as FormulaValue;
    if (!value.formula && !value.sharedFormula) {
      // Shared-formula dependent: ExcelJS's stream reader drops the formula, keep the cached value.
      if (value.result === undefined || value.result === null) return "skip";
      cell.value = value.result as never;
      return "shared";
    }
  }
  return "ok";
}

export function mapRow(row: ExcelRow, strings: PackedStrings | null): MappedRow {
  const cells: [number, WorkbookCell][] = [];
  let sharedFormulaValues = 0;
  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const state = prepare(cell, strings);
    if (state === "skip") return;
    if (state === "shared") sharedFormulaValues++;
    const fortune = valueFromExcelCell(cell);
    if (!fortune) return;
    const converted = cellFromFortune(fortune, undefined);
    if (converted) cells.push([colNumber - 1, converted]);
  });
  const mapped: MappedRow = { row: row.number - 1, cells, cellCount: row.cellCount, hasValues: row.hasValues, sharedFormulaValues };
  if (row.height) mapped.height = row.height;
  return mapped;
}
