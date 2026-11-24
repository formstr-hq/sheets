// Converts an ExcelJS cell to a FortuneSheet cell. Shared by the in-memory importer (xlsx.ts) and
// the streaming importer (xlsxStream), so both produce identical cells. Pure: type-only imports
// from exceljs/@fortune-sheet, so a worker can use it without pulling in the DOM helpers.
import type * as ExcelJS from "exceljs";
import type { Cell } from "./legacy/fortuneFormat";
import {
  formatNumberForDisplay,
  isDateFormat,
  toDateSerial as toSerial,
} from "./workbook/format";

export function argbToHex(value: unknown): string | undefined {
  if (!value) return undefined;
  if (typeof value === "string") {
    if (/^[0-9a-f]{8}$/i.test(value)) return `#${value.slice(2)}`;
    if (/^#[0-9a-f]{6}$/i.test(value)) return value;
  }
  if (typeof value === "object" && value && "argb" in value) {
    const argb = (value as { argb?: string }).argb;
    if (argb && /^[0-9a-f]{8}$/i.test(argb)) return `#${argb.slice(2)}`;
  }
  return undefined;
}

export function cellStyle(cell: ExcelJS.Cell, format: string | undefined): Partial<Cell> {
  const style: Partial<Cell> = {};
  const font = cell.font;
  if (font?.bold) style.bl = 1;
  if (font?.italic) style.it = 1;
  if (font?.underline) style.un = 1;
  if (font?.strike) style.cl = 1;
  if (font?.size) style.fs = font.size;
  if (font?.name) style.ff = font.name;
  const fontColor = argbToHex(font?.color);
  if (fontColor) style.fc = fontColor;
  const fillColor = argbToHex(
    cell.fill && cell.fill.type === "pattern" ? cell.fill.fgColor : undefined,
  );
  if (fillColor) style.bg = fillColor;
  const horizontal = cell.alignment?.horizontal;
  if (horizontal === "left") style.ht = 1;
  else if (horizontal === "center") style.ht = 0;
  else if (horizontal === "right") style.ht = 2;
  const vertical = cell.alignment?.vertical;
  if (vertical === "top") style.vt = 1;
  else if (vertical === "middle") style.vt = 0;
  else if (vertical === "bottom") style.vt = 2;
  if (cell.alignment?.wrapText) style.tb = "2";
  if (format && format !== "General" && format !== "@") {
    style.ct = {
      fa: format,
      t: isDateFormat(format, cell.value) ? "d" : "n",
    };
  }
  return style;
}

export function valueFromExcelCell(cell: ExcelJS.Cell): Cell | null {
  if (cell.master && cell.master.address && cell.master.address !== cell.address) return null;
  const raw = cell.value;
  const format = typeof cell.numFmt === "string" ? cell.numFmt : undefined;

  if (raw === null || raw === undefined) return null;

  if (typeof raw === "object" && raw && "formula" in raw) {
    const formulaValue = raw as {
      formula?: string;
      sharedFormula?: string;
      result?: unknown;
      text?: string;
    };
    const formula = formulaValue.formula ?? formulaValue.sharedFormula ?? "";
    const result = formulaValue.result;
    const isDate = result instanceof Date || isDateFormat(format, result);
    const errorText =
      result && typeof result === "object" && "error" in result
        ? String((result as { error: unknown }).error)
        : null;
    const numericResult =
      typeof result === "number"
        ? result
        : result instanceof Date
          ? toSerial(result)
          : Number(result);
    const hasResult = result !== undefined && result !== null;
    const value =
      errorText ??
      (hasResult
        ? Number.isFinite(numericResult)
          ? numericResult
          : String(result)
        : 0);
    const display = errorText
      ? errorText
      : typeof value === "number"
        ? formatNumberForDisplay(value, format)
        : result instanceof Date
          ? result.toISOString().slice(0, 10)
          : String(result ?? "");
    return {
      ...cellStyle(cell, format),
      f: formula.startsWith("=") ? formula : `=${formula}`,
      v: value,
      m: display,
      ct: {
        fa: format || (isDate ? "yyyy-mm-dd" : "General"),
        t: isDate ? "d" : typeof value === "string" ? "s" : "n",
      },
    };
  }

  if (raw instanceof Date) {
    const serial = toSerial(raw);
    return {
      ...cellStyle(cell, format),
      v: serial,
      m: cell.text || raw.toISOString().slice(0, 10),
      ct: { fa: format || "yyyy-mm-dd", t: "d" },
    };
  }

  if (typeof raw === "number") {
    return {
      ...cellStyle(cell, format),
      v: raw,
      m: formatNumberForDisplay(raw, format),
      ct: { fa: format || "General", t: isDateFormat(format, raw) ? "d" : "n" },
    };
  }

  if (typeof raw === "boolean") {
    return {
      ...cellStyle(cell, format),
      v: raw,
      m: raw ? "TRUE" : "FALSE",
      ct: { fa: "General", t: "b" },
    };
  }

  if (typeof raw === "object" && raw && "richText" in raw) {
    const text = (raw as { richText: Array<{ text: string }> }).richText
      .map((part) => part.text)
      .join("");
    return {
      ...cellStyle(cell, format),
      v: text,
      m: text,
      ct: { fa: "General", t: "s" },
    };
  }

  if (typeof raw === "object" && raw && "error" in raw) {
    // An error value such as #N/A (ExcelJS gives { error }); shown as its text, like formula errors.
    const text = String((raw as { error: unknown }).error);
    return {
      ...cellStyle(cell, format),
      v: text,
      m: text,
      ct: { fa: "General", t: "s" },
    };
  }

  if (typeof raw === "object" && raw && "text" in raw) {
    const text = String((raw as { text: unknown }).text ?? "");
    return {
      ...cellStyle(cell, format),
      v: text,
      m: text,
      ct: { fa: "General", t: "s" },
    };
  }

  const text = String(raw);
  return {
    ...cellStyle(cell, format),
    v: text,
    m: text,
    ct: { fa: format && format !== "General" ? format : "General", t: "s" },
  };
}
