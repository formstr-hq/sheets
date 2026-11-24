import type { CellPrimitive, WorkbookCell } from "./types";

export const DEFAULT_FONT_FAMILIES = [
  "Inter",
  "Arial",
  "Tahoma",
  "Verdana",
  "Times New Roman",
] as const;

const DATE_SERIAL_OFFSET = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;

export function toDateSerial(date: Date): number {
  return (
    (Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()) -
      DATE_SERIAL_OFFSET) /
    DAY_MS
  );
}

export function fromDateSerial(serial: number): Date {
  return new Date(DATE_SERIAL_OFFSET + serial * DAY_MS);
}

export function isDateFormat(format: string | undefined, value?: unknown): boolean {
  if (value instanceof Date) return true;
  if (!format) return false;
  return /(yy|dd|d{1,2}\/|h:mm|am\/pm|\[h\])/i.test(format);
}

// Number.prototype.toLocaleString builds a new Intl.NumberFormat on every call, which dominates
// import time for big sheets; formatters are cached by their only two options instead (same output).
const formatters = new Map<number, Intl.NumberFormat>();
function numberFormatter(decimals: number, useGrouping: boolean): Intl.NumberFormat {
  const key = decimals * 2 + (useGrouping ? 1 : 0);
  let formatter = formatters.get(key);
  if (!formatter) {
    formatter = new Intl.NumberFormat("en-US", {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
      useGrouping,
    });
    formatters.set(key, formatter);
  }
  return formatter;
}

export function formatNumberForDisplay(
  value: number,
  format: string | undefined,
): string {
  if (!format || format === "General" || format === "@") return String(value);
  if (isDateFormat(format, value)) {
    return fromDateSerial(value).toISOString().slice(0, 10);
  }
  const percent = format.includes("%");
  const currency = format.includes("$");
  const decimalsMatch = format.match(/\.(0+)/);
  const decimals = decimalsMatch ? decimalsMatch[1].length : 0;
  const useGrouping = /#,##0|,/.test(format);
  const display = percent ? value * 100 : value;
  let text = numberFormatter(decimals, useGrouping).format(display);
  if (currency) text = `$${text}`;
  if (percent) text = `${text}%`;
  return text;
}

export function displayValue(
  value: CellPrimitive | undefined,
  format: string | undefined,
  display?: string,
): string {
  if (display != null && display !== "") return display;
  if (value == null) return "";
  if (typeof value === "number") return formatNumberForDisplay(value, format);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

export function cellDisplay(cell: WorkbookCell): string {
  return displayValue(cell.value, cell.style?.numberFormat, cell.display);
}
