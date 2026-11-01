import type { Cell, CellWithRowAndCol, Sheet } from "@fortune-sheet/core";
import { DEFAULT_COLS, DEFAULT_ROWS } from "../constants";
import { randomId } from "../utils";
import { celldataToMatrix } from "../sheetData";
import {
  DEFAULT_FONT_FAMILIES,
  displayValue,
  fromDateSerial,
} from "./format";
import {
  WORKBOOK_SCHEMA,
  WORKBOOK_VERSION,
  type CellKind,
  type CellPrimitive,
  type WorkbookBorder,
  type WorkbookBorderType,
  type WorkbookCell,
  type WorkbookCellStyle,
  type WorkbookConditionalRule,
  type WorkbookDocument,
  type WorkbookRange,
  type WorkbookSheet,
  type WorkbookValidation,
} from "./types";

type FortuneCellMeta = {
  fs?: "user" | "tag";
  pubkey?: string;
  npub?: string;
  name?: string;
  picture?: string;
};

type FortuneCell = Cell & {
  ct?: {
    fa?: string;
    t?: string;
    s?: FortuneCellMeta;
  };
};

type DataVerification = Record<string, unknown>;

function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

function parseCellKey(key: string): { row: number; col: number } | null {
  const [row, col] = key.split(":").map(Number);
  if (!Number.isFinite(row) || !Number.isFinite(col)) return null;
  return { row, col };
}

function ensureFormula(formula: string): string {
  return formula.startsWith("=") ? formula : `=${formula}`;
}

function fontFamilyFromFortune(ff: number | string | undefined): string | undefined {
  if (typeof ff === "string") return ff || undefined;
  if (typeof ff === "number") return DEFAULT_FONT_FAMILIES[ff] ?? undefined;
  return undefined;
}

function fontFamilyToFortune(ff: string | undefined): number | string | undefined {
  if (!ff) return undefined;
  const index = DEFAULT_FONT_FAMILIES.findIndex(
    (family) => family.toLowerCase() === ff.toLowerCase(),
  );
  return index >= 0 ? index : ff;
}

function typeFromFortune(t: string | undefined): CellKind | undefined {
  switch (t) {
    case "n":
      return "number";
    case "b":
      return "boolean";
    case "d":
      return "date";
    case "e":
      return "error";
    case "s":
    case "g":
      return "text";
    default:
      return undefined;
  }
}

function typeToFortune(t: CellKind | undefined): string {
  switch (t) {
    case "number":
      return "n";
    case "boolean":
      return "b";
    case "date":
      return "d";
    case "error":
      return "e";
    default:
      return "s";
  }
}

function styleFromFortune(cell: FortuneCell): WorkbookCellStyle | undefined {
  const style: WorkbookCellStyle = {};
  if (cell.bl) style.bold = true;
  if (cell.it) style.italic = true;
  if (cell.un) style.underline = true;
  if (cell.cl) style.strike = true;
  if (cell.fs) style.fontSize = cell.fs;
  const fontFamily = fontFamilyFromFortune(cell.ff);
  if (fontFamily) style.fontFamily = fontFamily;
  if (cell.fc) style.color = cell.fc;
  if (cell.bg) style.background = cell.bg;
  if (cell.ht === 0) style.align = "center";
  else if (cell.ht === 1) style.align = "left";
  else if (cell.ht === 2) style.align = "right";
  if (cell.vt === 0) style.verticalAlign = "middle";
  else if (cell.vt === 1) style.verticalAlign = "top";
  else if (cell.vt === 2) style.verticalAlign = "bottom";
  if (cell.tb === "2") style.wrap = true;
  if (cell.tr != null && cell.tr !== "") {
    const rotation = Number(cell.tr);
    if (Number.isFinite(rotation)) style.rotation = rotation;
  }
  if (cell.ct?.fa && cell.ct.fa !== "General") {
    style.numberFormat = cell.ct.fa;
  }
  return Object.keys(style).length > 0 ? style : undefined;
}

function validationFromFortune(raw: unknown): WorkbookValidation | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const item = raw as Record<string, unknown>;
  if (typeof item.type !== "string") return undefined;
  const validation: WorkbookValidation = { type: item.type };
  if (item.type === "dropdown" && typeof item.value1 === "string") {
    validation.options = item.value1
      .split(",")
      .map((option) => option.trim())
      .filter(Boolean);
  }
  const params: Record<string, string | boolean> = {};
  for (const key of [
    "type2",
    "value1",
    "value2",
    "validity",
    "remote",
    "prohibitInput",
    "hintShow",
    "hintValue",
    "checked",
  ]) {
    const value = item[key];
    if (typeof value === "string" || typeof value === "boolean") params[key] = value;
  }
  if (Object.keys(params).length > 0) validation.params = params;
  return validation;
}

function validationToFortune(
  validation: WorkbookValidation,
  rangeText: string,
): Record<string, unknown> {
  const params = validation.params ?? {};
  return {
    type: validation.type,
    type2: typeof params.type2 === "string" ? params.type2 : "",
    rangeTxt: rangeText,
    value1:
      validation.options?.join(",") ??
      (typeof params.value1 === "string" ? params.value1 : ""),
    value2: typeof params.value2 === "string" ? params.value2 : "",
    validity: typeof params.validity === "string" ? params.validity : "",
    remote: params.remote === true,
    prohibitInput: params.prohibitInput === true,
    hintShow: params.hintShow === true,
    hintValue: typeof params.hintValue === "string" ? params.hintValue : "",
    checked: params.checked === true,
  };
}

function cellFromFortune(
  cell: Cell | null,
  validation: unknown,
  link?: { url: string; label?: string },
): WorkbookCell | null {
  if (!cell) return null;
  const fortune = cell as FortuneCell;
  const meta = fortune.ct?.s;
  const canonical: WorkbookCell = {};
  if (fortune.v !== undefined) canonical.value = fortune.v as CellPrimitive;
  if (fortune.f) canonical.formula = ensureFormula(fortune.f);
  if (fortune.m !== undefined) canonical.display = String(fortune.m);
  const style = styleFromFortune(fortune);
  if (style) canonical.style = style;
  if (meta?.fs === "user") {
    canonical.person = {
      pubkey: meta.pubkey ?? "",
      npub: meta.npub,
      name: meta.name,
      picture: meta.picture,
    };
    canonical.type = "text";
    if (canonical.value === undefined) canonical.value = canonical.display ?? "";
  } else if (meta?.fs === "tag") {
    canonical.tag = String(fortune.m ?? fortune.v ?? "");
    canonical.type = "text";
    if (canonical.value === undefined) canonical.value = canonical.tag;
  } else {
    const type = typeFromFortune(fortune.ct?.t);
    if (type) canonical.type = type;
    else if (typeof canonical.value === "boolean") canonical.type = "boolean";
  }
  if (canonical.type === "date" && typeof canonical.value === "number") {
    canonical.display = fromDateSerial(canonical.value).toISOString().slice(0, 10);
  }
  const parsedValidation = validationFromFortune(validation);
  if (parsedValidation) canonical.validation = parsedValidation;
  if (link) canonical.link = link;
  const meaningful =
    canonical.value !== undefined ||
    canonical.formula !== undefined ||
    canonical.style !== undefined ||
    canonical.person !== undefined ||
    canonical.tag !== undefined ||
    canonical.validation !== undefined ||
    canonical.link !== undefined;
  return meaningful ? canonical : null;
}

function cellFromFortuneToFortune(cell: WorkbookCell): FortuneCell {
  const fortune: FortuneCell = {};
  const style = cell.style;
  const numberFormat = style?.numberFormat;
  if (cell.formula) fortune.f = ensureFormula(cell.formula);
  const value = cell.value;
  if (value !== undefined) fortune.v = value === null ? "" : value;
  else if (cell.formula) fortune.v = 0;
  fortune.m = displayValue(value, numberFormat, cell.display);
  if (cell.person) {
    fortune.ct = {
      fa: numberFormat ?? "General",
      t: "s",
      s: {
        fs: "user",
        pubkey: cell.person.pubkey,
        npub: cell.person.npub,
        name: cell.person.name,
        picture: cell.person.picture,
      },
    };
  } else if (cell.tag) {
    fortune.ct = {
      fa: "General",
      t: "s",
      s: { fs: "tag" },
    };
    if (fortune.v === undefined) fortune.v = cell.tag;
    fortune.m = cell.tag;
  } else {
    fortune.ct = {
      fa: numberFormat ?? "General",
      t: typeToFortune(cell.type),
    };
  }
  if (style) {
    if (style.bold) fortune.bl = 1;
    if (style.italic) fortune.it = 1;
    if (style.underline) fortune.un = 1;
    if (style.strike) fortune.cl = 1;
    if (style.fontSize) fortune.fs = style.fontSize;
    const ff = fontFamilyToFortune(style.fontFamily);
    if (ff !== undefined) fortune.ff = ff;
    if (style.color) fortune.fc = style.color;
    if (style.background) fortune.bg = style.background;
    if (style.align === "center") fortune.ht = 0;
    else if (style.align === "left") fortune.ht = 1;
    else if (style.align === "right") fortune.ht = 2;
    if (style.verticalAlign === "middle") fortune.vt = 0;
    else if (style.verticalAlign === "top") fortune.vt = 1;
    else if (style.verticalAlign === "bottom") fortune.vt = 2;
    if (style.wrap) fortune.tb = "2";
    if (style.rotation != null) fortune.tr = String(style.rotation);
  }
  return fortune;
}

function rangesFromFortune(cellrange: unknown): WorkbookRange[] {
  if (!Array.isArray(cellrange)) return [];
  const ranges: WorkbookRange[] = [];
  for (const range of cellrange) {
    if (!range || typeof range !== "object") continue;
    const selection = range as { row?: number[]; column?: number[] };
    if (!Array.isArray(selection.row) || !Array.isArray(selection.column)) continue;
    const [r1, r2] = selection.row;
    const [c1, c2] = selection.column;
    if ([r1, r2, c1, c2].every((value) => Number.isFinite(value))) {
      ranges.push({ r1, c1, r2, c2 });
    }
  }
  return ranges;
}

function conditionalFromFortune(raw: unknown): WorkbookConditionalRule[] {
  if (!Array.isArray(raw)) return [];
  const rules: WorkbookConditionalRule[] = [];
  for (const rawRule of raw) {
    if (!rawRule || typeof rawRule !== "object") continue;
    const rule = rawRule as {
      cellrange?: unknown;
      format?: { textColor?: string | null; cellColor?: string | null };
      conditionName?: string;
      conditionValue?: unknown[];
    };
    const ranges = rangesFromFortune(rule.cellrange);
    if (ranges.length === 0 || !rule.conditionName) continue;
    rules.push({
      ranges,
      condition: {
        type: rule.conditionName,
        values: (rule.conditionValue ?? []).map((value) => String(value)),
      },
      format: {
        textColor: rule.format?.textColor ?? undefined,
        cellColor: rule.format?.cellColor ?? undefined,
      },
    });
  }
  return rules;
}

function conditionalToFortune(
  rules: WorkbookConditionalRule[] | undefined,
): unknown[] {
  if (!rules || rules.length === 0) return [];
  return rules.map((rule) => ({
    type: "default",
    cellrange: rule.ranges.map((range) => ({
      row: [range.r1, range.r2],
      column: [range.c1, range.c2],
    })),
    format: {
      textColor: rule.format.textColor ?? null,
      cellColor: rule.format.cellColor ?? null,
    },
    conditionName: rule.condition.type,
    conditionRange: [],
    conditionValue: rule.condition.values,
  }));
}

function borderTypeFromFortune(value: string | undefined): WorkbookBorderType {
  switch (value) {
    case "border-outside":
      return "outside";
    case "border-inside":
      return "inside";
    case "border-top":
      return "top";
    case "border-bottom":
      return "bottom";
    case "border-left":
      return "left";
    case "border-right":
      return "right";
    case "border-none":
      return "none";
    default:
      return "all";
  }
}

function borderTypeToFortune(value: WorkbookBorderType): string {
  switch (value) {
    case "outside":
      return "border-outside";
    case "inside":
      return "border-inside";
    case "top":
      return "border-top";
    case "bottom":
      return "border-bottom";
    case "left":
      return "border-left";
    case "right":
      return "border-right";
    case "none":
      return "border-none";
    default:
      return "border-all";
  }
}

function bordersFromFortune(raw: unknown): WorkbookBorder[] {
  if (!Array.isArray(raw)) return [];
  const result: WorkbookBorder[] = [];
  for (const item of raw) {
    if (!item || typeof item !== "object") continue;
    const border = item as {
      rangeType?: string;
      borderType?: string;
      color?: string;
      style?: string;
      range?: unknown;
      value?: { row_index?: number; col_index?: number };
    };
    if (border.rangeType === "cell" && border.value) {
      const row = border.value.row_index;
      const col = border.value.col_index;
      if (Number.isFinite(row) && Number.isFinite(col) && row != null && col != null) {
        result.push({
          range: { r1: row, c1: col, r2: row, c2: col },
          type: borderTypeFromFortune(border.borderType),
          color: border.color,
          style: border.style,
        });
      }
    } else {
      const ranges = rangesFromFortune(border.range);
      for (const range of ranges) {
        result.push({
          range,
          type: borderTypeFromFortune(border.borderType),
          color: border.color,
          style: border.style,
        });
      }
    }
  }
  return result;
}

function bordersToFortune(borders: WorkbookBorder[] | undefined): unknown[] {
  if (!borders || borders.length === 0) return [];
  return borders.map((border) => ({
    rangeType: "range",
    borderType: borderTypeToFortune(border.type),
    color: border.color ?? "#000000",
    style: border.style ?? "1",
    range: [
      {
        row: [border.range.r1, border.range.r2],
        column: [border.range.c1, border.range.c2],
      },
    ],
  }));
}

function linksFromFortune(sheet: Sheet): Record<string, { url: string; label?: string }> {
  const links: Record<string, { url: string; label?: string }> = {};
  const hyperlinks = (sheet as Sheet & {
    hyperlink?: Record<string, { linkType?: string; linkAddress?: string }>;
  }).hyperlink;
  if (!hyperlinks) return links;
  for (const [key, value] of Object.entries(hyperlinks)) {
    if (!value?.linkAddress) continue;
    if (value.linkType === "cellrange" || value.linkType === "sheet") continue;
    links[key] = { url: value.linkAddress };
  }
  return links;
}

function frozenFromFortune(sheet: Sheet): WorkbookSheet["frozen"] {
  const frozen = sheet.frozen;
  if (!frozen) return undefined;
  const rowFocus = frozen.range?.row_focus ?? 0;
  const columnFocus = frozen.range?.column_focus ?? 0;
  switch (frozen.type) {
    case "rangeRow":
    case "row":
      return { rows: Math.max(1, rowFocus), columns: 0 };
    case "rangeColumn":
    case "column":
      return { rows: 0, columns: Math.max(1, columnFocus) };
    case "rangeBoth":
    case "both":
      return { rows: Math.max(1, rowFocus), columns: Math.max(1, columnFocus) };
    default:
      return undefined;
  }
}

function frozenToFortune(
  frozen: WorkbookSheet["frozen"],
): Sheet["frozen"] | undefined {
  if (!frozen) return undefined;
  if (frozen.rows > 0 && frozen.columns > 0) {
    return {
      type: "rangeBoth",
      range: { row_focus: frozen.rows, column_focus: frozen.columns },
    };
  }
  if (frozen.rows > 0) {
    return { type: "rangeRow", range: { row_focus: frozen.rows, column_focus: 0 } };
  }
  if (frozen.columns > 0) {
    return { type: "rangeColumn", range: { row_focus: 0, column_focus: frozen.columns } };
  }
  return undefined;
}

export interface WorkbookFromFortuneOptions {
  title: string;
  sheets: Sheet[];
  createdAt?: number;
  updatedAt?: number;
  activeSheetId?: string;
}

export function fortuneSheetsToWorkbook(
  options: WorkbookFromFortuneOptions,
): WorkbookDocument {
  const now = Date.now();
  const createdAt = options.createdAt ?? now;
  const updatedAt = options.updatedAt ?? now;
  const sheets: WorkbookSheet[] = options.sheets.map((sheet, index) => {
    const matrix = sheet.data
      ? sheet.data
      : celldataToMatrix(sheet.celldata, sheet.row ?? DEFAULT_ROWS, sheet.column ?? DEFAULT_COLS);
    const rowCount = Math.max(sheet.row ?? 0, matrix.length, 1);
    const colCount = Math.max(
      sheet.column ?? 0,
      matrix.reduce((max, row) => Math.max(max, row.length), 0),
      1,
    );
    const verification = (sheet.dataVerification ?? {}) as DataVerification;
    const links = linksFromFortune(sheet);
    const cells: Record<string, WorkbookCell> = {};
    for (let row = 0; row < matrix.length; row++) {
      const matrixRow = matrix[row];
      for (let col = 0; col < matrixRow.length; col++) {
        const key = cellKey(row, col);
        const cell = cellFromFortune(
          matrixRow[col],
          verification[key],
          links[key],
        );
        if (cell) cells[key] = cell;
      }
    }
    for (const [key, link] of Object.entries(links)) {
      if (cells[key]) continue;
      const parsed = parseCellKey(key);
      if (!parsed) continue;
      cells[key] = { link };
    }
    const config = sheet.config ?? {};
    const columnWidths: Record<number, number> = {};
    for (const [key, value] of Object.entries(config.columnlen ?? {})) {
      if (Number.isFinite(Number(key)) && typeof value === "number") {
        columnWidths[Number(key)] = value;
      }
    }
    const rowHeights: Record<number, number> = {};
    for (const [key, value] of Object.entries(config.rowlen ?? {})) {
      if (Number.isFinite(Number(key)) && typeof value === "number") {
        rowHeights[Number(key)] = value;
      }
    }
    const hiddenRows = Object.keys(config.rowhidden ?? {})
      .map(Number)
      .filter(Number.isFinite);
    const hiddenColumns = Object.keys(config.colhidden ?? {})
      .map(Number)
      .filter(Number.isFinite);
    const merges = Object.values(config.merge ?? {}).map((merge) => ({
      r: merge.r,
      c: merge.c,
      rs: merge.rs,
      cs: merge.cs,
    }));
    const sheetData: WorkbookSheet = {
      id: sheet.id || randomId(10),
      name: sheet.name || `Sheet ${index + 1}`,
      order: sheet.order ?? index,
      rowCount,
      colCount,
      cells,
      hidden: sheet.hide === 1 ? true : undefined,
      color: sheet.color,
      showGridLines:
        typeof sheet.showGridLines === "boolean"
          ? sheet.showGridLines
          : sheet.showGridLines === 1
            ? true
            : undefined,
      defaultRowHeight: sheet.defaultRowHeight,
      defaultColWidth: sheet.defaultColWidth,
      columnWidths: Object.keys(columnWidths).length > 0 ? columnWidths : undefined,
      rowHeights: Object.keys(rowHeights).length > 0 ? rowHeights : undefined,
      hiddenRows: hiddenRows.length > 0 ? hiddenRows : undefined,
      hiddenColumns: hiddenColumns.length > 0 ? hiddenColumns : undefined,
      merges: merges.length > 0 ? merges : undefined,
      frozen: frozenFromFortune(sheet),
      conditionalFormats: conditionalFromFortune(sheet.luckysheet_conditionformat_save),
      borders: bordersFromFortune(config.borderInfo),
    };
    return sheetData;
  });
  const activeSheetId =
    options.activeSheetId ??
    options.sheets.find((sheet) => sheet.status === 1)?.id ??
    sheets[0]?.id ??
    "";
  return {
    schema: WORKBOOK_SCHEMA,
    version: WORKBOOK_VERSION,
    title: options.title,
    activeSheetId,
    createdAt,
    updatedAt,
    sheets,
  };
}

export function workbookToFortuneSheets(document: WorkbookDocument): Sheet[] {
  const sorted = [...document.sheets].sort((a, b) => a.order - b.order);
  return sorted.map((sheet) => {
    const celldata: CellWithRowAndCol[] = [];
    const dataVerification: Record<string, unknown> = {};
    const hyperlink: Record<string, { linkType: string; linkAddress: string }> = {};
    for (const [key, canonical] of Object.entries(sheet.cells)) {
      const parsed = parseCellKey(key);
      if (!parsed) continue;
      const link = canonical.link;
      const fortune = cellFromFortuneToFortune(canonical);
      celldata.push({ r: parsed.row, c: parsed.col, v: fortune });
      if (canonical.validation) {
        const rangeText = `${columnLetter(parsed.col)}${parsed.row + 1}`;
        dataVerification[key] = validationToFortune(canonical.validation, rangeText);
      }
      if (link) {
        hyperlink[key] = { linkType: "webpage", linkAddress: link.url };
      }
    }
    const config: Sheet["config"] = {
      columnlen: sheet.columnWidths,
      rowlen: sheet.rowHeights,
      rowhidden: sheet.hiddenRows
        ? Object.fromEntries(sheet.hiddenRows.map((row) => [row, 0]))
        : undefined,
      colhidden: sheet.hiddenColumns
        ? Object.fromEntries(sheet.hiddenColumns.map((col) => [col, 0]))
        : undefined,
      merge: sheet.merges
        ? Object.fromEntries(
            sheet.merges.map((merge) => [`${merge.r}_${merge.c}`, merge]),
          )
        : undefined,
      borderInfo: bordersToFortune(sheet.borders),
    };
    return {
      id: sheet.id,
      name: sheet.name,
      order: sheet.order,
      row: Math.max(DEFAULT_ROWS, sheet.rowCount),
      column: Math.max(DEFAULT_COLS, sheet.colCount),
      status: sheet.id === document.activeSheetId ? 1 : 0,
      hide: sheet.hidden ? 1 : undefined,
      color: sheet.color,
      showGridLines: sheet.showGridLines === false ? false : 1,
      defaultRowHeight: sheet.defaultRowHeight,
      defaultColWidth: sheet.defaultColWidth,
      celldata,
      config,
      dataVerification:
        Object.keys(dataVerification).length > 0 ? dataVerification : undefined,
      luckysheet_conditionformat_save: conditionalToFortune(sheet.conditionalFormats),
      frozen: frozenToFortune(sheet.frozen),
      hyperlink: Object.keys(hyperlink).length > 0 ? hyperlink : undefined,
    } satisfies Sheet;
  });
}

export function createEmptyWorkbook(title: string): WorkbookDocument {
  const now = Date.now();
  const id = randomId(10);
  return {
    schema: WORKBOOK_SCHEMA,
    version: WORKBOOK_VERSION,
    title,
    activeSheetId: id,
    createdAt: now,
    updatedAt: now,
    sheets: [
      {
        id,
        name: "Sheet 1",
        order: 0,
        rowCount: DEFAULT_ROWS,
        colCount: DEFAULT_COLS,
        cells: {},
        showGridLines: true,
      },
    ],
  };
}

export function createWorkbookFromFortune(
  title: string,
  sheets: Sheet[],
): WorkbookDocument {
  if (sheets.length === 0) return createEmptyWorkbook(title);
  return fortuneSheetsToWorkbook({ title, sheets });
}

export function collectReferencedPubkeys(document: WorkbookDocument): string[] {
  const pubkeys = new Set<string>();
  for (const sheet of document.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      if (cell.person?.pubkey) pubkeys.add(cell.person.pubkey);
    }
  }
  return Array.from(pubkeys);
}

function columnLetter(index: number): string {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function normalizeCanonicalSheet(sheet: WorkbookSheet, index: number): WorkbookSheet {
  const cells: Record<string, WorkbookCell> = {};
  for (const [key, cell] of Object.entries(sheet.cells ?? {})) {
    if (!parseCellKey(key) || !cell) continue;
    const next = { ...cell };
    if (next.formula) next.formula = ensureFormula(next.formula);
    cells[key] = next;
  }
  return {
    ...sheet,
    id: sheet.id || randomId(10),
    name: sheet.name || `Sheet ${index + 1}`,
    order: Number.isFinite(sheet.order) ? sheet.order : index,
    rowCount: Math.max(sheet.rowCount ?? 0, DEFAULT_ROWS),
    colCount: Math.max(sheet.colCount ?? 0, DEFAULT_COLS),
    cells,
  };
}

export function normalizeWorkbook(
  input: unknown,
  fallbackTitle = "Untitled spreadsheet",
): WorkbookDocument {
  if (!input || typeof input !== "object") {
    return createEmptyWorkbook(fallbackTitle);
  }
  const raw = input as Partial<WorkbookDocument> & {
    sheets?: unknown;
    schema?: string;
  };
  if (Array.isArray(raw.sheets) && raw.sheets.length > 0) {
    const first = raw.sheets[0] as unknown as Record<string, unknown> | null;
    if (first && typeof first === "object" && "cells" in first) {
      const sheets = (raw.sheets as WorkbookSheet[]).map((sheet, index) =>
        normalizeCanonicalSheet(sheet, index),
      );
      const activeSheetId =
        typeof raw.activeSheetId === "string" &&
        sheets.some((sheet) => sheet.id === raw.activeSheetId)
          ? raw.activeSheetId
          : sheets[0].id;
      const createdAt = typeof raw.createdAt === "number" ? raw.createdAt : Date.now();
      const updatedAt = typeof raw.updatedAt === "number" ? raw.updatedAt : createdAt;
      return {
        schema: WORKBOOK_SCHEMA,
        version: typeof raw.version === "number" ? raw.version : WORKBOOK_VERSION,
        title:
          typeof raw.title === "string" && raw.title.trim()
            ? raw.title
            : fallbackTitle,
        activeSheetId,
        createdAt,
        updatedAt,
        sheets,
      };
    }
    if (first && typeof first === "object" && ("celldata" in first || "data" in first)) {
      return fortuneSheetsToWorkbook({
        title:
          typeof raw.title === "string" && raw.title.trim()
            ? raw.title
            : fallbackTitle,
        sheets: raw.sheets as Sheet[],
        createdAt: typeof raw.createdAt === "number" ? raw.createdAt : Date.now(),
        updatedAt: typeof raw.updatedAt === "number" ? raw.updatedAt : Date.now(),
        activeSheetId:
          typeof raw.activeSheetId === "string" ? raw.activeSheetId : undefined,
      });
    }
  }
  return createEmptyWorkbook(fallbackTitle);
}

