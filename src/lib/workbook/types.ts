export const WORKBOOK_SCHEMA = "formstr.sheets.workbook";
export const WORKBOOK_VERSION = 1;

export type CellPrimitive = string | number | boolean | null;
export type CellKind = "text" | "number" | "boolean" | "date" | "error";

export interface WorkbookPerson {
  pubkey: string;
  npub?: string;
  name?: string;
  picture?: string;
}

export interface WorkbookCellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  background?: string;
  align?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  wrap?: boolean;
  rotation?: number;
  numberFormat?: string;
}

export interface WorkbookValidation {
  type: string;
  options?: string[];
  params?: Record<string, string | boolean>;
}

export interface WorkbookLink {
  url: string;
  label?: string;
}

export interface WorkbookCell {
  value?: CellPrimitive;
  formula?: string;
  display?: string;
  type?: CellKind;
  style?: WorkbookCellStyle;
  person?: WorkbookPerson;
  tag?: string;
  validation?: WorkbookValidation;
  link?: WorkbookLink;
}

export interface WorkbookMerge {
  r: number;
  c: number;
  rs: number;
  cs: number;
}

export interface WorkbookRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export type WorkbookBorderType =
  | "all"
  | "outside"
  | "inside"
  | "top"
  | "bottom"
  | "left"
  | "right"
  | "none";

export interface WorkbookBorder {
  range: WorkbookRange;
  type: WorkbookBorderType;
  color?: string;
  style?: string;
}

export interface WorkbookConditionalRule {
  ranges: WorkbookRange[];
  condition: {
    type: string;
    values: string[];
  };
  format: {
    textColor?: string;
    cellColor?: string;
  };
}

export interface WorkbookSheet {
  id: string;
  name: string;
  order: number;
  rowCount: number;
  colCount: number;
  hidden?: boolean;
  color?: string;
  showGridLines?: boolean;
  defaultRowHeight?: number;
  defaultColWidth?: number;
  cells: Record<string, WorkbookCell>;
  columnWidths?: Record<number, number>;
  rowHeights?: Record<number, number>;
  hiddenRows?: number[];
  hiddenColumns?: number[];
  merges?: WorkbookMerge[];
  frozen?: {
    rows: number;
    columns: number;
  };
  conditionalFormats?: WorkbookConditionalRule[];
  borders?: WorkbookBorder[];
}

export interface WorkbookDocument {
  schema: typeof WORKBOOK_SCHEMA;
  version: number;
  title: string;
  activeSheetId: string;
  createdAt: number;
  updatedAt: number;
  sheets: WorkbookSheet[];
}
