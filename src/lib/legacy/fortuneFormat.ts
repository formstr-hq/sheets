// The shape of a sheet in FortuneSheet's data format ("celldata", "config", "luckysheet_*" ...).
//
// FortuneSheet is no longer part of the editor, but sheets saved before the switch (and .xlsx import,
// templates and the streaming importer) still speak this format, so the converter in
// workbook/adapter.ts reads and writes it. These types are copied from @fortune-sheet/core 1.0.4
// (MIT License, Copyright (c) 2022 ruilisi) with the parts nobody here uses left out.

export type CellStyle = {
  bl?: number;
  it?: number;
  ff?: number | string;
  fs?: number;
  fc?: string;
  ht?: number;
  vt?: number;
  tb?: string;
  cl?: number;
  un?: number;
  tr?: string;
};

export type Cell = {
  v?: string | number | boolean;
  m?: string | number;
  mc?: { r: number; c: number; rs?: number; cs?: number };
  f?: string;
  ct?: { fa?: string; t?: string; s?: any };
  qp?: number;
  spl?: any;
  bg?: string;
  lo?: number;
  rt?: number;
  ps?: { left: number | null; top: number | null; width: number | null; height: number | null; value: string; isShow: boolean };
  hl?: { r: number; c: number; id: string };
} & CellStyle;

export type CellWithRowAndCol = { r: number; c: number; v: Cell | null };

export type CellMatrix = (Cell | null)[][];

export type SheetConfig = {
  merge?: Record<string, { r: number; c: number; rs: number; cs: number }>;
  rowlen?: Record<string, number>;
  columnlen?: Record<string, number>;
  rowhidden?: Record<string, number>;
  colhidden?: Record<string, number>;
  customHeight?: Record<string, number>;
  customWidth?: Record<string, number>;
  borderInfo?: any[];
  authority?: any;
  rowReadOnly?: Record<number, number>;
  colReadOnly?: Record<number, number>;
};

export type Image = { id: string; width: number; height: number; left: number; top: number; src: string };

export type Sheet = {
  name: string;
  config?: SheetConfig;
  order?: number;
  color?: string;
  data?: CellMatrix;
  celldata?: CellWithRowAndCol[];
  id?: string;
  images?: Image[];
  zoomRatio?: number;
  column?: number;
  row?: number;
  addRows?: number;
  status?: number;
  hide?: number;
  luckysheet_select_save?: any[];
  luckysheet_selection_range?: { row: number[]; column: number[] }[];
  calcChain?: any[];
  defaultRowHeight?: number;
  defaultColWidth?: number;
  showGridLines?: boolean | number;
  pivotTable?: any;
  isPivotTable?: boolean;
  filter?: Record<string, any>;
  filter_select?: { row: number[]; column: number[] };
  luckysheet_conditionformat_save?: any[];
  luckysheet_alternateformat_save?: any[];
  dataVerification?: any;
  hyperlink?: Record<string, { linkType: string; linkAddress: string }>;
  dynamicArray_compute?: any;
  dynamicArray?: any[];
  frozen?: {
    type: "row" | "column" | "both" | "rangeRow" | "rangeColumn" | "rangeBoth";
    range?: { row_focus: number; column_focus: number };
  };
};
