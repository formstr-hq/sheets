import type { Cell, Sheet } from "@fortune-sheet/core";
import { DEFAULT_COLS, DEFAULT_ROWS } from "./constants";
import { colName, randomId } from "./utils";
import type { SheetDocument, SheetTemplate } from "./types";
import type { FormstrCell } from "./sheetData";
import { fortuneSheetsToWorkbook } from "./workbook";

const INK = "#141719";
const PAPER = "#FAF8F5";
const MUTED = "#6B6560";
const LINE = "#E7E1D8";
const EMBER = "#FF5A36";

function text(value: string, extra: Partial<Cell> = {}): Cell {
  return {
    v: value,
    m: value,
    ct: { fa: "General", t: "s" },
    ...extra,
  };
}

function numberCell(value: number, format = "#,##0.00"): Cell {
  const display = format.includes("$")
    ? `$${value.toLocaleString("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      })}`
    : value.toFixed(2);
  return {
    v: value,
    m: display,
    ct: { fa: format, t: "n" },
  };
}

function formulaCell(
  formula: string,
  value: string | number,
  display: string,
  format = "General",
  type = "n",
): Cell {
  return {
    f: formula.startsWith("=") ? formula : `=${formula}`,
    v: value,
    m: display,
    ct: { fa: format, t: type },
  };
}

function tagCell(value: string): FormstrCell {
  return {
    v: value,
    m: value,
    ct: { fa: "General", t: "s", s: { fs: "tag" } },
  };
}

function dateSerial(year: number, month: number, day: number): number {
  return Math.round(
    (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86_400_000,
  );
}

function dateFormula(year: number, month: number, day: number): Cell {
  const serial = dateSerial(year, month, day);
  const display = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  return {
    f: `DATE(${year},${month},${day})`,
    v: serial,
    m: display,
    ct: { fa: "yyyy-mm-dd", t: "d" },
  };
}

function headerCell(value: string, align: 0 | 1 | 2 = 1): Cell {
  return text(value, {
    bg: "#F3EFE8",
    fc: "#57534E",
    bl: 1,
    fs: 10,
    ht: align,
    vt: 0,
  });
}

function titleCell(value: string): Cell {
  return text(value, {
    bl: 1,
    fs: 20,
    fc: INK,
    ht: 1,
    vt: 0,
  });
}

function dropdownValidation(
  rows: number[],
  col: number,
  options: string[],
): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  const rangeTxt = `${colName(col)}${rows[0] + 1}:${colName(col)}${rows[rows.length - 1] + 1}`;
  for (const row of rows) {
    result[`${row}_${col}`] = {
      type: "dropdown",
      type2: "",
      rangeTxt,
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

function statusRules(
  rows: number[],
  col: number,
  statuses: Array<{ value: string; bg: string; fg: string }>,
): unknown[] {
  return statuses.map((status) => ({
    type: "default",
    cellrange: [{ row: [rows[0], rows[rows.length - 1]], column: [col, col] }],
    format: { textColor: status.fg, cellColor: status.bg },
    conditionName: "equal",
    conditionRange: [],
    conditionValue: [status.value],
  }));
}

function sheet(params: {
  name: string;
  rows?: number;
  cols?: number;
  values: Array<{ r: number; c: number; cell: Cell }>;
  columnWidths?: Record<number, number>;
  rowHeights?: Record<number, number>;
  merges?: Array<{ r: number; c: number; rs: number; cs: number }>;
  dataVerification?: Record<string, unknown>;
  conditionalFormats?: unknown[];
}): Sheet {
  const rowHeights = params.rowHeights ?? {};
  const colWidths = params.columnWidths ?? {};
  return {
    id: randomId(10),
    name: params.name,
    row: params.rows ?? DEFAULT_ROWS,
    column: params.cols ?? DEFAULT_COLS,
    status: 1,
    order: 0,
    celldata: params.values.map((item) => ({ r: item.r, c: item.c, v: item.cell })),
    config: {
      rowlen: Object.fromEntries(Object.entries(rowHeights).map(([key, value]) => [key, value])),
      columnlen: Object.fromEntries(Object.entries(colWidths).map(([key, value]) => [key, value])),
      merge: Object.fromEntries(
        (params.merges ?? []).map((merge) => [
          `${merge.r}_${merge.c}`,
          { r: merge.r, c: merge.c, rs: merge.rs, cs: merge.cs },
        ]),
      ),
      borderInfo: [],
    },
    dataVerification: params.dataVerification,
    luckysheet_conditionformat_save: params.conditionalFormats,
    showGridLines: 1,
    defaultRowHeight: 30,
    defaultColWidth: 116,
  };
}

function blankDocument(title: string): SheetDocument {
  const now = Date.now();
  return fortuneSheetsToWorkbook({
    title,
    createdAt: now,
    updatedAt: now,
    sheets: [
      sheet({
        name: "Sheet 1",
        values: [],
        columnWidths: { 0: 160, 1: 180, 2: 160, 3: 160 },
      }),
    ],
  });
}

function projectDocument(title: string): SheetDocument {
  const statuses = [
    { value: "Done", bg: "#DCFCE7", fg: "#166534" },
    { value: "In progress", bg: "#FEF3C7", fg: "#92400E" },
    { value: "Blocked", bg: "#FEE2E2", fg: "#991B1B" },
    { value: "Not started", bg: "#F1F5F9", fg: "#475569" },
  ];
  const rows = [3, 4, 5, 6, 7, 8];
  const values: Array<{ r: number; c: number; cell: Cell }> = [
    { r: 0, c: 0, cell: titleCell("Project tracker") },
    { r: 2, c: 0, cell: headerCell("Task") },
    { r: 2, c: 1, cell: headerCell("Owner") },
    { r: 2, c: 2, cell: headerCell("Status", 0) },
    { r: 2, c: 3, cell: headerCell("Due", 0) },
    { r: 2, c: 4, cell: headerCell("Priority", 0) },
    { r: 2, c: 5, cell: headerCell("Notes") },
    { r: 2, c: 6, cell: headerCell("Progress", 0) },
    { r: 2, c: 7, cell: headerCell("Value", 0) },
    { r: 3, c: 0, cell: text("Design the empty state") },
    { r: 3, c: 2, cell: text("Done", { bg: "#DCFCE7", fc: "#166534", ht: 0 }) },
    { r: 3, c: 3, cell: dateFormula(2026, 11, 5) },
    { r: 3, c: 4, cell: tagCell("High") },
    { r: 3, c: 5, cell: text("First-run experience") },
    { r: 4, c: 0, cell: text("Wire up relay settings") },
    { r: 4, c: 2, cell: text("In progress", { bg: "#FEF3C7", fc: "#92400E", ht: 0 }) },
    { r: 4, c: 3, cell: dateFormula(2026, 11, 9) },
    { r: 4, c: 4, cell: tagCell("High") },
    { r: 4, c: 5, cell: text("NIP-65 aware") },
    { r: 5, c: 0, cell: text("Invite flow copy") },
    { r: 5, c: 2, cell: text("Not started", { ht: 0 }) },
    { r: 5, c: 3, cell: dateFormula(2026, 11, 12) },
    { r: 5, c: 4, cell: tagCell("Medium") },
    { r: 5, c: 5, cell: text("Needs review") },
    { r: 6, c: 0, cell: text("Mobile polish") },
    { r: 6, c: 2, cell: text("Not started", { ht: 0 }) },
    { r: 6, c: 3, cell: dateFormula(2026, 11, 18) },
    { r: 6, c: 4, cell: tagCell("Low") },
    { r: 7, c: 0, cell: text("Launch checklist") },
    { r: 7, c: 2, cell: text("Blocked", { bg: "#FEE2E2", fc: "#991B1B", ht: 0 }) },
    { r: 7, c: 3, cell: dateFormula(2026, 11, 21) },
    { r: 7, c: 4, cell: tagCell("High") },
    { r: 7, c: 5, cell: text("Waiting on domain") },
    { r: 8, c: 0, cell: text("Retro notes") },
    { r: 8, c: 2, cell: text("Not started", { ht: 0 }) },
    { r: 8, c: 3, cell: dateFormula(2026, 11, 25) },
    { r: 8, c: 4, cell: tagCell("Low") },
    { r: 8, c: 6, cell: text("Done", { fc: MUTED, ht: 2, fs: 10 }) },
    {
      r: 8,
      c: 7,
      cell: formulaCell('COUNTIF(C4:C9,"Done")', 2, "2"),
    },
    { r: 9, c: 6, cell: text("Open", { fc: MUTED, ht: 2, fs: 10 }) },
    {
      r: 9,
      c: 7,
      cell: formulaCell("COUNTA(C4:C9)-COUNTIF(C4:C9,\"Done\")", 4, "4"),
    },
    { r: 10, c: 6, cell: text("Progress", { fc: MUTED, ht: 2, fs: 10 }) },
    {
      r: 10,
      c: 7,
      cell: formulaCell(
        'IF(COUNTA(C4:C9)=0,0,COUNTIF(C4:C9,"Done")/COUNTA(C4:C9))',
        0.333,
        "33.3%",
        "0.0%",
      ),
    },
  ];

  const now = Date.now();
  return fortuneSheetsToWorkbook({
    title,
    createdAt: now,
    updatedAt: now,
    sheets: [
      sheet({
        name: "Projects",
        values,
        columnWidths: {
          0: 232,
          1: 150,
          2: 128,
          3: 112,
          4: 104,
          5: 220,
          6: 92,
          7: 82,
        },
        rowHeights: { 0: 44, 2: 32 },
        merges: [{ r: 0, c: 0, rs: 1, cs: 8 }],
        dataVerification: dropdownValidation(rows, 2, [
          "Not started",
          "In progress",
          "Blocked",
          "Done",
        ]),
        conditionalFormats: statusRules(rows, 2, statuses),
      }),
    ],
  });
}

function budgetDocument(title: string): SheetDocument {
  const rows = [3, 4, 5, 6, 7, 8];
  const data = [
    ["Design", 4200, 3800],
    ["Engineering", 12800, 11950],
    ["Hosting", 1800, 1540],
    ["Marketing", 3200, 2680],
    ["Tools", 1600, 1600],
    ["Buffer", 900, 0],
  ];
  const values: Array<{ r: number; c: number; cell: Cell }> = [
    { r: 0, c: 0, cell: titleCell("Launch budget") },
    { r: 2, c: 0, cell: headerCell("Category") },
    { r: 2, c: 1, cell: headerCell("Planned", 2) },
    { r: 2, c: 2, cell: headerCell("Actual", 2) },
    { r: 2, c: 3, cell: headerCell("Remaining", 2) },
    { r: 2, c: 4, cell: headerCell("Used", 2) },
    ...data.flatMap(([category, planned, actual], index) => {
      const r = 3 + index;
      return [
        { r, c: 0, cell: text(String(category)) },
        { r, c: 1, cell: numberCell(Number(planned), "$#,##0.00") },
        { r, c: 2, cell: numberCell(Number(actual), "$#,##0.00") },
        {
          r,
          c: 3,
          cell: formulaCell(
            `B${r + 1}-C${r + 1}`,
            Number(planned) - Number(actual),
            `$${(Number(planned) - Number(actual)).toFixed(2)}`,
            "$#,##0.00",
          ),
        },
        {
          r,
          c: 4,
          cell: formulaCell(
            `IF(B${r + 1}=0,0,C${r + 1}/B${r + 1})`,
            Number(planned) === 0 ? 0 : Number(actual) / Number(planned),
            `${((Number(actual) / Number(planned || 1)) * 100).toFixed(1)}%`,
            "0.0%",
          ),
        },
      ];
    }),
    { r: 9, c: 0, cell: text("Total", { bl: 1, bg: "#F3EFE8" }) },
    {
      r: 9,
      c: 1,
      cell: formulaCell("SUM(B4:B9)", 24300, "$24,300.00", "$#,##0.00"),
    },
    {
      r: 9,
      c: 2,
      cell: formulaCell("SUM(C4:C9)", 21570, "$21,570.00", "$#,##0.00"),
    },
    {
      r: 9,
      c: 3,
      cell: formulaCell("SUM(D4:D9)", 2730, "$2,730.00", "$#,##0.00"),
    },
    {
      r: 9,
      c: 4,
      cell: formulaCell("IF(B10=0,0,C10/B10)", 0.8877, "88.8%", "0.0%"),
    },
    { r: 11, c: 0, cell: text("Tip: change Actual and watch Remaining + Used recalculate.", { fc: MUTED, fs: 10, it: 1 }) },
  ];
  const now = Date.now();
  return fortuneSheetsToWorkbook({
    title,
    createdAt: now,
    updatedAt: now,
    sheets: [
      sheet({
        name: "Budget",
        values,
        columnWidths: { 0: 210, 1: 130, 2: 130, 3: 140, 4: 100 },
        rowHeights: { 0: 44, 2: 32 },
        merges: [{ r: 0, c: 0, rs: 1, cs: 5 }],
        dataVerification: dropdownValidation(rows, 0, data.map((row) => String(row[0]))),
      }),
    ],
  });
}

function calendarDocument(title: string): SheetDocument {
  const rows = [3, 4, 5, 6, 7, 8];
  const entries = [
    ["Nov 4", "Private sheets launch note", "Blog", "Draft"],
    ["Nov 6", "Behind the encryption", "Newsletter", "Idea"],
    ["Nov 9", "Collaboration walkthrough", "YouTube", "In progress"],
    ["Nov 12", "Template roundup", "Blog", "Not started"],
    ["Nov 16", "Nostr for teams", "Newsletter", "Idea"],
    ["Nov 20", "Roadmap AMA", "X", "Not started"],
  ];
  const values: Array<{ r: number; c: number; cell: Cell }> = [
    { r: 0, c: 0, cell: titleCell("Content calendar") },
    { r: 2, c: 0, cell: headerCell("Date", 0) },
    { r: 2, c: 1, cell: headerCell("Title") },
    { r: 2, c: 2, cell: headerCell("Channel", 0) },
    { r: 2, c: 3, cell: headerCell("Status", 0) },
    { r: 2, c: 4, cell: headerCell("Owner") },
    { r: 2, c: 5, cell: headerCell("Notes") },
    ...entries.flatMap(([date, entryTitle, channel, status], index) => {
      const r = 3 + index;
      return [
        { r, c: 0, cell: text(String(date), { ht: 0, fc: MUTED }) },
        { r, c: 1, cell: text(String(entryTitle), { ...(status === "In progress" ? { bl: 1 } : {}) }) },
        { r, c: 2, cell: text(String(channel), { ht: 0 }) },
        { r, c: 3, cell: text(String(status), { ht: 0 }) },
        { r, c: 4, cell: text("") },
        { r, c: 5, cell: tagCell(index % 2 === 0 ? "Series" : "One-off") },
      ];
    }),
  ];
  const now = Date.now();
  return fortuneSheetsToWorkbook({
    title,
    createdAt: now,
    updatedAt: now,
    sheets: [
      sheet({
        name: "Calendar",
        values,
        columnWidths: { 0: 100, 1: 260, 2: 128, 3: 140, 4: 150, 5: 120 },
        rowHeights: { 0: 44, 2: 32 },
        merges: [{ r: 0, c: 0, rs: 1, cs: 6 }],
        dataVerification: {
          ...dropdownValidation(rows, 2, ["Blog", "Newsletter", "X", "YouTube"]),
          ...dropdownValidation(rows, 3, ["Idea", "Draft", "In progress", "Not started"]),
        },
        conditionalFormats: statusRules(rows, 3, [
          { value: "In progress", bg: "#FEF3C7", fg: "#92400E" },
          { value: "Draft", bg: "#DBEAFE", fg: "#1E40AF" },
          { value: "Idea", bg: "#F1F5F9", fg: "#475569" },
        ]),
      }),
    ],
  });
}

export const TEMPLATES: SheetTemplate[] = [
  {
    id: "blank",
    name: "Blank",
    description: "A clean grid with room to grow.",
    color: "#64748B",
    build: blankDocument,
  },
  {
    id: "project",
    name: "Project tracker",
    description: "Owners, status dropdowns and priority pills.",
    color: EMBER,
    build: projectDocument,
  },
  {
    id: "budget",
    name: "Launch budget",
    description: "Currency formats and live formulas.",
    color: "#10B981",
    build: budgetDocument,
  },
  {
    id: "calendar",
    name: "Content calendar",
    description: "Channels, statuses and tag pills.",
    color: "#6366F1",
    build: calendarDocument,
  },
];

export function templateById(id: string): SheetTemplate {
  return TEMPLATES.find((template) => template.id === id) ?? TEMPLATES[0];
}

export const TEMPLATE_BORDER = LINE;
export const TEMPLATE_INK = INK;
export const TEMPLATE_PAPER = PAPER;
