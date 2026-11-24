// Toolbar for large sheet mode, laid out like the main editor's (FortuneSheet) toolbar. Buttons
// change the selected cells' style or number format; the ones that need the whole sheet in memory
// (filters, conditional formats, freeze, ...) are shown disabled with the reason.
import { useState, type ReactNode } from "react";
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownToLine,
  ArrowUpToLine,
  Baseline,
  Bold,
  Eraser,
  Filter,
  FoldVertical,
  Grid2x2,
  Italic,
  Link,
  ListChecks,
  MessageSquare,
  PaintBucket,
  Paintbrush,
  Percent,
  Redo2,
  RotateCw,
  Search,
  Sigma,
  Snowflake,
  Strikethrough,
  TableCellsMerge,
  Underline,
  Undo2,
  WrapText,
  DollarSign,
  Palette,
  Tag,
  UserRound,
} from "lucide-react";
import type { WorkbookCellStyle } from "../lib/workbook";
import { DEFAULT_FONT_FAMILIES } from "../lib/workbook/format";
import { cx } from "../lib/utils";

export type StylePatch = Partial<WorkbookCellStyle>;

export const NUMBER_FORMATS: { label: string; format: string | undefined }[] = [
  { label: "Automatic", format: undefined },
  { label: "Plain text", format: "@" },
  { label: "Number", format: "#,##0.00" },
  { label: "Percent", format: "0.00%" },
  { label: "Currency", format: "$#,##0.00" },
  { label: "Date", format: "yyyy-mm-dd" },
  { label: "Scientific", format: "0.00E+00" },
  { label: "Accounting", format: "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)" },
  { label: "Time", format: "h:mm:ss AM/PM" },
  { label: "Time (24 hour)", format: "hh:mm:ss" },
  { label: "Date and time", format: "yyyy-mm-dd hh:mm:ss" },
  { label: "Long date", format: "mmmm d, yyyy" },
  { label: "Euro", format: "[$€-407] #,##0.00" },
  { label: "British pound", format: "[$£-809]#,##0.00" },
  { label: "Japanese yen", format: "[$¥-411]#,##0" },
  { label: "Indian rupee", format: "[$₹-4009]#,##0.00" },
  { label: "Swiss franc", format: "[$CHF-807] #,##0.00" },
  { label: "Negative in red", format: "#,##0.00;[Red]-#,##0.00" },
];
const ROTATIONS = ["No rotation", "Tilt up", "Tilt down", "Stacked", "Rotate up", "Rotate down"];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 30, 36];
const NOT_HERE = "Not available in large sheet mode yet";

function Button({
  title,
  active,
  disabled,
  onClick,
  children,
}: {
  title: string;
  active?: boolean;
  disabled?: boolean;
  onClick?: () => void;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      className={cx("big-tool", active && "is-active")}
      title={disabled ? `${title}: ${NOT_HERE}` : title}
      disabled={disabled}
      onMouseDown={(event) => event.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  );
}

const Divider = () => <span className="big-tool-divider" />;

export function BigSheetToolbar({
  style,
  canEdit,
  canUndo,
  canRedo,
  painting,
  onUndo,
  onRedo,
  onPaint,
  onClear,
  onStyle,
  freeze,
  zoom,
  onQuickFormula,
  onSearch,
  onLink,
  onComment,
  onValidation,
  onPerson,
  onTagPill,
  onConditional,
  filtered,
  onFilter,
  onBorder,
  onMerge,
  canSplit,
  onFormat,
  onDecimals,
}: {
  /** Style of the selected (anchor) cell, to show which buttons are on. */
  style: WorkbookCellStyle;
  canEdit: boolean;
  canUndo: boolean;
  canRedo: boolean;
  painting: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onPaint: () => void;
  onClear: () => void;
  onStyle: (patch: StylePatch) => void;
  /** Frozen panes: how many rows/columns are frozen now, the choices, and how to apply one. */
  freeze: { rows: number; columns: number; options: { label: string; rows: number; columns: number }[]; apply: (rows: number, columns: number) => void };
  /** Writes SUM/AVERAGE/... of each selected column just below the selection. */
  onQuickFormula: (name: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN") => void;
  onSearch: () => void;
  onLink: () => void;
  onComment: () => void;
  onValidation: () => void;
  onPerson: () => void;
  onTagPill: () => void;
  onConditional: () => void;
  /** Is a filter on? */
  filtered: boolean;
  onFilter: () => void;
  /** Zoom as a fraction (1 = 100%) and how to change it. */
  zoom: { value: number; set: (zoom: number) => void };
  /** Puts a border on the selection. */
  onBorder: (type: "all" | "outside" | "inside" | "top" | "bottom" | "left" | "right" | "none", color: string, style: string) => void;
  onMerge: (how: "all" | "rows" | "columns" | "split") => void;
  /** Is there a merged block in the selection (so "Unmerge" does something)? */
  canSplit: boolean;
  onFormat: (format: string | undefined) => void;
  onDecimals: (delta: number) => void;
}) {
  const off = !canEdit;
  const [borderColor, setBorderColor] = useState("#000000");
  const [borderStyle, setBorderStyle] = useState("thin");
  const toggle = (key: "bold" | "italic" | "strike" | "underline") => onStyle({ [key]: !style[key] });
  return (
    <div className="big-toolbar" role="toolbar">
      <Button title="Tag a person" disabled={off} onClick={onPerson}>
        <UserRound size={16} />
      </Button>
      <Button title="Add a tag pill" disabled={off} onClick={onTagPill}>
        <Tag size={16} />
      </Button>
      <Divider />
      <Button title="Undo (Ctrl+Z)" disabled={off || !canUndo} onClick={onUndo}>
        <Undo2 size={16} />
      </Button>
      <Button title="Redo (Ctrl+Y)" disabled={off || !canRedo} onClick={onRedo}>
        <Redo2 size={16} />
      </Button>
      <Divider />
      <Button title="Paint format" active={painting} disabled={off} onClick={onPaint}>
        <Paintbrush size={16} />
      </Button>
      <Button title="Clear format" disabled={off} onClick={onClear}>
        <Eraser size={16} />
      </Button>
      <Divider />
      <Button title="Format as currency" disabled={off} onClick={() => onFormat("$#,##0.00")}>
        <DollarSign size={16} />
      </Button>
      <Button title="Format as percent" disabled={off} onClick={() => onFormat("0.00%")}>
        <Percent size={16} />
      </Button>
      <Button title="Decrease decimal places" disabled={off} onClick={() => onDecimals(-1)}>
        <span className="big-tool-text">.0</span>
      </Button>
      <Button title="Increase decimal places" disabled={off} onClick={() => onDecimals(1)}>
        <span className="big-tool-text">.00</span>
      </Button>
      <select
        className="big-tool-select"
        title="Number format"
        disabled={off}
        value={NUMBER_FORMATS.find((item) => item.format === style.numberFormat)?.label ?? (style.numberFormat ? "Custom" : "Automatic")}
        onChange={(event) => onFormat(NUMBER_FORMATS.find((item) => item.label === event.target.value)?.format)}
      >
        {NUMBER_FORMATS.map((item) => (
          <option key={item.label}>{item.label}</option>
        ))}
        {style.numberFormat && !NUMBER_FORMATS.some((item) => item.format === style.numberFormat) && <option>Custom</option>}
      </select>
      <Divider />
      <select
        className="big-tool-select is-wide"
        title="Font"
        disabled={off}
        value={style.fontFamily ?? DEFAULT_FONT_FAMILIES[0]}
        onChange={(event) => onStyle({ fontFamily: event.target.value })}
      >
        {[...new Set([...DEFAULT_FONT_FAMILIES, ...(style.fontFamily ? [style.fontFamily] : [])])].map((font) => (
          <option key={font}>{font}</option>
        ))}
      </select>
      <Divider />
      <select
        className="big-tool-select"
        title="Font size"
        disabled={off}
        value={style.fontSize ?? 11}
        onChange={(event) => onStyle({ fontSize: Number(event.target.value) })}
      >
        {[...new Set([...FONT_SIZES, style.fontSize ?? 11])].sort((a, b) => a - b).map((size) => (
          <option key={size} value={size}>
            {size}
          </option>
        ))}
      </select>
      <Divider />
      <Button title="Bold (Ctrl+B)" active={!!style.bold} disabled={off} onClick={() => toggle("bold")}>
        <Bold size={16} />
      </Button>
      <Button title="Italic (Ctrl+I)" active={!!style.italic} disabled={off} onClick={() => toggle("italic")}>
        <Italic size={16} />
      </Button>
      <Button title="Strikethrough" active={!!style.strike} disabled={off} onClick={() => toggle("strike")}>
        <Strikethrough size={16} />
      </Button>
      <Button title="Underline (Ctrl+U)" active={!!style.underline} disabled={off} onClick={() => toggle("underline")}>
        <Underline size={16} />
      </Button>
      <Divider />
      <label className={cx("big-tool", off && "is-disabled")} title="Text color">
        <Baseline size={16} style={{ color: style.color }} />
        <input type="color" disabled={off} value={style.color ?? "#000000"} onChange={(event) => onStyle({ color: event.target.value })} />
      </label>
      <label className={cx("big-tool", off && "is-disabled")} title="Fill color">
        <PaintBucket size={16} style={{ color: style.background }} />
        <input
          type="color"
          disabled={off}
          value={style.background ?? "#ffffff"}
          onChange={(event) => onStyle({ background: event.target.value })}
        />
      </label>
      <label className="big-tool is-select" title="Zoom (Ctrl + wheel)">
        <select className="big-tool-select" value={String(zoom.value)} onChange={(event) => zoom.set(Number(event.target.value))}>
          {[0.5, 0.75, 0.9, 1, 1.25, 1.5, 2].map((z) => (
            <option key={z} value={String(z)}>
              {Math.round(z * 100)}%
            </option>
          ))}
          {![0.5, 0.75, 0.9, 1, 1.25, 1.5, 2].includes(zoom.value) && <option value={String(zoom.value)}>{Math.round(zoom.value * 100)}%</option>}
        </select>
      </label>
      <label className={cx("big-tool is-select", off && "is-disabled")} title="Borders">
        <Grid2x2 size={16} />
        <select
          className="big-tool-select"
          disabled={off}
          value=""
          onChange={(event) => event.target.value && onBorder(event.target.value as Parameters<typeof onBorder>[0], borderColor, borderStyle)}
        >
          <option value="" disabled>
            Borders
          </option>
          <option value="all">All borders</option>
          <option value="outside">Outer borders</option>
          <option value="inside">Inner borders</option>
          <option value="top">Top border</option>
          <option value="bottom">Bottom border</option>
          <option value="left">Left border</option>
          <option value="right">Right border</option>
          <option value="none">No borders</option>
        </select>
      </label>
      <label className={cx("big-tool is-select", off && "is-disabled")} title="Border line style">
        <select className="big-tool-select" disabled={off} value={borderStyle} onChange={(event) => setBorderStyle(event.target.value)}>
          <option value="thin">Thin</option>
          <option value="medium">Medium</option>
          <option value="thick">Thick</option>
          <option value="dashed">Dashed</option>
          <option value="dotted">Dotted</option>
        </select>
      </label>
      <label className={cx("big-tool", off && "is-disabled")} title="Border color">
        <span style={{ width: 14, height: 14, border: "2px solid " + borderColor, borderRadius: 3 }} />
        <input type="color" disabled={off} value={borderColor} onChange={(event) => setBorderColor(event.target.value)} />
      </label>
      <label className={cx("big-tool is-select", off && "is-disabled")} title="Merge cells">
        <TableCellsMerge size={16} />
        <select
          className="big-tool-select"
          disabled={off}
          value=""
          onChange={(event) => event.target.value && onMerge(event.target.value as Parameters<typeof onMerge>[0])}
        >
          <option value="" disabled>
            Merge
          </option>
          <option value="all">Merge all</option>
          <option value="rows">Merge each row</option>
          <option value="columns">Merge each column</option>
          <option value="split" disabled={!canSplit}>
            Unmerge
          </option>
        </select>
      </label>
      <Divider />
      <Button title="Align left" active={style.align === "left"} disabled={off} onClick={() => onStyle({ align: "left" })}>
        <AlignLeft size={16} />
      </Button>
      <Button title="Align center" active={style.align === "center"} disabled={off} onClick={() => onStyle({ align: "center" })}>
        <AlignCenter size={16} />
      </Button>
      <Button title="Align right" active={style.align === "right"} disabled={off} onClick={() => onStyle({ align: "right" })}>
        <AlignRight size={16} />
      </Button>
      <Button title="Align top" active={style.verticalAlign === "top"} disabled={off} onClick={() => onStyle({ verticalAlign: "top" })}>
        <ArrowUpToLine size={16} />
      </Button>
      <Button
        title="Align middle"
        active={style.verticalAlign === "middle"}
        disabled={off}
        onClick={() => onStyle({ verticalAlign: "middle" })}
      >
        <FoldVertical size={16} />
      </Button>
      <Button
        title="Align bottom"
        active={style.verticalAlign === "bottom"}
        disabled={off}
        onClick={() => onStyle({ verticalAlign: "bottom" })}
      >
        <ArrowDownToLine size={16} />
      </Button>
      <Button title="Wrap text" active={!!style.wrap} disabled={off} onClick={() => onStyle({ wrap: !style.wrap })}>
        <WrapText size={16} />
      </Button>
      <label className={cx("big-tool is-select", off && "is-disabled")} title="Text rotation">
        <RotateCw size={16} />
        <select
          className="big-tool-select"
          disabled={off}
          value=""
          onChange={(event) => event.target.value !== "" && onStyle({ rotation: Number(event.target.value) || undefined })}
        >
          <option value="" disabled>
            {style.rotation ? ROTATIONS[style.rotation] : "Rotate"}
          </option>
          {ROTATIONS.map((label, code) => (
            <option key={label} value={code}>
              {label}
            </option>
          ))}
        </select>
      </label>
      <Divider />
      <label className="big-tool is-select" title="Freeze rows and columns">
        <Snowflake size={16} />
        <select
          className="big-tool-select"
          value=""
          onChange={(event) => {
            const choice = freeze.options[Number(event.target.value)];
            if (choice) freeze.apply(choice.rows, choice.columns);
          }}
        >
          <option value="" disabled>
            {freeze.rows || freeze.columns
              ? `Frozen: ${[freeze.rows ? `${freeze.rows} row${freeze.rows > 1 ? "s" : ""}` : "", freeze.columns ? `${freeze.columns} col${freeze.columns > 1 ? "s" : ""}` : ""].filter(Boolean).join(", ")}`
              : "Freeze"}
          </option>
          {freeze.options.map((item, index) => (
            <option key={item.label} value={index}>
              {item.label}
            </option>
          ))}
        </select>
      </label>
      <Button title="Conditional formatting" disabled={off} onClick={onConditional}>
        <Palette size={16} />
      </Button>
      <Button title="Filter by the values of the active column" active={filtered} disabled={off} onClick={onFilter}>
        <Filter size={16} />
      </Button>
      <Button title="Insert link" disabled={off} onClick={onLink}>
        <Link size={16} />
      </Button>
      <Button title="Comment" disabled={off} onClick={onComment}>
        <MessageSquare size={16} />
      </Button>
      <label className={cx("big-tool is-select", off && "is-disabled")} title="Quick formula: total of the selected columns, written below them">
        <Sigma size={16} />
        <select
          className="big-tool-select"
          disabled={off}
          value=""
          onChange={(event) => event.target.value && onQuickFormula(event.target.value as Parameters<typeof onQuickFormula>[0])}
        >
          <option value="" disabled>
            Σ
          </option>
          <option value="SUM">Sum</option>
          <option value="AVERAGE">Average</option>
          <option value="COUNT">Count</option>
          <option value="MAX">Max</option>
          <option value="MIN">Min</option>
        </select>
      </label>
      <Button title="Data validation" disabled={off} onClick={onValidation}>
        <ListChecks size={16} />
      </Button>
      <Button title="Find (Ctrl+F)" onClick={onSearch}>
        <Search size={16} />
      </Button>
    </div>
  );
}

/** Changes the number of decimals of a format (adds a plain number format when there is none). */
export function shiftDecimals(format: string | undefined, delta: number): string {
  const base = !format || format === "General" || format === "@" ? "#,##0" : format;
  const match = base.match(/\.(0+)/);
  const decimals = Math.max(0, Math.min(10, (match ? match[1].length : 0) + delta));
  const zeros = decimals ? `.${"0".repeat(decimals)}` : "";
  if (match) return base.replace(/\.0+/, zeros);
  return base.replace(/(0)(?!.*0)/, `$1${zeros}`);
}
