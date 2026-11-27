// Toolbar for large sheet mode, laid out like the main editor's (FortuneSheet) toolbar. Buttons
// change the selected cells' style or number format; the ones that need the whole sheet in memory
// (filters, conditional formats, freeze, ...) are shown disabled with the reason.
import type { ReactNode } from "react";
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
];
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
  onFormat: (format: string | undefined) => void;
  onDecimals: (delta: number) => void;
}) {
  const off = !canEdit;
  const toggle = (key: "bold" | "italic" | "strike" | "underline") => onStyle({ [key]: !style[key] });
  return (
    <div className="big-toolbar" role="toolbar">
      <Button title="Tag a person" disabled>
        <UserRound size={16} />
      </Button>
      <Button title="Add a tag pill" disabled>
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
      <Button title="Borders" disabled>
        <Grid2x2 size={16} />
      </Button>
      <Button title="Merge cells" disabled>
        <TableCellsMerge size={16} />
      </Button>
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
      <Button title="Text rotation" disabled>
        <RotateCw size={16} />
      </Button>
      <Divider />
      <Button title="Freeze" disabled>
        <Snowflake size={16} />
      </Button>
      <Button title="Conditional formatting" disabled>
        <Palette size={16} />
      </Button>
      <Button title="Filter" disabled>
        <Filter size={16} />
      </Button>
      <Button title="Insert link" disabled>
        <Link size={16} />
      </Button>
      <Button title="Comment" disabled>
        <MessageSquare size={16} />
      </Button>
      <Button title="Formulas" disabled>
        <Sigma size={16} />
      </Button>
      <Button title="Data validation" disabled>
        <ListChecks size={16} />
      </Button>
      <Button title="Search" disabled>
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
