// Exporting a big sheet as an .xlsx file without ever holding the whole sheet: the chunks are read one
// after another and each one's rows are written straight into a streaming zip (fflate), as inline
// strings so no shared-strings table has to be built first. Styles are collected while the rows go by
// and written last (a zip's entries can come in any order).
//
// Covered: values, formulas (with their last result), number formats, fonts, fills, alignment,
// rotation, borders, merged cells, column widths, row heights, hidden rows and columns, frozen panes,
// tab names/colours/hidden, hyperlinks, list/number/date/length validation, value-based conditional
// formatting, data bars and colour scales. Not covered: comments, person tags (their text is kept).
import { Zip, ZipDeflate, strToU8 } from "fflate";
import type { WorkbookCell, WorkbookConditionalRule } from "../workbook";
import { columnName } from "@formstr/sheets-grid";
import { edgesIn, edgeKey, type Border } from "@formstr/sheets-grid";
import type { BigSheetStore, TabLayout } from "./bigSheet";

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** Escapes text for XML and drops characters XML cannot hold. */
export function esc(text: string): string {
  // eslint-disable-next-line no-control-regex
  return text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** "#aabbcc" as the ARGB Excel wants ("FFAABBCC"); undefined for anything else. */
export function argb(color: string | undefined): string | undefined {
  const m = /^#?([0-9a-f]{6})$/i.exec((color ?? "").trim());
  return m ? `FF${m[1].toUpperCase()}` : undefined;
}

/** A sheet name Excel accepts: no []:*?/\, at most 31 characters, different from the others. */
export function sheetNames(names: string[]): string[] {
  const used = new Set<string>();
  return names.map((name) => {
    const base = (name.replace(/[[\]:*?/\\]/g, "_").trim() || "Sheet").slice(0, 31);
    let candidate = base;
    for (let n = 2; used.has(candidate.toLowerCase()); n++) candidate = `${base.slice(0, 31 - String(n).length - 1)} ${n}`;
    used.add(candidate.toLowerCase());
    return candidate;
  });
}

const BUILTIN_FORMATS: Record<string, number> = {
  General: 0,
  "0": 1,
  "0.00": 2,
  "#,##0": 3,
  "#,##0.00": 4,
  "0%": 9,
  "0.00%": 10,
  "0.00E+00": 11,
  "mm-dd-yy": 14,
  "h:mm AM/PM": 18,
  "h:mm:ss AM/PM": 19,
  "h:mm": 20,
  "h:mm:ss": 21,
  "@": 49,
};

const BORDER_STYLE: Record<string, string> = { thin: "thin", medium: "medium", thick: "thick", dashed: "dashed", dotted: "dotted" };

type Sides = { top?: { style: string; color?: string }; bottom?: { style: string; color?: string }; left?: { style: string; color?: string }; right?: { style: string; color?: string } };

/** Collects the distinct looks of cells and hands out the style numbers (cellXfs indexes) for them. */
export class StyleTable {
  private fonts: string[] = ['<font><sz val="11"/><name val="Arial"/></font>'];
  private fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  private borders: string[] = ["<border><left/><right/><top/><bottom/><diagonal/></border>"];
  private formats = new Map<string, number>();
  private xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  private lookup = new Map<string, number>();
  private fontIds = new Map<string, number>();
  private fillIds = new Map<string, number>();
  private borderIds = new Map<string, number>();
  readonly dxfs: string[] = [];

  private font(style: NonNullable<WorkbookCell["style"]>): number {
    const parts = [
      style.bold ? "<b/>" : "",
      style.italic ? "<i/>" : "",
      style.strike ? "<strike/>" : "",
      style.underline ? "<u/>" : "",
      `<sz val="${style.fontSize ?? 11}"/>`,
      argb(style.color) ? `<color rgb="${argb(style.color)}"/>` : "",
      `<name val="${esc(style.fontFamily ?? "Arial")}"/>`,
    ].join("");
    const xml = `<font>${parts}</font>`;
    let id = this.fontIds.get(xml);
    if (id === undefined) {
      id = this.fonts.push(xml) - 1;
      this.fontIds.set(xml, id);
    }
    return id;
  }

  private fill(color: string | undefined): number {
    const rgb = argb(color);
    if (!rgb) return 0;
    const xml = `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill>`;
    let id = this.fillIds.get(xml);
    if (id === undefined) {
      id = this.fills.push(xml) - 1;
      this.fillIds.set(xml, id);
    }
    return id;
  }

  private border(sides: Sides | undefined): number {
    if (!sides || !(sides.top || sides.bottom || sides.left || sides.right)) return 0;
    const side = (name: string, value: { style: string; color?: string } | undefined) =>
      value ? `<${name} style="${BORDER_STYLE[value.style] ?? "thin"}"><color rgb="${argb(value.color) ?? "FF000000"}"/></${name}>` : `<${name}/>`;
    const xml = `<border>${side("left", sides.left)}${side("right", sides.right)}${side("top", sides.top)}${side("bottom", sides.bottom)}<diagonal/></border>`;
    let id = this.borderIds.get(xml);
    if (id === undefined) {
      id = this.borders.push(xml) - 1;
      this.borderIds.set(xml, id);
    }
    return id;
  }

  private format(code: string | undefined): number {
    if (!code || code === "General") return 0;
    if (code in BUILTIN_FORMATS) return BUILTIN_FORMATS[code];
    let id = this.formats.get(code);
    if (id === undefined) {
      id = 164 + this.formats.size;
      this.formats.set(code, id);
    }
    return id;
  }

  /** The style number for a cell's style (and the borders around it); 0 when it has none. */
  id(style: WorkbookCell["style"] | undefined, sides?: Sides): number {
    const hasStyle = style && Object.keys(style).length > 0;
    const hasSides = sides && (sides.top || sides.bottom || sides.left || sides.right);
    if (!hasStyle && !hasSides) return 0;
    const s = style ?? {};
    const key = JSON.stringify([s, sides ?? null]);
    const known = this.lookup.get(key);
    if (known !== undefined) return known;
    const rotation = s.rotation ? ({ 1: 45, 2: 135, 3: 255, 4: 90, 5: 180 } as Record<number, number>)[s.rotation] : undefined;
    const alignment =
      s.align || s.verticalAlign || s.wrap || rotation
        ? `<alignment${s.align ? ` horizontal="${s.align}"` : ""}${s.verticalAlign ? ` vertical="${s.verticalAlign === "middle" ? "center" : s.verticalAlign}"` : ""}${s.wrap ? ' wrapText="1"' : ""}${rotation ? ` textRotation="${rotation}"` : ""}/>`
        : "";
    const fontId = this.font(s);
    const fillId = this.fill(s.background);
    const borderId = this.border(sides);
    const numFmtId = this.format(s.numberFormat);
    const xf = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ""} applyFont="1"${fillId ? ' applyFill="1"' : ""}${borderId ? ' applyBorder="1"' : ""}${alignment ? ' applyAlignment="1">' + alignment + "</xf>" : "/>"}`;
    const id = this.xfs.push(xf) - 1;
    this.lookup.set(key, id);
    return id;
  }

  /** A formatting block for conditional formatting rules ("differential" style); returns its number. */
  dxf(format: WorkbookConditionalRule["format"]): number {
    const font = argb(format.textColor) ? `<font><color rgb="${argb(format.textColor)}"/></font>` : "";
    const fill = argb(format.cellColor) ? `<fill><patternFill patternType="solid"><fgColor rgb="${argb(format.cellColor)}"/><bgColor rgb="${argb(format.cellColor)}"/></patternFill></fill>` : "";
    return this.dxfs.push(`<dxf>${font}${fill}</dxf>`) - 1;
  }

  xml(): string {
    const formats = [...this.formats].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`).join("");
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS_MAIN}">` +
      (formats ? `<numFmts count="${this.formats.size}">${formats}</numFmts>` : "") +
      `<fonts count="${this.fonts.length}">${this.fonts.join("")}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join("")}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join("")}</borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join("")}</cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
      `<dxfs count="${this.dxfs.length}">${this.dxfs.join("")}</dxfs></styleSheet>`
    );
  }
}

/** One cell as XML (`ref` like "B7"); null when it has nothing to write. */
export function cellXml(ref: string, cell: WorkbookCell | undefined, styleId: number): string | null {
  if (!cell && !styleId) return null;
  const s = styleId ? ` s="${styleId}"` : "";
  if (!cell || (cell.value === undefined && !cell.formula)) return styleId ? `<c r="${ref}"${s}/>` : null;
  const value = cell.value;
  const formula = cell.formula ? `<f>${esc(cell.formula.replace(/^=/, ""))}</f>` : "";
  if (cell.type === "error" || (typeof value === "string" && /^#(REF|DIV\/0|NAME|VALUE|N\/A|NUM|NULL|CIRC)[!?]?$/.test(value) && cell.formula)) {
    return `<c r="${ref}"${s} t="e">${formula}<v>${esc(String(value ?? "#VALUE!"))}</v></c>`;
  }
  if (typeof value === "number" && Number.isFinite(value)) return `<c r="${ref}"${s}>${formula}<v>${value}</v></c>`;
  if (typeof value === "boolean") return `<c r="${ref}"${s} t="b">${formula}<v>${value ? 1 : 0}</v></c>`;
  if (value === undefined || value === null) return `<c r="${ref}"${s}>${formula}</c>`;
  const text = String(value);
  if (cell.formula) return `<c r="${ref}"${s} t="str">${formula}<v>${esc(text)}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(text)}</t></is></c>`;
}

const OPERATORS: Record<string, string> = {
  between: "between",
  notBetween: "notBetween",
  eq: "equal",
  neq: "notEqual",
  gt: "greaterThan",
  gte: "greaterThanOrEqual",
  lt: "lessThan",
  lte: "lessThanOrEqual",
};

function dateSerial(text: string): number | null {
  const m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(text.trim());
  return m ? Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / 86_400_000 + 25569 : null;
}

/** The <dataValidation> element for a cell's rule, or null for rules Excel has no direct form of. */
export function validationXml(ref: string, rule: NonNullable<WorkbookCell["validation"]>): string | null {
  const op = OPERATORS[String(rule.params?.op ?? "between")] ?? "between";
  const a = String(rule.params?.a ?? "");
  const b = String(rule.params?.b ?? "");
  const head = `<dataValidation allowBlank="1" showErrorMessage="1" sqref="${ref}"`;
  if (rule.type === "list") {
    const options = (rule.options ?? []).join(",").replace(/"/g, '""');
    return options.length > 250 ? null : `${head} type="list"><formula1>"${esc(options)}"</formula1></dataValidation>`;
  }
  if (rule.type === "number" || rule.type === "integer" || rule.type === "length") {
    const type = rule.type === "number" ? "decimal" : rule.type === "integer" ? "whole" : "textLength";
    const second = op === "between" || op === "notBetween" ? `<formula2>${esc(b)}</formula2>` : "";
    return `${head} type="${type}" operator="${op}"><formula1>${esc(a)}</formula1>${second}</dataValidation>`;
  }
  if (rule.type === "date") {
    const from = dateSerial(a);
    const to = dateSerial(b);
    if (from === null) return null;
    const second = op === "between" || op === "notBetween" ? `<formula2>${to ?? from}</formula2>` : "";
    return `${head} type="date" operator="${op}"><formula1>${from}</formula1>${second}</dataValidation>`;
  }
  return null;
}

/** The <conditionalFormatting> blocks for a tab's rules (value rules, bars and scales). Priorities count from `first`. */
export function conditionalXml(rules: readonly WorkbookConditionalRule[], styles: StyleTable): string {
  let priority = 1;
  const blocks: string[] = [];
  for (const rule of rules) {
    const sqref = rule.ranges.map((r) => `${columnName(r.c1)}${r.r1 + 1}:${columnName(r.c2)}${r.r2 + 1}`).join(" ");
    const topLeft = `${columnName(rule.ranges[0].c1)}${rule.ranges[0].r1 + 1}`;
    const [v1 = "", v2 = ""] = rule.condition.values;
    let inner: string | null = null;
    const type = rule.condition.type;
    if (type === "dataBar") {
      inner = `<cfRule type="dataBar" priority="${priority}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${argb(rule.format.cellColor) ?? "FF5B9BD5"}"/></dataBar></cfRule>`;
    } else if (type === "colorScale") {
      inner = `<cfRule type="colorScale" priority="${priority}"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="${argb(rule.format.cellColor) ?? "FFFFFFFF"}"/><color rgb="${argb(rule.format.toColor) ?? "FF63BE7B"}"/></colorScale></cfRule>`;
    } else {
      const dxfId = styles.dxf(rule.format);
      if (OPERATORS[type]) {
        const second = type === "between" ? `<formula>${esc(v2)}</formula>` : "";
        const quoted = (x: string) => (Number.isFinite(Number(x)) && x.trim() !== "" ? x : `"${esc(x.replace(/"/g, '""'))}"`);
        inner = `<cfRule type="cellIs" dxfId="${dxfId}" priority="${priority}" operator="${OPERATORS[type]}"><formula>${quoted(v1)}</formula>${second}</cfRule>`;
      } else if (type === "contains") {
        inner = `<cfRule type="containsText" dxfId="${dxfId}" priority="${priority}" operator="containsText" text="${esc(v1)}"><formula>NOT(ISERROR(SEARCH("${esc(v1.replace(/"/g, '""'))}",${topLeft})))</formula></cfRule>`;
      } else if (type === "empty") {
        inner = `<cfRule type="containsBlanks" dxfId="${dxfId}" priority="${priority}"><formula>LEN(TRIM(${topLeft}))=0</formula></cfRule>`;
      } else if (type === "notEmpty") {
        inner = `<cfRule type="notContainsBlanks" dxfId="${dxfId}" priority="${priority}"><formula>LEN(TRIM(${topLeft}))&gt;0</formula></cfRule>`;
      }
    }
    if (inner) {
      blocks.push(`<conditionalFormatting sqref="${sqref}">${inner}</conditionalFormatting>`);
      priority++;
    }
  }
  return blocks.join("");
}

export interface ExportOptions {
  title?: string;
  /** Called after each chunk: parts done of all parts. */
  onProgress?: (done: number, total: number) => void;
  /** Return true to stop (the promise then rejects with "cancelled"). */
  cancelled?: () => boolean;
}

/** Writes the whole workbook (every tab, read chunk by chunk) as an .xlsx and returns it as a Blob. */
export async function exportBigSheet(store: BigSheetStore, options: ExportOptions = {}): Promise<Blob> {
  const parts: Uint8Array[] = [];
  let failure: Error | null = null;
  const zip = new Zip((error, data) => {
    if (error) failure = error;
    else parts.push(data);
  });
  const encoder = new TextEncoder();
  const files: { name: string; push: (text: string, final?: boolean) => void }[] = [];
  const open = (name: string) => {
    const file = new ZipDeflate(name, { level: 4 });
    zip.add(file);
    const entry = { name, push: (text: string, final = false) => file.push(encoder.encode(text), final) };
    files.push(entry);
    return entry;
  };
  const whole = (name: string, text: string) => open(name).push(text, true);

  const styles = new StyleTable();
  const names = sheetNames(store.tabs.map((tab) => tab.name));
  const total = store.tabs.reduce((sum, tab) => sum + Math.max(1, tab.chunks.length), 0);
  let done = 0;
  const hiddenTab = (i: number) => !!store.tabMeta(i).hidden;

  for (let t = 0; t < store.tabs.length; t++) {
    const info = store.tabs[t];
    const layout: TabLayout = info.layout;
    const sheet = open(`xl/worksheets/sheet${t + 1}.xml`);
    const links: { ref: string; url: string }[] = [];
    const validations: string[] = [];
    const frozen = layout.frozen;
    const merged = layout.merges ?? [];
    const hiddenRows = new Set([...(layout.hiddenRows ?? []), ...(layout.filter?.hidden ?? [])]);
    const hiddenCols = new Set(layout.hiddenColumns ?? []);
    const heights = layout.rowHeights ?? {};
    const borders = (layout.borders ?? []) as Border[];

    sheet.push(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
        (store.tabMeta(t).color && argb(store.tabMeta(t).color) ? `<sheetPr><tabColor rgb="${argb(store.tabMeta(t).color)}"/></sheetPr>` : "") +
        `<sheetViews><sheetView workbookViewId="0"${t === 0 ? ' tabSelected="1"' : ""}${layout.zoom && layout.zoom !== 1 ? ` zoomScale="${Math.round(layout.zoom * 100)}"` : ""}>` +
        (frozen && (frozen.rows || frozen.columns)
          ? `<pane${frozen.columns ? ` xSplit="${frozen.columns}"` : ""}${frozen.rows ? ` ySplit="${frozen.rows}"` : ""} topLeftCell="${columnName(frozen.columns)}${frozen.rows + 1}" activePane="${frozen.rows && frozen.columns ? "bottomRight" : frozen.rows ? "bottomLeft" : "topRight"}" state="frozen"/>`
          : "") +
        `</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>`,
    );
    const cols = new Set<number>([...Object.keys(layout.columnWidths ?? {}).map(Number), ...hiddenCols]);
    if (cols.size > 0) {
      const xml = [...cols]
        .sort((a, b) => a - b)
        .map((c) => `<col min="${c + 1}" max="${c + 1}" width="${Math.max(0.5, ((layout.columnWidths?.[c] ?? 116) / 7.5)).toFixed(2)}" customWidth="1"${hiddenCols.has(c) ? ' hidden="1"' : ""}/>`)
        .join("");
      sheet.push(`<cols>${xml}</cols>`);
    }
    sheet.push("<sheetData>");

    const writeRow = (row: number, cells: [number, WorkbookCell | undefined, number][]) => {
      const attrs = `${heights[row] ? ` ht="${(heights[row] * 0.75).toFixed(2)}" customHeight="1"` : ""}${hiddenRows.has(row) ? ' hidden="1"' : ""}`;
      const xml = cells
        .sort((a, b) => a[0] - b[0])
        .map(([col, cell, style]) => cellXml(`${columnName(col)}${row + 1}`, cell, style))
        .filter((x): x is string => x !== null)
        .join("");
      if (!xml && !attrs) return;
      sheet.push(`<row r="${row + 1}"${attrs}>${xml}</row>`);
    };

    const rowsWithSize = [...new Set([...Object.keys(heights).map(Number), ...hiddenRows])].sort((a, b) => a - b);
    let sizePos = 0;
    for (let c = 0; c < Math.max(1, info.chunks.length); c++) {
      if (options.cancelled?.()) throw new Error("cancelled");
      if (failure) throw failure;
      if (info.chunks.length > 0) await store.load(t, c);
      const entry = info.chunks[c];
      const start = entry?.rowStart ?? 0;
      const end = c === info.chunks.length - 1 || !entry ? Math.max(info.rows, start + (entry?.rowCount ?? 0)) : start + entry.rowCount;
      const byRow = new Map<number, Map<number, WorkbookCell>>();
      for (const [row, col, cell] of store.cellsOfChunk(t, c)) {
        let line = byRow.get(row);
        if (!line) byRow.set(row, (line = new Map()));
        line.set(col, cell);
        if (cell.link) links.push({ ref: `${columnName(col)}${row + 1}`, url: cell.link.url });
        if (cell.validation && validations.length < 5000) {
          const xml = validationXml(`${columnName(col)}${row + 1}`, cell.validation);
          if (xml) validations.push(xml);
        }
      }
      // Cells that only have borders still need a place in the file.
      let edges: ReturnType<typeof edgesIn> | null = null;
      if (borders.length > 0 && end > start) {
        edges = edgesIn(borders, start, 0, end - 1, Math.max(info.cols - 1, 0));
        for (const key of [...edges.h.keys(), ...edges.v.keys()]) {
          const row = Math.floor(key / 32768);
          const col = key % 32768;
          for (const r of [row, row - 1]) {
            if (r >= start && r < end && !byRow.get(r)?.has(col)) {
              const line = byRow.get(r) ?? new Map<number, WorkbookCell>();
              byRow.set(r, line);
              if (!line.has(col)) line.set(col, {});
            }
          }
          if (col > 0 && !byRow.get(row)?.has(col - 1) && row >= start && row < end) {
            const line = byRow.get(row) ?? new Map<number, WorkbookCell>();
            byRow.set(row, line);
            if (!line.has(col - 1)) line.set(col - 1, {});
          }
        }
      }
      const sidesOf = (row: number, col: number): Sides | undefined => {
        if (!edges) return undefined;
        const pick = (map: Map<number, { style: string; color: string }>, r: number, c: number) => map.get(edgeKey(r, c));
        const sides: Sides = {};
        const top = pick(edges.h, row, col);
        const bottom = pick(edges.h, row + 1, col);
        const left = pick(edges.v, row, col);
        const right = pick(edges.v, row, col + 1);
        if (top) sides.top = top;
        if (bottom) sides.bottom = bottom;
        if (left) sides.left = left;
        if (right) sides.right = right;
        return sides.top || sides.bottom || sides.left || sides.right ? sides : undefined;
      };
      // Rows with a height or hidden flag but no cells are written too, in order with the others.
      const rows = new Set(byRow.keys());
      while (sizePos < rowsWithSize.length && rowsWithSize[sizePos] < end) {
        if (rowsWithSize[sizePos] >= start) rows.add(rowsWithSize[sizePos]);
        sizePos++;
      }
      for (const row of [...rows].sort((a, b) => a - b)) {
        const line = byRow.get(row) ?? new Map<number, WorkbookCell>();
        writeRow(
          row,
          [...line].map(([col, cell]) => [col, Object.keys(cell).length === 0 ? undefined : cell, styles.id(cell.style, sidesOf(row, col))]),
        );
      }
      options.onProgress?.(++done, total);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    sheet.push("</sheetData>");
    if (merged.length > 0) {
      sheet.push(
        `<mergeCells count="${merged.length}">${merged.map((m) => `<mergeCell ref="${columnName(m.c)}${m.r + 1}:${columnName(m.c + m.cs - 1)}${m.r + m.rs}"/>`).join("")}</mergeCells>`,
      );
    }
    if (layout.conditionalFormats?.length) sheet.push(conditionalXml(layout.conditionalFormats, styles));
    if (validations.length > 0) sheet.push(`<dataValidations count="${validations.length}">${validations.join("")}</dataValidations>`);
    if (links.length > 0) {
      sheet.push(`<hyperlinks>${links.map((link, i) => `<hyperlink ref="${link.ref}" r:id="rId${i + 1}"/>`).join("")}</hyperlinks>`);
      whole(
        `xl/worksheets/_rels/sheet${t + 1}.xml.rels`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${links
          .map((link, i) => `<Relationship Id="rId${i + 1}" Type="${NS_REL}/hyperlink" Target="${esc(link.url)}" TargetMode="External"/>`)
          .join("")}</Relationships>`,
      );
    }
    sheet.push(`<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`, true);
  }

  whole(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${store.tabs
      .map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
      .join("")}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  );
  whole(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  whole(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><bookViews><workbookView/></bookViews><sheets>${store.tabs
      .map((_, i) => `<sheet name="${esc(names[i])}" sheetId="${i + 1}"${hiddenTab(i) ? ' state="hidden"' : ""} r:id="rId${i + 1}"/>`)
      .join("")}</sheets></workbook>`,
  );
  whole(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${store.tabs
      .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
      .join("")}<Relationship Id="rId${store.tabs.length + 1}" Type="${NS_REL}/styles" Target="styles.xml"/></Relationships>`,
  );
  whole("xl/styles.xml", styles.xml());
  zip.end();
  if (failure) throw failure;
  void strToU8;
  return new Blob(parts as BlobPart[], { type: MIME });
}
