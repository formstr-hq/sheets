// The grid itself: a canvas that draws the cells on screen, a scroll area that gives it a scrollbar,
// and the mouse/keyboard handling that drives the selection. Cell contents come from a lazy data
// source, so the same grid shows 10 cells or 25 million.
import {
  advance,
  click,
  isSelected,
  move,
  selectAll,
  single,
  wholeCol,
  wholeRow,
  type Cell,
  type Direction,
  type Range,
  type Selection,
} from "./selection";
import { edgeKey, edgesIn, type Border, type EdgeStyle } from "./borders";
import { MergeIndex, type Merge } from "./merges";
import { SizeIndex, type SizeOverride } from "./sizes";
import { scrollExtent, toScrollPx, toSheetPx, visible } from "./viewport";

export interface GridCellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  background?: string;
  align?: "left" | "center" | "right";
  fontSize?: number;
  fontFamily?: string;
  wrap?: boolean;
  verticalAlign?: "top" | "middle" | "bottom";
  /** A data bar behind the text: how much of the cell's width it fills (0 to 1) and its colour. */
  bar?: { fraction: number; color: string };
  /** FortuneSheet codes: 0 none, 1 up 45°, 2 down 45°, 3 stacked letters, 4 up 90°, 5 down 90°. */
  rotation?: number;
}

export interface GridCell {
  text: string;
  style?: GridCellStyle;
  /** The cell is a hyperlink (drawn blue and underlined). */
  link?: boolean;
  /** A tick box: true = ticked. The cell's text is not drawn. */
  checkbox?: boolean;
  /** Drawn as a rounded label instead of plain text: a person (with an initial) or a tag. */
  pill?: { label: string; bg: string; fg: string; avatar: boolean };
  /** The cell holds a number (shown to the right unless aligned otherwise, whatever its format looks like). */
  numeric?: boolean;
  /** The cell has a comment (a small mark in its corner). */
  note?: boolean;
  /** The cell holds something its validation rule does not allow (a red mark in the top-left corner). */
  invalid?: boolean;
}

/** What the grid reads. `cell` returns undefined while the cell's data is still loading. */
export interface GridData {
  rowCount: number;
  colCount: number;
  cell(row: number, col: number): GridCell | null | undefined;
  /** Asks for the rows on screen to be loaded; called on every scroll. */
  ensure?(firstRow: number, endRow: number): void;
  /** Calls back when cells changed or finished loading. Returns an unsubscribe function. */
  subscribe?(listener: () => void): () => void;
}

export interface GridTheme {
  font: string;
  background: string;
  text: string;
  line: string;
  headBackground: string;
  headText: string;
  headSelected: string;
  selection: string;
  selectionFill: string;
  loading: string;
}

export const LIGHT_THEME: GridTheme = {
  font: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
  background: "#ffffff",
  text: "#1f2328",
  line: "#e1e4e8",
  headBackground: "#f6f8fa",
  headText: "#57606a",
  headSelected: "#dbeafe",
  selection: "#2563eb",
  selectionFill: "rgba(37, 99, 235, 0.12)",
  loading: "#f3f4f6",
};

export interface GridOptions {
  data: GridData;
  rowHeight?: number;
  colWidth?: number;
  headerWidth?: number;
  headerHeight?: number;
  theme?: GridTheme;
  /** Called when the user starts editing (Enter/F2/double-click/typing). `text` is set when typing replaces the cell. */
  onEdit?(cell: Cell, text: string | undefined): void;
  onSelect?(selection: Selection): void;
  /** Delete/Backspace on the selection. */
  onClear?(selection: Selection): void;
  /**
   * Ctrl+C / Ctrl+X on the last selected range. Return the text for the clipboard (or a promise of it
   * when cells still have to load); return undefined to do nothing. `cut` is true for Ctrl+X.
   */
  onCopy?(range: Range, cut: boolean): string | Promise<string | undefined> | undefined;
  /** Ctrl+V: the clipboard's text and the selection it goes into. */
  onPaste?(text: string, selection: Selection): void;
  /** A tick box in a cell was clicked: flip it. */
  onToggle?(cell: Cell): void;
  /** The fill handle (small square at the selection's corner) was dragged: fill `source` out to `extended` (which contains it). */
  onFillDrag?(source: Range, extended: Range): void;
  /** The selection's border was dragged: move the cells of `source` so that its top-left lands at `target`'s. */
  onMoveDrag?(source: Range, target: Range): void;
  /** Ctrl+D (down) / Ctrl+R (right). */
  onFill?(direction: "down" | "right", selection: Selection): void;
  /** Cells to look at for Ctrl+arrow (does the cell hold anything?). */
  hasData?(row: number, col: number): boolean;
  onResize?(kind: "row" | "col", index: number, size: number): void;
  /**
   * A click on a cell while a formula is being typed. Return true when it was used (to insert the
   * cell's address): the grid then leaves the selection and focus alone.
   */
  onReference?(cell: Cell): boolean;
  /** Ctrl+wheel or Ctrl +/-/0 changed the zoom (1 = 100%). */
  onZoom?(zoom: number): void;
  /** Right-click: where (page coordinates) and on which row/column or header. The selection has already been adjusted. */
  onContextMenu?(info: { x: number; y: number; row: number; col: number; onRowHead: boolean; onColHead: boolean }): void;
}

export function columnName(index: number): string {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

const MIN_SIZE = 8;

/** A run of rows (or columns) drawn together: their pixel positions and the screen strip they may paint in. */
interface Band {
  first: number;
  end: number;
  /** Screen position (px) of each index from `first`. */
  pos: number[];
  start: number;
  stop: number;
}

export class Grid {
  readonly rows: SizeIndex;
  readonly cols: SizeIndex;
  selection: Selection | null = null;
  frozenRows = 0;
  frozenCols = 0;
  /** 1 = 100%. Everything is laid out in unzoomed px and drawn scaled. */
  zoom = 1;
  readonly merges = new MergeIndex();
  private borders: readonly Border[] = [];

  private readonly options: GridOptions;
  private readonly theme: GridTheme;
  private headW: number;
  private readonly baseHeadW: number;
  private readonly headH: number;
  private readonly root: HTMLDivElement;
  private readonly scroller: HTMLDivElement;
  private readonly spacer: HTMLDivElement;
  private readonly canvas: HTMLCanvasElement;
  private readonly ctx: CanvasRenderingContext2D;
  private width = 0;
  private height = 0;
  private frame = 0;
  private data: GridData;
  private unsubscribe?: () => void;
  private observer?: ResizeObserver;
  /** While the fill handle is dragged: the range it would fill. */
  private fillPreview: Range | null = null;
  private drag: { kind: "cells" | "row" | "col"; from: number } | { kind: "fill"; source: Range } | { kind: "move"; source: Range; grab: Cell } | { kind: "resize"; axis: "row" | "col"; index: number; start: number; size: number } | null = null;

  constructor(parent: HTMLElement, options: GridOptions) {
    this.options = options;
    this.data = options.data;
    this.theme = options.theme ?? LIGHT_THEME;
    this.baseHeadW = options.headerWidth ?? 46;
    this.headW = this.baseHeadW;
    this.headH = options.headerHeight ?? 28;
    this.rows = new SizeIndex(options.data.rowCount, options.rowHeight ?? 30);
    this.cols = new SizeIndex(options.data.colCount, options.colWidth ?? 116);

    this.root = document.createElement("div");
    this.root.className = "sheets-grid";
    this.root.tabIndex = 0;
    this.root.style.cssText = "position:relative;width:100%;height:100%;overflow:hidden;outline:none;";
    this.scroller = document.createElement("div");
    this.scroller.style.cssText = "position:absolute;inset:0;overflow:auto;";
    this.spacer = document.createElement("div");
    this.spacer.style.cssText = "position:absolute;top:0;left:0;width:1px;height:1px;";
    this.canvas = document.createElement("canvas");
    this.canvas.style.cssText = "position:absolute;top:0;left:0;pointer-events:none;";
    this.scroller.append(this.spacer);
    this.root.append(this.canvas, this.scroller);
    parent.append(this.root);
    const ctx = this.canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas is not available in this browser.");
    this.ctx = ctx;

    this.scroller.addEventListener("scroll", this.schedule);
    this.scroller.addEventListener("pointerdown", this.onPointerDown);
    this.scroller.addEventListener("pointermove", this.onPointerMove);
    this.scroller.addEventListener("pointerup", this.onPointerUp);
    this.scroller.addEventListener("dblclick", this.onDoubleClick);
    this.scroller.addEventListener("contextmenu", this.onContextMenu);
    this.scroller.addEventListener("wheel", this.onWheel, { passive: false });
    this.root.addEventListener("keydown", this.onKeyDown);
    // Firefox sends clipboard events to the page, not to the focused element, so listen on the document.
    document.addEventListener("copy", this.onCopy);
    document.addEventListener("cut", this.onCopy);
    document.addEventListener("paste", this.onPaste);
    this.observer = new ResizeObserver(() => this.resize());
    this.observer.observe(this.root);
    this.attach(this.data);
    this.resize();
  }

  /** Switches to another data source (another tab). Selection and scroll reset. */
  setData(data: GridData): void {
    this.unsubscribe?.();
    this.data = data;
    this.rows.replace([]);
    this.cols.replace([]);
    this.rows.setLength(data.rowCount);
    this.cols.setLength(data.colCount);
    this.frozenRows = 0;
    this.frozenCols = 0;
    this.merges.set([]);
    this.borders = [];
    this.selection = null;
    this.scroller.scrollTop = 0;
    this.scroller.scrollLeft = 0;
    this.attach(data);
    this.updateSpacer();
    this.schedule();
  }

  private attach(data: GridData): void {
    this.unsubscribe = data.subscribe?.(this.schedule);
  }

  destroy(): void {
    this.unsubscribe?.();
    this.observer?.disconnect();
    cancelAnimationFrame(this.frame);
    document.removeEventListener("copy", this.onCopy);
    document.removeEventListener("cut", this.onCopy);
    document.removeEventListener("paste", this.onPaste);
    this.root.remove();
  }

  focus(): void {
    this.root.focus({ preventScroll: true });
  }

  /** Redraws on the next animation frame (several calls in a row draw once). */
  readonly schedule = (): void => {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.draw();
    });
  };

  select(selection: Selection | null, reveal = true): void {
    // A selection never cuts a merged block in half.
    this.selection = selection && this.merges.size > 0 ? this.merges.expandSelection(selection) : selection;
    selection = this.selection;
    if (selection && reveal) this.reveal(selection.active);
    if (selection) this.options.onSelect?.(selection);
    this.schedule();
  }

  /** Scrolls so `cell` is fully on screen (cells in the frozen rows and columns are always on screen). */
  reveal(cell: Cell): void {
    const v = this.view();
    if (cell.row >= v.fr) {
      const win = Math.max(1, v.viewH - v.fh);
      const total = this.rows.total - v.fh;
      const y = this.rows.offset(cell.row) - v.fh;
      const top = v.rowTop - v.fh;
      const size = this.rows.size(cell.row);
      if (y < top) this.scroller.scrollTop = toScrollPx(y, total, win) * this.zoom;
      else if (y + size > top + win) this.scroller.scrollTop = toScrollPx(y + size - win, total, win) * this.zoom;
    }
    if (cell.col >= v.fc) {
      const win = Math.max(1, v.viewW - v.fw);
      const total = this.cols.total - v.fw;
      const x = this.cols.offset(cell.col) - v.fw;
      const left = v.left - v.fw;
      const size = this.cols.size(cell.col);
      if (x < left) this.scroller.scrollLeft = toScrollPx(x, total, win) * this.zoom;
      else if (x + size > left + win) this.scroller.scrollLeft = toScrollPx(x + size - win, total, win) * this.zoom;
    }
  }

  /** Where `cell` is on screen, relative to the grid's top-left corner (for placing an editor). */
  cellRect(cell: Cell): { x: number; y: number; width: number; height: number } {
    const v = this.view();
    const block = this.merges.at(cell.row, cell.col);
    const row = block?.r ?? cell.row;
    const col = block?.c ?? cell.col;
    const z = this.zoom;
    const rect = {
      x: this.headW + (col < v.fc ? this.cols.offset(col) : v.fw + this.cols.offset(col) - v.left),
      y: this.headH + (row < v.fr ? this.rows.offset(row) : v.fh + this.rows.offset(row) - v.rowTop),
      width: block ? this.cols.offset(block.c + block.cs) - this.cols.offset(block.c) : this.cols.size(col),
      height: block ? this.rows.offset(block.r + block.rs) - this.rows.offset(block.r) : this.rows.size(row),
    };
    return { x: rect.x * z, y: rect.y * z, width: rect.width * z, height: rect.height * z };
  }

  scrollToRow(row: number): void {
    const target = Math.max(0, Math.min(this.rows.length - 1, row));
    const v = this.view();
    this.scroller.scrollTop = toScrollPx(Math.max(0, this.rows.offset(target) - v.fh), this.rows.total - v.fh, Math.max(1, v.viewH - v.fh)) * this.zoom;
    this.schedule();
  }

  /** Zooms the sheet (0.25 to 2). Keeps the top-left of what is shown in place. */
  setZoom(zoom: number, notify = false): void {
    const next = Math.round(Math.min(2, Math.max(0.25, zoom)) * 100) / 100;
    if (next === this.zoom) return;
    const top = this.scroller.scrollTop / this.zoom;
    const left = this.scroller.scrollLeft / this.zoom;
    this.zoom = next;
    this.resize();
    this.scroller.scrollTop = top * next;
    this.scroller.scrollLeft = left * next;
    this.schedule();
    if (notify) this.options.onZoom?.(next);
  }

  /** Keeps the first `rows` rows and `cols` columns on screen while the rest scrolls. */
  setFrozen(rows: number, cols: number): void {
    this.frozenRows = Math.max(0, Math.floor(rows));
    this.frozenCols = Math.max(0, Math.floor(cols));
    this.updateSpacer();
    this.schedule();
  }

  /** Cells merged into one (top-left cell, rows and columns spanned). */
  setMerges(merges: readonly Merge[]): void {
    this.merges.set(merges);
    this.schedule();
  }

  /** Borders, as the list of "put this border on that range" requests in the order they were made. */
  setBorders(borders: readonly Border[]): void {
    this.borders = borders;
    this.schedule();
  }

  /** Replaces the row heights / column widths that differ from the default (size 0 hides). Used to apply what was saved with the sheet. */
  setSizes(kind: "row" | "col", overrides: readonly SizeOverride[]): void {
    (kind === "row" ? this.rows : this.cols).replace(overrides);
    this.updateSpacer();
    this.schedule();
  }

  // --- layout -----------------------------------------------------------------------------

  private resize(): void {
    // The canvas covers what the scroll area shows without its scrollbars.
    const screenW = Math.max(1, this.scroller.clientWidth);
    const screenH = Math.max(1, this.scroller.clientHeight);
    // Layout is in unzoomed px: the same screen shows more (or less) of the sheet.
    this.width = screenW / this.zoom;
    this.height = screenH / this.zoom;
    const dpr = window.devicePixelRatio || 1;
    this.canvas.width = Math.floor(screenW * dpr);
    this.canvas.height = Math.floor(screenH * dpr);
    this.canvas.style.width = `${screenW}px`;
    this.canvas.style.height = `${screenH}px`;
    this.ctx.setTransform(dpr * this.zoom, 0, 0, dpr * this.zoom, 0, 0);
    this.updateSpacer();
    this.schedule();
  }

  /** Frozen rows/columns as they apply now (ignored when they would leave no room to scroll). */
  private frozenPx() {
    const viewW = this.width - this.headW;
    const viewH = this.height - this.headH;
    let fr = Math.min(this.frozenRows, this.rows.length);
    let fc = Math.min(this.frozenCols, this.cols.length);
    let fh = this.rows.offset(fr);
    let fw = this.cols.offset(fc);
    if (fh > viewH - 30) {
      fr = 0;
      fh = 0;
    }
    if (fw > viewW - 30) {
      fc = 0;
      fw = 0;
    }
    return { fr, fc, fh, fw };
  }

  /** The scroll area is as long as the part of the sheet that scrolls (capped), offset by the headers and frozen rows. */
  private updateSpacer(): void {
    // Wide enough for the biggest row number ("25,000,000").
    this.headW = Math.max(this.baseHeadW, Math.ceil(this.rows.length.toLocaleString().length * 7 + 16));
    const viewW = this.width - this.headW;
    const viewH = this.height - this.headH;
    const { fh, fw } = this.frozenPx();
    this.spacer.style.height = `${(scrollExtent(this.rows.total - fh, Math.max(1, viewH - fh)) + this.headH + fh) * this.zoom}px`;
    this.spacer.style.width = `${(scrollExtent(this.cols.total - fw, Math.max(1, viewW - fw)) + this.headW + fw) * this.zoom}px`;
  }

  private view() {
    const viewW = this.width - this.headW;
    const viewH = this.height - this.headH;
    const { fr, fc, fh, fw } = this.frozenPx();
    const winH = Math.max(1, viewH - fh);
    const winW = Math.max(1, viewW - fw);
    /** Sheet position shown just below the frozen rows / right of the frozen columns. */
    const rowTop = fh + toSheetPx(this.scroller.scrollTop / this.zoom, this.rows.total - fh, winH);
    const left = fw + toSheetPx(this.scroller.scrollLeft / this.zoom, this.cols.total - fw, winW);
    const rows = visible(this.rows, rowTop, winH);
    const cols = visible(this.cols, left, winW);
    if (rows.first < fr) Object.assign(rows, { first: fr, shift: 0 });
    if (cols.first < fc) Object.assign(cols, { first: fc, shift: 0 });
    return { viewW, viewH, fr, fc, fh, fw, rowTop, left, rows, cols };
  }

  /** The row and column bands on screen: the frozen ones (if any) and the scrolling ones, each with pixel positions. */
  private bands(v: ReturnType<Grid["view"]>): { rowBands: Band[]; colBands: Band[] } {
    const make = (sizes: SizeIndex, first: number, end: number, origin: number, start: number, stop: number): Band => {
      const pos: number[] = [];
      let at = origin;
      for (let i = first; i < end; i++) {
        pos.push(at);
        at += sizes.size(i);
      }
      return { first, end, pos, start, stop };
    };
    const rowBands: Band[] = [];
    const colBands: Band[] = [];
    if (v.fr > 0) rowBands.push(make(this.rows, 0, v.fr, this.headH, this.headH, this.headH + v.fh));
    rowBands.push(make(this.rows, v.rows.first, v.rows.end, this.headH + v.fh - v.rows.shift, this.headH + v.fh, this.headH + v.viewH));
    if (v.fc > 0) colBands.push(make(this.cols, 0, v.fc, this.headW, this.headW, this.headW + v.fw));
    colBands.push(make(this.cols, v.cols.first, v.cols.end, this.headW + v.fw - v.cols.shift, this.headW + v.fw, this.headW + v.viewW));
    return { rowBands, colBands };
  }

  // --- drawing ----------------------------------------------------------------------------

  private draw(): void {
    // The sheet may have grown (an edit past the last row) since the last frame.
    if (this.rows.length !== this.data.rowCount || this.cols.length !== this.data.colCount) {
      this.rows.setLength(this.data.rowCount);
      this.cols.setLength(this.data.colCount);
      this.updateSpacer();
    }
    const { ctx, theme } = this;
    const v = this.view();
    const { rowBands, colBands } = this.bands(v);
    const fonts = (style: GridCellStyle | undefined) =>
      `${style?.italic ? "italic " : ""}${style?.bold ? "700 " : ""}${Math.min(style?.fontSize ?? 11, 24) * (96 / 72)}px ${style?.fontFamily ? `${style.fontFamily}, ` : ""}${theme.font}`;

    if (v.fr > 0) this.data.ensure?.(0, v.fr);
    // Only the rows that are really shown are loaded: a filter can hide thousands of rows in a row.
    let from = -1;
    for (let r = v.rows.first; r < v.rows.end; r++) {
      if (this.rows.isHidden(r)) {
        if (from >= 0) this.data.ensure?.(from, r);
        from = -1;
      } else if (from < 0) from = r;
    }
    if (from >= 0) this.data.ensure?.(from, v.rows.end);

    ctx.clearRect(0, 0, this.width, this.height);
    ctx.fillStyle = theme.background;
    ctx.fillRect(0, 0, this.width, this.height);
    ctx.textBaseline = "middle";

    // Each block (frozen corner, frozen rows, frozen columns, the scrolling part) is clipped to itself.
    for (const rb of rowBands) {
      for (const cb of colBands) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(cb.start, rb.start, cb.stop - cb.start, rb.stop - rb.start);
        ctx.clip();
        this.drawBlock(rb, cb, fonts);
        this.drawSelection(rb, cb);
        ctx.restore();
      }
    }
    // A heavier line marks where the frozen part ends.
    ctx.fillStyle = theme.headText;
    if (v.fr > 0) ctx.fillRect(this.headW, this.headH + v.fh - 1, v.viewW, 2);
    if (v.fc > 0) ctx.fillRect(this.headW + v.fw - 1, this.headH, 2, v.viewH);

    this.drawHeaders(rowBands, colBands);
  }

  /** Cells and grid lines of one block. */
  private drawBlock(rb: Band, cb: Band, fonts: (style: GridCellStyle | undefined) => string): void {
    const { ctx, theme } = this;
    // Merged blocks reaching into this block; their cells are drawn as one after the grid lines.
    const blocks = this.merges.intersecting(rb.first, cb.first, rb.end - 1, cb.end - 1);
    const covered = (r: number, c: number) => blocks.some((m) => r >= m.r && r < m.r + m.rs && c >= m.c && c < m.c + m.cs);
    for (let r = rb.first; r < rb.end; r++) {
      const h = this.rows.size(r);
      if (h === 0) continue;
      const top = rb.pos[r - rb.first];
      for (let c = cb.first; c < cb.end; c++) {
        const w = this.cols.size(c);
        if (w === 0) continue;
        if (blocks.length > 0 && covered(r, c)) continue;
        const left = cb.pos[c - cb.first];
        const cell = this.data.cell(r, c);
        const style = cell?.style;
        if (cell === undefined) {
          ctx.fillStyle = theme.loading;
          ctx.fillRect(left, top, w, h);
        } else if (style?.background) {
          ctx.fillStyle = style.background;
          ctx.fillRect(left, top, w, h);
        }
        if (style?.bar && cell !== undefined) {
          ctx.fillStyle = style.bar.color;
          ctx.fillRect(left + 1, top + 3, Math.max(1, (w - 2) * style.bar.fraction), h - 6);
        }
        if (cell?.checkbox !== undefined) this.drawCheckbox(cell.checkbox, left, top, w, h);
        else if (cell?.pill) this.drawPill(cell, left, top, w, h);
        else if (cell?.text) this.drawText(cell, left, top, w, h, fonts(style));
        if (cell?.note) this.drawNoteMark(left, top, w);
        if (cell?.invalid) this.drawInvalidMark(left, top);
      }
    }
    ctx.strokeStyle = theme.line;
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let i = 0; i < rb.pos.length; i++) {
      const size = this.rows.size(rb.first + i);
      if (size === 0) continue;
      const line = Math.round(rb.pos[i] + size) - 0.5;
      ctx.moveTo(cb.start, line);
      ctx.lineTo(cb.stop, line);
    }
    for (let i = 0; i < cb.pos.length; i++) {
      const size = this.cols.size(cb.first + i);
      if (size === 0) continue;
      const line = Math.round(cb.pos[i] + size) - 0.5;
      ctx.moveTo(line, rb.start);
      ctx.lineTo(line, rb.stop);
    }
    ctx.stroke();
    for (const m of blocks) {
      const left = cb.pos[0] + this.cols.offset(m.c) - this.cols.offset(cb.first);
      const top = rb.pos[0] + this.rows.offset(m.r) - this.rows.offset(rb.first);
      const w = this.cols.offset(m.c + m.cs) - this.cols.offset(m.c);
      const h = this.rows.offset(m.r + m.rs) - this.rows.offset(m.r);
      const cell = this.data.cell(m.r, m.c);
      ctx.fillStyle = cell?.style?.background ?? (cell === undefined ? theme.loading : theme.background);
      ctx.fillRect(left, top, w, h);
      if (cell?.text) this.drawText(cell, left, top, w, h, fonts(cell.style));
      if (cell?.note) this.drawNoteMark(left, top, w);
      ctx.strokeStyle = theme.line;
      ctx.lineWidth = 1;
      ctx.strokeRect(Math.round(left) - 0.5, Math.round(top) - 0.5, Math.round(w) + 1, Math.round(h) + 1);
    }
    if (this.borders.length > 0) this.drawBorders(rb, cb, blocks);
  }

  /** The borders in one block, drawn over the grid lines. */
  private drawBorders(rb: Band, cb: Band, blocks: readonly Merge[]): void {
    const { ctx } = this;
    const edges = edgesIn(this.borders, rb.first, cb.first, rb.end - 1, cb.end - 1);
    if (edges.h.size === 0 && edges.v.size === 0) return;
    const lineAt = (map: Map<number, EdgeStyle>, row: number, col: number, x1: number, y1: number, x2: number, y2: number) => {
      const edge = map.get(edgeKey(row, col));
      if (!edge) return;
      // Lines inside a merged block are not drawn.
      const horizontal = map === edges.h;
      if (blocks.some((m) => (horizontal ? row > m.r && row < m.r + m.rs && col >= m.c && col < m.c + m.cs : col > m.c && col < m.c + m.cs && row >= m.r && row < m.r + m.rs))) return;
      ctx.strokeStyle = edge.color;
      ctx.lineWidth = edge.style === "thick" ? 3 : edge.style === "medium" ? 2 : 1;
      ctx.setLineDash(edge.style === "dashed" ? [5, 3] : edge.style === "dotted" ? [1.5, 2.5] : []);
      ctx.beginPath();
      ctx.moveTo(x1, y1);
      ctx.lineTo(x2, y2);
      ctx.stroke();
    };
    const rowY = (r: number) => Math.round(rb.pos[0] + this.rows.offset(r) - this.rows.offset(rb.first)) - 0.5;
    const colX = (c: number) => Math.round(cb.pos[0] + this.cols.offset(c) - this.cols.offset(cb.first)) - 0.5;
    for (let r = rb.first; r <= rb.end; r++) {
      for (let c = cb.first; c < cb.end; c++) {
        if (this.cols.size(c) === 0) continue;
        lineAt(edges.h, r, c, colX(c) + 0.5, rowY(r), colX(c + 1) + 0.5, rowY(r));
      }
    }
    for (let c = cb.first; c <= cb.end; c++) {
      for (let r = rb.first; r < rb.end; r++) {
        if (this.rows.size(r) === 0) continue;
        lineAt(edges.v, r, c, colX(c), rowY(r) + 0.5, colX(c), rowY(r + 1) + 0.5);
      }
    }
    ctx.setLineDash([]);
  }

  private drawText(cell: GridCell, left: number, top: number, w: number, h: number, font: string): void {
    const { ctx, theme } = this;
    const style = cell.style;
    ctx.save();
    ctx.beginPath();
    ctx.rect(left, top, w, h);
    ctx.clip();
    ctx.font = font;
    ctx.fillStyle = style?.color ?? (cell.link ? "#1a56db" : theme.text);
    const rotation = style?.rotation ?? 0;
    if (rotation >= 1 && rotation <= 5) {
      this.drawRotated(cell, left, top, w, h, rotation);
      ctx.restore();
      return;
    }
    const align = style?.align ?? ((cell.numeric ?? /^[-+]?[\d.,]+%?$/.test(cell.text)) ? "right" : "left");
    ctx.textAlign = align;
    const pad = 4;
    const tx = align === "left" ? left + pad : align === "right" ? left + w - pad : left + w / 2;
    const lines = style?.wrap ? wrapText(ctx, cell.text, w - pad * 2) : [cell.text.split("\n")[0]];
    const lineH = Math.min(style?.fontSize ?? 11, 24) * (96 / 72) * 1.3;
    const block = lines.length * lineH;
    const valign = style?.verticalAlign;
    const startY =
      valign === "top"
        ? top + lineH / 2 + 2
        : valign === "bottom"
          ? top + h - block + lineH / 2 - 2
          : style?.wrap
            ? top + Math.max(lineH / 2 + 2, (h - block) / 2 + lineH / 2)
            : top + h / 2;
    lines.forEach((line, i) => {
      const ty = startY + i * lineH;
      ctx.fillText(line, tx, ty);
      const width = ctx.measureText(line).width;
      const from = align === "left" ? tx : align === "right" ? tx - width : tx - width / 2;
      if (style?.underline || cell.link) ctx.fillRect(from, ty + lineH / 2 - 2, width, 1);
      if (style?.strike) ctx.fillRect(from, ty, width, 1);
    });
    ctx.restore();
  }

  /** A 14px box, with a tick when checked, centred in the cell. */
  private drawCheckbox(checked: boolean, left: number, top: number, w: number, h: number): void {
    const { ctx } = this;
    const size = Math.min(14, h - 6);
    const x = Math.round(left + w / 2 - size / 2) + 0.5;
    const y = Math.round(top + h / 2 - size / 2) + 0.5;
    ctx.save();
    ctx.lineWidth = 1.5;
    ctx.strokeStyle = checked ? this.theme.selection : "#6b7280";
    ctx.fillStyle = checked ? this.theme.selection : this.theme.background;
    ctx.beginPath();
    ctx.rect(x, y, size, size);
    ctx.fill();
    ctx.stroke();
    if (checked) {
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(x + size * 0.22, y + size * 0.52);
      ctx.lineTo(x + size * 0.43, y + size * 0.73);
      ctx.lineTo(x + size * 0.78, y + size * 0.28);
      ctx.stroke();
    }
    ctx.restore();
  }

  /** A person or tag shown as a rounded label (with a round initial for a person). */
  private drawPill(cell: GridCell, left: number, top: number, w: number, h: number): void {
    const { ctx } = this;
    const pill = cell.pill;
    if (!pill || w < 14 || h < 10) return;
    ctx.save();
    ctx.beginPath();
    ctx.rect(left + 1, top + 1, w - 2, h - 2);
    ctx.clip();
    const font = "500 11px Inter, 'Segoe UI', sans-serif";
    ctx.font = font;
    const padX = 7;
    const avatar = pill.avatar ? 14 : 0;
    const gap = pill.avatar ? 5 : 0;
    const room = Math.max(8, w - 4 - padX * 2 - avatar - gap - 4);
    let text = pill.label;
    while (text.length > 1 && ctx.measureText(text).width > room) text = text.slice(0, -1);
    if (text !== pill.label) text += "…";
    const height = Math.min(h - 6, 20);
    const width = Math.min(w - 4, Math.ceil(ctx.measureText(text).width + padX * 2 + avatar + gap));
    const x = left + 2;
    const y = top + (h - height) / 2;
    const r = height / 2;
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + width, y, x + width, y + height, r);
    ctx.arcTo(x + width, y + height, x, y + height, r);
    ctx.arcTo(x, y + height, x, y, r);
    ctx.arcTo(x, y, x + width, y, r);
    ctx.closePath();
    ctx.fillStyle = pill.bg;
    ctx.fill();
    let textX = x + padX;
    ctx.textBaseline = "middle";
    ctx.textAlign = "left";
    if (pill.avatar) {
      const cx = textX + avatar / 2;
      ctx.beginPath();
      ctx.arc(cx, y + height / 2, avatar / 2, 0, Math.PI * 2);
      ctx.fillStyle = pill.fg;
      ctx.fill();
      ctx.fillStyle = pill.bg;
      ctx.font = "700 8px Inter, 'Segoe UI', sans-serif";
      ctx.textAlign = "center";
      ctx.fillText((pill.label.trim()[0] ?? "?").toUpperCase(), cx, y + height / 2 + 0.5);
      ctx.textAlign = "left";
      textX += avatar + gap;
    }
    ctx.fillStyle = pill.fg;
    ctx.font = font;
    ctx.fillText(text, textX, y + height / 2 + 0.5);
    ctx.restore();
  }

  /** The small red triangle in the top-left corner of a cell whose content breaks its validation rule. */
  private drawInvalidMark(left: number, top: number): void {
    const { ctx } = this;
    ctx.fillStyle = "#e5484d";
    ctx.beginPath();
    ctx.moveTo(left + 1, top + 1);
    ctx.lineTo(left + 8, top + 1);
    ctx.lineTo(left + 1, top + 8);
    ctx.fill();
  }

  /** The little red corner that marks a cell with a comment. */
  private drawNoteMark(left: number, top: number, w: number): void {
    const { ctx } = this;
    ctx.fillStyle = "#e5484d";
    ctx.beginPath();
    ctx.moveTo(left + w - 8, top + 1);
    ctx.lineTo(left + w - 1, top + 1);
    ctx.lineTo(left + w - 1, top + 8);
    ctx.fill();
  }

  /** Text turned 45° or 90°, or stacked letter by letter, inside the cell (the caller has clipped to the cell). */
  private drawRotated(cell: GridCell, left: number, top: number, w: number, h: number, rotation: number): void {
    const { ctx } = this;
    const style = cell.style;
    ctx.textAlign = "center";
    const lineH = Math.min(style?.fontSize ?? 11, 24) * (96 / 72) * 1.2;
    const text = cell.text.split("\n")[0];
    if (rotation === 3) {
      const letters = [...text];
      const total = letters.length * lineH;
      let y = top + Math.max(lineH / 2, (h - total) / 2 + lineH / 2);
      for (const letter of letters) {
        if (y > top + h) break;
        ctx.fillText(letter, left + w / 2, y);
        y += lineH;
      }
      return;
    }
    const angle = rotation === 1 ? -Math.PI / 4 : rotation === 2 ? Math.PI / 4 : rotation === 4 ? -Math.PI / 2 : Math.PI / 2;
    ctx.translate(left + w / 2, top + h / 2);
    ctx.rotate(angle);
    ctx.fillText(text, 0, 0);
  }

  private drawSelection(rb: Band, cb: Band): void {
    const sel = this.selection;
    if (!sel) return;
    const { ctx, theme } = this;
    const rectOf = (r1: number, c1: number, r2: number, c2: number) => {
      // Rows/columns outside the block are clamped to its edge so a huge range still draws.
      const ra = Math.max(r1, rb.first);
      const rz = Math.min(r2, rb.end - 1);
      const ca = Math.max(c1, cb.first);
      const cz = Math.min(c2, cb.end - 1);
      if (ra > rz || ca > cz) return null;
      const top = rb.pos[ra - rb.first];
      const left = cb.pos[ca - cb.first];
      return {
        left,
        top,
        width: cb.pos[cz - cb.first] + this.cols.size(cz) - left,
        height: rb.pos[rz - rb.first] + this.rows.size(rz) - top,
      };
    };
    for (const range of sel.ranges) {
      const rect = rectOf(range.r1, range.c1, range.r2, range.c2);
      if (!rect) continue;
      ctx.fillStyle = theme.selectionFill;
      ctx.fillRect(rect.left, rect.top, rect.width, rect.height);
      ctx.strokeStyle = theme.selection;
      ctx.lineWidth = 2;
      ctx.strokeRect(rect.left + 1, rect.top + 1, rect.width - 2, rect.height - 2);
    }
    // The fill handle sits at the corner of the last range; while dragging, the area it would fill is outlined.
    const last = sel.ranges[sel.ranges.length - 1];
    const corner = rectOf(last.r2, last.c2, last.r2, last.c2);
    if (corner && this.options.onFillDrag) {
      ctx.fillStyle = theme.selection;
      ctx.fillRect(corner.left + corner.width - 4, corner.top + corner.height - 4, 7, 7);
    }
    if (this.fillPreview) {
      const box = rectOf(this.fillPreview.r1, this.fillPreview.c1, this.fillPreview.r2, this.fillPreview.c2);
      if (box) {
        ctx.setLineDash([4, 3]);
        ctx.strokeStyle = theme.selection;
        ctx.lineWidth = 2;
        ctx.strokeRect(box.left + 1, box.top + 1, box.width - 2, box.height - 2);
        ctx.setLineDash([]);
      }
    }
    // The active cell is drawn without the tint so it reads as "where typing goes".
    const block = this.merges.at(sel.active.row, sel.active.col);
    const active = block
      ? rectOf(block.r, block.c, block.r + block.rs - 1, block.c + block.cs - 1)
      : rectOf(sel.active.row, sel.active.col, sel.active.row, sel.active.col);
    if (active) {
      ctx.fillStyle = theme.background;
      const cell = this.data.cell(sel.active.row, sel.active.col);
      ctx.fillRect(active.left + 2, active.top + 2, active.width - 4, active.height - 4);
      if (cell?.checkbox !== undefined) this.drawCheckbox(cell.checkbox, active.left, active.top, active.width, active.height);
      else if (cell?.pill) this.drawPill(cell, active.left, active.top, active.width, active.height);
      else if (cell?.text) {
        this.drawText(cell, active.left, active.top, active.width, active.height, `${cell.style?.bold ? "700 " : ""}${Math.min(cell.style?.fontSize ?? 11, 24) * (96 / 72)}px ${theme.font}`);
      }
    }
  }

  private drawHeaders(rowBands: Band[], colBands: Band[]): void {
    const { ctx, theme } = this;
    const sel = this.selection;
    const rowSelected = (r: number) => !!sel?.ranges.some((g) => r >= g.r1 && r <= g.r2);
    const colSelected = (c: number) => !!sel?.ranges.some((g) => c >= g.c1 && c <= g.c2);
    ctx.font = `12px ${theme.font}`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    for (const band of rowBands) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, band.start, this.headW, band.stop - band.start);
      ctx.clip();
      for (let r = band.first; r < band.end; r++) {
        const h = this.rows.size(r);
        if (h === 0) continue;
        const top = band.pos[r - band.first];
        ctx.fillStyle = rowSelected(r) ? theme.headSelected : theme.headBackground;
        ctx.fillRect(0, top, this.headW, h);
        ctx.fillStyle = theme.headText;
        ctx.fillText((r + 1).toLocaleString(), this.headW / 2, top + h / 2);
        ctx.fillStyle = theme.line;
        ctx.fillRect(0, Math.round(top + h) - 1, this.headW, 1);
      }
      ctx.restore();
    }

    for (const band of colBands) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(band.start, 0, band.stop - band.start, this.headH);
      ctx.clip();
      for (let c = band.first; c < band.end; c++) {
        const w = this.cols.size(c);
        if (w === 0) continue;
        const left = band.pos[c - band.first];
        ctx.fillStyle = colSelected(c) ? theme.headSelected : theme.headBackground;
        ctx.fillRect(left, 0, w, this.headH);
        ctx.fillStyle = theme.headText;
        ctx.fillText(columnName(c), left + w / 2, this.headH / 2);
        ctx.fillStyle = theme.line;
        ctx.fillRect(Math.round(left + w) - 1, 0, 1, this.headH);
      }
      ctx.restore();
    }

    ctx.fillStyle = theme.headBackground;
    ctx.fillRect(0, 0, this.headW, this.headH);
    ctx.fillStyle = theme.line;
    ctx.fillRect(0, this.headH - 1, this.width, 1);
    ctx.fillRect(this.headW - 1, 0, 1, this.height);
  }

  // --- mouse ------------------------------------------------------------------------------

  /** Row/column under a point (in the grid's own coordinates), or the header zone it is in. */
  private hit(clientX: number, clientY: number) {
    const rect = this.root.getBoundingClientRect();
    const px = (clientX - rect.left) / this.zoom;
    const py = (clientY - rect.top) / this.zoom;
    const v = this.view();
    const relY = Math.max(0, py - this.headH);
    const relX = Math.max(0, px - this.headW);
    // Sheet position under the pointer: frozen rows/columns map 1:1, the rest follows the scroll.
    const sheetY = relY < v.fh ? relY : v.rowTop + (relY - v.fh);
    const sheetX = relX < v.fw ? relX : v.left + (relX - v.fw);
    const onColHead = py < this.headH;
    const onRowHead = px < this.headW;
    const row = this.rows.indexAt(sheetY);
    const col = this.cols.indexAt(sheetX);
    return { px, py, row, col, onColHead, onRowHead, sheetX, sheetY };
  }

  /** The edge of a row/column header the pointer is on (within 4px), for resizing. */
  private edge(hit: ReturnType<Grid["hit"]>): { axis: "row" | "col"; index: number } | null {
    if (hit.onColHead && !hit.onRowHead) {
      const x = hit.sheetX;
      const end = this.cols.offset(hit.col) + this.cols.size(hit.col);
      const startEdge = this.cols.offset(hit.col);
      if (Math.abs(x - end) <= 4) return { axis: "col", index: hit.col };
      if (Math.abs(x - startEdge) <= 4 && hit.col > 0) return { axis: "col", index: this.cols.nextVisible(hit.col - 1, -1) };
    }
    if (hit.onRowHead && !hit.onColHead) {
      const y = hit.sheetY;
      const end = this.rows.offset(hit.row) + this.rows.size(hit.row);
      const startEdge = this.rows.offset(hit.row);
      if (Math.abs(y - end) <= 4) return { axis: "row", index: hit.row };
      if (Math.abs(y - startEdge) <= 4 && hit.row > 0) return { axis: "row", index: this.rows.nextVisible(hit.row - 1, -1) };
    }
    return null;
  }

  /** The selected range when the point (grid px, unzoomed) is on its border (not at the fill handle), for dragging it. */
  private borderHit(px: number, py: number): Range | null {
    const sel = this.selection;
    if (!sel || !this.options.onMoveDrag) return null;
    const range = sel.ranges[sel.ranges.length - 1];
    if (sel.ranges.length > 1 || range.r2 >= this.rows.length || range.c2 >= this.cols.length) return null;
    // A whole row, column or sheet is not moved by its border.
    if ((range.r1 === 0 && range.r2 >= this.rows.length - 1) || (range.c1 === 0 && range.c2 >= this.cols.length - 1)) return null;
    const z = this.zoom;
    const a = this.cellRect({ row: range.r1, col: range.c1 });
    const b = this.cellRect({ row: range.r2, col: range.c2 });
    const x1 = a.x / z;
    const y1 = a.y / z;
    const x2 = (b.x + b.width) / z;
    const y2 = (b.y + b.height) / z;
    const near = 4;
    const inOuter = px >= x1 - near && px <= x2 + near && py >= y1 - near && py <= y2 + near;
    const inInner = px > x1 + near && px < x2 - near && py > y1 + near && py < y2 - near;
    if (!inOuter || inInner) return null;
    // The corner is the fill handle's.
    if (Math.abs(px - x2) <= 7 && Math.abs(py - y2) <= 7) return null;
    return range;
  }

  /** The selected range when the point (grid px, unzoomed) is on its fill handle. */
  private fillHandleHit(px: number, py: number): Range | null {
    const sel = this.selection;
    if (!sel || !this.options.onFillDrag) return null;
    const range = sel.ranges[sel.ranges.length - 1];
    // Only for a range that is on screen with a visible bottom-right cell.
    if (range.r2 >= this.rows.length || range.c2 >= this.cols.length || this.rows.size(range.r2) === 0 || this.cols.size(range.c2) === 0) return null;
    const z = this.zoom;
    const box = this.cellRect({ row: range.r2, col: range.c2 });
    const x = (box.x + box.width) / z;
    const y = (box.y + box.height) / z;
    return Math.abs(px - x) <= 6 && Math.abs(py - y) <= 6 ? range : null;
  }

  private onPointerDown = (event: PointerEvent): void => {
    if (event.button !== 0) return;
    const hit = this.hit(event.clientX, event.clientY);
    // The scrollbars sit inside the scroller; ignore presses on them.
    if (hit.px >= this.width || hit.py >= this.height) return;
    if (!hit.onRowHead && !hit.onColHead && this.options.onToggle) {
      const here = this.data.cell(hit.row, hit.col);
      if (here?.checkbox !== undefined) {
        // The box sits in the middle of the cell; a click on it ticks it (and still selects the cell).
        const box = this.cellRect({ row: hit.row, col: hit.col });
        const mid = (box.x + box.width / 2) / this.zoom;
        const midY = (box.y + box.height / 2) / this.zoom;
        if (Math.abs(hit.px - mid) <= 9 && Math.abs(hit.py - midY) <= 9) this.options.onToggle({ row: hit.row, col: hit.col });
      }
    }
    const border = this.fillHandleHit(hit.px, hit.py) ? null : this.borderHit(hit.px, hit.py);
    if (border) {
      this.focus();
      this.drag = { kind: "move", source: border, grab: { row: hit.row, col: hit.col } };
      this.fillPreview = null;
      this.scroller.setPointerCapture(event.pointerId);
      event.preventDefault();
      return;
    }
    const handle = this.fillHandleHit(hit.px, hit.py);
    if (handle) {
      this.focus();
      this.drag = { kind: "fill", source: handle };
      this.fillPreview = null;
      this.scroller.setPointerCapture(event.pointerId);
      event.preventDefault();
      return;
    }
    if (!hit.onRowHead && !hit.onColHead && this.options.onReference?.({ row: hit.row, col: hit.col })) {
      event.preventDefault();
      return;
    }
    this.focus();
    const mode = event.shiftKey ? "extend" : event.ctrlKey || event.metaKey ? "add" : "replace";
    const bounds = { rows: this.rows, cols: this.cols };
    const resize = this.edge(hit);
    if (resize) {
      const size = resize.axis === "col" ? this.cols.size(resize.index) : this.rows.size(resize.index);
      this.drag = { kind: "resize", axis: resize.axis, index: resize.index, start: resize.axis === "col" ? event.clientX : event.clientY, size };
    } else if (hit.onRowHead && hit.onColHead) {
      this.select(selectAll(bounds), false);
    } else if (hit.onColHead) {
      this.drag = { kind: "col", from: hit.col };
      this.select(wholeCol(bounds, hit.col), false);
    } else if (hit.onRowHead) {
      this.drag = { kind: "row", from: hit.row };
      this.select(wholeRow(bounds, hit.row), false);
    } else {
      this.drag = { kind: "cells", from: 0 };
      this.select(click(this.selection, { row: hit.row, col: hit.col }, mode), false);
    }
    this.scroller.setPointerCapture(event.pointerId);
    event.preventDefault();
  };

  private onPointerMove = (event: PointerEvent): void => {
    const hit = this.hit(event.clientX, event.clientY);
    const drag = this.drag;
    if (!drag) {
      const over = this.edge(hit);
      this.scroller.style.cursor = over ? (over.axis === "col" ? "col-resize" : "row-resize") : this.fillHandleHit(hit.px, hit.py) ? "crosshair" : this.borderHit(hit.px, hit.py) ? "move" : "";
      return;
    }
    const bounds = { rows: this.rows, cols: this.cols };
    if (drag.kind === "move") {
      const dr = Math.max(-drag.source.r1, Math.min(this.rows.length - 1 - drag.source.r2, hit.row - drag.grab.row));
      const dc = Math.max(-drag.source.c1, Math.min(this.cols.length - 1 - drag.source.c2, hit.col - drag.grab.col));
      this.fillPreview = dr === 0 && dc === 0 ? null : { r1: drag.source.r1 + dr, r2: drag.source.r2 + dr, c1: drag.source.c1 + dc, c2: drag.source.c2 + dc };
      this.schedule();
      return;
    }
    if (drag.kind === "fill") {
      const src = drag.source;
      const below = hit.row - src.r2;
      const right = hit.col - src.c2;
      // Extend along the axis the pointer moved furthest on (down or right only).
      if (below > 0 && below >= right) this.fillPreview = { ...src, r2: hit.row };
      else if (right > 0) this.fillPreview = { ...src, c2: hit.col };
      else this.fillPreview = null;
      this.schedule();
      return;
    }
    if (drag.kind === "resize") {
      const delta = ((drag.axis === "col" ? event.clientX : event.clientY) - drag.start) / this.zoom;
      const size = Math.max(MIN_SIZE, Math.round(drag.size + delta));
      (drag.axis === "col" ? this.cols : this.rows).set(drag.index, size);
      this.updateSpacer();
      this.schedule();
    } else if (drag.kind === "col") {
      this.select(wholeCol(bounds, drag.from, hit.col), false);
    } else if (drag.kind === "row") {
      this.select(wholeRow(bounds, drag.from, hit.row), false);
    } else if (this.selection) {
      this.select(click(this.selection, { row: hit.row, col: hit.col }, "extend"), false);
    }
  };

  private onPointerUp = (event: PointerEvent): void => {
    const drag = this.drag;
    this.drag = null;
    if (this.scroller.hasPointerCapture(event.pointerId)) this.scroller.releasePointerCapture(event.pointerId);
    if (drag?.kind === "move") {
      const preview = this.fillPreview;
      this.fillPreview = null;
      this.schedule();
      if (preview) {
        this.options.onMoveDrag?.(drag.source, preview);
        this.select({ ranges: [preview], active: { row: preview.r1, col: preview.c1 }, anchor: { row: preview.r1, col: preview.c1 } }, false);
      }
      return;
    }
    if (drag?.kind === "fill") {
      const preview = this.fillPreview;
      this.fillPreview = null;
      this.schedule();
      if (preview) {
        this.options.onFillDrag?.(drag.source, preview);
        this.select({ ranges: [preview], active: { row: preview.r1, col: preview.c1 }, anchor: { row: preview.r1, col: preview.c1 } }, false);
      }
      return;
    }
    if (drag?.kind === "resize") {
      const size = drag.axis === "col" ? this.cols.size(drag.index) : this.rows.size(drag.index);
      this.options.onResize?.(drag.axis === "col" ? "col" : "row", drag.index, size);
    }
  };

  private onWheel = (event: WheelEvent): void => {
    if (!event.ctrlKey && !event.metaKey) return;
    event.preventDefault();
    this.setZoom(this.zoom * (event.deltaY < 0 ? 1.1 : 1 / 1.1), true);
  };

  private onContextMenu = (event: MouseEvent): void => {
    if (!this.options.onContextMenu) return;
    const hit = this.hit(event.clientX, event.clientY);
    if (hit.px >= this.width || hit.py >= this.height) return;
    event.preventDefault();
    this.focus();
    const bounds = { rows: this.rows, cols: this.cols };
    const inside = this.selection?.ranges.some((g) => {
      if (hit.onRowHead && hit.onColHead) return false;
      if (hit.onColHead) return hit.col >= g.c1 && hit.col <= g.c2 && g.r1 === 0 && g.r2 >= this.rows.length - 1;
      if (hit.onRowHead) return hit.row >= g.r1 && hit.row <= g.r2 && g.c1 === 0 && g.c2 >= this.cols.length - 1;
      return hit.row >= g.r1 && hit.row <= g.r2 && hit.col >= g.c1 && hit.col <= g.c2;
    });
    // Right-clicking outside the selection selects what was clicked, like other spreadsheets.
    if (!inside) {
      if (hit.onRowHead && hit.onColHead) this.select(selectAll(bounds), false);
      else if (hit.onColHead) this.select(wholeCol(bounds, hit.col), false);
      else if (hit.onRowHead) this.select(wholeRow(bounds, hit.row), false);
      else this.select(single({ row: hit.row, col: hit.col }), false);
    }
    this.options.onContextMenu({ x: event.clientX, y: event.clientY, row: hit.row, col: hit.col, onRowHead: hit.onRowHead, onColHead: hit.onColHead });
  };

  private onDoubleClick = (event: MouseEvent): void => {
    const hit = this.hit(event.clientX, event.clientY);
    if (hit.onColHead || hit.onRowHead) return;
    this.options.onEdit?.({ row: hit.row, col: hit.col }, undefined);
  };

  // --- keyboard ---------------------------------------------------------------------------

  private onKeyDown = (event: KeyboardEvent): void => {
    const sel = this.selection;
    const mod = event.ctrlKey || event.metaKey;
    const bounds = { rows: this.rows, cols: this.cols };
    const arrows: Record<string, Direction> = { ArrowUp: "up", ArrowDown: "down", ArrowLeft: "left", ArrowRight: "right" };
    if (mod && (event.key === "+" || event.key === "=" || event.key === "-" || event.key === "0")) {
      event.preventDefault();
      this.setZoom(event.key === "0" ? 1 : this.zoom * (event.key === "-" ? 1 / 1.1 : 1.1), true);
      return;
    }
    if (mod && event.key.toLowerCase() === "a") {
      event.preventDefault();
      this.select(selectAll(bounds), false);
      return;
    }
    if (!sel) {
      if (arrows[event.key] || event.key === "Tab") {
        event.preventDefault();
        this.select(single({ row: 0, col: 0 }));
      }
      return;
    }
    if (mod && !event.shiftKey && !event.altKey && (event.key.toLowerCase() === "d" || event.key.toLowerCase() === "r")) {
      event.preventDefault(); // the browser would bookmark the page
      this.options.onFill?.(event.key.toLowerCase() === "d" ? "down" : "right", sel);
      return;
    }
    const dir = arrows[event.key];
    if (dir) {
      event.preventDefault();
      // From a merged block, moving starts at the block's edge in that direction.
      const block = !event.shiftKey ? this.merges.at(sel.active.row, sel.active.col) : undefined;
      const from: Selection = block
        ? single({
            row: dir === "down" ? block.r + block.rs - 1 : block.r,
            col: dir === "right" ? block.c + block.cs - 1 : block.c,
          })
        : sel;
      this.select(move(bounds, from, dir, { extend: event.shiftKey, jump: mod, hasData: this.options.hasData }));
      return;
    }
    const page = Math.max(1, Math.floor((this.height - this.headH) / (this.rows.size(sel.active.row) || 30)) - 1);
    switch (event.key) {
      case "PageDown":
      case "PageUp": {
        event.preventDefault();
        const dr = event.key === "PageDown" ? page : -page;
        const row = this.rows.nextVisible(Math.max(0, Math.min(this.rows.length - 1, sel.active.row + dr)), dr > 0 ? -1 : 1);
        this.select(event.shiftKey ? click(sel, { row, col: sel.active.col }, "extend") : single({ row, col: sel.active.col }));
        return;
      }
      case "Home":
      case "End": {
        event.preventDefault();
        const end = event.key === "End";
        const row = mod ? (end ? this.rows.length - 1 : 0) : sel.active.row;
        const col = end && !mod ? this.cols.length - 1 : mod && end ? this.cols.length - 1 : 0;
        this.select(single({ row: this.rows.nextVisible(row, end ? -1 : 1), col: this.cols.nextVisible(col, end ? -1 : 1) }));
        return;
      }
      case "Tab":
        event.preventDefault();
        this.select(advance(bounds, sel, event.shiftKey ? "previous" : "next"));
        return;
      case "Enter":
        event.preventDefault();
        if (this.options.onEdit && !event.shiftKey && (event.altKey || event.metaKey)) return;
        this.options.onEdit?.(sel.active, undefined);
        return;
      case "F2":
        event.preventDefault();
        this.options.onEdit?.(sel.active, undefined);
        return;
      case "Escape":
        return;
      case "Delete":
      case "Backspace":
        event.preventDefault();
        this.options.onClear?.(sel);
        return;
      default:
        if (!mod && !event.altKey && event.key.length === 1) {
          event.preventDefault();
          this.options.onEdit?.(sel.active, event.key);
        }
    }
  };

  // --- clipboard ---------------------------------------------------------------------------

  /** Does the grid have the keyboard (and not a text box on top of it)? */
  private hasFocus(): boolean {
    const active = document.activeElement;
    return active === this.root || (!!active && this.root.contains(active));
  }

  private onCopy = (event: ClipboardEvent): void => {
    const sel = this.selection;
    if (!sel || !this.options.onCopy || !this.hasFocus()) return;
    const range = sel.ranges[sel.ranges.length - 1];
    const result = this.options.onCopy(range, event.type === "cut");
    if (result === undefined) return;
    event.preventDefault();
    if (typeof result === "string") {
      event.clipboardData?.setData("text/plain", result);
    } else {
      // The cells had to load first, so the text arrives after the event; write it the async way.
      void result.then((text) => (text === undefined ? undefined : navigator.clipboard?.writeText(text))).catch(() => {});
    }
  };

  private onPaste = (event: ClipboardEvent): void => {
    const sel = this.selection;
    if (!sel || !this.options.onPaste || !this.hasFocus()) return;
    const text = event.clipboardData?.getData("text/plain");
    event.preventDefault();
    if (text !== undefined && text !== "") this.options.onPaste(text, sel);
  };

  /** After an edit commits: Enter moves down within the selection, Tab to the right. */
  afterEdit(direction: "down" | "next" | "none"): void {
    if (!this.selection || direction === "none") return;
    this.select(advance({ rows: this.rows, cols: this.cols }, this.selection, direction));
    this.focus();
  }

  isSelected(row: number, col: number): boolean {
    return !!this.selection && isSelected(this.selection, row, col);
  }
}

function wrapText(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const lines: string[] = [];
  for (const paragraph of text.split("\n")) {
    let line = "";
    for (const word of paragraph.split(" ")) {
      const next = line ? `${line} ${word}` : word;
      if (line && ctx.measureText(next).width > width) {
        lines.push(line);
        line = word;
      } else line = next;
    }
    lines.push(line);
  }
  return lines;
}
