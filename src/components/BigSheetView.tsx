// Scrollable grid for sheets too big for the main editor. Only the rows on screen are drawn and
// only the chunks holding them are downloaded (BigSheetStore), so an 11M-cell sheet scrolls in a
// normal browser tab. Editing is switched by BIG_SHEET_EDITING.
import { useCallback, useEffect, useMemo, useReducer, useRef, useState, type ReactNode } from "react";
import { AlertTriangle, Loader2, Save } from "lucide-react";
import { BIG_SHEET_EDITING } from "../lib/constants";
import { BigSheetStore } from "../lib/v2/bigSheet";
import type { SheetRootV2 } from "../lib/v2/root";
import { uploadProgress } from "../lib/v2/progress";
import type { LibraryEntry } from "../lib/types";
import type { WorkbookCell } from "../lib/workbook";
import { cellDisplay } from "../lib/workbook/format";
import { BigSheetToolbar, shiftDecimals, type StylePatch } from "./BigSheetToolbar";
import { useLibrary } from "../providers/LibraryProvider";
import { useToast } from "../providers/ToastProvider";
import { cx } from "../lib/utils";

// Same geometry as the main editor's grid (Editor.tsx <Workbook>), so both modes look alike.
const ROW_H = 30;
const COL_W = 116;
const HEAD_W = 46;
const HEAD_H = 28;
/** Browsers cap element heights (~17M px in Firefox); taller sheets scroll proportionally. */
const MAX_SCROLL_PX = 10_000_000;

function columnName(index: number): string {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

function shown(cell: WorkbookCell | null | undefined): string {
  return cell ? cellDisplay(cell) : "";
}

/** A cell after a style change: the cached display text is dropped so the new number format shows. */
function restyle(cell: WorkbookCell | undefined, patch: StylePatch): WorkbookCell {
  const { display: _display, ...rest } = cell ?? {};
  void _display;
  const style = { ...rest.style, ...patch };
  for (const key of Object.keys(style) as (keyof typeof style)[]) if (style[key] === undefined) delete style[key];
  return { ...rest, style };
}

export function BigSheetView({ entry, root }: { entry: LibraryEntry; root: SheetRootV2 }) {
  const library = useLibrary();
  const { toast } = useToast();
  // Created once per root: recreating it (e.g. when the server list loads) would drop unsaved edits.
  const serversRef = useRef(library.blossomServers);
  const store = useMemo(() => new BigSheetStore(root, serversRef.current), [root]);
  const [ready, setReady] = useState(false);
  const [openError, setOpenError] = useState("");
  const [tab, setTab] = useState(0);
  const [, redraw] = useReducer((n: number) => n + 1, 0);
  const [view, setView] = useState({ top: 0, left: 0, width: 800, height: 600 });
  const [selected, setSelected] = useState<{ row: number; col: number } | null>(null);
  /** Other corner of a shift-selected range (equals `selected` for a single cell). */
  const [anchor, setAnchor] = useState<{ row: number; col: number } | null>(null);
  /** Style copied by the paint-format button, applied to the next selection. */
  const [painting, setPainting] = useState<WorkbookCell["style"] | null>(null);
  const [editing, setEditing] = useState<{ row: number; col: number; text: string } | null>(null);
  const [saving, setSaving] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canEdit = BIG_SHEET_EDITING && !!entry.editKey;

  useEffect(() => {
    let cancelled = false;
    setReady(false);
    store
      .open()
      .then(() => !cancelled && setReady(true))
      .catch((error) => !cancelled && setOpenError(error instanceof Error ? error.message : String(error)));
    const off = store.subscribe(redraw);
    return () => {
      cancelled = true;
      off();
    };
  }, [store]);

  const measure = useCallback(() => {
    const el = scrollRef.current;
    if (el) setView({ top: el.scrollTop, left: el.scrollLeft, width: el.clientWidth, height: el.clientHeight });
  }, []);
  useEffect(() => {
    measure();
    window.addEventListener("resize", measure);
    return () => window.removeEventListener("resize", measure);
  }, [measure, ready]);

  const info = store.tabs[tab];
  const rows = info?.rows ?? 0;
  const cols = info?.cols ?? 0;
  const totalPx = rows * ROW_H;
  const scrollPx = Math.min(totalPx, MAX_SCROLL_PX);
  const scale = totalPx > MAX_SCROLL_PX ? (totalPx - view.height) / (MAX_SCROLL_PX - view.height) : 1;
  const virtualTop = view.top * scale;
  const firstRow = Math.max(0, Math.min(rows - 1, Math.floor(virtualTop / ROW_H)));
  const rowOffset = virtualTop - firstRow * ROW_H;
  const visibleRows = Math.ceil((view.height - HEAD_H) / ROW_H) + 1;
  const firstCol = Math.floor(view.left / COL_W);
  const visibleCols = Math.ceil((view.width - HEAD_W) / COL_W) + 1;
  const lastRow = Math.min(rows, firstRow + visibleRows);
  const lastCol = Math.min(cols, firstCol + visibleCols);

  // Fetch the chunks under the visible rows.
  useEffect(() => {
    if (!ready || !info || info.chunks.length === 0) return;
    const from = store.chunkFor(tab, firstRow);
    const to = store.chunkFor(tab, Math.max(firstRow, lastRow - 1));
    for (let c = from; c <= to; c++) void store.load(tab, c).catch(() => {});
  }, [ready, store, tab, info, firstRow, lastRow]);

  const save = useCallback(async () => {
    if (store.dirtyCount === 0 || saving) return;
    setSaving(true);
    try {
      const target = library.bigSheetUpload();
      if (!target.signer || target.servers.length === 0) throw new Error("Sign in and tick a storage server to save.");
      const next = await store.save({
        servers: target.servers,
        replicas: target.replicas,
        signer: target.signer,
        onProgress: (event) =>
          uploadProgress.report({
            phase: "uploading",
            message: "Saving your changes…",
            tabs: [],
            bytesDone: event.bytesDone,
            bytesTotal: event.bytesTotal,
            blobsDone: event.blobsDone,
            blobsTotal: 0,
            servers: target.servers,
          }),
      });
      await library.publishBigRoot(entry, next);
      uploadProgress.dismiss();
      toast("Saved.", "success");
    } catch (error) {
      uploadProgress.dismiss();
      toast(error instanceof Error ? error.message : "Couldn't save.", "error");
    } finally {
      setSaving(false);
    }
  }, [entry, library, saving, store, toast]);

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        void save();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [save]);

  useEffect(() => {
    const warn = (event: BeforeUnloadEvent) => {
      if (store.dirtyCount > 0) event.preventDefault();
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [store]);

  const commit = async () => {
    if (!editing) return;
    const { row, col, text } = editing;
    setEditing(null);
    await store.setCell(tab, row, col, text);
  };

  const range = () => {
    if (!selected) return null;
    const a = anchor ?? selected;
    return {
      r1: Math.min(a.row, selected.row),
      r2: Math.max(a.row, selected.row),
      c1: Math.min(a.col, selected.col),
      c2: Math.max(a.col, selected.col),
    };
  };
  const inRange = (row: number, col: number) => {
    const area = range();
    return !!area && row >= area.r1 && row <= area.r2 && col >= area.c1 && col <= area.c2;
  };
  const editRange = (change: (cell: WorkbookCell | undefined) => WorkbookCell | undefined) => {
    const area = range();
    if (!area || !canEdit) return;
    store.edit(tab, area.r1, area.c1, area.r2, area.c2, change).catch((error) =>
      toast(error instanceof Error ? error.message : String(error), "error"),
    );
  };
  const applyStyle = (patch: StylePatch) => editRange((cell) => restyle(cell, patch));
  const select = (row: number, col: number, extend: boolean) => {
    setSelected({ row, col });
    if (!extend) setAnchor({ row, col });
    if (painting !== null && !extend) {
      const style = painting;
      setPainting(null);
      if (canEdit) {
        store
          .edit(tab, row, col, row, col, (cell) => ({ ...(cell ?? {}), style: style ? { ...style } : undefined }))
          .catch(() => {});
      }
    }
  };

  const selectionLabel = () => {
    const area = range();
    if (!area) return "";
    const first = `${columnName(area.c1)}${area.r1 + 1}`;
    return area.r1 === area.r2 && area.c1 === area.c2 ? first : `${first}:${columnName(area.c2)}${area.r2 + 1}`;
  };

  /** Scrolls so `row` is on screen (used by Go to row). */
  const goToRow = (row: number) => {
    const el = scrollRef.current;
    if (!el) return;
    const target = Math.max(0, Math.min(rows - 1, row));
    el.scrollTop = (target * ROW_H) / scale;
    select(target, selected?.col ?? 0, false);
  };

  if (openError) {
    return (
      <div className="editor-error editor-error-inline">
        <AlertTriangle size={24} />
        <h2>Couldn't open this sheet</h2>
        <p>{openError}</p>
      </div>
    );
  }
  if (!ready || !info) {
    return (
      <div className="editor-loading editor-loading-inline">
        <Loader2 size={22} className="spin" />
        <p>Loading the first rows…</p>
      </div>
    );
  }

  const cells: ReactNode[] = [];
  for (let row = firstRow; row < lastRow; row++) {
    const y = HEAD_H + (row - firstRow) * ROW_H - rowOffset;
    const loaded = store.isLoaded(tab, row);
    cells.push(
      <div key={`h${row}`} className="big-row-head" style={{ top: y, height: ROW_H, width: HEAD_W }}>
        {(row + 1).toLocaleString()}
      </div>,
    );
    for (let col = firstCol; col < lastCol; col++) {
      const x = HEAD_W + col * COL_W - view.left;
      const cell = loaded ? store.cell(tab, row, col) : undefined;
      const isSel = selected?.row === row && selected.col === col;
      const st = cell?.style;
      if (editing && editing.row === row && editing.col === col) {
        cells.push(
          <input
            key={`e${row}:${col}`}
            className="big-cell-input"
            style={{ top: y, left: x, width: COL_W, height: ROW_H }}
            autoFocus
            value={editing.text}
            onChange={(event) => setEditing({ ...editing, text: event.target.value })}
            onBlur={() => void commit()}
            onKeyDown={(event) => {
              if (event.key === "Enter") void commit().then(() => select(row + 1, col, false));
              if (event.key === "Escape") setEditing(null);
            }}
          />,
        );
        continue;
      }
      cells.push(
        <div
          key={`${row}:${col}`}
          className={cx(
            "big-cell",
            isSel && "is-selected",
            !isSel && inRange(row, col) && "is-in-range",
            !loaded && "is-loading",
          )}
          style={{
            top: y,
            left: x,
            width: COL_W,
            height: ROW_H,
            fontWeight: st?.bold ? 700 : undefined,
            fontStyle: st?.italic ? "italic" : undefined,
            textDecoration: [st?.underline && "underline", st?.strike && "line-through"].filter(Boolean).join(" ") || undefined,
            fontFamily: st?.fontFamily,
            fontSize: st?.fontSize ? `${Math.min(st.fontSize, 18)}pt` : undefined,
            textAlign: st?.align,
            whiteSpace: st?.wrap ? "normal" : undefined,
            background: st?.background,
            color: st?.color,
          }}
          onClick={(event) => select(row, col, event.shiftKey)}
          onDoubleClick={() => canEdit && loaded && setEditing({ row, col, text: cell?.formula ? `=${cell.formula.replace(/^=/, "")}` : shown(cell) })}
        >
          {loaded ? shown(cell) : ""}
        </div>,
      );
    }
  }
  for (let col = firstCol; col < lastCol; col++) {
    cells.push(
      <div key={`c${col}`} className="big-col-head" style={{ left: HEAD_W + col * COL_W - view.left, width: COL_W, height: HEAD_H }}>
        {columnName(col)}
      </div>,
    );
  }

  const selectedCell = selected ? store.cell(tab, selected.row, selected.col) : undefined;

  return (
    <div
      className="big-sheet"
      tabIndex={-1}
      onKeyDown={(event) => {
        if (editing || !selected) return;
        const mod = event.ctrlKey || event.metaKey;
        const key = event.key.toLowerCase();
        if (mod && canEdit && (key === "z" || key === "y")) {
          event.preventDefault();
          if (key === "y" || event.shiftKey) store.redo();
          else store.undo();
          return;
        }
        if (mod && canEdit && (key === "b" || key === "i" || key === "u")) {
          event.preventDefault();
          const flag = key === "b" ? "bold" : key === "i" ? "italic" : "underline";
          applyStyle({ [flag]: !selectedCell?.style?.[flag] });
          return;
        }
        const move = (dr: number, dc: number) => {
          event.preventDefault();
          select(
            Math.max(0, Math.min(rows - 1, selected.row + dr)),
            Math.max(0, Math.min(cols - 1, selected.col + dc)),
            event.shiftKey && event.key !== "Tab",
          );
        };
        if (event.key === "ArrowDown") move(1, 0);
        else if (event.key === "ArrowUp") move(-1, 0);
        else if (event.key === "ArrowRight" || event.key === "Tab") move(0, 1);
        else if (event.key === "ArrowLeft") move(0, -1);
        else if (canEdit && (event.key === "Enter" || event.key === "F2")) {
          event.preventDefault();
          setEditing({ ...selected, text: shown(selectedCell) });
        } else if (canEdit && (event.key === "Delete" || event.key === "Backspace")) {
          editRange((cell) => (cell?.style ? { style: cell.style } : undefined));
        } else if (canEdit && event.key.length === 1 && !mod) {
          setEditing({ ...selected, text: event.key });
          event.preventDefault();
        }
      }}
    >
      <BigSheetToolbar
        style={selectedCell?.style ?? {}}
        canEdit={canEdit && !!selected}
        canUndo={store.canUndo}
        canRedo={store.canRedo}
        painting={painting !== null}
        onUndo={() => store.undo()}
        onRedo={() => store.redo()}
        onPaint={() => setPainting(painting !== null ? null : { ...(selectedCell?.style ?? {}) })}
        onClear={() =>
          editRange((cell) =>
            cell ? { ...(cell.value !== undefined ? { value: cell.value } : {}), ...(cell.formula ? { formula: cell.formula } : {}) } : undefined,
          )
        }
        onStyle={applyStyle}
        onFormat={(format) => applyStyle({ numberFormat: format })}
        onDecimals={(delta) => editRange((cell) => restyle(cell, { numberFormat: shiftDecimals(cell?.style?.numberFormat, delta) }))}
      />
      <div className="big-sheet-bar">
        <span className="big-sheet-ref">{selectionLabel()}</span>
        <span className="big-sheet-fx">fx</span>
        <span className="big-sheet-formula">
          {selectedCell?.formula ? `=${selectedCell.formula.replace(/^=/, "")}` : shown(selectedCell)}
        </span>
        <input
          className="big-sheet-goto"
          placeholder="Go to row"
          inputMode="numeric"
          onKeyDown={(event) => {
            if (event.key === "Enter") goToRow(Number((event.target as HTMLInputElement).value.replace(/\D/g, "")) - 1);
          }}
        />
        <span className="big-sheet-note">
          {rows.toLocaleString()} rows · large sheet mode{canEdit ? "" : " · read-only"}
        </span>
        {canEdit && (
          <button type="button" className="btn btn-primary btn-sm" disabled={saving || store.dirtyCount === 0} onClick={() => void save()}>
            {saving ? <Loader2 size={14} className="spin" /> : <Save size={14} />}
            {store.dirtyCount ? `Save ${store.dirtyCount} changed chunk${store.dirtyCount === 1 ? "" : "s"}` : "Saved"}
          </button>
        )}
      </div>
      <div ref={scrollRef} className="big-sheet-scroll" onScroll={measure}>
        {/* Sticky viewport first, then the spacer that gives the scrollbars their length. */}
        <div className="big-sheet-viewport" style={{ width: view.width, height: view.height }}>
          {cells}
          <div className="big-corner" style={{ width: HEAD_W, height: HEAD_H }} />
        </div>
        <div style={{ height: scrollPx + HEAD_H, width: HEAD_W + cols * COL_W, marginTop: -view.height }} />
      </div>
      <div className="big-sheet-tabs">
        {store.tabs.map((item, index) => (
          <button
            key={item.name + index}
            type="button"
            className={cx("big-sheet-tab", index === tab && "is-active")}
            onClick={() => {
              setTab(index);
              setSelected(null);
              setAnchor(null);
              setEditing(null);
              if (scrollRef.current) scrollRef.current.scrollTop = 0;
            }}
          >
            {item.name}
          </button>
        ))}
      </div>
    </div>
  );
}
