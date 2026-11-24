// Grid for sheets too big for the main editor. The canvas grid (@formstr/sheets-grid) draws only the
// rows on screen and only the chunks holding them are downloaded (BigSheetStore), so an 11M-cell
// sheet scrolls in a normal browser tab. Editing is switched by BIG_SHEET_EDITING.
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { addBorder, columnName, Grid, type BorderType, type GridData, type Selection } from "@formstr/sheets-grid";
import { AlertTriangle, Loader2, Save } from "lucide-react";
import { BIG_SHEET_EDITING } from "../lib/constants";
import { BigSheetStore } from "../lib/v2/bigSheet";
import { sortArea } from "../lib/v2/bigSort";
import { editStructure, renameTabAndFormulas } from "../lib/v2/bigStructure";
import { calculateTyped, FormulaIndex, pointRect, recalcVolatileInChunk, refreshDependents, type ChangedRect } from "../lib/v2/bigFormulas";
import {
  copyBlock,
  copyBlockNow,
  dragFill,
  moveCells,
  fillArea,
  fromTsv,
  isLoadedArea,
  pasteBlock,
  pasteText,
  type Block,
} from "../lib/v2/bigClipboard";
import type { SheetRootV2 } from "../lib/v2/root";
import { uploadProgress } from "../lib/v2/progress";
import type { LibraryEntry } from "../lib/types";
import type { WorkbookCell } from "../lib/workbook";
import { cellDisplay } from "../lib/workbook/format";
import { BigFindBar } from "./BigFindBar";
import { exportBigSheet } from "../lib/v2/bigExport";
import { downloadBlob } from "../lib/xlsx";
import { fileSafeName } from "../lib/utils";
import { PersonDialog, TagDialog, type PersonChoice } from "./CellMetaDialogs";
import { useAuth } from "../providers/AuthProvider";
import { useProfiles } from "../providers/ProfileProvider";
import { colorFromString, shortNpub, toNpub } from "../lib/utils";
import type { UserRef } from "../lib/types";
import { BigValidationDialog } from "./BigValidationDialog";
import { BigConditionalDialog } from "./BigConditionalDialog";
import { BigFilterDialog } from "./BigFilterDialog";
import { BigSortDialog } from "./BigSortDialog";
import { formatFor, numericExtent, SCALE_TYPES, withoutArea } from "../lib/v2/bigConditional";
import { validate } from "../lib/v2/bigValidation";
import { helpFor, tokenizeFormula } from "../lib/v2/formulaHelp";
import { BigSheetToolbar, shiftDecimals, type StylePatch } from "./BigSheetToolbar";
import { useLibrary } from "../providers/LibraryProvider";
import { useToast } from "../providers/ToastProvider";
import { cx } from "../lib/utils";

/** The formula engine's number formatter (scientific, accounting, times, dates, [$€-407]...), once it has loaded. */
let formatter: ((value: number, format: string) => string) | null = null;
void import("@formstr/sheets-formula").then((engine) => {
  formatter = (value, format) => engine.formatValue(value, format);
});
/** Formats the simple, common cases are handled by cellDisplay; the rest go to the engine's formatter. */
const SIMPLE_FORMAT = /^(General|@|#,##0(\.0+)?|0(\.0+)?%?|#,##0(\.0+)?%|\$#,##0(\.0+)?|yyyy-mm-dd)$/;

function shown(cell: WorkbookCell | null | undefined): string {
  if (!cell) return "";
  const format = cell.style?.numberFormat;
  if (formatter && typeof cell.value === "number" && !cell.display && format && !SIMPLE_FORMAT.test(format)) {
    try {
      return formatter(cell.value, format);
    } catch {
      // fall through to the basic formatter
    }
  }
  return cellDisplay(cell);
}

/** "B7" or "A1:C5" as a range (0-based), or null when it is not a cell reference. */
function parseReference(text: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(text.trim());
  if (!match) return null;
  const col = (letters: string) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const r1 = Number(match[2]) - 1;
  const c1 = col(match[1]);
  const r2 = match[4] ? Number(match[4]) - 1 : r1;
  const c2 = match[3] ? col(match[3]) : c1;
  if (r1 < 0 || r2 < 0) return null;
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

/** Colours for the cells a formula points at (the same cell always gets the same colour). */
const REF_COLORS = ["#2563eb", "#dc2626", "#16a34a", "#9333ea", "#ea580c", "#0891b2", "#c026d3", "#65a30d"];

/** A cell after a style change: the cached display text is dropped so the new number format shows. */
function restyle(cell: WorkbookCell | undefined, patch: StylePatch): WorkbookCell {
  const { display: _display, ...rest } = cell ?? {};
  void _display;
  const style = { ...rest.style, ...patch };
  for (const key of Object.keys(style) as (keyof typeof style)[]) if (style[key] === undefined) delete style[key];
  return { ...rest, style };
}

/** Gives the grid what the tab saved about its look: column widths, row heights, hidden rows/columns, frozen panes. */
function applyLayout(grid: Grid, store: BigSheetStore, tab: number): void {
  const layout = store.tabs[tab]?.layout ?? {};
  const overrides = (sizes?: Record<number, number>, hidden?: number[]) => [
    ...Object.entries(sizes ?? {}).map(([index, size]) => ({ index: Number(index), size })),
    ...(hidden ?? []).map((index) => ({ index, size: 0 })),
  ];
  grid.setSizes("col", overrides(layout.columnWidths, layout.hiddenColumns));
  // Rows hidden by a filter are hidden like any other hidden row.
  grid.setSizes("row", overrides(layout.rowHeights, [...(layout.hiddenRows ?? []), ...(layout.filter?.hidden ?? [])]));
  grid.setFrozen(layout.frozen?.rows ?? 0, layout.frozen?.columns ?? 0);
  grid.setZoom(layout.zoom ?? 1);
  grid.setMerges(layout.merges ?? []);
  grid.setBorders(layout.borders ?? []);
}

/** The store's tab as the grid reads it. */
function tabData(store: BigSheetStore, tab: number): GridData {
  return {
    get rowCount() {
      return store.tabs[tab]?.rows ?? 0;
    },
    get colCount() {
      return store.tabs[tab]?.cols ?? 0;
    },
    cell(row, col) {
      if (!store.isLoaded(tab, row)) return undefined;
      const cell = store.cell(tab, row, col);
      // Conditional formatting colours the cell while it is drawn; nothing is stored in the cell.
      const rule = formatFor(store.tabs[tab]?.layout.conditionalFormats, row, col, cell);
      const style = rule
        ? { ...cell?.style, ...(rule.textColor ? { color: rule.textColor } : {}), ...(rule.cellColor ? { background: rule.cellColor } : {}), ...(rule.bar ? { bar: rule.bar } : {}) }
        : cell?.style;
      if (!cell) return rule ? { text: "", style } : null;
      // A cell that no longer fits its validation rule (typed before the rule, pasted, filled) is marked.
      const invalid = !!cell.validation && cell.value !== undefined && validate(cell.validation, shown(cell)) !== null;
      const checkbox = cell.validation?.type === "checkbox" ? cell.value === true || String(cell.value).toUpperCase() === "TRUE" : undefined;
      const pill = cell.person
        ? { label: cell.person.name || shown(cell) || shortNpub(cell.person.pubkey), ...colorFromString(cell.person.name || shown(cell) || cell.person.pubkey), avatar: true }
        : cell.tag
          ? { label: cell.tag, ...colorFromString(cell.tag), avatar: false }
          : undefined;
      return { text: shown(cell), style, link: !!cell.link, note: !!cell.comment, invalid, numeric: typeof cell.value === "number", pill, checkbox };
    },
    ensure(first, end) {
      const info = store.tabs[tab];
      if (!info || info.chunks.length === 0 || end <= first) return;
      const from = store.chunkFor(tab, first);
      const to = store.chunkFor(tab, Math.max(first, end - 1));
      for (let c = from; c <= to; c++) void store.load(tab, c).catch(() => {});
    },
    subscribe: (listener) => store.subscribe(listener),
  };
}

export function BigSheetView({ entry, root }: { entry: LibraryEntry; root: SheetRootV2 }) {
  const library = useLibrary();
  const auth = useAuth();
  const profiles = useProfiles();
  const { toast } = useToast();
  // Created once per root: recreating it (e.g. when the server list loads) would drop unsaved edits.
  const serversRef = useRef(library.blossomServers);
  const store = useMemo(() => new BigSheetStore(root, serversRef.current), [root]);
  const [ready, setReady] = useState(false);
  const [openError, setOpenError] = useState("");
  const [tab, setTab] = useState(0);
  const [, setVersion] = useState(0);
  const [selection, setSelection] = useState<Selection | null>(null);
  /** Style copied by the paint-format button, applied to the next selection. */
  const [painting, setPainting] = useState<WorkbookCell["style"] | null>(null);
  const [editing, setEditing] = useState<{ row: number; col: number; text: string } | null>(null);
  const [saving, setSaving] = useState(false);
  const [validating, setValidating] = useState(false);
  const [metaDialog, setMetaDialog] = useState<"person" | "tag" | null>(null);
  /** Manual calculation: edits do not refresh other formulas until F9 (or the Recalculate button). */
  const [manualCalc, setManualCalc] = useState(false);
  const [stale, setStale] = useState(false);
  const manualRef = useRef(false);
  manualRef.current = manualCalc;
  const [filtering, setFiltering] = useState(false);
  const [sortDialog, setSortDialog] = useState(false);
  /** Right-click menu of a sheet tab (page position of its bottom-left, which tab). */
  const [tabMenu, setTabMenu] = useState<{ x: number; y: number; index: number } | null>(null);
  /** Function names for formula autocomplete (loaded the first time a formula is typed). */
  const [functionList, setFunctionList] = useState<string[] | null>(null);
  const [suggestAt, setSuggestAt] = useState(0);
  const [formatting, setFormatting] = useState(false);
  /** What is typed in the name box (null: show the selection). */
  const [nameDraft, setNameDraft] = useState<string | null>(null);
  /** The find bar (Ctrl+F), with the replace row when opened by Ctrl+H. */
  const [finding, setFinding] = useState<{ replace: boolean } | null>(null);
  /** Right-click menu: where it is on the page. */
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  /** Progress of a formula that has to read several chunks, e.g. =SUM(A:A). */
  const [calculating, setCalculating] = useState<string | null>(null);
  const hostRef = useRef<HTMLDivElement>(null);
  const gridRef = useRef<Grid | null>(null);
  /** Set by Escape so the blur that follows the input unmounting does not save the edit. */
  const cancelled = useRef(false);
  /** The cells of the last copy, so pasting them keeps formulas and formats. */
  const clipboard = useRef<Block | null>(null);
  const canEdit = BIG_SHEET_EDITING && !!entry.editKey;
  // The grid's callbacks are created once, so they read the latest state through these refs.
  const live = useRef({ canEdit, editing, painting, tab });
  live.current = { canEdit, editing, painting, tab };

  useEffect(() => {
    let cancelled = false;
    setReady(false);
    store
      .open()
      .then(() => !cancelled && setReady(true))
      .catch((error) => !cancelled && setOpenError(error instanceof Error ? error.message : String(error)));
    // Toolbar state (undo/redo, dirty count) follows the store; the grid redraws itself.
    const off = store.subscribe(() => setVersion((n) => n + 1));
    return () => {
      cancelled = true;
      off();
    };
  }, [store]);

  const info = store.tabs[tab];

  // Which cells hold formulas, so a change can refresh the formulas that read it.
  const index = useMemo(() => new FormulaIndex(), [store]);
  // Refreshes run one after another, so two quick edits never recalculate the same cells at once.
  const refreshQueue = useRef<Promise<void>>(Promise.resolve());
  const refresh = useCallback(
    (areas: ChangedRect[], recalcInside = false, force = false, all = false): Promise<void> => {
      const job = refreshQueue.current.then(async () => {
        if (manualRef.current && !force) {
          setStale(true);
          return;
        }
        try {
          const count = await refreshDependents(
            store,
            index,
            areas,
            (done, total, phase) =>
              setCalculating(
                phase === "scan"
                  ? `Looking for formulas… ${done} of ${total} parts`
                  : total > 1
                    ? `Updating formulas… ${done.toLocaleString()} of ${total.toLocaleString()}`
                    : null,
              ),
            recalcInside,
            { all },
          );
          if (all) setStale(false);
          if (count > 0) setVersion((n) => n + 1);
        } catch (error) {
          toast(error instanceof Error ? error.message : "Couldn't update formulas.", "error");
        } finally {
          setCalculating(null);
        }
      });
      refreshQueue.current = job;
      return job;
    },
    [index, store, toast],
  );
  const refreshRef = useRef(refresh);
  refreshRef.current = refresh;
  useEffect(() => {
    if (!ready) return;
    // Each chunk is read for formulas as it arrives; NOW(), RAND() and friends get fresh values right away.
    const learn = (tabNo: number, chunk: number) =>
      void index
        .scanChunk(store, tabNo, chunk)
        .then(() => recalcVolatileInChunk(store, tabNo, chunk))
        .catch(() => {});
    store.onChunkLoaded = learn;
    for (const [tabNo, chunk] of store.loadedChunks()) learn(tabNo, chunk);
    return () => {
      store.onChunkLoaded = undefined;
    };
  }, [ready, store, index]);

  const range = selection?.ranges[selection.ranges.length - 1] ?? null;
  const active = selection?.active ?? null;

  const editRange = useCallback(
    (change: (cell: WorkbookCell | undefined) => WorkbookCell | undefined, target?: Selection | null) => {
      const area = (target ?? selection)?.ranges[(target ?? selection)!.ranges.length - 1];
      if (!area || !canEdit) return;
      store
        .edit(tab, area.r1, area.c1, area.r2, area.c2, change)
        .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
    },
    [canEdit, selection, store, tab, toast],
  );

  // One grid for the life of the view; tabs swap its data.
  useEffect(() => {
    const host = hostRef.current;
    if (!ready || !host || !store.tabs[0]) return;
    /** Clears values (keeps formats), then refreshes the formulas that read the cleared cells. */
    const clearArea = (area: { r1: number; c1: number; r2: number; c2: number }) => {
      if (!live.current.canEdit) return;
      const tabNo = live.current.tab;
      store
        .edit(tabNo, area.r1, area.c1, area.r2, area.c2, (cell) => (cell?.style ? { style: cell.style } : undefined))
        .then(() => refreshRef.current([{ tab: tabNo, r1: area.r1, c1: area.c1, r2: area.r2, c2: area.c2 }]))
        .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
    };
    const grid = new Grid(host, {
      data: tabData(store, live.current.tab),
      onSelect: (next) => {
        setSelection(next);
        const style = live.current.painting;
        if (style !== null && live.current.canEdit) {
          setPainting(null);
          const area = next.ranges[next.ranges.length - 1];
          store
            .edit(live.current.tab, area.r1, area.c1, area.r2, area.c2, (cell) => ({ ...(cell ?? {}), style: style ? { ...style } : undefined }))
            .catch(() => {});
        }
      },
      onEdit: (cell, text) => {
        if (!live.current.canEdit || !store.isLoaded(live.current.tab, cell.row)) return;
        const existing = store.cell(live.current.tab, cell.row, cell.col);
        const initial = text ?? (existing?.formula ? `=${existing.formula.replace(/^=/, "")}` : shown(existing));
        cancelled.current = false;
        setEditing({ row: cell.row, col: cell.col, text: initial });
      },
      onClear: (next) => clearArea(next.ranges[next.ranges.length - 1]),
      onReference: (cell) => {
        // While a formula waits for an operand ("=", "=A1+", "=SUM("), a click inserts that cell's address.
        const current = live.current.editing;
        if (!current || !/^=(?:.*[-+*/^&(,<>=:;])?$/.test(current.text)) return false;
        setEditing({ ...current, text: `${current.text}${columnName(cell.col)}${cell.row + 1}` });
        return true;
      },
      onZoom: (zoom) => {
        // Zoom is a view setting: it is remembered with the sheet when it can be edited.
        setVersion((n) => n + 1);
        if (live.current.canEdit) store.setLayout(live.current.tab, { zoom: zoom === 1 ? undefined : zoom }).catch(() => {});
      },
      onContextMenu: (info) => setMenu({ x: info.x, y: info.y }),
      onResize: (kind, index, size) => {
        // Resizing in a read-only sheet only lasts until the page is closed.
        if (!live.current.canEdit) return;
        const tabNo = live.current.tab;
        const layout = store.tabs[tabNo].layout;
        const patch = kind === "col" ? { columnWidths: { ...layout.columnWidths, [index]: size } } : { rowHeights: { ...layout.rowHeights, [index]: size } };
        store.setLayout(tabNo, patch).catch(() => {});
      },
      onCopy: (area, cut) => {
        if (cut && !live.current.canEdit) return undefined;
        const tabNo = live.current.tab;
        const finish = (block: Block) => {
          clipboard.current = block;
          if (cut) clearArea(area);
          return block.text;
        };
        try {
          // Everything in memory: answer inside the copy event. Otherwise download first.
          if (isLoadedArea(store, tabNo, area)) return finish(copyBlockNow(store, tabNo, area));
          return copyBlock(store, tabNo, area).then(finish, (error) => {
            toast(error instanceof Error ? error.message : String(error), "error");
            return undefined;
          });
        } catch (error) {
          toast(error instanceof Error ? error.message : String(error), "error");
          return undefined;
        }
      },
      onPaste: (text, next) => {
        if (!live.current.canEdit) return;
        const tabNo = live.current.tab;
        const area = next.ranges[next.ranges.length - 1];
        const copied = clipboard.current;
        const normal = text.replace(/\r\n?/g, "\n");
        const job =
          copied && copied.text === normal
            ? pasteBlock(store, tabNo, area.r1, area.c1, copied, area)
            : pasteText(store, tabNo, area.r1, area.c1, fromTsv(normal), area);
        job
          .then(async (done) => {
            const written = { tab: tabNo, ...done };
            gridRef.current?.select({
              ranges: [done],
              active: { row: done.r1, col: done.c1 },
              anchor: { row: done.r1, col: done.c1 },
            });
            await refreshRef.current([written], true);
          })
          .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
      },
      onToggle: (cell) => {
        if (!live.current.canEdit) return;
        const tabNo = live.current.tab;
        const current = store.cell(tabNo, cell.row, cell.col);
        if (current?.validation?.type !== "checkbox") return;
        const ticked = current.value === true || String(current.value).toUpperCase() === "TRUE";
        store
          .edit(tabNo, cell.row, cell.col, cell.row, cell.col, (before) => ({ ...(before ?? {}), value: !ticked }))
          .then(() => refreshRef.current([{ tab: tabNo, r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col }]))
          .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
      },
      onMoveDrag: (source, target) => {
        if (!live.current.canEdit) return;
        const tabNo = live.current.tab;
        moveCells(store, tabNo, source, target)
          .then(async (done) => {
            if (!done) return;
            // Both the cells that left and the cells that arrived can change formulas that read them.
            await refreshRef.current(
              [
                { tab: tabNo, ...source },
                { tab: tabNo, ...done },
              ],
              true,
            );
          })
          .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
      },
      onFillDrag: (source, extended) => {
        if (!live.current.canEdit) return;
        const tabNo = live.current.tab;
        dragFill(store, tabNo, source, extended)
          .then(async (done) => {
            if (done) await refreshRef.current([{ tab: tabNo, ...done }], true);
          })
          .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
      },
      onFill: (direction, next) => {
        if (!live.current.canEdit) return;
        const tabNo = live.current.tab;
        fillArea(store, tabNo, next.ranges[next.ranges.length - 1], direction)
          .then(async (done) => {
            if (done) await refreshRef.current([{ tab: tabNo, ...done }], true);
          })
          .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
      },
    });
    gridRef.current = grid;
    applyLayout(grid, store, live.current.tab);
    return () => {
      grid.destroy();
      gridRef.current = null;
    };
    // The tab is read through `live`; switching tabs calls setData below, not a rebuild.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ready, store]);

  const switchTab = (index: number) => {
    setTab(index);
    setSelection(null);
    setEditing(null);
    gridRef.current?.setData(tabData(store, index));
    if (gridRef.current) applyLayout(gridRef.current, store, index);
  };

  // --- sheet tabs ------------------------------------------------------------------------------

  /** Shows tab `next` after the tab list changed (tab numbers may have moved, so formulas are looked up again). */
  const showTabs = (next: number, renumbered: boolean) => {
    if (renumbered) index.forget();
    const grid = gridRef.current;
    setTab(next);
    setSelection(null);
    setEditing(null);
    if (grid) {
      grid.setData(tabData(store, next));
      applyLayout(grid, store, next);
    }
    setVersion((n) => n + 1);
  };
  /** Runs a tab change after any formula work that is still going. */
  const tabJob = (work: () => Promise<void> | void) => {
    const job = refreshQueue.current.then(async () => {
      try {
        await work();
      } catch (error) {
        toast(error instanceof Error ? error.message : "Couldn't change the tabs.", "error");
      } finally {
        setCalculating(null);
      }
    });
    refreshQueue.current = job;
  };
  const visibleTabs = () => store.tabs.map((_, i) => i).filter((i) => !store.tabMeta(i).hidden);
  const uniqueName = (base: string) => {
    const taken = new Set(store.tabs.map((t) => t.name.toLowerCase()));
    if (!taken.has(base.toLowerCase())) return base;
    for (let n = 2; ; n++) if (!taken.has(`${base} ${n}`.toLowerCase())) return `${base} ${n}`;
  };
  const addTab = () =>
    tabJob(() => {
      if (!canEdit) return;
      const next = store.addTab(uniqueName(`Sheet${store.tabs.length + 1}`));
      showTabs(next, false);
    });
  const renameTab = (at: number) => {
    if (!canEdit) return;
    const asked = window.prompt("Tab name:", store.tabs[at].name);
    if (asked === null) return;
    const name = asked.trim();
    if (!name || name.length > 31 || /[[\]:\\/?*']/.test(name)) return toast("Use 1–31 characters, without [ ] : \\ / ? * or quotes.", "error");
    if (store.tabs.some((t, i) => i !== at && t.name.toLowerCase() === name.toLowerCase())) return toast("Another tab already has that name.", "error");
    tabJob(async () => {
      setCalculating("Updating formulas…");
      await renameTabAndFormulas(store, index, at, name, (done, total) => setCalculating(`Looking for formulas… ${done} of ${total} parts`));
      setVersion((n) => n + 1);
    });
  };
  const deleteTab = (at: number) => {
    if (!canEdit) return;
    if (store.tabs.length <= 1) return toast("A workbook needs at least one tab.", "error");
    if (!window.confirm(`Delete the tab "${store.tabs[at].name}"? This can't be undone.`)) return;
    tabJob(() => {
      const order = store.tabs.map((_, i) => i).filter((i) => i !== at);
      const shownBefore = tab;
      store.remapTabs(order);
      const next = shownBefore === at ? Math.min(at, order.length - 1) : shownBefore > at ? shownBefore - 1 : shownBefore;
      const visible = visibleTabs();
      showTabs(visible.includes(next) ? next : (visible[0] ?? 0), true);
    });
  };
  const moveTab = (at: number, delta: number) => {
    if (!canEdit) return;
    const to = at + delta;
    if (to < 0 || to >= store.tabs.length) return;
    tabJob(() => {
      const order = store.tabs.map((_, i) => i);
      [order[at], order[to]] = [order[to], order[at]];
      store.remapTabs(order);
      showTabs(order.indexOf(tab), true);
    });
  };
  const copyTab = (at: number) => {
    if (!canEdit) return;
    tabJob(() => {
      const next = store.duplicateTab(at, uniqueName(`Copy of ${store.tabs[at].name}`.slice(0, 31)));
      showTabs(next, true);
    });
  };
  const hideTab = (at: number, hidden: boolean) => {
    if (!canEdit) return;
    try {
      store.setTabHidden(at, hidden);
    } catch (error) {
      return toast(error instanceof Error ? error.message : "Couldn't hide that tab.", "error");
    }
    if (hidden && at === tab) showTabs(visibleTabs()[0] ?? 0, false);
    else setVersion((n) => n + 1);
  };
  const TAB_COLORS = ["#e5484d", "#f76b15", "#f5d90a", "#30a46c", "#0090ff", "#8e4ec6", "#6b7280"];

  /** Keeps the first rows/columns in view while scrolling; remembered with the sheet when it can be edited. */
  const freeze = (rows: number, columns: number) => {
    gridRef.current?.setFrozen(rows, columns);
    if (canEdit) store.setLayout(tab, { frozen: rows > 0 || columns > 0 ? { rows, columns } : undefined }).catch(() => {});
    setVersion((n) => n + 1);
  };
  /** Puts a border on the selection (or takes them off with "none"). */
  const border = (type: BorderType, color: string, style: string) => {
    if (!canEdit || !range) return;
    const layout = store.tabs[tab].layout;
    const next = addBorder(layout.borders ?? [], { ...range, type, color, style });
    store.setLayout(tab, { borders: next.length > 0 ? next : undefined }).then(() => gridRef.current && applyLayout(gridRef.current, store, tab)).catch(() => {});
  };

  /**
   * Merges the selection into one cell (it keeps the top-left value), or splits merged cells.
   * "all" merges everything, "rows" each row on its own, "columns" each column on its own.
   */
  const merge = async (how: "all" | "rows" | "columns" | "split") => {
    if (!canEdit || !range) return;
    const layout = store.tabs[tab].layout;
    const touches = (m: { r: number; c: number; rs: number; cs: number }) =>
      m.r <= range.r2 && m.r + m.rs - 1 >= range.r1 && m.c <= range.c2 && m.c + m.cs - 1 >= range.c1;
    const kept = (layout.merges ?? []).filter((m) => !touches(m));
    const added: { r: number; c: number; rs: number; cs: number }[] = [];
    if (how === "all") added.push({ r: range.r1, c: range.c1, rs: range.r2 - range.r1 + 1, cs: range.c2 - range.c1 + 1 });
    if (how === "rows") for (let r = range.r1; r <= range.r2; r++) added.push({ r, c: range.c1, rs: 1, cs: range.c2 - range.c1 + 1 });
    if (how === "columns") for (let c = range.c1; c <= range.c2; c++) added.push({ r: range.r1, c, rs: range.r2 - range.r1 + 1, cs: 1 });
    const blocks = added.filter((m) => m.rs * m.cs > 1);
    // Nothing to merge and nothing to split: leave the sheet as it is (no "unsaved changes").
    if (blocks.length === 0 && kept.length === (layout.merges ?? []).length) return;
    if (blocks.length > 20_000) {
      toast("That would merge too many cells at once.", "error");
      return;
    }
    try {
      // The other cells of a block give up their values; the top-left one is kept.
      for (const m of blocks) {
        const clear = (cell: WorkbookCell | undefined) => (cell?.style ? { style: cell.style } : undefined);
        if (m.rs > 1) await store.edit(tab, m.r + 1, m.c, m.r + m.rs - 1, m.c + m.cs - 1, clear);
        if (m.cs > 1) await store.edit(tab, m.r, m.c + 1, m.r, m.c + m.cs - 1, clear);
      }
      await store.setLayout(tab, { merges: [...kept, ...blocks].length > 0 ? [...kept, ...blocks] : undefined });
      if (gridRef.current) applyLayout(gridRef.current, store, tab);
      await refresh(blocks.map((m) => ({ tab, r1: m.r, c1: m.c, r2: m.r + m.rs - 1, c2: m.c + m.cs - 1 })));
    } catch (error) {
      toast(error instanceof Error ? error.message : String(error), "error");
    }
  };
  const mergedHere = !!range && !!store.tabs[tab]?.layout.merges?.some((m) => m.r <= range.r2 && m.r + m.rs - 1 >= range.r1 && m.c <= range.c2 && m.c + m.cs - 1 >= range.c1);

  /** Hides or shows the rows/columns of the selection (showing also reveals the hidden ones just beside it). */
  const setHidden = (axis: "rows" | "columns", hide: boolean) => {
    if (!canEdit || !range) return;
    const key = axis === "rows" ? "hiddenRows" : "hiddenColumns";
    const from = axis === "rows" ? range.r1 : range.c1;
    const to = axis === "rows" ? range.r2 : range.c2;
    const count = axis === "rows" ? store.tabs[tab].rows : store.tabs[tab].cols;
    const current = store.tabs[tab].layout[key] ?? [];
    let next: number[];
    if (hide) {
      if (to - from + 1 > 200_000) return toast("That is too many to hide at once.", "error");
      const set = new Set(current);
      for (let i = from; i <= to; i++) set.add(i);
      if (set.size >= count) return toast("Can't hide every " + (axis === "rows" ? "row" : "column") + ".", "error");
      next = [...set].sort((a, b) => a - b);
    } else {
      next = current.filter((i) => i < from - 1 || i > to + 1);
    }
    store
      .setLayout(tab, { [key]: next.length > 0 ? next : undefined })
      .then(() => gridRef.current && applyLayout(gridRef.current, store, tab))
      .catch(() => {});
  };
  /**
   * Inserts or deletes the selected rows/columns (as many as are selected). Formulas follow; it can't be
   * undone, so deleting asks first. Runs after any formula refresh that is still going.
   */
  const structure = (kind: "insert" | "delete", axis: "row" | "col", where: "before" | "after" = "before") => {
    if (!canEdit || !range) return;
    const first = axis === "row" ? range.r1 : range.c1;
    const last = axis === "row" ? range.r2 : range.c2;
    const count = last - first + 1;
    const word = axis === "row" ? "row" : "column";
    if (
      kind === "delete" &&
      !window.confirm(`Delete ${count.toLocaleString()} ${word}${count === 1 ? "" : "s"}? This can't be undone.`)
    ) {
      return;
    }
    const at = kind === "insert" && where === "after" ? last + 1 : first;
    const tabNo = tab;
    const job = refreshQueue.current.then(async () => {
      try {
        setCalculating("Updating…");
        await editStructure(store, index, tabNo, { kind, axis, index: at, count }, (done, total, phase) =>
          setCalculating(phase === "scan" ? `Looking for formulas… ${done} of ${total} parts` : `Updating formulas… ${done.toLocaleString()} of ${total.toLocaleString()}`),
        );
        const grid = gridRef.current;
        if (grid) {
          applyLayout(grid, store, tabNo);
          const rows = store.tabs[tabNo].rows;
          const cols = store.tabs[tabNo].cols;
          const row = axis === "row" ? Math.min(at, rows - 1) : (active?.row ?? 0);
          const col = axis === "col" ? Math.min(at, cols - 1) : (active?.col ?? 0);
          const area =
            kind === "insert"
              ? axis === "row"
                ? { r1: at, r2: at + count - 1, c1: range.c1, c2: range.c2 }
                : { r1: range.r1, r2: range.r2, c1: at, c2: at + count - 1 }
              : { r1: row, r2: row, c1: col, c2: col };
          grid.select({ ranges: [area], active: { row: area.r1, col: area.c1 }, anchor: { row: area.r1, col: area.c1 } }, false);
        }
        setVersion((n) => n + 1);
      } catch (error) {
        toast(error instanceof Error ? error.message : "Couldn't change the sheet.", "error");
      } finally {
        setCalculating(null);
      }
    });
    refreshQueue.current = job;
  };

  /** SUM / AVERAGE / ... of each selected column, written in the cell under the selection. */
  const quickFormula = async (name: string) => {
    if (!canEdit || !range) return;
    const tabNo = tab;
    const row = range.r2 + 1;
    if (range.r2 === range.r1 && range.c2 === range.c1) {
      toast("Select the cells to total first, then pick a formula.", "error");
      return;
    }
    if (row >= store.tabs[tabNo].rows) {
      toast("There is no row below the selection to put the result in.", "error");
      return;
    }
    try {
      setCalculating("Calculating…");
      await store.load(tabNo, store.chunkFor(tabNo, row));
      const written: ChangedRect[] = [];
      for (let col = range.c1; col <= range.c2; col++) {
        const formula = `${name}(${columnName(col)}${range.r1 + 1}:${columnName(col)}${range.r2 + 1})`;
        const result = await calculateTyped(store, tabNo, row, col, formula);
        await store.edit(tabNo, row, col, row, col, (cell) => ({
          ...(cell?.style ? { style: cell.style } : {}),
          formula,
          value: result.value,
          ...(result.error ? { type: "error" as const } : {}),
        }));
        written.push(pointRect({ tab: tabNo, row, col }));
      }
      await refresh(written);
      gridRef.current?.select({ ranges: [{ r1: row, r2: row, c1: range.c1, c2: range.c2 }], active: { row, col: range.c1 }, anchor: { row, col: range.c1 } });
    } catch (error) {
      toast(error instanceof Error ? error.message : "Couldn't calculate that.", "error");
    } finally {
      setCalculating(null);
    }
  };

  /** Everyone who can be tagged: the owner, the people the sheet is shared with, and you. */
  const peopleChoices = useMemo<PersonChoice[]>(() => {
    const map = new Map<string, PersonChoice>();
    if (entry.ownerPubkey) {
      const profile = profiles.profiles.get(entry.ownerPubkey);
      map.set(entry.ownerPubkey, {
        pubkey: entry.ownerPubkey,
        npub: toNpub(entry.ownerPubkey),
        name: profile?.displayName || profile?.name || entry.ownerName,
        picture: profile?.picture || entry.ownerPicture,
        role: "owner",
      });
    }
    for (const person of entry.sharedWith ?? []) {
      const profile = profiles.profiles.get(person.pubkey);
      map.set(person.pubkey, {
        pubkey: person.pubkey,
        npub: person.npub,
        name: profile?.displayName || profile?.name || person.name,
        picture: profile?.picture || person.picture,
        role: person.role,
      });
    }
    if (auth.pubkey && !map.has(auth.pubkey)) {
      const profile = profiles.profiles.get(auth.pubkey);
      map.set(auth.pubkey, {
        pubkey: auth.pubkey,
        npub: auth.npub || toNpub(auth.pubkey),
        name: profile?.displayName || profile?.name,
        picture: profile?.picture,
        role: entry.role === "owner" ? "owner" : entry.role === "editor" ? "editor" : "viewer",
      });
    }
    return Array.from(map.values());
  }, [auth.npub, auth.pubkey, entry, profiles.profiles]);

  /** Writes a person into the selected cells (their name is the cell's text). */
  const applyPerson = (user: UserRef) => {
    const label = user.name || shortNpub(user.pubkey);
    editRange((cell) => {
      const { tag: _tag, ...rest } = cell ?? {};
      void _tag;
      return { ...rest, value: label, person: { pubkey: user.pubkey, npub: user.npub, name: user.name, picture: user.picture } };
    });
  };
  /** Writes a tag into the selected cells. */
  const applyTag = (tag: string) => {
    editRange((cell) => {
      const { person: _person, ...rest } = cell ?? {};
      void _person;
      return { ...rest, value: tag, tag };
    });
  };

  /** Writes the whole workbook as an .xlsx file, a part at a time, and downloads it. */
  const exporting = useRef(false);
  const downloadXlsx = async () => {
    if (exporting.current) return;
    exporting.current = true;
    try {
      const blob = await exportBigSheet(store, {
        title: entry.title,
        onProgress: (done, total) => setCalculating(`Preparing the .xlsx… ${done} of ${total} parts`),
      });
      downloadBlob(`${fileSafeName(entry.title || "sheet")}.xlsx`, blob);
      toast("The .xlsx file is ready. It has your latest edits, saved or not.", "success");
    } catch (error) {
      toast(error instanceof Error ? error.message : "Couldn't prepare the .xlsx file.", "error");
    } finally {
      exporting.current = false;
      setCalculating(null);
    }
  };

  /** Turns "example.com" into a link a browser may follow (http, https or mailto only), or null. */
  const cleanUrl = (text: string): string | null => {
    const trimmed = text.trim();
    if (!trimmed) return null;
    const withScheme = /^[a-z][a-z0-9+.-]*:/i.test(trimmed) ? trimmed : `https://${trimmed}`;
    return /^(https?:|mailto:)/i.test(withScheme) ? withScheme : null;
  };

  /** Asks for an address and links the selected cells to it (an empty answer leaves things as they are). */
  const editLink = () => {
    if (!canEdit || !range) return;
    const current = active ? store.cell(tab, active.row, active.col)?.link?.url : undefined;
    const asked = window.prompt("Link address (https://…, or an email as mailto:name@example.com):", current ?? "https://");
    if (asked === null) return;
    const url = cleanUrl(asked);
    if (!url) return toast("Only web (http/https) and mailto links are allowed.", "error");
    editRange((cell) => ({ ...(cell ?? {}), link: { url }, ...(cell?.value === undefined && !cell?.formula ? { value: url } : {}) }));
  };
  const removeLink = () => editRange((cell) => (cell?.link ? (({ link: _link, ...rest }) => (Object.keys(rest).length > 0 ? rest : undefined))(cell) : cell));
  const openLink = () => {
    const url = active ? store.cell(tab, active.row, active.col)?.link?.url : undefined;
    if (url && cleanUrl(url)) window.open(url, "_blank", "noopener,noreferrer");
  };

  /** Adds, edits or (with an empty answer) deletes the comment of the active cell. */
  const editComment = () => {
    if (!canEdit || !active) return;
    const current = store.cell(tab, active.row, active.col)?.comment ?? "";
    const asked = window.prompt("Comment for this cell (leave empty to delete it):", current);
    if (asked === null) return;
    const text = asked.trim();
    store
      .edit(tab, active.row, active.col, active.row, active.col, (cell) => {
        const { comment: _comment, ...rest } = cell ?? {};
        void _comment;
        const next = text ? { ...rest, comment: text } : rest;
        return Object.keys(next).length > 0 ? next : undefined;
      })
      .catch((error) => toast(error instanceof Error ? error.message : String(error), "error"));
  };

  /** Sorts the selected rows by the column of the active cell. */
  const sort = (descending: boolean, options: { col?: number; header?: boolean } = {}) => {
    if (!canEdit || !range || !active) return;
    const tabNo = tab;
    // With a header row the first selected row stays where it is.
    const area = options.header ? { ...range, r1: range.r1 + 1 } : range;
    const key = Math.min(Math.max(options.col ?? active.col, range.c1), range.c2);
    sortArea(store, tabNo, area, key, descending, store.tabs[tabNo].layout.merges ?? [])
      .then(async (moved) => {
        if (moved) await refresh([{ tab: tabNo, r1: area.r1, c1: area.c1, r2: area.r2, c2: area.c2 }]);
      })
      .catch((error) => toast(error instanceof Error ? error.message : "Couldn't sort.", "error"));
  };

  /** Asks for a size and gives it to every selected row (or column). */
  const askSize = (axis: "rows" | "columns") => {
    if (!canEdit || !range) return;
    const rowsAxis = axis === "rows";
    const from = rowsAxis ? range.r1 : range.c1;
    const to = rowsAxis ? range.r2 : range.c2;
    const asked = window.prompt(rowsAxis ? "Row height in pixels (8–545):" : "Column width in pixels (8–2038):", String(rowsAxis ? 30 : 116));
    if (asked === null) return;
    const size = Math.round(Number(asked));
    const max = rowsAxis ? 545 : 2038;
    if (!Number.isFinite(size) || size < 8 || size > max) return toast(`Enter a number from 8 to ${max}.`, "error");
    if (to - from + 1 > 200_000) return toast("That is too many to resize at once.", "error");
    const key = rowsAxis ? "rowHeights" : "columnWidths";
    const next = { ...(store.tabs[tab].layout[key] ?? {}) };
    for (let i = from; i <= to; i++) next[i] = size;
    store
      .setLayout(tab, { [key]: next })
      .then(() => gridRef.current && applyLayout(gridRef.current, store, tab))
      .catch(() => {});
  };
  const hiddenNear = (axis: "rows" | "columns") => {
    if (!range) return false;
    const list = store.tabs[tab]?.layout[axis === "rows" ? "hiddenRows" : "hiddenColumns"] ?? [];
    const from = axis === "rows" ? range.r1 : range.c1;
    const to = axis === "rows" ? range.r2 : range.c2;
    return list.some((i) => i >= from - 1 && i <= to + 1);
  };

  const frozen = store.tabs[tab]?.layout.frozen ?? { rows: 0, columns: 0 };
  const freezeOptions = [
    { label: "No frozen rows or columns", rows: 0, columns: 0 },
    { label: "1 row", rows: 1, columns: 0 },
    { label: "2 rows", rows: 2, columns: 0 },
    { label: "1 column", rows: 0, columns: 1 },
    { label: "1 row and 1 column", rows: 1, columns: 1 },
    ...(active
      ? [
          { label: `Up to row ${active.row + 1}`, rows: active.row + 1, columns: frozen.columns },
          { label: `Up to column ${columnName(active.col)}`, rows: frozen.rows, columns: active.col + 1 },
        ]
      : []),
  ];

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

  const commit = async (move: "down" | "next" | "none") => {
    if (!editing) return;
    const { row, col, text } = editing;
    setEditing(null);
    if (text.trim().startsWith("=") && text.trim().length > 1) {
      const formula = text.trim().slice(1);
      try {
        setCalculating("Calculating…");
        const result = await calculateTyped(store, tab, row, col, formula, (done, total) =>
          setCalculating(total > 1 ? `Calculating… chunk ${done} of ${total}` : "Calculating…"),
        );
        await store.edit(tab, row, col, row, col, (cell) => ({
          ...(cell?.style ? { style: cell.style } : {}),
          formula,
          value: result.value,
          ...(result.error ? { type: "error" as const } : {}),
        }));
      } catch (error) {
        toast(error instanceof Error ? error.message : "Couldn't calculate that formula.", "error");
      } finally {
        setCalculating(null);
      }
    } else {
      const rule = store.cell(tab, row, col)?.validation;
      const problem = rule ? validate(rule, text) : null;
      if (problem) {
        // Rejected: the cell keeps its old content and the editor stays open for a fix.
        toast(problem, "error");
        setEditing({ row, col, text });
        return;
      }
      await store.setCell(tab, row, col, text);
    }
    await refresh([pointRect({ tab, row, col })]);
    gridRef.current?.afterEdit(move);
    gridRef.current?.focus();
  };

  const undo = () => {
    const cells = store.undo();
    if (store.structuralUndo) {
      // Rows, columns or tabs came back: show them again and look for formulas afresh.
      const visible = visibleTabs();
      const next = Math.min(tab, store.tabs.length - 1);
      showTabs(visible.includes(next) ? next : (visible[0] ?? 0), true);
      return;
    }
    void refresh(cells.map(pointRect));
  };
  const redo = () => void refresh(store.redo().map(pointRect));

  const applyStyle = (patch: StylePatch) => editRange((cell) => restyle(cell, patch));

  const selectionLabel = () => {
    if (!range) return "";
    const first = `${columnName(range.c1)}${range.r1 + 1}`;
    return range.r1 === range.r2 && range.c1 === range.c2 ? first : `${first}:${columnName(range.c2)}${range.r2 + 1}`;
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

  /** Functions that start like the word being typed after "=" (or after an operator or bracket). */
  const suggestions = (() => {
    if (!editing || !functionList || !editing.text.startsWith("=")) return [] as string[];
    const word = /(?:^=|[-+*/^&(,<>=:;\s])([A-Za-z][A-Za-z0-9_.]*)$/.exec(editing.text)?.[1];
    if (!word) return [];
    const upper = word.toUpperCase();
    const found = functionList.filter((name) => name.startsWith(upper) && name !== upper);
    return found.slice(0, 8);
  })();
  /** Which function the cursor is in while a formula is typed, with its arguments. */
  const help = editing ? helpFor(editing.text) : null;
  /** Replaces the word being typed with the chosen function and an opening bracket. */
  const acceptSuggestion = (name: string) => {
    if (!editing) return;
    setEditing({ ...editing, text: editing.text.replace(/[A-Za-z][A-Za-z0-9_.]*$/, `${name}(`) });
    setSuggestAt(0);
  };

  /** Count / sum / average of the numbers in the selection (only the parts already loaded, up to 50,000 cells). */
  const totals = (() => {
    if (!range || (range.r1 === range.r2 && range.c1 === range.c2)) return "";
    const last = Math.min(range.r2, range.r1 + 49_999);
    let count = 0;
    let sum = 0;
    let filled = 0;
    let cells = 0;
    outer: for (let r = range.r1; r <= last; r++) {
      if (!store.isLoaded(tab, r)) continue;
      for (let c = range.c1; c <= range.c2; c++) {
        if (++cells > 50_000) break outer;
        const cell = store.cell(tab, r, c);
        if (cell?.value === undefined || cell.value === "") continue;
        filled++;
        if (typeof cell.value === "number" && Number.isFinite(cell.value)) {
          count++;
          sum += cell.value;
        }
      }
    }
    if (filled === 0) return "";
    const fmt = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: 4 });
    return count > 0 ? `Count ${filled.toLocaleString()} · Sum ${fmt(sum)} · Avg ${fmt(sum / count)}` : `Count ${filled.toLocaleString()}`;
  })();

  const selectedCell = active ? store.cell(tab, active.row, active.col) : undefined;
  const rows = info.rows;
  const grid = gridRef.current;
  const editRect = editing && grid ? grid.cellRect(editing) : null;

  return (
    <div
      className="big-sheet"
      tabIndex={-1}
      onKeyDown={(event) => {
        if (editing) return;
        const mod = event.ctrlKey || event.metaKey;
        const key = event.key.toLowerCase();
        if (event.key === "F9") {
          event.preventDefault();
          void refresh([], false, true, true);
          return;
        }
        if (mod && !event.shiftKey && !event.altKey && (key === "f" || key === "h")) {
          event.preventDefault();
          setFinding({ replace: key === "h" });
          return;
        }
        if (mod && canEdit && active && (event.key === ";" || event.key === ":")) {
          // Ctrl+; puts today's date in the cell, Ctrl+Shift+; the time.
          event.preventDefault();
          const now = new Date();
          const pad = (n: number) => String(n).padStart(2, "0");
          const text =
            event.key === ";"
              ? `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`
              : `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
          void store.setCell(tab, active.row, active.col, text).then(() => refresh([pointRect({ tab, row: active.row, col: active.col })]));
          return;
        }
        if (mod && canEdit && (key === "z" || key === "y")) {
          event.preventDefault();
          if (key === "y" || event.shiftKey) redo();
          else undo();
          return;
        }
        if (mod && canEdit && selection && (key === "b" || key === "i" || key === "u")) {
          event.preventDefault();
          const flag = key === "b" ? "bold" : key === "i" ? "italic" : "underline";
          applyStyle({ [flag]: !selectedCell?.style?.[flag] });
        }
      }}
    >
      <BigSheetToolbar
        style={selectedCell?.style ?? {}}
        canEdit={canEdit && !!selection}
        canUndo={store.canUndo}
        canRedo={store.canRedo}
        painting={painting !== null}
        onUndo={undo}
        onRedo={redo}
        onPaint={() => setPainting(painting !== null ? null : { ...(selectedCell?.style ?? {}) })}
        onClear={() =>
          editRange((cell) =>
            cell ? { ...(cell.value !== undefined ? { value: cell.value } : {}), ...(cell.formula ? { formula: cell.formula } : {}) } : undefined,
          )
        }
        onStyle={applyStyle}
        onValidation={() => setValidating(true)}
        onPerson={() => setMetaDialog("person")}
        onTagPill={() => setMetaDialog("tag")}
        onConditional={() => setFormatting(true)}
        filtered={!!store.tabs[tab].layout.filter}
        onFilter={() => setFiltering(true)}
        onLink={editLink}
        onComment={editComment}
        onQuickFormula={(name) => void quickFormula(name)}
        onSearch={() => setFinding({ replace: false })}
        zoom={{ value: gridRef.current?.zoom ?? 1, set: (z) => gridRef.current?.setZoom(z, true) }}
        onBorder={border}
        onMerge={(how) => void merge(how)}
        canSplit={mergedHere}
        freeze={{ rows: frozen.rows, columns: frozen.columns, options: freezeOptions, apply: freeze }}
        onFormat={(format) => applyStyle({ numberFormat: format })}
        onDecimals={(delta) => editRange((cell) => restyle(cell, { numberFormat: shiftDecimals(cell?.style?.numberFormat, delta) }))}
      />
      <div className="big-sheet-bar">
        <input
          className="big-sheet-ref"
          aria-label="Go to cell"
          value={nameDraft ?? selectionLabel()}
          onFocus={(event) => {
            setNameDraft(selectionLabel());
            event.target.select();
          }}
          onChange={(event) => setNameDraft(event.target.value)}
          onBlur={() => setNameDraft(null)}
          onKeyDown={(event) => {
            if (event.key === "Escape") {
              setNameDraft(null);
              gridRef.current?.focus();
            } else if (event.key === "Enter") {
              event.preventDefault();
              const target = parseReference(nameDraft ?? "");
              if (!target) return toast("Type a cell like B7 or a range like A1:C5.", "error");
              const r1 = Math.min(target.r1, rows - 1);
              const c1 = Math.min(target.c1, info.cols - 1);
              const area = { r1, c1, r2: Math.min(target.r2, rows - 1), c2: Math.min(target.c2, info.cols - 1) };
              gridRef.current?.select({ ranges: [area], active: { row: r1, col: c1 }, anchor: { row: r1, col: c1 } });
              setNameDraft(null);
              gridRef.current?.focus();
            }
          }}
        />
        <span className="big-sheet-fx">fx</span>
        <span className="big-sheet-formula">
          {selectedCell?.formula
            ? [{ text: "=", kind: "operator" as const, color: undefined as number | undefined }, ...tokenizeFormula(selectedCell.formula.replace(/^=/, ""))].map((token, i) => (
                <span key={i} className={`ft-${token.kind}`} style={token.kind === "ref" ? { color: REF_COLORS[(token.color ?? 0) % REF_COLORS.length] } : undefined}>
                  {token.text}
                </span>
              ))
            : shown(selectedCell)}
        </span>
        <button
          type="button"
          className="big-sheet-calc"
          title={manualCalc ? "Manual calculation: formulas update only when you recalculate (F9). Click to switch to automatic." : "Automatic calculation. Click to switch to manual."}
          onClick={() => setManualCalc((on) => !on)}
        >
          {manualCalc ? "Calc: manual" : "Calc: auto"}
        </button>
        {manualCalc && (
          <button type="button" className="big-sheet-calc is-stale" disabled={!!calculating} onClick={() => void refresh([], false, true, true)} title="Recalculate every formula (F9)">
            {stale ? "Recalculate (F9) •" : "Recalculate (F9)"}
          </button>
        )}
        <input
          className="big-sheet-goto"
          placeholder="Go to row"
          inputMode="numeric"
          onKeyDown={(event) => {
            if (event.key !== "Enter") return;
            const row = Math.max(0, Math.min(rows - 1, Number((event.target as HTMLInputElement).value.replace(/\D/g, "")) - 1));
            gridRef.current?.scrollToRow(row);
            gridRef.current?.select({ ranges: [{ r1: row, r2: row, c1: active?.col ?? 0, c2: active?.col ?? 0 }], active: { row, col: active?.col ?? 0 }, anchor: { row, col: active?.col ?? 0 } });
          }}
        />
        <span className="big-sheet-note">
          {calculating ?? (store.tabs[tab].layout.filter && !selectedCell?.comment ? `Filter on column ${columnName(store.tabs[tab].layout.filter.col)} ${store.tabs[tab].layout.filter.summary} · ${store.tabs[tab].layout.filter.hidden.length.toLocaleString()} rows hidden` : selectedCell?.comment ? `💬 ${selectedCell.comment}` : totals || `${rows.toLocaleString()} rows · large sheet mode${canEdit ? "" : " · read-only"}`)}
        </span>
        <button type="button" className="big-sheet-calc" disabled={!!calculating} onClick={() => void downloadXlsx()} title="Download every tab as an .xlsx file">
          Download .xlsx
        </button>
        {canEdit && (
          <button type="button" className="btn btn-primary btn-sm" disabled={saving || store.dirtyCount === 0} onClick={() => void save()}>
            {saving ? <Loader2 size={14} className="spin" /> : <Save size={14} />}
            {store.dirtyCount ? `Save ${store.dirtyCount} changed chunk${store.dirtyCount === 1 ? "" : "s"}` : "Saved"}
          </button>
        )}
      </div>
      <PersonDialog open={metaDialog === "person"} onClose={() => setMetaDialog(null)} onPick={applyPerson} people={peopleChoices} />
      <TagDialog open={metaDialog === "tag"} onClose={() => setMetaDialog(null)} onPick={applyTag} />
      {sortDialog && range && active && (
        <BigSortDialog
          firstCol={range.c1}
          lastCol={range.c2}
          activeCol={active.col}
          rows={range.r2 - range.r1 + 1}
          onClose={() => setSortDialog(false)}
          onSort={({ col, descending, header }) => {
            setSortDialog(false);
            sort(descending, { col, header });
          }}
        />
      )}
      {filtering && active && (
        <BigFilterDialog
          key={`${tab}:${active.col}`}
          store={store}
          tab={tab}
          col={store.tabs[tab].layout.filter?.col ?? active.col}
          columnLabel={columnName(store.tabs[tab].layout.filter?.col ?? active.col)}
          current={store.tabs[tab].layout.filter}
          onClose={() => setFiltering(false)}
          onApply={(filter) => {
            setFiltering(false);
            store
              .setLayout(tab, { filter })
              .then(() => gridRef.current && applyLayout(gridRef.current, store, tab))
              .catch(() => {});
          }}
          onClear={() => {
            setFiltering(false);
            store
              .setLayout(tab, { filter: undefined })
              .then(() => gridRef.current && applyLayout(gridRef.current, store, tab))
              .catch(() => {});
          }}
        />
      )}
      {formatting && range && (
        <BigConditionalDialog
          rangeLabel={selectionLabel()}
          hasRules={(store.tabs[tab].layout.conditionalFormats ?? []).some((rule) =>
            rule.ranges.some((r) => r.r1 <= range.r2 && r.r2 >= range.r1 && r.c1 <= range.c2 && r.c2 >= range.c1),
          )}
          onClose={() => setFormatting(false)}
          onSave={(condition, format) => {
            setFormatting(false);
            const area = { r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c2 };
            const tabNo = tab;
            const add = (final: typeof condition) => {
              const rules = store.tabs[tabNo].layout.conditionalFormats ?? [];
              // A new rule goes first so it wins over older ones on the same cells.
              return store.setLayout(tabNo, { conditionalFormats: [{ ranges: [area], condition: final, format }, ...rules] });
            };
            if (!SCALE_TYPES.includes(condition.type)) {
              add(condition).catch(() => {});
              return;
            }
            // Bars and scales compare with the lowest and highest number of the range: read it once now.
            setCalculating("Reading the range…");
            numericExtent(store, tabNo, area)
              .then((extent) => {
                if (!extent) return toast("There are no numbers in the selection to colour.", "error");
                return add({ type: condition.type, values: [String(extent.min), String(extent.max)] });
              })
              .catch((error) => toast(error instanceof Error ? error.message : "Couldn't read the range.", "error"))
              .finally(() => setCalculating(null));
          }}
          onClearHere={() => {
            setFormatting(false);
            const left = withoutArea(store.tabs[tab].layout.conditionalFormats ?? [], range);
            store.setLayout(tab, { conditionalFormats: left.length > 0 ? left : undefined }).catch(() => {});
          }}
        />
      )}
      {validating && range && (
        <BigValidationDialog
          current={selectedCell?.validation}
          rangeLabel={selectionLabel()}
          onClose={() => setValidating(false)}
          onSave={(rule) => {
            setValidating(false);
            editRange((cell) => ({ ...(cell ?? {}), validation: rule }));
          }}
          onRemove={() => {
            setValidating(false);
            editRange((cell) => {
              if (!cell?.validation) return cell;
              const { validation: _validation, ...rest } = cell;
              void _validation;
              return Object.keys(rest).length > 0 ? rest : undefined;
            });
          }}
        />
      )}
      {finding && (
        <BigFindBar
          key={finding.replace ? "replace" : "find"}
          store={store}
          tab={tab}
          canEdit={canEdit}
          withReplace={finding.replace}
          onGo={(match) => {
            const grid = gridRef.current;
            if (!grid) return;
            const at = { row: match.row, col: match.col };
            grid.select({ ranges: [{ r1: at.row, c1: at.col, r2: at.row, c2: at.col }], active: at, anchor: at });
          }}
          onReplaced={(changed) => void refresh(changed.map((m) => pointRect({ tab, row: m.row, col: m.col })))}
          onClose={() => {
            setFinding(null);
            gridRef.current?.focus();
          }}
        />
      )}
      <div className="big-sheet-scroll big-sheet-canvas">
        <div ref={hostRef} className="big-sheet-host" />
        {editing && editRect && help && (
          <div className="big-help" style={{ top: Math.max(0, editRect.y - 30), left: editRect.x }} role="status">
            <b>{help.name}</b>(
            {help.args.map((arg, i) => (
              <span key={i}>
                {i > 0 && ", "}
                <span className={i === help.current ? "big-help-current" : undefined}>{arg}</span>
              </span>
            ))}
            )
          </div>
        )}
        {editing && editRect && suggestions.length > 0 && (
          <ul className="big-suggest" style={{ top: editRect.y + editRect.height, left: editRect.x }} role="listbox">
            {suggestions.map((name, index) => (
              <li
                key={name}
                role="option"
                aria-selected={index === suggestAt}
                className={index === suggestAt ? "is-active" : undefined}
                onPointerDown={(event) => {
                  event.preventDefault(); // keep the editor focused
                  acceptSuggestion(name);
                }}
              >
                {name}
              </li>
            ))}
          </ul>
        )}
        {editing && (
          <datalist id="big-validation-options">
            {(store.cell(tab, editing.row, editing.col)?.validation?.options ?? []).map((option) => (
              <option key={option} value={option} />
            ))}
          </datalist>
        )}
        {editing && editRect && (
          <input
            className="big-cell-input"
            style={{ top: editRect.y, left: editRect.x, width: Math.max(editRect.width, 120), height: editRect.height }}
            autoFocus
            value={editing.text}
            list={store.cell(tab, editing.row, editing.col)?.validation?.type === "list" ? "big-validation-options" : undefined}
            onChange={(event) => {
              setEditing({ ...editing, text: event.target.value });
              setSuggestAt(0);
              if (event.target.value.startsWith("=") && !functionList) {
                void import("@formstr/sheets-formula").then((engine) => setFunctionList(engine.functionNames()));
              }
            }}
            onBlur={() => {
              if (cancelled.current) cancelled.current = false;
              else void commit("none");
            }}
            onKeyDown={(event) => {
              if (suggestions.length > 0) {
                if (event.key === "ArrowDown" || event.key === "ArrowUp") {
                  event.preventDefault();
                  setSuggestAt((at) => (at + (event.key === "ArrowDown" ? 1 : suggestions.length - 1)) % suggestions.length);
                  return;
                }
                if (event.key === "Tab" || event.key === "Enter") {
                  event.preventDefault();
                  acceptSuggestion(suggestions[Math.min(suggestAt, suggestions.length - 1)]);
                  return;
                }
              }
              if (event.key === "Enter") {
                event.preventDefault();
                void commit("down");
              } else if (event.key === "Tab") {
                event.preventDefault();
                void commit("next");
              } else if (event.key === "Escape") {
                cancelled.current = true;
                setEditing(null);
                gridRef.current?.focus();
              }
            }}
          />
        )}
      </div>
      {menu && range && (
        <>
          <div className="big-menu-backdrop" onPointerDown={() => setMenu(null)} onContextMenu={(event) => (event.preventDefault(), setMenu(null))} />
          <div className="big-menu" style={{ top: menu.y, left: menu.x }} role="menu">
            {(
              [
                [`Hide row${range.r2 > range.r1 ? "s" : ""} ${range.r1 + 1}${range.r2 > range.r1 ? "–" + (range.r2 + 1) : ""}`, () => setHidden("rows", true), true],
                [`Hide column${range.c2 > range.c1 ? "s" : ""} ${columnName(range.c1)}${range.c2 > range.c1 ? "–" + columnName(range.c2) : ""}`, () => setHidden("columns", true), true],
                [`Insert ${range.r2 - range.r1 + 1 > 1 ? range.r2 - range.r1 + 1 + " rows" : "row"} above`, () => structure("insert", "row", "before"), true],
                [`Insert ${range.r2 - range.r1 + 1 > 1 ? range.r2 - range.r1 + 1 + " rows" : "row"} below`, () => structure("insert", "row", "after"), true],
                [`Delete row${range.r2 > range.r1 ? "s" : ""}`, () => structure("delete", "row"), true],
                [`Insert ${range.c2 - range.c1 + 1 > 1 ? range.c2 - range.c1 + 1 + " columns" : "column"} left`, () => structure("insert", "col", "before"), true],
                [`Insert ${range.c2 - range.c1 + 1 > 1 ? range.c2 - range.c1 + 1 + " columns" : "column"} right`, () => structure("insert", "col", "after"), true],
                [`Delete column${range.c2 > range.c1 ? "s" : ""}`, () => structure("delete", "col"), true],
                [`Sort rows A → Z by column ${columnName(Math.min(Math.max(active?.col ?? range.c1, range.c1), range.c2))}`, () => sort(false), range.r2 > range.r1],
                [`Sort rows Z → A by column ${columnName(Math.min(Math.max(active?.col ?? range.c1, range.c1), range.c2))}`, () => sort(true), range.r2 > range.r1],
                ["Insert or edit link…", editLink, true],
                ["Open link", openLink, !!selectedCell?.link],
                ["Remove link", removeLink, !!selectedCell?.link],
                [selectedCell?.comment ? "Edit comment…" : "Add comment…", editComment, true],
                ["Sort range…", () => setSortDialog(true), range.r2 > range.r1],
                ["Set row height…", () => askSize("rows"), true],
                ["Set column width…", () => askSize("columns"), true],
                ["Show hidden rows here", () => setHidden("rows", false), hiddenNear("rows")],
                ["Show hidden columns here", () => setHidden("columns", false), hiddenNear("columns")],
              ] as [string, () => void, boolean][]
            ).map(([label, run, enabled]) => (
              <button
                key={label}
                type="button"
                role="menuitem"
                disabled={!canEdit || !enabled}
                onClick={() => {
                  setMenu(null);
                  run();
                }}
              >
                {label}
              </button>
            ))}
          </div>
        </>
      )}
      {tabMenu && (
        <>
          <div className="big-menu-backdrop" onPointerDown={() => setTabMenu(null)} onContextMenu={(event) => (event.preventDefault(), setTabMenu(null))} />
          <div className="big-menu" style={{ bottom: window.innerHeight - tabMenu.y, left: tabMenu.x }} role="menu">
            {(
              [
                ["Rename…", () => renameTab(tabMenu.index), true],
                ["Make a copy", () => copyTab(tabMenu.index), true],
                ["Delete", () => deleteTab(tabMenu.index), store.tabs.length > 1],
                ["Move left", () => moveTab(tabMenu.index, -1), tabMenu.index > 0],
                ["Move right", () => moveTab(tabMenu.index, 1), tabMenu.index < store.tabs.length - 1],
                ["Hide tab", () => hideTab(tabMenu.index, true), visibleTabs().length > 1],
                ["Remove colour", () => (store.setTabColor(tabMenu.index, undefined), setVersion((n) => n + 1)), !!store.tabMeta(tabMenu.index).color],
              ] as [string, () => void, boolean][]
            ).map(([label, run, enabled]) => (
              <button key={label} type="button" role="menuitem" disabled={!canEdit || !enabled} onClick={() => (setTabMenu(null), run())}>
                {label}
              </button>
            ))}
            <div className="big-menu-swatches">
              {TAB_COLORS.map((color) => (
                <button
                  key={color}
                  type="button"
                  aria-label={`Colour ${color}`}
                  className="big-menu-swatch"
                  style={{ background: color }}
                  disabled={!canEdit}
                  onClick={() => {
                    store.setTabColor(tabMenu.index, color);
                    setTabMenu(null);
                    setVersion((n) => n + 1);
                  }}
                />
              ))}
            </div>
            {store.tabs.some((_, i) => store.tabMeta(i).hidden) && (
              <>
                <div className="big-menu-title">Hidden tabs</div>
                {store.tabs.map((item, i) =>
                  store.tabMeta(i).hidden ? (
                    <button key={item.name + i} type="button" role="menuitem" disabled={!canEdit} onClick={() => (setTabMenu(null), hideTab(i, false))}>
                      Show "{item.name}"
                    </button>
                  ) : null,
                )}
              </>
            )}
          </div>
        </>
      )}
      <div className="big-sheet-tabs">
        {store.tabs.map((item, index) =>
          store.tabMeta(index).hidden ? null : (
            <button
              key={item.name + index}
              type="button"
              className={cx("big-sheet-tab", index === tab && "is-active")}
              style={store.tabMeta(index).color ? { boxShadow: `inset 0 -3px 0 ${store.tabMeta(index).color}` } : undefined}
              onClick={() => switchTab(index)}
              onDoubleClick={() => renameTab(index)}
              onContextMenu={(event) => {
                event.preventDefault();
                setTabMenu({ x: event.clientX, y: event.clientY, index });
              }}
            >
              {item.name}
            </button>
          ),
        )}
        {canEdit && (
          <button type="button" className="big-sheet-tab" title="Add a tab" aria-label="Add a tab" onClick={addTab}>
            +
          </button>
        )}
      </div>
    </div>
  );
}
