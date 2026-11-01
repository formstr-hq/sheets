import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { useLocation, useNavigate, useParams } from "react-router-dom";
import { Workbook, type WorkbookInstance } from "@fortune-sheet/react";
import type { Hooks, Settings, Sheet } from "@fortune-sheet/core";
import type { Event } from "nostr-tools";
import {
  AlertTriangle,
  ArrowLeft,
  CheckCircle2,
  CloudOff,
  CloudUpload,
  Copy,
  Eye,
  FileDown,
  FileSpreadsheet,
  KeyRound,
  Loader2,
  LockKeyhole,
  MoreHorizontal,
  RefreshCw,
  Share2,
  Tag,
  Trash2,
  UserRound,
} from "lucide-react";
import { Avatar } from "./Avatar";
import { Dialog } from "./Dialog";
import { Menu, MenuItem, MenuSeparator } from "./Menu";
import { LogoMark } from "./Logo";
import { PersonDialog, TagDialog, type PersonChoice } from "./CellMetaDialogs";
import { ShareDialog } from "./ShareDialog";
import { useAuth } from "../providers/AuthProvider";
import { useLibrary } from "../providers/LibraryProvider";
import { useProfiles } from "../providers/ProfileProvider";
import { useToast } from "../providers/ToastProvider";
import { drawCellPill } from "../lib/cellRender";
import { saveDocCache } from "../lib/local";
import { sheetsForUpdate } from "../lib/sheetData";
import {
  collectReferencedPubkeys,
  fortuneSheetsToWorkbook,
  workbookToFortuneSheets,
} from "../lib/workbook";
import { exportActiveSheetCsv, exportSpreadsheet } from "../lib/xlsx";
import { formatRelativeTime, shortNpub, toNpub } from "../lib/utils";
import type { LibraryEntry, SaveState, SheetDocument, UserRef } from "../lib/types";

const TOOLBAR_ITEMS = [
  "undo",
  "redo",
  "|",
  "format-painter",
  "clear-format",
  "|",
  "currency-format",
  "percentage-format",
  "number-decrease",
  "number-increase",
  "format",
  "|",
  "font",
  "|",
  "font-size",
  "|",
  "bold",
  "italic",
  "strike-through",
  "underline",
  "|",
  "font-color",
  "background",
  "border",
  "merge-cell",
  "|",
  "horizontal-align",
  "vertical-align",
  "text-wrap",
  "text-rotation",
  "|",
  "freeze",
  "conditionFormat",
  "filter",
  "link",
  "comment",
  "quick-formula",
  "dataVerification",
  "search",
];

const CELL_CONTEXT_MENU = [
  "copy",
  "paste",
  "|",
  "insert-row",
  "insert-column",
  "delete-row",
  "delete-column",
  "delete-cell",
  "hide-row",
  "hide-column",
  "set-row-height",
  "set-column-width",
  "|",
  "clear",
  "sort",
  "orderAZ",
  "orderZA",
  "filter",
  "data",
  "cell-format",
];

const SHEET_TAB_CONTEXT_MENU = ["delete", "copy", "rename", "color", "hide", "|", "move"];

interface SelectionRange {
  row: number[];
  column: number[];
}

type RemoteUpdate = { doc: SheetDocument; event: Event };

function roleRank(role: LibraryEntry["role"]): number {
  return role === "owner" ? 2 : role === "editor" ? 1 : 0;
}

function saveLabel(state: SaveState, lastSavedAt?: number): string {
  switch (state) {
    case "saving":
      return "Saving…";
    case "unsaved":
      return "Unsaved changes";
    case "offline":
      return "Saved locally";
    case "readonly":
      return "View only";
    case "saved":
      return lastSavedAt ? `Saved ${formatRelativeTime(lastSavedAt).toLowerCase()}` : "Saved";
    default:
      return "Loading…";
  }
}

export function Editor() {
  const { naddr = "" } = useParams<{ naddr: string }>();
  const location = useLocation();
  const navigate = useNavigate();
  const auth = useAuth();
  const library = useLibrary();
  const profiles = useProfiles();
  const { toast } = useToast();

  const existing = library.getEntryByNaddr(naddr);
  const parsedLink = useMemo(
    () => library.parseShareLink(naddr, location.hash),
    [library.parseShareLink, location.hash, naddr],
  );
  const entry = useMemo<LibraryEntry | null>(() => {
    if (!existing) return parsedLink;
    if (!parsedLink) return existing;
    return {
      ...existing,
      ...parsedLink,
      role: roleRank(parsedLink.role) > roleRank(existing.role) ? parsedLink.role : existing.role,
      editKey: parsedLink.editKey ?? existing.editKey,
      viewKey: parsedLink.viewKey || existing.viewKey,
      title: existing.title || parsedLink.title,
      ownerPubkey: existing.ownerPubkey || parsedLink.ownerPubkey,
      ownerName: existing.ownerName ?? parsedLink.ownerName,
      ownerPicture: existing.ownerPicture ?? parsedLink.ownerPicture,
      color: existing.color ?? parsedLink.color,
      sharedWith: existing.sharedWith,
      source: existing.source ?? parsedLink.source,
      lastEventAt: existing.lastEventAt ?? parsedLink.lastEventAt,
      lastEventId: existing.lastEventId ?? parsedLink.lastEventId,
    };
  }, [existing, parsedLink]);
  const entryKey = entry ? `${entry.coordinate}|${entry.viewKey}|${entry.editKey ?? ""}` : "";
  const entryRef = useRef<LibraryEntry | null>(entry);

  const [sheetDocument, setSheetDocument] = useState<SheetDocument | null>(null);
  const documentRef = useRef<SheetDocument | null>(null);
  const [initialSheets, setInitialSheets] = useState<Sheet[]>([]);
  const [workbookVersion, setWorkbookVersion] = useState(0);
  const [title, setTitle] = useState("");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [saveState, setSaveState] = useState<SaveState>("idle");
  const saveStateRef = useRef<SaveState>("idle");
  const [lastSavedAt, setLastSavedAt] = useState<number | undefined>(undefined);
  const [dirty, setDirty] = useState(false);
  const dirtyRef = useRef(false);
  const [remoteUpdate, setRemoteUpdate] = useState<RemoteUpdate | null>(null);
  const [shareOpen, setShareOpen] = useState(false);
  const [personOpen, setPersonOpen] = useState(false);
  const [tagOpen, setTagOpen] = useState(false);
  const [deleteOpen, setDeleteOpen] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const selectionRef = useRef<SelectionRange[]>([]);

  const workbookRef = useRef<WorkbookInstance>(null);
  const sheetsRef = useRef<Sheet[]>([]);
  const loadedRef = useRef(false);
  const applyingRemoteRef = useRef(false);
  const saveTimerRef = useRef<number | null>(null);
  const loadTimerRef = useRef<number | null>(null);
  const saveInFlightRef = useRef(false);
  const queuedSaveRef = useRef(false);
  const lastEventAtRef = useRef<number | undefined>(undefined);
  const lastEventIdRef = useRef<string | undefined>(undefined);
  const remoteCheckRef = useRef(false);
  const saveNowRef = useRef<() => Promise<void>>(async () => {});

  const canEdit = !!entry?.editKey;

  const updateSaveState = useCallback((next: SaveState) => {
    saveStateRef.current = next;
    setSaveState(next);
  }, []);

  useEffect(() => {
    entryRef.current = entry;
  }, [entry]);

  const buildCurrentWorkbook = useCallback(
    (currentDoc: SheetDocument, updatedAt = Date.now()) => {
      const activeId = sheetsRef.current.find((sheet) => sheet.status === 1)?.id;
      const sheets =
        sheetsRef.current.length > 0
          ? sheetsRef.current
          : workbookToFortuneSheets(currentDoc);
      return fortuneSheetsToWorkbook({
        title: currentDoc.title,
        createdAt: currentDoc.createdAt,
        updatedAt,
        activeSheetId: activeId ?? currentDoc.activeSheetId,
        sheets,
      });
    },
    [],
  );

  const markDirty = useCallback(() => {
    if (!dirtyRef.current) {
      dirtyRef.current = true;
      setDirty(true);
    }
    updateSaveState("unsaved");
  }, [updateSaveState]);

  const saveNow = useCallback(async () => {
    const currentEntry = entryRef.current;
    const currentDoc = documentRef.current;
    if (!currentEntry?.editKey || !currentDoc) {
      if (currentEntry && !currentEntry.editKey) updateSaveState("readonly");
      return;
    }
    if (saveInFlightRef.current) {
      queuedSaveRef.current = true;
      return;
    }
    saveInFlightRef.current = true;
    if (saveTimerRef.current) {
      window.clearTimeout(saveTimerRef.current);
      saveTimerRef.current = null;
    }
    updateSaveState("saving");
    const docToSave: SheetDocument = buildCurrentWorkbook(currentDoc, Date.now());
    documentRef.current = docToSave;
    try {
      const { event } = await library.saveDocument(currentEntry, docToSave);
      lastEventAtRef.current = event.created_at * 1000;
      lastEventIdRef.current = event.id;
      setSheetDocument(docToSave);
      setTitle(docToSave.title);
      setLastSavedAt(Date.now());
      dirtyRef.current = false;
      setDirty(false);
      updateSaveState("saved");
      if (currentEntry.title !== docToSave.title) {
        void library.updateTitle(currentEntry, docToSave.title).catch(() => {});
      }
    } catch (err) {
      dirtyRef.current = true;
      setDirty(true);
      updateSaveState("offline");
      toast(
        err instanceof Error ? err.message : "Saved on this device. We'll retry when relays respond.",
        "error",
      );
    } finally {
      saveInFlightRef.current = false;
      if (queuedSaveRef.current) {
        queuedSaveRef.current = false;
        window.setTimeout(() => void saveNowRef.current(), 350);
      }
    }
  }, [buildCurrentWorkbook, library.saveDocument, library.updateTitle, toast, updateSaveState]);

  saveNowRef.current = saveNow;

  const scheduleSave = useCallback(() => {
    if (!entryRef.current?.editKey) return;
    updateSaveState("unsaved");
    if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    saveTimerRef.current = window.setTimeout(() => {
      saveTimerRef.current = null;
      void saveNowRef.current();
    }, 1400);
  }, [updateSaveState]);

  const handleChange = useCallback(
    (data: Sheet[]) => {
      sheetsRef.current = data;
      if (!loadedRef.current || applyingRemoteRef.current || !entryRef.current?.editKey) return;
      markDirty();
      scheduleSave();
    },
    [markDirty, scheduleSave],
  );

  const applyRemote = useCallback(
    (remoteDoc: SheetDocument, event: Event) => {
      const currentEntry = entryRef.current;
      if (!currentEntry) return;
      const activeId = sheetsRef.current.find((sheet) => sheet.status === 1)?.id;
      const prepared: SheetDocument =
        activeId && remoteDoc.sheets.some((sheet) => sheet.id === activeId)
          ? { ...remoteDoc, activeSheetId: activeId }
          : remoteDoc;
      const fortuneSheets = workbookToFortuneSheets(prepared);
      applyingRemoteRef.current = true;
      setSheetDocument(prepared);
      documentRef.current = prepared;
      setTitle(prepared.title);
      sheetsRef.current = fortuneSheets;
      workbookRef.current?.updateSheet(sheetsForUpdate(fortuneSheets));
      lastEventAtRef.current = event.created_at * 1000;
      lastEventIdRef.current = event.id;
      saveDocCache(currentEntry, prepared, {
        savedAt: event.created_at * 1000,
        eventCreatedAt: event.created_at * 1000,
        eventId: event.id,
        pending: false,
      });
      dirtyRef.current = false;
      setDirty(false);
      updateSaveState(currentEntry.editKey ? "saved" : "readonly");
      setRemoteUpdate(null);
      const referenced = collectReferencedPubkeys(prepared);
      if (referenced.length > 0) void profiles.ensureProfiles(referenced);
      window.setTimeout(() => {
        applyingRemoteRef.current = false;
      }, 500);
      toast("Loaded the latest version.", "info");
    },
    [profiles.ensureProfiles, toast, updateSaveState],
  );

  useEffect(() => {
    if (!entry) return;
    let cancelled = false;
    setLoading(true);
    setError("");
    setRemoteUpdate(null);
    setSheetDocument(null);
    documentRef.current = null;
    setInitialSheets([]);
    setTitle("");
    selectionRef.current = [];
    updateSaveState("idle");
    if (saveTimerRef.current) {
      window.clearTimeout(saveTimerRef.current);
      saveTimerRef.current = null;
    }
    if (loadTimerRef.current) window.clearTimeout(loadTimerRef.current);

    const currentEntry = entry;
    library
      .loadDocument(currentEntry)
      .then((result) => {
        if (cancelled) return;
        const doc = result.doc;
        const fortuneSheets = workbookToFortuneSheets(doc);
        setSheetDocument(doc);
        documentRef.current = doc;
        setTitle(doc.title);
        setInitialSheets(fortuneSheets);
        setWorkbookVersion((version) => version + 1);
        sheetsRef.current = fortuneSheets;
        lastEventAtRef.current = result.event
          ? result.event.created_at * 1000
          : currentEntry.lastEventAt;
        lastEventIdRef.current = result.event?.id ?? currentEntry.lastEventId;
        setLastSavedAt(doc.updatedAt);
        const pending = !!result.pending;
        dirtyRef.current = pending;
        setDirty(pending);
        updateSaveState(
          result.localOnly || pending
            ? currentEntry.editKey
              ? "offline"
              : "readonly"
            : currentEntry.editKey
              ? "saved"
              : "readonly",
        );
        loadedRef.current = false;
        loadTimerRef.current = window.setTimeout(() => {
          loadedRef.current = true;
        }, 1000);
        setLoading(false);
        const referenced = collectReferencedPubkeys(doc);
        if (referenced.length > 0) void profiles.ensureProfiles(referenced);
        if (currentEntry.source === "link" && auth.pubkey) library.addLinkEntry(currentEntry);
        if (pending && currentEntry.editKey) {
          window.setTimeout(() => void saveNowRef.current(), 1200);
        }
      })
      .catch((err) => {
        if (cancelled) return;
        setError(err instanceof Error ? err.message : "Couldn't open this spreadsheet.");
        setLoading(false);
      });

    return () => {
      cancelled = true;
      if (loadTimerRef.current) window.clearTimeout(loadTimerRef.current);
    };
  }, [
    auth.pubkey,
    entryKey,
    library.addLinkEntry,
    library.loadDocument,
    profiles.ensureProfiles,
    updateSaveState,
  ]);

  useEffect(() => {
    if (!entry) return;
    const keys = [entry.ownerPubkey, ...entry.sharedWith.map((person) => person.pubkey)].filter(Boolean);
    if (keys.length > 0) void profiles.ensureProfiles(keys);
  }, [entry, profiles.ensureProfiles]);

  useEffect(() => {
    const currentEntry = entryRef.current;
    if (!currentEntry) return;
    return library.subscribeToEntry(currentEntry, (event) => {
      if (event.id === lastEventIdRef.current) return;
      if (saveInFlightRef.current || applyingRemoteRef.current) return;
      if (lastEventAtRef.current && event.created_at * 1000 <= lastEventAtRef.current) return;
      const targetEntry = entryRef.current;
      if (!targetEntry) return;
      if (remoteCheckRef.current) return;
      remoteCheckRef.current = true;
      void library
        .fetchRemoteDocument(targetEntry)
        .then((result) => {
          if (!result) return;
          if (result.event.id === lastEventIdRef.current) return;
          if (lastEventAtRef.current && result.event.created_at * 1000 <= lastEventAtRef.current) return;
          const currentDoc = documentRef.current;
          const currentSheets = currentDoc
            ? JSON.stringify(buildCurrentWorkbook(currentDoc, currentDoc.updatedAt).sheets)
            : "";
          const remoteSheets = JSON.stringify(result.doc.sheets);
          if (currentSheets === remoteSheets) {
            lastEventAtRef.current = result.event.created_at * 1000;
            lastEventIdRef.current = result.event.id;
            return;
          }
          if (!dirtyRef.current) applyRemote(result.doc, result.event);
          else setRemoteUpdate(result);
        })
        .catch(() => {})
        .finally(() => {
          remoteCheckRef.current = false;
        });
    });
  }, [applyRemote, buildCurrentWorkbook, entryKey, library.fetchRemoteDocument, library.subscribeToEntry]);

  useEffect(() => {
    const onBeforeUnload = (event: BeforeUnloadEvent) => {
      if (!dirtyRef.current) return;
      event.preventDefault();
      event.returnValue = "";
    };
    const onVisibility = () => {
      if (document.visibilityState === "hidden" && dirtyRef.current) void saveNowRef.current();
    };
    const onOnline = () => {
      if (saveStateRef.current === "offline" && dirtyRef.current) void saveNowRef.current();
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        void saveNowRef.current();
      }
    };
    window.addEventListener("beforeunload", onBeforeUnload);
    document.addEventListener("visibilitychange", onVisibility);
    window.addEventListener("online", onOnline);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      window.removeEventListener("beforeunload", onBeforeUnload);
      document.removeEventListener("visibilitychange", onVisibility);
      window.removeEventListener("online", onOnline);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, []);

  useEffect(() => {
    return () => {
      const currentEntry = entryRef.current;
      const currentDoc = documentRef.current;
      if (currentEntry?.editKey && dirtyRef.current && currentDoc) {
        void library
          .saveDocument(currentEntry, buildCurrentWorkbook(currentDoc, Date.now()))
          .catch(() => {});
      }
    };
  }, [buildCurrentWorkbook, library.saveDocument]);

  const handleTitleChange = (value: string) => {
    if (!entryRef.current?.editKey) return;
    setTitle(value);
    if (!value.trim()) return;
    if (documentRef.current) {
      documentRef.current = { ...documentRef.current, title: value };
    }
    markDirty();
    scheduleSave();
  };

  const openMetaDialog = (mode: "person" | "tag") => {
    if (!entryRef.current?.editKey) {
      toast("You have view-only access to this sheet.", "info");
      return;
    }
    const ranges = workbookRef.current?.getSelection();
    if (!ranges || ranges.length === 0) {
      toast("Select one or more cells first.", "info");
      return;
    }
    selectionRef.current = ranges.map((range) => ({ row: range.row, column: range.column }));
    if (mode === "person") setPersonOpen(true);
    else setTagOpen(true);
  };

  const applyPerson = (user: UserRef) => {
    const ranges = selectionRef.current;
    if (ranges.length === 0) return;
    const label = user.name || shortNpub(user.pubkey);
    const calls: Parameters<WorkbookInstance["batchCallApis"]>[0] = [];
    for (const range of ranges) {
      const startRow = Math.min(range.row[0], range.row[1]);
      const endRow = Math.max(range.row[0], range.row[1]);
      const startCol = Math.min(range.column[0], range.column[1]);
      const endCol = Math.max(range.column[0], range.column[1]);
      for (let row = startRow; row <= endRow; row++) {
        for (let column = startCol; column <= endCol; column++) {
          calls.push({
            name: "setCellValue",
            args: [
              row,
              column,
              {
                v: label,
                m: label,
                ct: {
                  fa: "General",
                  t: "s",
                  s: {
                    fs: "user",
                    pubkey: user.pubkey,
                    npub: user.npub,
                    name: user.name,
                    picture: user.picture,
                  },
                },
              },
            ],
          });
        }
      }
    }
    workbookRef.current?.batchCallApis(calls);
    toast("Person tag added.", "success");
  };

  const applyTag = (tag: string) => {
    const ranges = selectionRef.current;
    if (ranges.length === 0) return;
    const calls: Parameters<WorkbookInstance["batchCallApis"]>[0] = [];
    for (const range of ranges) {
      const startRow = Math.min(range.row[0], range.row[1]);
      const endRow = Math.max(range.row[0], range.row[1]);
      const startCol = Math.min(range.column[0], range.column[1]);
      const endCol = Math.max(range.column[0], range.column[1]);
      for (let row = startRow; row <= endRow; row++) {
        for (let column = startCol; column <= endCol; column++) {
          calls.push({
            name: "setCellValue",
            args: [
              row,
              column,
              {
                v: tag,
                m: tag,
                ct: { fa: "General", t: "s", s: { fs: "tag" } },
              },
            ],
          });
        }
      }
    }
    workbookRef.current?.batchCallApis(calls);
    toast("Tag added.", "success");
  };

  const copyLink = async (mode: "view" | "edit") => {
    if (!entry) return;
    try {
      await navigator.clipboard.writeText(library.buildShareLink(entry, mode));
      toast(`${mode === "edit" ? "Edit" : "View"} link copied.`, "success");
    } catch {
      toast("Couldn't copy the link.", "error");
    }
  };

  const downloadXlsx = async () => {
    const currentDoc = documentRef.current;
    if (!currentDoc) return;
    try {
      await exportSpreadsheet(buildCurrentWorkbook(currentDoc, Date.now()));
      toast("Workbook exported.", "success");
    } catch (err) {
      toast(err instanceof Error ? err.message : "Couldn't export this workbook.", "error");
    }
  };

  const downloadCsv = () => {
    const currentDoc = documentRef.current;
    if (!currentDoc) return;
    try {
      exportActiveSheetCsv(buildCurrentWorkbook(currentDoc, Date.now()));
      toast("Current sheet exported.", "success");
    } catch (err) {
      toast(err instanceof Error ? err.message : "Couldn't export this sheet.", "error");
    }
  };

  const deleteSheet = async () => {
    if (!entry) return;
    setDeleting(true);
    try {
      await library.deleteSheet(entry);
      toast(entry.role === "owner" ? "Spreadsheet deleted." : "Removed from your sheets.", "success");
      navigate("/");
    } catch (err) {
      toast(err instanceof Error ? err.message : "Couldn't delete the sheet.", "error");
    } finally {
      setDeleting(false);
    }
  };

  const peopleChoices = useMemo<PersonChoice[]>(() => {
    const map = new Map<string, PersonChoice>();
    if (entry?.ownerPubkey) {
      const profile = profiles.profiles.get(entry.ownerPubkey);
      map.set(entry.ownerPubkey, {
        pubkey: entry.ownerPubkey,
        npub: toNpub(entry.ownerPubkey),
        name: profile?.displayName || profile?.name || entry.ownerName,
        picture: profile?.picture || entry.ownerPicture,
        role: "owner",
      });
    }
    for (const person of entry?.sharedWith ?? []) {
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
        role: entry?.role === "owner" ? "owner" : entry?.role === "editor" ? "editor" : "viewer",
      });
    }
    return Array.from(map.values());
  }, [auth.npub, auth.pubkey, entry, profiles.profiles]);

  const collaboratorAvatars = useMemo(() => {
    const map = new Map<string, { pubkey: string; name?: string; picture?: string }>();
    if (entry?.ownerPubkey) {
      const profile = profiles.profiles.get(entry.ownerPubkey);
      map.set(entry.ownerPubkey, {
        pubkey: entry.ownerPubkey,
        name: profile?.displayName || profile?.name || entry.ownerName,
        picture: profile?.picture || entry.ownerPicture,
      });
    }
    for (const person of entry?.sharedWith ?? []) {
      const profile = profiles.profiles.get(person.pubkey);
      map.set(person.pubkey, {
        pubkey: person.pubkey,
        name: profile?.displayName || profile?.name || person.name,
        picture: profile?.picture || person.picture,
      });
    }
    if (auth.pubkey && !map.has(auth.pubkey)) {
      const profile = profiles.profiles.get(auth.pubkey);
      map.set(auth.pubkey, {
        pubkey: auth.pubkey,
        name: profile?.displayName || profile?.name,
        picture: profile?.picture,
      });
    }
    return Array.from(map.values());
  }, [auth.pubkey, entry, profiles.profiles]);

  const hooks = useMemo<Hooks>(
    () => ({
      afterRenderCell: (cell, info, ctx) => {
        drawCellPill(cell, info, ctx, (pubkey) => profiles.profileFor(pubkey));
      },
    }),
    [profiles.profiles, profiles.profileFor],
  );

  const customToolbarItems = useMemo<NonNullable<Settings["customToolbarItems"]>>(
    () => [
      {
        key: "formstr-person",
        tooltip: "Tag a person",
        icon: <UserRound size={16} />,
        children: <span className="sr-only">Tag a person</span>,
        onClick: () => openMetaDialog("person"),
      },
      {
        key: "formstr-tag",
        tooltip: "Add a tag pill",
        icon: <Tag size={16} />,
        children: <span className="sr-only">Add a tag pill</span>,
        onClick: () => openMetaDialog("tag"),
      },
    ],
    [],
  );

  if (!entry && library.loading) {
    return (
      <div className="editor-loading">
        <LogoMark size={42} />
        <Loader2 size={20} className="spin" />
        <p>Opening your spreadsheet…</p>
      </div>
    );
  }

  if (!entry) {
    return (
      <div className="editor-error">
        <LogoMark size={42} />
        <h1>This sheet isn't available</h1>
        <p>
          The link may be missing its access key, or this account doesn't have a
          grant for it.
        </p>
        <div className="empty-actions">
          <button type="button" className="btn btn-primary" onClick={() => navigate("/")}>
            Back to sheets
          </button>
          {!auth.pubkey && (
            <button type="button" className="btn btn-outline" onClick={() => navigate("/")}>
              Sign in
            </button>
          )}
        </div>
      </div>
    );
  }

  const statusIcon: Record<SaveState, ReactNode> = {
    idle: <Loader2 size={14} className="spin" />,
    unsaved: <CloudUpload size={14} />,
    saving: <Loader2 size={14} className="spin" />,
    saved: <CheckCircle2 size={14} />,
    offline: <CloudOff size={14} />,
    readonly: <Eye size={14} />,
  };

  const roleLabel =
    entry.role === "owner" ? "Owner" : entry.role === "editor" ? "Editor" : "Viewer";
  const workbookKey = `${entry.coordinate}:${workbookVersion}`;

  return (
    <div className="editor-page">
      <header className="editor-topbar">
        <div className="editor-topbar-left">
          <button type="button" className="icon-btn" onClick={() => navigate("/")} aria-label="Back to sheets">
            <ArrowLeft size={18} />
          </button>
          <LogoMark size={28} />
          <div className="editor-title-block">
            <input
              className="editor-title-input"
              value={title}
              maxLength={80}
              readOnly={!canEdit}
              onChange={(event) => handleTitleChange(event.target.value)}
              onBlur={() => {
                if (!title.trim()) {
                  const fallback = documentRef.current?.title || entry.title || "Untitled spreadsheet";
                  setTitle(fallback);
                  if (documentRef.current) {
                    documentRef.current = { ...documentRef.current, title: fallback };
                  }
                }
              }}
              onKeyDown={(event) => {
                if (event.key === "Enter") (event.target as HTMLInputElement).blur();
              }}
              aria-label="Spreadsheet title"
            />
            <div className="editor-subline">
              <span className={`role-chip role-${entry.role}`}>{roleLabel}</span>
              {entry.source === "link" && (
                <span className="badge badge-muted">
                  <KeyRound size={11} /> Link access
                </span>
              )}
              {!auth.pubkey && <span className="badge badge-muted">Not signed in</span>}
              <span className="muted">{entry.title === title ? "Encrypted sheet" : "Encrypted sheet"}</span>
            </div>
          </div>
        </div>

        <div className="editor-topbar-right">
          <div className="collaborator-stack" title="People with access">
            {collaboratorAvatars.slice(0, 3).map((person) => (
              <Avatar
                key={person.pubkey}
                size={26}
                name={person.name || shortNpub(person.pubkey)}
                picture={person.picture}
                ring
              />
            ))}
            {collaboratorAvatars.length > 3 && (
              <span className="collaborator-more">+{collaboratorAvatars.length - 3}</span>
            )}
          </div>

          <div className={`sync-status sync-${saveState}`}>
            {statusIcon[saveState]}
            <span>{saveLabel(saveState, lastSavedAt)}</span>
            {saveState === "offline" && (
              <button type="button" className="sync-retry" onClick={() => void saveNow()}>
                Retry
              </button>
            )}
          </div>

          <button type="button" className="btn btn-outline" onClick={() => setShareOpen(true)}>
            <Share2 size={15} /> Share
          </button>

          <Menu
            width={252}
            trigger={
              <button type="button" className="icon-btn" aria-label="More actions">
                <MoreHorizontal size={18} />
              </button>
            }
          >
            {(close) => (
              <>
                <MenuItem
                  icon={<CloudUpload size={15} />}
                  label="Save now"
                  shortcut="⌘S"
                  disabled={!canEdit || saveState === "saving"}
                  onClick={() => {
                    close();
                    void saveNow();
                  }}
                />
                <MenuSeparator />
                <MenuItem
                  icon={<FileSpreadsheet size={15} />}
                  label="Download .xlsx"
                  onClick={() => {
                    close();
                    void downloadXlsx();
                  }}
                />
                <MenuItem
                  icon={<FileDown size={15} />}
                  label="Download current sheet (.csv)"
                  onClick={() => {
                    close();
                    downloadCsv();
                  }}
                />
                <MenuSeparator />
                <MenuItem
                  icon={<Copy size={15} />}
                  label="Copy view link"
                  onClick={() => {
                    close();
                    void copyLink("view");
                  }}
                />
                {entry.editKey && (
                  <MenuItem
                    icon={<KeyRound size={15} />}
                    label="Copy edit link"
                    onClick={() => {
                      close();
                      void copyLink("edit");
                    }}
                  />
                )}
                <MenuSeparator />
                <MenuItem
                  icon={<Trash2 size={15} />}
                  label={entry.role === "owner" ? "Delete spreadsheet" : "Remove from my sheets"}
                  danger
                  onClick={() => {
                    close();
                    setDeleteOpen(true);
                  }}
                />
              </>
            )}
          </Menu>
        </div>
      </header>

      {remoteUpdate && (
        <div className="banner banner-info editor-banner">
          <RefreshCw size={16} />
          <span>A collaborator saved a newer version while you were editing.</span>
          <button
            type="button"
            className="btn btn-outline btn-sm"
            onClick={() => applyRemote(remoteUpdate.doc, remoteUpdate.event)}
          >
            Load latest
          </button>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => {
              setRemoteUpdate(null);
              if (dirtyRef.current) void saveNow();
            }}
          >
            Keep mine
          </button>
        </div>
      )}

      {error ? (
        <div className="editor-error editor-error-inline">
          <AlertTriangle size={24} />
          <h2>Couldn't open this sheet</h2>
          <p>{error}</p>
          <div className="empty-actions">
            <button type="button" className="btn btn-primary" onClick={() => window.location.reload()}>
              Try again
            </button>
            <button type="button" className="btn btn-outline" onClick={() => navigate("/")}>
              Back to sheets
            </button>
          </div>
        </div>
      ) : loading || !sheetDocument ? (
        <div className="editor-loading editor-loading-inline">
          <Loader2 size={22} className="spin" />
          <p>Decrypting your sheet…</p>
        </div>
      ) : (
        <div className="editor-workbook">
          <Workbook
            key={workbookKey}
            ref={workbookRef}
            data={initialSheets}
            row={200}
            column={26}
            rowHeaderWidth={46}
            columnHeaderHeight={28}
            defaultColWidth={116}
            defaultRowHeight={30}
            defaultFontSize={11}
            lang="en"
            currency="$"
            allowEdit={canEdit}
            showToolbar
            showFormulaBar
            showSheetTabs
            toolbarItems={TOOLBAR_ITEMS}
            cellContextMenu={CELL_CONTEXT_MENU}
            sheetTabContextMenu={SHEET_TAB_CONTEXT_MENU}
            hooks={hooks}
            customToolbarItems={customToolbarItems}
            onChange={handleChange}
          />
        </div>
      )}

      <footer className="editor-statusbar">
        <span>
          <LockKeyhole size={12} />
          End-to-end encrypted · NIP-44
        </span>
        <span className="statusbar-dot" />
        <span>{auth.relays.length} relays</span>
        <span className="statusbar-dot" />
        <span>{dirty ? "Local changes pending" : "Up to date"}</span>
        <span className="statusbar-grow" />
        {entry.role === "owner" && <span>{entry.sharedWith.length} collaborator{entry.sharedWith.length === 1 ? "" : "s"}</span>}
        {entry.role !== "owner" && (
          <span>
            {entry.ownerPubkey
              ? `Owner: ${entry.ownerName || shortNpub(entry.ownerPubkey)}`
              : "Shared via link"}
          </span>
        )}
      </footer>

      <ShareDialog open={shareOpen} onClose={() => setShareOpen(false)} entry={entry} />
      <PersonDialog
        open={personOpen}
        onClose={() => {
          setPersonOpen(false);
              }}
        onPick={applyPerson}
        people={peopleChoices}
      />
      <TagDialog
        open={tagOpen}
        onClose={() => {
          setTagOpen(false);
              }}
        onPick={applyTag}
      />

      <Dialog
        open={deleteOpen}
        onClose={() => (deleting ? undefined : setDeleteOpen(false))}
        title={entry.role === "owner" ? "Delete this spreadsheet?" : "Remove from your sheets?"}
        subtitle={
          entry.role === "owner"
            ? "A deletion request is published to relays. People who already opened the sheet may still hold copies."
            : "This removes it from your library. The owner's access grants are unchanged."
        }
        size="sm"
        footer={
          <div className="dialog-footer-actions dialog-footer-end">
            <button type="button" className="btn btn-ghost" disabled={deleting} onClick={() => setDeleteOpen(false)}>
              Cancel
            </button>
            <button type="button" className="btn btn-danger" disabled={deleting} onClick={deleteSheet}>
              {deleting ? <Loader2 size={15} className="spin" /> : <Trash2 size={15} />}
              {entry.role === "owner" ? "Delete" : "Remove"}
            </button>
          </div>
        }
      >
        <p className="dialog-copy">
          {entry.title
            ? `“${entry.title}”${entry.role === "owner" ? " will be removed from your library." : " will no longer appear in Shared with me."}`
            : ""}
        </p>
      </Dialog>
    </div>
  );
}
