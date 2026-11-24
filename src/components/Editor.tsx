import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { useLocation, useNavigate, useParams } from "react-router-dom";
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
  HardDrive,
  Info,
  KeyRound,
  Loader2,
  LockKeyhole,
  MoreHorizontal,
  RefreshCw,
  Share2,
  Trash2,
} from "lucide-react";
import { AUTOSAVE_ENABLED } from "../lib/constants";
import type { BigSheetStore } from "../lib/v2/bigSheet";
import { OPEN_MAX_CELLS, SheetTooLargeToOpenError, type SheetRootV2, type SheetSize } from "../lib/v2/root";
import { BigSheetView } from "./BigSheetView";
import { Avatar } from "./Avatar";
import { Dialog } from "./Dialog";
import { Menu, MenuItem, MenuSeparator } from "./Menu";
import { LogoMark } from "./Logo";
import { ShareDialog } from "./ShareDialog";
import { StorageServersDialog } from "./StorageServersDialog";
import { FileInfoDialog } from "./FileInfoDialog";
import { useAuth } from "../providers/AuthProvider";
import { useLibrary } from "../providers/LibraryProvider";
import { useProfiles } from "../providers/ProfileProvider";
import { useToast } from "../providers/ToastProvider";
import { saveDocCache } from "../lib/local";
import { collectReferencedPubkeys } from "../lib/workbook";
import { exportActiveSheetCsv, exportSpreadsheet } from "../lib/xlsx";
import { formatRelativeTime, shortNpub } from "../lib/utils";
import type { LibraryEntry, SaveState, SheetDocument } from "../lib/types";

type RemoteUpdate = { doc: SheetDocument; event: Event };

/** Do two documents hold the same sheets (ignoring when they were saved)? */
function sameContent(a: SheetDocument, b: SheetDocument): boolean {
  return JSON.stringify(a.sheets) === JSON.stringify(b.sheets);
}

function roleRank(role: LibraryEntry["role"]): number {
  return role === "owner" ? 2 : role === "editor" ? 1 : 0;
}

function saveLabel(state: SaveState, lastSavedAt?: number): string {
  switch (state) {
    case "saving":
      return "Saving…";
    case "unsaved":
      return AUTOSAVE_ENABLED ? "Unsaved changes" : "Unsaved changes - press Ctrl+S to save";
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
  const [title, setTitle] = useState("");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  // Set when the sheet is too big for the editor: shown instead of the grid, nothing is downloaded.
  const [tooLarge, setTooLarge] = useState<SheetSize | null>(null);
  const [bigRoot, setBigRoot] = useState<SheetRootV2 | null>(null);
  const [saveState, setSaveState] = useState<SaveState>("idle");
  const saveStateRef = useRef<SaveState>("idle");
  const [lastSavedAt, setLastSavedAt] = useState<number | undefined>(undefined);
  const [dirty, setDirty] = useState(false);
  const dirtyRef = useRef(false);
  const [remoteUpdate, setRemoteUpdate] = useState<RemoteUpdate | null>(null);
  const [shareOpen, setShareOpen] = useState(false);
  const [deleteOpen, setDeleteOpen] = useState(false);
  const [storageOpen, setStorageOpen] = useState(false);
  const [infoDoc, setInfoDoc] = useState<SheetDocument | null>(null);
  const [deleting, setDeleting] = useState(false);
  // The grid's store: it holds the sheet being edited (every tab, every edit).
  const storeRef = useRef<BigSheetStore | null>(null);
  const applyingRemoteRef = useRef(false);
  const saveTimerRef = useRef<number | null>(null);
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

  /** The workbook as it is right now: the grid's store (with every unsaved edit) under the title the editor holds. */
  const buildCurrentWorkbook = useCallback(async (currentDoc: SheetDocument, updatedAt = Date.now()): Promise<SheetDocument> => {
    const store = storeRef.current;
    if (!store) return { ...currentDoc, updatedAt };
    const built = await store.toDocument();
    return { ...built, title: currentDoc.title, createdAt: currentDoc.createdAt, updatedAt };
  }, []);

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
    let docToSave: SheetDocument = currentDoc;
    try {
      docToSave = await buildCurrentWorkbook(currentDoc, Date.now());
      documentRef.current = docToSave;
      const { event } = await library.saveDocument(currentEntry, docToSave);
      lastEventAtRef.current = event.created_at * 1000;
      lastEventIdRef.current = event.id;
      storeRef.current?.markSaved();
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
    // With autosave off a change only marks the sheet unsaved; the user saves explicitly.
    if (!AUTOSAVE_ENABLED) return;
    if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    saveTimerRef.current = window.setTimeout(() => {
      saveTimerRef.current = null;
      void saveNowRef.current();
    }, 1400);
  }, [updateSaveState]);

  /** The grid has (or no longer has) unsaved changes. */
  const handleGridDirty = useCallback(
    (isDirty: boolean) => {
      if (!entryRef.current?.editKey || applyingRemoteRef.current) return;
      if (isDirty) {
        markDirty();
        scheduleSave();
      } else {
        dirtyRef.current = false;
        setDirty(false);
        if (saveStateRef.current === "unsaved") updateSaveState("saved");
      }
    },
    [markDirty, scheduleSave, updateSaveState],
  );
  const gridHost = useMemo(
    () => ({
      onStore: (store: BigSheetStore) => {
        storeRef.current = store;
      },
      onDirtyChange: handleGridDirty,
      save: () => saveNowRef.current(),
      saving: saveState === "saving",
    }),
    [handleGridDirty, saveState],
  );

  const applyRemote = useCallback(
    (remoteDoc: SheetDocument, event: Event) => {
      const currentEntry = entryRef.current;
      if (!currentEntry) return;
      const activeId = storeRef.current?.root.activeSheetId;
      const prepared: SheetDocument =
        activeId && remoteDoc.sheets.some((sheet) => sheet.id === activeId)
          ? { ...remoteDoc, activeSheetId: activeId }
          : remoteDoc;
      applyingRemoteRef.current = true;
      // A new document makes the grid open it afresh.
      setSheetDocument(prepared);
      documentRef.current = prepared;
      setTitle(prepared.title);
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
    setTooLarge(null);
    setBigRoot(null);
    setRemoteUpdate(null);
    setSheetDocument(null);
    documentRef.current = null;
    setTitle("");
    storeRef.current = null;
    updateSaveState("idle");
    if (saveTimerRef.current) {
      window.clearTimeout(saveTimerRef.current);
      saveTimerRef.current = null;
    }

    const currentEntry = entry;
    library
      .loadDocument(currentEntry)
      .then((result) => {
        if (cancelled) return;
        const doc = result.doc;
        setSheetDocument(doc);
        documentRef.current = doc;
        setTitle(doc.title);
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
        setLoading(false);
        const referenced = collectReferencedPubkeys(doc);
        if (referenced.length > 0) void profiles.ensureProfiles(referenced);
        if (currentEntry.source === "link" && auth.pubkey) library.addLinkEntry(currentEntry);
        if (AUTOSAVE_ENABLED && pending && currentEntry.editKey) {
          window.setTimeout(() => void saveNowRef.current(), 1200);
        }
      })
      .catch((err) => {
        if (cancelled) return;
        setTooLarge(err instanceof SheetTooLargeToOpenError ? err.size : null);
        setBigRoot(err instanceof SheetTooLargeToOpenError ? (err.root ?? null) : null);
        if (err instanceof SheetTooLargeToOpenError && err.root) setTitle(err.root.title);
        setError(err instanceof Error ? err.message : "Couldn't open this spreadsheet.");
        setLoading(false);
      });

    return () => {
      cancelled = true;
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
        .then(async (result) => {
          if (!result) return;
          if (result.event.id === lastEventIdRef.current) return;
          if (lastEventAtRef.current && result.event.created_at * 1000 <= lastEventAtRef.current) return;
          const currentDoc = documentRef.current;
          const current = currentDoc ? await buildCurrentWorkbook(currentDoc, currentDoc.updatedAt) : null;
          if (current && sameContent(current, result.doc)) {
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
      if (AUTOSAVE_ENABLED && document.visibilityState === "hidden" && dirtyRef.current) void saveNowRef.current();
    };
    const onOnline = () => {
      if (AUTOSAVE_ENABLED && saveStateRef.current === "offline" && dirtyRef.current) void saveNowRef.current();
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
        void buildCurrentWorkbook(currentDoc, Date.now()).then((built) => {
          if (AUTOSAVE_ENABLED) {
            void library.saveDocument(currentEntry, built).catch(() => {});
          } else {
            // No network save without the user's say-so: keep a local draft so leaving the page
            // does not lose the edits (too-big drafts do not fit in local storage and are skipped).
            saveDocCache(currentEntry, built, { pending: true });
          }
        });
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
      await exportSpreadsheet(await buildCurrentWorkbook(currentDoc, Date.now()));
      toast("Workbook exported.", "success");
    } catch (err) {
      toast(err instanceof Error ? err.message : "Couldn't export this workbook.", "error");
    }
  };

  const downloadCsv = async () => {
    const currentDoc = documentRef.current;
    if (!currentDoc) return;
    try {
      exportActiveSheetCsv(await buildCurrentWorkbook(currentDoc, Date.now()));
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

  return (
    <div className="editor-page">
      <header className="editor-topbar">
        <div className="editor-topbar-left">
          <button
            type="button"
            className="icon-btn"
            onClick={() => {
              if (
                !AUTOSAVE_ENABLED &&
                dirtyRef.current &&
                !window.confirm("You have unsaved changes. Leave without saving? (They stay as a local draft on this device.)")
              ) {
                return;
              }
              navigate("/");
            }}
            aria-label="Back to sheets"
          >
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
                    void downloadCsv();
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
                  icon={<Info size={15} />}
                  label="File info"
                  onClick={() => {
                    close();
                    const currentDoc = documentRef.current;
                    // Measure what is open right now, including unsaved edits.
                    if (currentDoc) void buildCurrentWorkbook(currentDoc, Date.now()).then(setInfoDoc);
                  }}
                />
                <MenuItem
                  icon={<HardDrive size={15} />}
                  label="Storage servers"
                  onClick={() => {
                    close();
                    setStorageOpen(true);
                  }}
                />
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

      {error && bigRoot && entry ? (
        <BigSheetView entry={entry} root={bigRoot} />
      ) : error && tooLarge ? (
        <div className="editor-error editor-error-inline">
          <AlertTriangle size={24} />
          <h2>Too large to edit in the browser yet</h2>
          <p>
            This sheet is stored and encrypted
            {tooLarge.cells !== null ? ` (${tooLarge.cells.toLocaleString()} cells` : ` (${(tooLarge.bytes / 1024 / 1024).toFixed(0)} MB compressed`}
            {`, ${tooLarge.tabs.length} tab${tooLarge.tabs.length === 1 ? "" : "s"})`}, but the editor can only open
            sheets up to {OPEN_MAX_CELLS.toLocaleString()} cells for now.
          </p>
          <table className="too-large-tabs">
            <tbody>
              {tooLarge.tabs.map((tab) => (
                <tr key={tab.name}>
                  <td>{tab.name}</td>
                  <td>{tab.cells !== null ? `${tab.cells.toLocaleString()} cells` : ""}</td>
                  <td>{tab.chunks ? `${tab.chunks} chunk${tab.chunks === 1 ? "" : "s"}` : "chunk index"}</td>
                  <td>{(tab.bytes / 1024 / 1024).toFixed(1)} MB</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div className="empty-actions">
            <button type="button" className="btn btn-outline" onClick={() => navigate("/")}>
              Back to sheets
            </button>
          </div>
        </div>
      ) : error ? (
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
        <BigSheetView entry={entry} document={sheetDocument} host={gridHost} />
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

      <StorageServersDialog open={storageOpen} onClose={() => setStorageOpen(false)} />
      <FileInfoDialog
        open={infoDoc !== null}
        onClose={() => setInfoDoc(null)}
        doc={infoDoc}
        lastSavedAt={lastSavedAt}
      />
    </div>
  );
}
