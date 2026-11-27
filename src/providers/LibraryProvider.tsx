import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import { nip19, type Event } from "nostr-tools";
import type { Sheet } from "../lib/legacy/fortuneFormat";
import {
  CLIENT_TAG,
  KIND_BLOSSOM_SERVERS,
  KIND_SHEET,
  KIND_SHEET_GRANT,
  KIND_SHEET_META,
  SHARE_HASH_PREFIX,
} from "../lib/constants";
import {
  buildSheetDeletionEvent,
  buildSheetEvent,
  decodeSheetEvent,
  generateDocKeys,
  sheetEventFromRoot,
  storedBytesOf,
  storedBytesOfEvent,
  type TabMemory,
} from "../lib/crypto";
import { importXlsxToBlossom } from "../lib/v2/streamImport";
import {
  coordinateFor,
  fetchBlossomServers,
  fetchEvents,
  fetchLatestEvent,
  fetchUserRelays,
  parseCoordinate,
  publishEvent,
  subscribeSheet,
} from "../lib/nostr";
import { createWorkbookFromFortune } from "../lib/workbook";
import { BlobUnavailableError, SheetTooLargeError, SheetTooLargeToOpenError, type SheetRootV2 } from "../lib/v2/root";
import { mergeServers } from "../lib/v2/blossom";
import { uploadProgress } from "../lib/v2/progress";
import { templateById } from "../lib/templates";
import {
  loadDocCache,
  loadHiddenCoordinates,
  loadLibraryCache,
  loadUploadPrefs,
  saveUploadPrefs as persistUploadPrefs,
  type UploadPrefs,
  saveDocCache,
  saveHiddenCoordinates,
  saveLibraryCache,
} from "../lib/local";
import { mergeRelays, randomId, toNpub } from "../lib/utils";
import { encodeNKeys, decodeNKeys } from "../lib/nkeys";
import type {
  Collaborator,
  LibraryEntry,
  LoadedDocument,
  SheetDocument,
  SheetGrantPayload,
  SheetIndexPayload,
  SheetMetaPayload,
} from "../lib/types";
import { useAuth, type ActiveSigner } from "./AuthProvider";
import { useProfiles } from "./ProfileProvider";

interface LibraryContextValue {
  entries: LibraryEntry[];
  loading: boolean;
  error: string | null;
  refresh: () => Promise<void>;
  createSheet: (templateId: string, title: string) => Promise<LibraryEntry>;
  importDocument: (title: string, sheets: Sheet[]) => Promise<LibraryEntry>;
  /** Big xlsx: read in a worker, encrypted and uploaded chunk by chunk, never held in memory whole. */
  importStreamed: (title: string, file: File, signal?: AbortSignal) => Promise<{ entry: LibraryEntry; dropped: string[] }>;
  getEntry: (idOrCoordinate: string) => LibraryEntry | undefined;
  getEntryByNaddr: (naddr: string) => LibraryEntry | undefined;
  parseShareLink: (naddr: string, hash: string) => LibraryEntry | null;
  addLinkEntry: (entry: LibraryEntry) => void;
  loadDocument: (entry: LibraryEntry) => Promise<LoadedDocument>;
  fetchRemoteDocument: (entry: LibraryEntry) => Promise<{ doc: SheetDocument; event: Event } | null>;
  saveDocument: (
    entry: LibraryEntry,
    document: SheetDocument,
  ) => Promise<{ event: Event; accepted: number }>;
  updateTitle: (entry: LibraryEntry, title: string) => Promise<void>;
  deleteSheet: (entry: LibraryEntry) => Promise<void>;
  shareSheet: (
    entry: LibraryEntry,
    collaborator: Collaborator,
    role: "viewer" | "editor",
  ) => Promise<void>;
  revokeAccess: (entry: LibraryEntry, pubkey: string) => Promise<void>;
  subscribeToEntry: (entry: LibraryEntry, onEvent: (event: Event) => void) => () => void;
  buildShareLink: (entry: LibraryEntry, mode: "view" | "edit") => string;
  /** The user's Blossom servers (their kind 10063 list), where large sheets are stored. */
  blossomServers: string[];
  refreshBlossomServers: () => Promise<string[]>;
  /**
   * Publishes the list. `removed` are servers the user took out; everything else already on
   * their profile is kept, so a stale or failed read cannot wipe servers added elsewhere.
   */
  saveBlossomServers: (servers: string[], removed: string[]) => Promise<string[]>;
  /** Which servers uploads go to, and whether to keep a second copy. Stored on this device. */
  uploadPrefs: UploadPrefs;
  saveUploadPrefs: (prefs: UploadPrefs) => void;
  /** Where a big sheet's edited chunks are uploaded, and who signs. */
  bigSheetUpload: () => { servers: string[]; replicas: number; signer?: ActiveSigner };
  /** Signs a big sheet's updated root with its edit key and publishes it. */
  publishBigRoot: (entry: LibraryEntry, root: SheetRootV2) => Promise<void>;
}

const LibraryContext = createContext<LibraryContextValue | undefined>(undefined);

function isHexKey(value: unknown): value is string {
  return typeof value === "string" && /^[0-9a-f]{64}$/i.test(value);
}

function normalizeCollaborator(raw: unknown): Collaborator | null {
  if (!raw || typeof raw !== "object") return null;
  const item = raw as Partial<Collaborator>;
  if (!isHexKey(item.pubkey) || (item.role !== "viewer" && item.role !== "editor")) return null;
  return {
    pubkey: item.pubkey,
    npub: item.npub || toNpub(item.pubkey),
    name: item.name,
    picture: item.picture,
    role: item.role,
    grantedAt: item.grantedAt,
  };
}

function normalizeEntry(raw: unknown): LibraryEntry | null {
  if (!raw || typeof raw !== "object") return null;
  const item = raw as Partial<LibraryEntry>;
  if (!isHexKey(item.viewKey)) return null;
  let editPubkey = item.editPubkey;
  let id = item.id;
  if ((!editPubkey || !id) && typeof item.coordinate === "string") {
    const parsed = parseCoordinate(item.coordinate);
    if (parsed) {
      editPubkey = editPubkey || parsed.pubkey;
      id = id || parsed.id;
    }
  }
  if (!editPubkey || !id) return null;
  if (item.editKey != null && !isHexKey(item.editKey)) return null;
  const coordinate =
    typeof item.coordinate === "string" && item.coordinate
      ? item.coordinate
      : coordinateFor(editPubkey, id);
  const role =
    item.role === "owner" || item.role === "editor" || item.role === "viewer"
      ? item.role
      : "viewer";
  const sharedWith = Array.isArray(item.sharedWith)
    ? item.sharedWith.map(normalizeCollaborator).filter((value): value is Collaborator => !!value)
    : [];
  return {
    id,
    coordinate,
    editPubkey,
    viewKey: item.viewKey,
    editKey: item.editKey,
    title: typeof item.title === "string" && item.title.trim() ? item.title : "Untitled spreadsheet",
    role,
    ownerPubkey: typeof item.ownerPubkey === "string" ? item.ownerPubkey : "",
    ownerName: item.ownerName,
    ownerPicture: item.ownerPicture,
    createdAt: typeof item.createdAt === "number" ? item.createdAt : Date.now(),
    updatedAt: typeof item.updatedAt === "number" ? item.updatedAt : Date.now(),
    lastOpenedAt: item.lastOpenedAt,
    sharedWith,
    template: item.template,
    color: item.color,
    relayHints: Array.isArray(item.relayHints) ? item.relayHints.filter((relay) => typeof relay === "string") : undefined,
    lastEventAt: item.lastEventAt,
    lastEventId: item.lastEventId,
    source: item.source ?? (role === "owner" ? "owned" : "link"),
  };
}

function grantToEntry(event: Event, payload: SheetGrantPayload): LibraryEntry | null {
  if (payload.revoked) return null;
  if (!isHexKey(payload.viewKey) || !isHexKey(payload.editPubkey)) return null;
  if (payload.role !== "viewer" && payload.role !== "editor") return null;
  const coordinate = payload.coordinate || coordinateFor(payload.editPubkey, payload.sheetId);
  const id = payload.sheetId || parseCoordinate(coordinate)?.id;
  if (!id) return null;
  const editKey =
    payload.role === "editor" && isHexKey(payload.editKey) ? payload.editKey : undefined;
  return {
    id,
    coordinate,
    editPubkey: payload.editPubkey,
    viewKey: payload.viewKey,
    editKey,
    title: payload.title || "Shared spreadsheet",
    role: editKey ? "editor" : "viewer",
    ownerPubkey: payload.ownerPubkey || event.pubkey,
    ownerName: payload.ownerName,
    createdAt: payload.grantedAt || event.created_at * 1000,
    updatedAt: payload.grantedAt || event.created_at * 1000,
    sharedWith: [],
    color: payload.color,
    relayHints: payload.relayHints,
    source: "grant",
  };
}

interface MetaLoadResult {
  entries: LibraryEntry[];
  deleted: Set<string>;
  timestamps: Map<string, number>;
}

async function loadMetaEntries(
  signer: ActiveSigner,
  pubkey: string,
  relays: string[],
): Promise<MetaLoadResult> {
  const events = await fetchEvents(
    relays,
    { kinds: [KIND_SHEET_META], authors: [pubkey], limit: 1000 },
    8000,
  );
  events.sort((a, b) => a.created_at - b.created_at);

  const entries = new Map<string, LibraryEntry>();
  const deleted = new Set<string>();
  const timestamps = new Map<string, number>();

  for (const event of events) {
    const d = event.tags.find((tag) => tag[0] === "d")?.[1];
    if (!d) continue;
    timestamps.set(d, Math.max(timestamps.get(d) ?? 0, event.created_at));
    try {
      const json = await signer.nip44Decrypt(pubkey, event.content);
      const parsed = JSON.parse(json) as SheetMetaPayload | SheetIndexPayload;
      if (d === "index") {
        const legacy = parsed as SheetIndexPayload;
        if (Array.isArray(legacy.entries)) {
          for (const raw of legacy.entries) {
            const entry = normalizeEntry(raw);
            if (entry && !deleted.has(entry.coordinate)) {
              entries.set(entry.coordinate, entry);
            }
          }
        }
        continue;
      }
      const meta = parsed as SheetMetaPayload;
      if (meta.deleted) {
        deleted.add(d);
        entries.delete(d);
        continue;
      }
      const entry = normalizeEntry(meta.entry ?? meta);
      if (entry) {
        entries.set(entry.coordinate, entry);
        deleted.delete(entry.coordinate);
      }
    } catch {
      // Skip metadata this signer can't decrypt.
    }
  }

  return { entries: Array.from(entries.values()), deleted, timestamps };
}

async function loadGrantEntries(
  signer: ActiveSigner,
  pubkey: string,
  relays: string[],
): Promise<LibraryEntry[]> {
  const events = await fetchEvents(
    relays,
    { kinds: [KIND_SHEET_GRANT], "#p": [pubkey], limit: 400 },
    8000,
  );
  events.sort((a, b) => b.created_at - a.created_at);
  const latest = new Map<string, Event>();
  for (const event of events) {
    const d = event.tags.find((tag) => tag[0] === "d")?.[1];
    if (!d) continue;
    const key = `${event.pubkey}:${d}`;
    if (!latest.has(key)) latest.set(key, event);
  }

  const entries: LibraryEntry[] = [];
  for (const event of latest.values()) {
    try {
      const json = await signer.nip44Decrypt(event.pubkey, event.content);
      const payload = JSON.parse(json) as SheetGrantPayload;
      if (payload.revoked) continue;
      const recipient = event.tags.find((tag) => tag[0] === "p")?.[1];
      if (recipient && recipient !== pubkey) continue;
      const entry = grantToEntry(event, payload);
      if (entry) entries.push(entry);
    } catch {
      // Skip grants this signer can't decrypt.
    }
  }
  return entries;
}

export function LibraryProvider({ children }: { children: ReactNode }) {
  const { pubkey, signer, relays } = useAuth();
  const { profileFor } = useProfiles();
  const [entries, setEntries] = useState<LibraryEntry[]>([]);
  const entriesRef = useRef<LibraryEntry[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const replaceEntries = useCallback(
    (next: LibraryEntry[]) => {
      entriesRef.current = next;
      setEntries(next);
      if (pubkey) saveLibraryCache(pubkey, next);
    },
    [pubkey],
  );

  const lastMetaAtRef = useRef(new Map<string, number>());

  // Blossom: the user's server list (kind 10063), warmed after sign-in, and a per-sheet
  // memory of blobs already uploaded so a save only sends the tabs that changed.
  const serversRef = useRef<string[]>([]);
  const [blossomServers, setBlossomServers] = useState<string[]>([]);
  const reuseRef = useRef(new Map<string, Map<string, TabMemory>>());

  const [uploadPrefs, setUploadPrefsState] = useState<UploadPrefs>({ off: [] });
  const prefsRef = useRef(uploadPrefs);
  useEffect(() => {
    const prefs = pubkey ? loadUploadPrefs(pubkey) : { off: [] };
    prefsRef.current = prefs;
    setUploadPrefsState(prefs);
  }, [pubkey]);

  const saveUploadPrefs = useCallback(
    (prefs: UploadPrefs) => {
      prefsRef.current = prefs;
      setUploadPrefsState(prefs);
      if (pubkey) persistUploadPrefs(pubkey, prefs);
    },
    [pubkey],
  );

  const reuseFor = useCallback((coordinate: string) => {
    let map = reuseRef.current.get(coordinate);
    if (!map) {
      map = new Map();
      reuseRef.current.set(coordinate, map);
    }
    return map;
  }, []);

  const fetchServers = useCallback(async () => {
    if (!pubkey) return [];
    const list = await fetchBlossomServers(relays, pubkey).catch(() => []);
    serversRef.current = list;
    setBlossomServers(list);
    return list;
  }, [pubkey, relays]);

  useEffect(() => {
    serversRef.current = [];
    setBlossomServers([]);
    void fetchServers();
  }, [fetchServers]);

  const saveBlossomServers = useCallback(
    async (servers: string[], removed: string[]) => {
      if (!signer || !pubkey) throw new Error("Please sign in first.");
      // Read the current list again so servers added elsewhere (e.g. in Formstr Drive) survive.
      const merged = mergeServers(await fetchServers(), servers, removed);
      const event = await signer.signEvent({
        kind: KIND_BLOSSOM_SERVERS,
        created_at: Math.floor(Date.now() / 1000),
        tags: merged.map((url) => ["server", url]),
        content: "",
      });
      await publishEvent(event, relays);
      serversRef.current = merged;
      setBlossomServers(merged);
      return merged;
    },
    [fetchServers, pubkey, relays, signer],
  );

  // Builds the sheet event. If the sheet is too big to store inline and we have not seen any
  // Blossom servers, look the list up again once (it may have been added since sign-in).
  // Remembers a sheet's stored size on its library entry (shown on the dashboard card).
  const sizesRef = useRef(
    new Map<string, number>(
      (() => {
        try {
          return Object.entries(JSON.parse(window.localStorage.getItem("formstr-sheets:sizes") ?? "{}") as Record<string, number>);
        } catch {
          return [];
        }
      })(),
    ),
  );
  const noteSize = useCallback(
    (coordinate: string, bytes: number) => {
      sizesRef.current.set(coordinate, bytes);
      try {
        window.localStorage.setItem("formstr-sheets:sizes", JSON.stringify(Object.fromEntries(sizesRef.current)));
      } catch {
        // Sizes are a convenience; the card just shows none.
      }
      const current = entriesRef.current.find((item) => item.coordinate === coordinate);
      if (!current || current.storedBytes === bytes) return;
      replaceEntries(
        entriesRef.current.map((item) => (item.coordinate === coordinate ? { ...item, storedBytes: bytes } : item)),
      );
    },
    [replaceEntries],
  );

  // Cards whose size is not known yet get it from the sheet's latest event (a small download; the blobs are not fetched).
  const sizeLookups = useRef(new Set<string>());
  useEffect(() => {
    const missing = entries.filter((entry) => !entry.storedBytes && !sizesRef.current.has(entry.coordinate) && !sizeLookups.current.has(entry.coordinate));
    if (missing.length === 0 || relays.length === 0) return;
    let stopped = false;
    void (async () => {
      for (const entry of missing) {
        if (stopped) return;
        sizeLookups.current.add(entry.coordinate);
        try {
          const event = await fetchLatestEvent(relays, { kinds: [KIND_SHEET], authors: [entry.editPubkey], "#d": [entry.id], limit: 1 }, 6000);
          const bytes = event ? storedBytesOfEvent(event, entry.viewKey) : null;
          if (bytes) noteSize(entry.coordinate, bytes);
        } catch {
          // Try again the next time the list loads.
          sizeLookups.current.delete(entry.coordinate);
        }
      }
    })();
    return () => {
      stopped = true;
    };
  }, [entries, noteSize, relays]);

  const buildEvent = useCallback(
    async (coordinate: string, params: Omit<Parameters<typeof buildSheetEvent>[0], "blossom">) => {
      const attempt = (servers: string[]) => {
        // Unticked servers are never uploaded to (they are still used to read sheets).
        const { off } = prefsRef.current;
        const targets = servers.filter((url) => !off.includes(url));
        return buildSheetEvent({
          ...params,
          onSize: (bytes) => noteSize(coordinate, bytes),
          blossom: {
            servers: targets,
            // Every ticked server gets every blob.
            replicas: targets.length,
            signer: signer ?? undefined,
            reuse: reuseFor(coordinate),
            onProgress: uploadProgress.report,
          },
        });
      };
      try {
        return await attempt(serversRef.current);
      } catch (error) {
        if (!(error instanceof SheetTooLargeError)) throw error;
        const servers = await fetchServers();
        if (servers.length === 0) throw error;
        return attempt(servers);
      }
    },
    [fetchServers, noteSize, reuseFor, signer],
  );

  const persistEntry = useCallback(
    async (entry: LibraryEntry, deleted = false) => {
      if (!signer || !pubkey || entry.source === "grant") return;
      const payload: SheetMetaPayload = {
        version: 1,
        entry,
        ...(deleted ? { deleted: true } : {}),
      };
      const content = await signer.nip44Encrypt(pubkey, JSON.stringify(payload));
      const now = Math.floor(Date.now() / 1000);
      const previous = lastMetaAtRef.current.get(entry.coordinate) ?? 0;
      const createdAt = Math.max(now, previous + 1);
      lastMetaAtRef.current.set(entry.coordinate, createdAt);
      const signed = await signer.signEvent({
        kind: KIND_SHEET_META,
        created_at: createdAt,
        tags: [
          ["d", entry.coordinate],
          ["client", CLIENT_TAG],
          ["encrypted", "nip44"],
        ],
        content,
      });
      await publishEvent(signed, relays);
    },
    [pubkey, relays, signer],
  );

  const hydrateUpdatedAt = useCallback(
    async (currentEntries: LibraryEntry[]) => {
      if (currentEntries.length === 0) return;
      const authors = Array.from(new Set(currentEntries.map((entry) => entry.editPubkey)));
      const ids = Array.from(new Set(currentEntries.map((entry) => entry.id)));
      const events = await fetchEvents(
        relays,
        { kinds: [KIND_SHEET], authors, "#d": ids, limit: 500 },
        7000,
      );
      const latest = new Map<string, Event>();
      for (const event of events) {
        const id = event.tags.find((tag) => tag[0] === "d")?.[1];
        if (!id) continue;
        const coordinate = coordinateFor(event.pubkey, id);
        const current = latest.get(coordinate);
        if (!current || event.created_at > current.created_at) latest.set(coordinate, event);
      }
      if (latest.size === 0) return;
      const next = entriesRef.current.map((entry) => {
        const event = latest.get(entry.coordinate);
        if (!event) return entry;
        return {
          ...entry,
          updatedAt: Math.max(entry.updatedAt, event.created_at * 1000),
          lastEventAt: event.created_at * 1000,
          lastEventId: event.id,
        };
      });
      replaceEntries(next);
    },
    [relays, replaceEntries],
  );

  const refresh = useCallback(async () => {
    if (!pubkey || !signer) return;
    setLoading(true);
    setError(null);
    const hidden = new Set(loadHiddenCoordinates(pubkey));
    try {
      const [metaResult, grantResult] = await Promise.allSettled([
        loadMetaEntries(signer, pubkey, relays),
        loadGrantEntries(signer, pubkey, relays),
      ]);

      if (metaResult.status === "fulfilled") {
        for (const [coordinate, timestamp] of metaResult.value.timestamps) {
          lastMetaAtRef.current.set(
            coordinate,
            Math.max(lastMetaAtRef.current.get(coordinate) ?? 0, timestamp),
          );
        }
      }

      const deleted = metaResult.status === "fulfilled" ? metaResult.value.deleted : new Set<string>();
      const indexEntries =
        metaResult.status === "fulfilled"
          ? metaResult.value.entries.filter(
              (entry) => !hidden.has(entry.coordinate) && !deleted.has(entry.coordinate),
            )
          : [];
      const grantEntries =
        grantResult.status === "fulfilled"
          ? grantResult.value.filter(
              (entry) =>
                !hidden.has(entry.coordinate) &&
                !deleted.has(entry.coordinate) &&
                !indexEntries.some((owned) => owned.coordinate === entry.coordinate),
            )
          : [];

      if (metaResult.status === "rejected" && grantResult.status === "rejected") {
        setError("Couldn't reach relays. Showing the copies saved on this device.");
      } else if (metaResult.status === "rejected") {
        setError("Couldn't load your library metadata. Shared sheets still loaded.");
      }

      const merged = [...indexEntries, ...grantEntries].sort(
        (a, b) => (b.lastOpenedAt ?? b.updatedAt) - (a.lastOpenedAt ?? a.updatedAt),
      );
      replaceEntries(merged);
      void hydrateUpdatedAt(merged).catch(() => {});
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't load your library.");
    } finally {
      setLoading(false);
    }
  }, [hydrateUpdatedAt, pubkey, relays, replaceEntries, signer]);

  useEffect(() => {
    if (!pubkey || !signer) {
      entriesRef.current = [];
      setEntries([]);
      setLoading(false);
      return;
    }
    const cached = loadLibraryCache(pubkey);
    entriesRef.current = cached;
    setEntries(cached);
    void refresh();
  }, [pubkey, signer, refresh]);

  const createFromDocument = useCallback(
    async (
      title: string,
      document: SheetDocument,
      template: string,
      color: string,
    ): Promise<LibraryEntry> => {
      if (!signer || !pubkey) throw new Error("Please sign in first.");
      const keys = generateDocKeys();
      const id = randomId(12);
      const now = Date.now();
      const finalDocument: SheetDocument = {
        ...document,
        title,
        createdAt: now,
        updatedAt: now,
      };
      const event = await buildEvent(coordinateFor(keys.editPubkey, id), {
        sheetId: id,
        document: finalDocument,
        viewKey: keys.viewKey,
        editKey: keys.editKey,
        createdAt: Math.floor(now / 1000),
      });
      const ownerProfile = profileFor(pubkey);
      const entry: LibraryEntry = {
        id,
        coordinate: coordinateFor(keys.editPubkey, id),
        editPubkey: keys.editPubkey,
        viewKey: keys.viewKey,
        editKey: keys.editKey,
        title,
        role: "owner",
        ownerPubkey: pubkey,
        ownerName: ownerProfile?.displayName || ownerProfile?.name,
        ownerPicture: ownerProfile?.picture,
        createdAt: now,
        updatedAt: now,
        sharedWith: [],
        template,
        color,
        relayHints: relays.slice(0, 3),
        lastEventAt: event.created_at * 1000,
        lastEventId: event.id,
        source: "owned",
        storedBytes: sizesRef.current.get(coordinateFor(keys.editPubkey, id)),
      };

      saveDocCache(entry, finalDocument, {
        savedAt: now,
        eventCreatedAt: event.created_at * 1000,
        eventId: event.id,
        pending: true,
      });

      const next = [entry, ...entriesRef.current.filter((item) => item.coordinate !== entry.coordinate)];
      replaceEntries(next);
      void publishEvent(event, relays).catch(() => {});
      void persistEntry(entry).catch(() => {
        setError("The sheet is saved on this device, but its library metadata didn't reach a relay.");
      });
      return entry;
    },
    [buildEvent, persistEntry, profileFor, pubkey, relays, replaceEntries, signer],
  );

  const createSheet = useCallback(
    async (templateId: string, title: string) => {
      const template = templateById(templateId);
      const document = template.build(title);
      return createFromDocument(title, document, template.id, template.color);
    },
    [createFromDocument],
  );

  const importDocument = useCallback(
    async (title: string, sheets: Sheet[]) => {
      const document = createWorkbookFromFortune(title, sheets);
      return createFromDocument(title, document, "import", "#FF5A36");
    },
    [createFromDocument],
  );

  const importStreamed = useCallback(
    async (title: string, file: File, signal?: AbortSignal): Promise<{ entry: LibraryEntry; dropped: string[] }> => {
      if (!signer || !pubkey) throw new Error("Please sign in first.");
      const servers = serversRef.current.length > 0 ? serversRef.current : await fetchServers();
      // Unticked servers are never uploaded to; every ticked one gets every blob.
      const targets = servers.filter((url) => !prefsRef.current.off.includes(url));
      const { root, summary } = await importXlsxToBlossom(file, {
        title,
        servers: targets,
        replicas: targets.length,
        signer,
        onProgress: uploadProgress.report,
        signal,
      });
      const keys = generateDocKeys();
      const id = randomId(12);
      const coordinate = coordinateFor(keys.editPubkey, id);
      const event = sheetEventFromRoot({ sheetId: id, viewKey: keys.viewKey, editKey: keys.editKey }, root);
      // No local copy of a streamed sheet exists (it would not fit), so the event must reach a relay.
      await publishEvent(event, relays);
      const ownerProfile = profileFor(pubkey);
      const entry: LibraryEntry = {
        id,
        coordinate,
        editPubkey: keys.editPubkey,
        viewKey: keys.viewKey,
        editKey: keys.editKey,
        title,
        role: "owner",
        ownerPubkey: pubkey,
        ownerName: ownerProfile?.displayName || ownerProfile?.name,
        ownerPicture: ownerProfile?.picture,
        createdAt: root.createdAt,
        updatedAt: root.updatedAt,
        sharedWith: [],
        template: "import",
        color: "#FF5A36",
        relayHints: relays.slice(0, 3),
        lastEventAt: event.created_at * 1000,
        lastEventId: event.id,
        source: "owned",
        storedBytes: storedBytesOf(root),
      };
      replaceEntries([entry, ...entriesRef.current.filter((item) => item.coordinate !== coordinate)]);
      void persistEntry(entry).catch(() => {
        setError("The sheet is stored, but its library metadata didn't reach a relay.");
      });
      return { entry, dropped: summary.dropped };
    },
    [fetchServers, persistEntry, profileFor, pubkey, relays, replaceEntries, signer],
  );

  const bigSheetUpload = useCallback(() => {
    const targets = serversRef.current.filter((url) => !prefsRef.current.off.includes(url));
    return { servers: targets, replicas: targets.length, signer: signer ?? undefined };
  }, [signer]);

  const publishBigRoot = useCallback(
    async (entry: LibraryEntry, root: SheetRootV2) => {
      if (!entry.editKey) throw new Error("You can only view this sheet.");
      const previous = entry.lastEventAt ? Math.floor(entry.lastEventAt / 1000) : 0;
      const createdAt = Math.max(Math.floor(Date.now() / 1000), previous + 1);
      const event = sheetEventFromRoot({ sheetId: entry.id, viewKey: entry.viewKey, editKey: entry.editKey, createdAt }, root);
      await publishEvent(event, relays);
      replaceEntries(
        entriesRef.current.map((item) =>
          item.coordinate === entry.coordinate
            ? {
                ...item,
                updatedAt: root.updatedAt,
                lastEventAt: event.created_at * 1000,
                lastEventId: event.id,
                storedBytes: storedBytesOf(root),
              }
            : item,
        ),
      );
    },
    [relays, replaceEntries],
  );

  const getEntry = useCallback(
    (idOrCoordinate: string) =>
      entriesRef.current.find(
        (entry) => entry.id === idOrCoordinate || entry.coordinate === idOrCoordinate,
      ),
    [],
  );

  const getEntryByNaddr = useCallback((naddr: string) => {
    try {
      const decoded = nip19.decode(naddr);
      if (decoded.type !== "naddr" || decoded.data.kind !== KIND_SHEET) return undefined;
      const coordinate = `${decoded.data.kind}:${decoded.data.pubkey}:${decoded.data.identifier}`;
      return entriesRef.current.find((entry) => entry.coordinate === coordinate);
    } catch {
      return undefined;
    }
  }, []);

  const parseShareLink = useCallback((naddr: string, hash: string): LibraryEntry | null => {
    try {
      const decoded = nip19.decode(naddr);
      if (decoded.type !== "naddr" || decoded.data.kind !== KIND_SHEET) return null;
      const coordinate = `${decoded.data.kind}:${decoded.data.pubkey}:${decoded.data.identifier}`;
      const existing = entriesRef.current.find((entry) => entry.coordinate === coordinate);
      if (hash.startsWith(SHARE_HASH_PREFIX)) {
        const keys = decodeNKeys(hash.slice(SHARE_HASH_PREFIX.length));
        if (!isHexKey(keys.viewKey)) return null;
        const editKey = keys.editKey && isHexKey(keys.editKey) ? keys.editKey : undefined;
        return {
          id: decoded.data.identifier,
          coordinate,
          editPubkey: decoded.data.pubkey,
          viewKey: keys.viewKey,
          editKey,
          title: existing?.title ?? "Shared spreadsheet",
          role: editKey ? "editor" : "viewer",
          ownerPubkey: existing?.ownerPubkey ?? "",
          ownerName: existing?.ownerName,
          createdAt: Date.now(),
          updatedAt: Date.now(),
          sharedWith: [],
          relayHints: decoded.data.relays,
          source: "link",
          lastEventAt: existing?.lastEventAt,
          lastEventId: existing?.lastEventId,
        };
      }
      return existing ?? null;
    } catch {
      return null;
    }
  }, []);

  const addLinkEntry = useCallback(
    (entry: LibraryEntry) => {
      const current = entriesRef.current;
      const existing = current.find((item) => item.coordinate === entry.coordinate);
      if (existing) {
        if (!existing.editKey && entry.editKey) {
          const updated: LibraryEntry = {
            ...existing,
            editKey: entry.editKey,
            role: "editor",
          };
          const next: LibraryEntry[] = current.map((item) =>
            item.coordinate === entry.coordinate ? updated : item,
          );
          replaceEntries(next);
          void persistEntry(updated).catch(() => {});
        }
        return;
      }
      const next = [entry, ...current];
      replaceEntries(next);
      void persistEntry(entry).catch(() => {});
    },
    [persistEntry, replaceEntries],
  );

  const loadDocument = useCallback(
    async (entry: LibraryEntry): Promise<LoadedDocument> => {
      const local = loadDocCache(entry);
      if (local?.pending && Date.now() - local.savedAt < 15_000) {
        return { doc: local.doc, event: null, localOnly: true, pending: true };
      }
      const event = await fetchLatestEvent(
        relays,
        { kinds: [KIND_SHEET], authors: [entry.editPubkey], "#d": [entry.id], limit: 1 },
        7000,
      );

      if (event) {
        const remoteTime = event.created_at * 1000;
        if (
          local &&
          (local.pending ||
            (local.savedAt > remoteTime && local.doc.updatedAt >= remoteTime))
        ) {
          return { doc: local.doc, event, localOnly: false, pending: !!local.pending };
        }
        let remoteDocument: SheetDocument;
        try {
          remoteDocument = await decodeSheetEvent(
            event,
            entry.viewKey,
            entry.title,
            { servers: serversRef.current, reuse: reuseFor(entry.coordinate) },
            true,
            (bytes) => noteSize(entry.coordinate, bytes),
          );
        } catch (error) {
          if (error instanceof BlobUnavailableError || error instanceof SheetTooLargeToOpenError) throw error;
          throw new Error("Couldn't decrypt this sheet. The access keys may have changed.");
        }
        if (entry.title !== remoteDocument.title) {
          const next = entriesRef.current.map((item) =>
            item.coordinate === entry.coordinate
              ? {
                  ...item,
                  title: remoteDocument.title,
                  updatedAt: remoteTime,
                  lastEventAt: remoteTime,
                  lastEventId: event.id,
                }
              : item,
          );
          replaceEntries(next);
        } else {
          const next = entriesRef.current.map((item) =>
            item.coordinate === entry.coordinate
              ? { ...item, updatedAt: remoteTime, lastEventAt: remoteTime, lastEventId: event.id }
              : item,
          );
          replaceEntries(next);
        }
        saveDocCache(entry, remoteDocument, {
          savedAt: remoteTime,
          eventCreatedAt: remoteTime,
          eventId: event.id,
          pending: false,
        });
        return { doc: remoteDocument, event, localOnly: false, pending: false };
      }

      if (local) {
        return { doc: local.doc, event: null, localOnly: true, pending: !!local.pending };
      }
      throw new Error("This sheet has no saved data yet.");
    },
    [relays, replaceEntries, reuseFor],
  );

  const fetchRemoteDocument = useCallback(
    async (entry: LibraryEntry) => {
      const event = await fetchLatestEvent(
        relays,
        { kinds: [KIND_SHEET], authors: [entry.editPubkey], "#d": [entry.id], limit: 1 },
        7000,
      );
      if (!event) return null;
      try {
        const doc = await decodeSheetEvent(
          event,
          entry.viewKey,
          entry.title,
          { servers: serversRef.current, reuse: reuseFor(entry.coordinate) },
          true,
          (bytes) => noteSize(entry.coordinate, bytes),
        );
        return { doc, event };
      } catch (error) {
        if (error instanceof BlobUnavailableError || error instanceof SheetTooLargeToOpenError) throw error;
        throw new Error("Couldn't decrypt the latest version. The access keys may have changed.");
      }
    },
    [relays, reuseFor],
  );

  const saveDocument = useCallback(
    async (entry: LibraryEntry, document: SheetDocument) => {
      if (!entry.editKey) throw new Error("You have view-only access to this sheet.");
      const now = Date.now();
      const nextDocument: SheetDocument = { ...document, updatedAt: now };
      const createdAt = Math.max(
        Math.floor(now / 1000),
        Math.floor((entry.lastEventAt ?? 0) / 1000) + 1,
      );
      // Keep the device copy first: building the event may upload to Blossom, and if that
      // fails the edits must survive a reload and be retried, like a failed relay publish.
      saveDocCache(entry, nextDocument, { savedAt: now, pending: true });
      let event: Event;
      try {
        event = await buildEvent(entry.coordinate, {
          sheetId: entry.id,
          document: nextDocument,
          viewKey: entry.viewKey,
          editKey: entry.editKey,
          createdAt,
        });
      } catch (err) {
        replaceEntries(
          entriesRef.current.map((item) =>
            item.coordinate === entry.coordinate ? { ...item, title: nextDocument.title, updatedAt: now } : item,
          ),
        );
        throw err;
      }

      try {
        const result = await publishEvent(event, relays);
        saveDocCache(entry, nextDocument, {
          savedAt: now,
          eventCreatedAt: event.created_at * 1000,
          eventId: event.id,
          pending: false,
        });
        const next = entriesRef.current.map((item) =>
          item.coordinate === entry.coordinate
            ? {
                ...item,
                title: nextDocument.title,
                updatedAt: now,
                lastEventAt: event.created_at * 1000,
                lastEventId: event.id,
              }
            : item,
        );
        replaceEntries(next);
        return { event, accepted: result.accepted };
      } catch (err) {
        saveDocCache(entry, nextDocument, {
          savedAt: now,
          eventCreatedAt: event.created_at * 1000,
          eventId: event.id,
          pending: true,
        });
        const next = entriesRef.current.map((item) =>
          item.coordinate === entry.coordinate
            ? { ...item, title: nextDocument.title, updatedAt: now, lastEventAt: event.created_at * 1000 }
            : item,
        );
        replaceEntries(next);
        throw err;
      }
    },
    [buildEvent, relays, replaceEntries],
  );

  const updateTitle = useCallback(
    async (entry: LibraryEntry, title: string) => {
      const current = entriesRef.current;
      const updated: LibraryEntry = { ...entry, title };
      const next = current.map((item) =>
        item.coordinate === entry.coordinate ? updated : item,
      );
      replaceEntries(next);
      if (entry.source !== "grant") {
        try {
          await persistEntry(updated);
        } catch {
          setError("The new title is saved locally, but its library metadata didn't reach a relay.");
        }
      }
    },
    [persistEntry, replaceEntries],
  );

  const deleteSheet = useCallback(
    async (entry: LibraryEntry) => {
      if (entry.role === "owner" && entry.editKey) {
        try {
          const deletion = buildSheetDeletionEvent({
            sheetId: entry.id,
            editKey: entry.editKey,
            eventId: entry.lastEventId,
          });
          await publishEvent(deletion, relays);
        } catch {
          // Local removal still happens; relays may keep old ciphertext.
        }
      }
      const next = entriesRef.current.filter((item) => item.coordinate !== entry.coordinate);
      replaceEntries(next);
      if (entry.source !== "grant") {
        try {
          await persistEntry(entry, true);
        } catch {
          // Local cache still has the change.
        }
      } else if (pubkey) {
        const currentHidden = loadHiddenCoordinates(pubkey);
        saveHiddenCoordinates(pubkey, Array.from(new Set([...currentHidden, entry.coordinate])));
      }
    },
    [persistEntry, pubkey, relays, replaceEntries],
  );

  const shareSheet = useCallback(
    async (entry: LibraryEntry, collaborator: Collaborator, role: "viewer" | "editor") => {
      if (!signer || !pubkey) throw new Error("Please sign in first.");
      if (entry.role !== "owner") throw new Error("Only the owner can manage access.");
      const now = Date.now();
      const recipientRelays = await fetchUserRelays(relays, collaborator.pubkey).catch(() => []);
      const grantRelays = mergeRelays(relays, entry.relayHints ?? [], recipientRelays);
      const payload: SheetGrantPayload = {
        version: 1,
        sheetId: entry.id,
        coordinate: entry.coordinate,
        editPubkey: entry.editPubkey,
        viewKey: entry.viewKey,
        editKey: role === "editor" ? entry.editKey : undefined,
        title: entry.title,
        role,
        ownerPubkey: pubkey,
        ownerName: entry.ownerName,
        grantedAt: now,
        color: entry.color,
        relayHints: entry.relayHints,
      };
      const content = await signer.nip44Encrypt(collaborator.pubkey, JSON.stringify(payload));
      const signed = await signer.signEvent({
        kind: KIND_SHEET_GRANT,
        created_at: Math.floor(now / 1000),
        tags: [
          ["d", `${entry.id}:${collaborator.pubkey}`],
          ["p", collaborator.pubkey],
          ["a", entry.coordinate],
          ["t", "sheet-grant"],
          ["client", CLIENT_TAG],
          ["encrypted", "nip44"],
        ],
        content,
      });
      await publishEvent(signed, grantRelays);

      const nextCollaborator: Collaborator = {
        ...collaborator,
        role,
        grantedAt: now,
      };
      const updated: LibraryEntry = {
        ...entry,
        sharedWith: [
          ...entry.sharedWith.filter((person) => person.pubkey !== collaborator.pubkey),
          nextCollaborator,
        ],
      };
      const next = entriesRef.current.map((item) =>
        item.coordinate === entry.coordinate ? updated : item,
      );
      replaceEntries(next);
      await persistEntry(updated);
    },
    [persistEntry, pubkey, relays, replaceEntries, signer],
  );

  const revokeAccess = useCallback(
    async (entry: LibraryEntry, targetPubkey: string) => {
      if (!signer || !pubkey) throw new Error("Please sign in first.");
      if (entry.role !== "owner") throw new Error("Only the owner can manage access.");
      const recipientRelays = await fetchUserRelays(relays, targetPubkey).catch(() => []);
      const grantRelays = mergeRelays(relays, entry.relayHints ?? [], recipientRelays);
      const dTag = `${entry.id}:${targetPubkey}`;
      const existing = await fetchLatestEvent(
        grantRelays,
        {
          kinds: [KIND_SHEET_GRANT],
          authors: [pubkey],
          "#d": [dTag],
          limit: 1,
        },
        5000,
      );
      const now = Date.now();
      const createdAt = Math.max(
        Math.floor(now / 1000),
        (existing?.created_at ?? 0) + 1,
      );
      const payload: SheetGrantPayload = {
        version: 1,
        sheetId: entry.id,
        coordinate: entry.coordinate,
        editPubkey: entry.editPubkey,
        viewKey: entry.viewKey,
        title: entry.title,
        role: "viewer",
        ownerPubkey: pubkey,
        ownerName: entry.ownerName,
        grantedAt: now,
        revoked: true,
        color: entry.color,
        relayHints: entry.relayHints,
      };
      const content = await signer.nip44Encrypt(targetPubkey, JSON.stringify(payload));
      const signed = await signer.signEvent({
        kind: KIND_SHEET_GRANT,
        created_at: createdAt,
        tags: [
          ["d", dTag],
          ["p", targetPubkey],
          ["a", entry.coordinate],
          ["t", "sheet-grant"],
          ["revoked", "1"],
          ["client", CLIENT_TAG],
          ["encrypted", "nip44"],
        ],
        content,
      });
      await publishEvent(signed, grantRelays);

      const updated: LibraryEntry = {
        ...entry,
        sharedWith: entry.sharedWith.filter((person) => person.pubkey !== targetPubkey),
      };
      const next = entriesRef.current.map((item) =>
        item.coordinate === entry.coordinate ? updated : item,
      );
      replaceEntries(next);
      await persistEntry(updated);
    },
    [persistEntry, pubkey, relays, replaceEntries, signer],
  );

  const subscribeToEntry = useCallback(
    (entry: LibraryEntry, onEvent: (event: Event) => void) =>
      subscribeSheet(mergeRelays(relays, entry.relayHints ?? []), entry.editPubkey, entry.id, onEvent),
    [relays],
  );

  const buildShareLink = useCallback(
    (entry: LibraryEntry, mode: "view" | "edit") => {
      const naddr = nip19.naddrEncode({
        kind: KIND_SHEET,
        pubkey: entry.editPubkey,
        identifier: entry.id,
        relays: entry.relayHints?.slice(0, 3) ?? relays.slice(0, 3),
      });
      const keys: Record<string, string> = { viewKey: entry.viewKey };
      if (mode === "edit" && entry.editKey) keys.editKey = entry.editKey;
      return `${window.location.origin}/s/${naddr}${SHARE_HASH_PREFIX}${encodeNKeys(keys)}`;
    },
    [relays],
  );

  const value = useMemo<LibraryContextValue>(
    () => ({
      entries: entries.map((entry) => (entry.storedBytes ? entry : { ...entry, storedBytes: sizesRef.current.get(entry.coordinate) })),
      loading,
      error,
      refresh,
      createSheet,
      importDocument,
      importStreamed,
      bigSheetUpload,
      publishBigRoot,
      getEntry,
      getEntryByNaddr,
      parseShareLink,
      addLinkEntry,
      loadDocument,
      fetchRemoteDocument,
      saveDocument,
      updateTitle,
      deleteSheet,
      shareSheet,
      revokeAccess,
      subscribeToEntry,
      buildShareLink,
      blossomServers,
      refreshBlossomServers: fetchServers,
      saveBlossomServers,
      uploadPrefs,
      saveUploadPrefs,
    }),
    [
      entries,
      loading,
      error,
      refresh,
      createSheet,
      importDocument,
      importStreamed,
      bigSheetUpload,
      publishBigRoot,
      getEntry,
      getEntryByNaddr,
      parseShareLink,
      addLinkEntry,
      loadDocument,
      fetchRemoteDocument,
      saveDocument,
      updateTitle,
      deleteSheet,
      shareSheet,
      revokeAccess,
      subscribeToEntry,
      buildShareLink,
      blossomServers,
      fetchServers,
      saveBlossomServers,
      uploadPrefs,
      saveUploadPrefs,
    ],
  );

  return <LibraryContext.Provider value={value}>{children}</LibraryContext.Provider>;
}

export function useLibrary(): LibraryContextValue {
  const context = useContext(LibraryContext);
  if (!context) throw new Error("useLibrary must be used within LibraryProvider");
  return context;
}
