import type { DocCache, LibraryEntry, SheetDocument } from "./types";

const PREFIX = "formstr-sheets:";

function read<T>(key: string): T | null {
  try {
    const raw = localStorage.getItem(PREFIX + key);
    return raw ? (JSON.parse(raw) as T) : null;
  } catch {
    return null;
  }
}

function write(key: string, value: unknown): boolean {
  try {
    localStorage.setItem(PREFIX + key, JSON.stringify(value));
    return true;
  } catch {
    // quota / private mode failures
    return false;
  }
}

function remove(key: string): void {
  try {
    localStorage.removeItem(PREFIX + key);
  } catch {
    // ignore
  }
}

export function loadLibraryCache(pubkey: string): LibraryEntry[] {
  return read<LibraryEntry[]>(`library:${pubkey}`) ?? [];
}

export function saveLibraryCache(pubkey: string, entries: LibraryEntry[]): void {
  write(`library:${pubkey}`, entries);
}

export function loadDocCache(entry: LibraryEntry): DocCache | null {
  return read<DocCache>(`doc:${entry.coordinate}`);
}

/**
 * Stores the device copy of a sheet. Returns false when it could not be stored
 * (usually because a large sheet exceeds the localStorage quota). In that case
 * any older copy is removed: left in place, an older copy marked pending would
 * be preferred on the next load over the newer version.
 */
export function saveDocCache(
  entry: LibraryEntry,
  doc: SheetDocument,
  meta: { savedAt?: number; eventCreatedAt?: number; eventId?: string; pending?: boolean } = {},
): boolean {
  const key = `doc:${entry.coordinate}`;
  const stored = write(key, {
    doc,
    savedAt: meta.savedAt ?? Date.now(),
    eventCreatedAt: meta.eventCreatedAt,
    eventId: meta.eventId,
    pending: meta.pending,
  } satisfies DocCache);
  if (!stored) remove(key);
  return stored;
}

export function removeDocCache(entry: LibraryEntry): void {
  remove(`doc:${entry.coordinate}`);
}

export function clearLibraryCache(pubkey: string): void {
  remove(`library:${pubkey}`);
}

export function loadHiddenCoordinates(pubkey: string): string[] {
  return read<string[]>(`hidden:${pubkey}`) ?? [];
}

export function saveHiddenCoordinates(pubkey: string, coordinates: string[]): void {
  write(`hidden:${pubkey}`, coordinates);
}

export function loadProfileCache(): Record<string, unknown> {
  return read<Record<string, unknown>>("profiles") ?? {};
}

export function saveProfileCache(profiles: Record<string, unknown>): void {
  write("profiles", profiles);
}

export function getTheme(): "light" | "dark" {
  try {
    const theme =
      localStorage.getItem(`${PREFIX}theme`) ??
      localStorage.getItem("formstr-sheets-theme");
    if (theme === "light" || theme === "dark") return theme;
  } catch {
    // ignore
  }
  return typeof window !== "undefined" &&
    window.matchMedia?.("(prefers-color-scheme: dark)").matches
    ? "dark"
    : "light";
}

export function setTheme(theme: "light" | "dark"): void {
  try {
    localStorage.setItem(`${PREFIX}theme`, theme);
    localStorage.removeItem("formstr-sheets-theme");
  } catch {
    // ignore
  }
  document.documentElement.dataset.theme = theme;
  document
    .querySelector('meta[name="theme-color"]')
    ?.setAttribute("content", theme === "dark" ? "#0d0f10" : "#faf8f5");
}

/**
 * Which of the user's Blossom servers uploads go to. Every ticked server gets a copy of every
 * blob, so ticking N servers means N-fold upload bandwidth and storage.
 */
export interface UploadPrefs {
  /** Servers the user unticked: still used to read sheets, never uploaded to. */
  off: string[];
}

export function loadUploadPrefs(pubkey: string): UploadPrefs {
  const stored = read<Partial<UploadPrefs>>(`upload-prefs:${pubkey}`);
  return {
    off: Array.isArray(stored?.off) ? stored.off.filter((url) => typeof url === "string") : [],
  };
}

export function saveUploadPrefs(pubkey: string, prefs: UploadPrefs): void {
  write(`upload-prefs:${pubkey}`, prefs);
}
