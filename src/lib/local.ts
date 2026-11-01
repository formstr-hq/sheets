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

function write(key: string, value: unknown): void {
  try {
    localStorage.setItem(PREFIX + key, JSON.stringify(value));
  } catch {
    // ignore quota / private mode failures
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

export function saveDocCache(
  entry: LibraryEntry,
  doc: SheetDocument,
  meta: { savedAt?: number; eventCreatedAt?: number; eventId?: string; pending?: boolean } = {},
): void {
  write(`doc:${entry.coordinate}`, {
    doc,
    savedAt: meta.savedAt ?? Date.now(),
    eventCreatedAt: meta.eventCreatedAt,
    eventId: meta.eventId,
    pending: meta.pending,
  } satisfies DocCache);
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
    const theme = localStorage.getItem(`${PREFIX}theme`);
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
  } catch {
    // ignore
  }
  document.documentElement.dataset.theme = theme;
  document
    .querySelector('meta[name="theme-color"]')
    ?.setAttribute("content", theme === "dark" ? "#0d0f10" : "#faf8f5");
}
