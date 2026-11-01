import { nip19 } from "nostr-tools";

export function cx(...parts: Array<string | false | null | undefined>): string {
  return parts.filter(Boolean).join(" ");
}

export function randomId(length = 12): string {
  const bytes = crypto.getRandomValues(new Uint8Array(length));
  let out = "";
  const alphabet = "abcdefghijklmnopqrstuvwxyz0123456789";
  for (const byte of bytes) out += alphabet[byte % alphabet.length];
  return out;
}

export function colName(index: number): string {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function colIndex(name: string): number {
  const clean = name.replace(/\$/g, "").toUpperCase();
  let out = 0;
  for (const char of clean) {
    out = out * 26 + (char.charCodeAt(0) - 64);
  }
  return out - 1;
}

export function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

export function parseCellKey(key: string): { row: number; col: number } | null {
  const [row, col] = key.split(":").map(Number);
  if (!Number.isFinite(row) || !Number.isFinite(col)) return null;
  return { row, col };
}

export function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

export function toNpub(pubkey: string): string {
  return nip19.npubEncode(pubkey);
}

export function parsePubkey(input: string): string {
  const value = input.trim();
  if (/^[0-9a-f]{64}$/i.test(value)) return value.toLowerCase();
  try {
    const decoded = nip19.decode(value);
    if (decoded.type === "npub") return decoded.data;
    if (decoded.type === "nprofile") return decoded.data.pubkey;
  } catch {
    throw new Error("Enter a valid npub or hex public key.");
  }
  throw new Error("Enter a valid npub or hex public key.");
}

export function shortKey(key: string, chars = 6): string {
  if (!key) return "";
  if (key.length <= chars * 2 + 3) return key;
  return `${key.slice(0, chars)}…${key.slice(-chars)}`;
}

export function shortNpub(pubkey: string): string {
  if (!pubkey) return "";
  try {
    return shortKey(nip19.npubEncode(pubkey), 8);
  } catch {
    return shortKey(pubkey, 6);
  }
}

export function formatRelativeTime(timestamp?: number): string {
  if (!timestamp) return "Just now";
  const diff = Date.now() - timestamp;
  const minute = 60_000;
  const hour = 60 * minute;
  const day = 24 * hour;
  if (diff < minute) return "Just now";
  if (diff < hour) return `${Math.floor(diff / minute)}m ago`;
  if (diff < day) return `${Math.floor(diff / hour)}h ago`;
  if (diff < 7 * day) return `${Math.floor(diff / day)}d ago`;
  return new Date(timestamp).toLocaleDateString(undefined, {
    month: "short",
    day: "numeric",
    year: new Date(timestamp).getFullYear() === new Date().getFullYear() ? undefined : "numeric",
  });
}

export function normalizeRelayUrl(url: string): string {
  try {
    const parsed = new URL(url);
    const path = parsed.pathname === "/" ? "" : parsed.pathname.replace(/\/$/, "");
    return `${parsed.protocol}//${parsed.host}${path}${parsed.search}`.toLowerCase();
  } catch {
    return url.trim().replace(/\/+$/, "").toLowerCase();
  }
}

export function mergeRelays(...lists: string[][]): string[] {
  const seen = new Map<string, string>();
  for (const list of lists) {
    for (const relay of list) {
      const url = relay.trim();
      if (!url) continue;
      const key = normalizeRelayUrl(url);
      if (!seen.has(key)) seen.set(key, url);
    }
  }
  return Array.from(seen.values());
}

export function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function isEditableTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  const tag = target.tagName.toLowerCase();
  return (
    tag === "input" ||
    tag === "textarea" ||
    tag === "select" ||
    target.isContentEditable
  );
}

export function downloadText(filename: string, text: string, type = "text/plain"): void {
  const blob = new Blob([text], { type: `${type};charset=utf-8` });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function fileSafeName(name: string): string {
  return (
    name
      .trim()
      .replace(/[\\/:*?"<>|]+/g, "-")
      .replace(/\s+/g, " ")
      .slice(0, 80) || "spreadsheet"
  );
}

export function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return `${parts[0][0]}${parts[parts.length - 1][0]}`.toUpperCase();
}

export function colorFromString(value: string): { bg: string; fg: string } {
  let hash = 0;
  for (let i = 0; i < value.length; i++) {
    hash = (hash * 31 + value.charCodeAt(i)) | 0;
  }
  const palette = [
    { bg: "#FFEDE6", fg: "#9A3412" },
    { bg: "#FEF3C7", fg: "#92400E" },
    { bg: "#DCFCE7", fg: "#166534" },
    { bg: "#DBEAFE", fg: "#1E40AF" },
    { bg: "#EDE9FE", fg: "#5B21B6" },
    { bg: "#FCE7F3", fg: "#9D174D" },
    { bg: "#CCFBF1", fg: "#115E59" },
    { bg: "#F1F5F9", fg: "#334155" },
  ];
  return palette[Math.abs(hash) % palette.length];
}
