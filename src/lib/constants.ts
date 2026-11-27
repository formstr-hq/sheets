export const APP_NAME = "Formstr Sheets";
export const APP_URL = "https://sheets.formstr.app";
export const CLIENT_TAG = "formstr-sheets";

export const KIND_SHEET = 33460;
export const KIND_SHEET_GRANT = 33461;
export const KIND_SHEET_META = 33462;
export const KIND_SHEET_INDEX = KIND_SHEET_META;
/** BUD-03: a user's list of Blossom servers, as `server` tags. */
export const KIND_BLOSSOM_SERVERS = 10063;

export const DEFAULT_RELAYS = [
  "wss://relay.formstr.app",
  "wss://relay.damus.io",
  "wss://relay.primal.net",
  "wss://nos.lol",
  "wss://relay.nostr.band",
];

/**
 * Automatic saving is off for now: every save of a big sheet can ask the signer for approval, so
 * saves happen only when the user presses Save (or Ctrl+S). Flip to true to restore autosave.
 */
export const AUTOSAVE_ENABLED = false;

/**
 * Editing of sheets too big for the main editor (BigSheetView). Off = big sheets open read-only
 * (still scrollable to every row). Edits store typed values; formulas are not recalculated.
 */
export const BIG_SHEET_EDITING = true;

/**
 * Opens small sheets in the big-sheet grid too (held in memory, saved inline) instead of FortuneSheet.
 * Opt-in while the new grid reaches parity: add `grid=new` to the address. Remove with FortuneSheet.
 */
export function newGridRequested(): boolean {
  try {
    return /[?&]grid=new/.test(window.location.href);
  } catch {
    return false;
  }
}

export const DEFAULT_ROWS = 200;
export const DEFAULT_COLS = 26;

export const DEFAULT_COL_WIDTH = 116;
export const ROW_HEADER_WIDTH = 46;
export const ROW_HEIGHT = 30;
export const HEADER_HEIGHT = 30;

export const SHEET_COLORS = [
  "#FF5A36",
  "#F59E0B",
  "#10B981",
  "#0EA5E9",
  "#6366F1",
  "#EC4899",
  "#8B5CF6",
  "#64748B",
];

export const CELL_COLORS = [
  "#141719",
  "#4B5563",
  "#9CA3AF",
  "#FFFFFF",
  "#EF4444",
  "#F97316",
  "#F59E0B",
  "#84CC16",
  "#10B981",
  "#06B6D4",
  "#3B82F6",
  "#6366F1",
  "#8B5CF6",
  "#EC4899",
  "#FDE68A",
  "#FECACA",
  "#A7F3D0",
  "#BAE6FD",
  "#DDD6FE",
  "#FBCFE8",
];

export const PILL_COLORS = [
  { bg: "#FFEDE6", fg: "#9A3412" },
  { bg: "#FEF3C7", fg: "#92400E" },
  { bg: "#DCFCE7", fg: "#166534" },
  { bg: "#DBEAFE", fg: "#1E40AF" },
  { bg: "#EDE9FE", fg: "#5B21B6" },
  { bg: "#FCE7F3", fg: "#9D174D" },
  { bg: "#CCFBF1", fg: "#115E59" },
  { bg: "#F1F5F9", fg: "#334155" },
];

export const FONT_FAMILIES = [
  { label: "Inter", value: "Inter" },
  { label: "System", value: "ui-sans-serif" },
  { label: "Mono", value: "ui-monospace" },
  { label: "Serif", value: "Georgia" },
] as const;

export const FONT_SIZES = [10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32] as const;

export const NUMBER_FORMATS = [
  { label: "Automatic", value: "auto" },
  { label: "Plain text", value: "plain" },
  { label: "Number", value: "number" },
  { label: "Currency", value: "currency" },
  { label: "Percent", value: "percent" },
  { label: "Date", value: "date" },
  { label: "Date & time", value: "datetime" },
  { label: "Time", value: "time" },
] as const;

export const SHARE_HASH_PREFIX = "#nkeys=";
