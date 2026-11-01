import { nip19 } from "nostr-tools";
import { KIND_SHEET } from "./constants";
import type { LibraryEntry } from "./types";

export function naddrForEntry(entry: LibraryEntry): string {
  return nip19.naddrEncode({
    kind: KIND_SHEET,
    pubkey: entry.editPubkey,
    identifier: entry.id,
    relays: entry.relayHints?.slice(0, 3),
  });
}

export function sheetPath(entry: LibraryEntry): string {
  return `/s/${naddrForEntry(entry)}`;
}
