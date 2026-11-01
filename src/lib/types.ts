import type { Event } from "nostr-tools";
import type { WorkbookDocument } from "./workbook";

export interface UserRef {
  pubkey: string;
  npub: string;
  name?: string;
  picture?: string;
}

export type SheetRole = "owner" | "editor" | "viewer";

export type SheetDocument = WorkbookDocument;

export interface Collaborator {
  pubkey: string;
  npub: string;
  name?: string;
  picture?: string;
  role: "viewer" | "editor";
  grantedAt?: number;
}

export interface LibraryEntry {
  id: string;
  coordinate: string;
  editPubkey: string;
  viewKey: string;
  editKey?: string;
  title: string;
  role: SheetRole;
  ownerPubkey: string;
  ownerName?: string;
  ownerPicture?: string;
  createdAt: number;
  updatedAt: number;
  lastOpenedAt?: number;
  sharedWith: Collaborator[];
  template?: string;
  color?: string;
  relayHints?: string[];
  lastEventAt?: number;
  lastEventId?: string;
  source?: "owned" | "grant" | "link";
}

export interface SheetIndexPayload {
  version: 1;
  entries: LibraryEntry[];
}

export interface SheetMetaPayload {
  version: 1;
  entry: LibraryEntry;
  deleted?: boolean;
}

export interface SheetGrantPayload {
  version: 1;
  sheetId: string;
  coordinate: string;
  editPubkey: string;
  viewKey: string;
  editKey?: string;
  title: string;
  role: "viewer" | "editor";
  ownerPubkey: string;
  ownerName?: string;
  grantedAt: number;
  revoked?: boolean;
  color?: string;
  relayHints?: string[];
}

export interface GrantRecord {
  event: Event;
  payload: SheetGrantPayload;
}

export interface Profile {
  pubkey: string;
  name?: string;
  displayName?: string;
  about?: string;
  picture?: string;
  nip05?: string;
  lud16?: string;
  updatedAt?: number;
}

export interface DocCache {
  doc: SheetDocument;
  savedAt: number;
  eventCreatedAt?: number;
  eventId?: string;
  pending?: boolean;
}

export interface LoadedDocument {
  doc: SheetDocument;
  event: Event | null;
  localOnly: boolean;
  pending: boolean;
}

export interface SheetTemplate {
  id: string;
  name: string;
  description: string;
  color: string;
  build: (title: string) => WorkbookDocument;
}

export type SaveState = "idle" | "unsaved" | "saving" | "saved" | "offline" | "readonly";
