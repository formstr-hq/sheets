import {
  finalizeEvent,
  generateSecretKey,
  getPublicKey,
  nip44,
  type Event,
  type EventTemplate,
} from "nostr-tools";
import { bytesToHex, hexToBytes } from "nostr-tools/utils";
import { KIND_SHEET, CLIENT_TAG } from "./constants";
import type { SheetDocument } from "./types";

export interface DocKeys {
  viewKey: string;
  editKey: string;
  editPubkey: string;
}

export function generateDocKeys(): DocKeys {
  const view = generateSecretKey();
  const edit = generateSecretKey();
  return {
    viewKey: bytesToHex(view),
    editKey: bytesToHex(edit),
    editPubkey: getPublicKey(edit),
  };
}

export function conversationKeyForSecret(secretHex: string): Uint8Array {
  const secret = hexToBytes(secretHex);
  return nip44.v2.utils.getConversationKey(secret, getPublicKey(secret));
}

export function encryptWithSecret(plaintext: string, secretHex: string): string {
  return nip44.v2.encrypt(plaintext, conversationKeyForSecret(secretHex));
}

export function decryptWithSecret(ciphertext: string, secretHex: string): string {
  return nip44.v2.decrypt(ciphertext, conversationKeyForSecret(secretHex));
}

export function buildSheetEvent(params: {
  sheetId: string;
  document: SheetDocument;
  viewKey: string;
  editKey: string;
  createdAt?: number;
}): Event {
  const createdAt = params.createdAt ?? Math.floor(Date.now() / 1000);
  const template: EventTemplate = {
    kind: KIND_SHEET,
    created_at: createdAt,
    tags: [
      ["d", params.sheetId],
      ["t", "formstr-sheet"],
      ["client", CLIENT_TAG],
      ["encrypted", "nip44"],
    ],
    content: encryptWithSecret(JSON.stringify(params.document), params.viewKey),
  };
  return finalizeEvent(template, hexToBytes(params.editKey));
}

export function buildSheetDeletionEvent(params: {
  sheetId: string;
  editKey: string;
  eventId?: string;
  createdAt?: number;
}): Event {
  const tags: string[][] = [["a", `${KIND_SHEET}:${getPublicKey(hexToBytes(params.editKey))}:${params.sheetId}`]];
  if (params.eventId) tags.push(["e", params.eventId]);
  const template: EventTemplate = {
    kind: 5,
    created_at: params.createdAt ?? Math.floor(Date.now() / 1000),
    tags,
    content: "Delete spreadsheet",
  };
  return finalizeEvent(template, hexToBytes(params.editKey));
}
