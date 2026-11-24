// Size information for the File info dialog: how big a sheet is and how it would be stored.

import { DEFAULT_CHUNK_SIZE } from "@formstr/drive-sdk";
import type { WorkbookDocument } from "../workbook";
import { MAX_CHUNK_BYTES, planTab, type PlannedChunk } from "./chunks";
import {
  INLINE_MAX_BYTES,
  buildBlossomRoot,
  buildInlineRoot,
  rootSize,
  type BlobRef,
  type ChunkEntry,
  type SheetStorage,
} from "./root";

const AES_GCM_TAG_BYTES = 16;
const utf8 = new TextEncoder();

export interface TabInfo {
  id: string;
  name: string;
  cells: number;
  /** The tab as plain JSON. */
  rawBytes: number;
  /** The tab compressed and encrypted, i.e. the total size of its Blossom blobs. */
  storedBytes: number;
  /** How many chunks (blobs) the tab is cut into. */
  chunks: number;
}

export interface FileInfo {
  tabs: TabInfo[];
  cells: number;
  /** The whole workbook as plain JSON: roughly what the old format put in one event. */
  rawBytes: number;
  storage: SheetStorage;
  /** Size of the Nostr event's content before encryption. */
  eventBytes: number;
  /** Total size of the encrypted Blossom files; 0 when the sheet is stored inline. */
  blobBytes: number;
  blobCount: number;
}

/** Ciphertext size of a plaintext of `plainBytes`: one 16-byte tag per 64 KiB segment. */
export function encryptedSize(plainBytes: number, segmentBytes = DEFAULT_CHUNK_SIZE): number {
  return plainBytes + Math.max(1, Math.ceil(plainBytes / segmentBytes)) * AES_GCM_TAG_BYTES;
}

/** Measures a workbook the way it would be saved right now. */
export async function measureWorkbook(doc: WorkbookDocument): Promise<FileInfo> {
  // Cut the tabs exactly as a save would (default chunk size, no earlier layout to follow).
  const plans: PlannedChunk[][] = [];
  for (const sheet of doc.sheets) plans.push(await planTab(sheet, { max: MAX_CHUNK_BYTES }));
  const tabs: TabInfo[] = doc.sheets.map((sheet, index) => ({
    id: sheet.id,
    name: sheet.name,
    cells: Object.keys(sheet.cells).length,
    rawBytes: utf8.encode(JSON.stringify(sheet)).length,
    storedBytes: plans[index].reduce((sum, chunk) => sum + encryptedSize(chunk.size), 0),
    chunks: plans[index].length,
  }));
  const base = { tabs, cells: tabs.reduce((sum, tab) => sum + tab.cells, 0), rawBytes: utf8.encode(JSON.stringify(doc)).length };

  const inline = rootSize(await buildInlineRoot(doc));
  if (inline <= INLINE_MAX_BYTES) {
    return { ...base, storage: "inline", eventBytes: inline, blobBytes: 0, blobCount: 0 };
  }

  // The event only holds pointers; size it with placeholder refs of the same shape.
  const placeholder = (chunk: { size: number }): BlobRef => ({
    x: "0".repeat(64),
    k: "0".repeat(64),
    n: encryptedSize(chunk.size),
    p: chunk.size,
    c: DEFAULT_CHUNK_SIZE,
    h: "0".repeat(64),
  });
  const lists: ChunkEntry[][] = plans.map((chunks) =>
    chunks.map((chunk) => ({ id: "00000000", ref: placeholder(chunk), rowStart: chunk.rowStart, rowCount: chunk.rowCount })),
  );
  return {
    ...base,
    storage: "blossom",
    eventBytes: rootSize(buildBlossomRoot(doc, lists, ["https://blossom.example.com"])),
    blobBytes: tabs.reduce((sum, tab) => sum + tab.storedBytes, 0),
    blobCount: tabs.reduce((sum, tab) => sum + tab.chunks, 0),
  };
}
