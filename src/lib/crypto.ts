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
import { normalizeWorkbook } from "./workbook";
import type { FileSigner } from "@formstr/drive-sdk";
import {
  FRAME_TYPE_TAB_INDEX,
  INLINE_MAX_BYTES,
  ROOT_MAX_BYTES,
  ROOT_VERSION,
  SheetTooLargeError,
  assertOpenable,
  buildBlossomRoot,
  buildInlineRoot,
  isSheetRootV2,
  parseRootV2,
  rootSize,
  serializeRoot,
  storageOf,
  type BlobRef,
  type ChunkEntry,
  type SheetRootV2,
} from "./v2/root";
import type { WorkbookDocument } from "./workbook";
import type { UploadProgress } from "./v2/progress";
import { chunkEntry, planTab, MIN_CHUNK_BYTES, type ChunkMemory, type PlannedChunk } from "./v2/chunks";
import {
  MAX_REPLICAS,
  PayloadTooLargeError,
  chunkLimitFor,
  downloadTabBody,
  learnServerLimit,
  prepareTab,
  sanitizeServers,
  uploadFrames,
  type BlossomEnv,
  type UploadedBlob,
} from "./v2/blossom";

/** How one worksheet was last stored: its chunks and who holds them. Lets a save skip unchanged chunks. */
export type TabMemory = ChunkMemory[];

export interface BlossomContext {
  /** Servers to upload to (the user's kind 10063 list, minus the ones they unticked). */
  servers: string[];
  /** How many of `servers` each blob is copied to. The app passes all of them. Defaults to one. */
  replicas?: number;
  /** Signs the upload authorization. Required to save sheets that do not fit inline. */
  signer?: FileSigner;
  /**
   * What is known to be on Blossom for this sheet, by tab id. Read to skip unchanged chunks (an
   * identical plaintext hash), and updated after a successful save or load.
   */
  reuse?: Map<string, TabMemory>;
  env?: BlossomEnv;
  /** Told how the upload is going (shown as the upload card). Silent for inline saves. */
  onProgress?: (progress: UploadProgress) => void;
  /** Root size above which chunk lists move into index blobs. Defaults to the 48 KB limit; for tests. */
  rootMaxBytes?: number;
}

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

/** Gives up shrinking chunks below this (after repeated 413s the server is not usable). */
const MAX_PLAN_ATTEMPTS = 8;

/**
 * Cuts every tab into chunks, uploads the chunks that changed in one batch (one authorization),
 * and returns a root that points at them. A 413 from a server shrinks the chunk limit and the
 * tabs are cut again.
 */
async function buildBlossomRootFor(
  doc: SheetDocument,
  context: BlossomContext | undefined,
  inlineBytes: number,
): Promise<SheetRootV2> {
  const servers = sanitizeServers(context?.servers ?? []);
  if (!context?.signer || servers.length === 0) throw new SheetTooLargeError(inlineBytes);
  const signer = context.signer;

  const replicas = Math.max(1, context.replicas ?? MAX_REPLICAS);
  const needed = Math.min(replicas, servers.length);
  // A chunk is only reused if enough of the servers we are uploading to already hold it.
  const canReuse = (chunk: ChunkMemory) => servers.filter((url) => chunk.servers.includes(url)).length >= needed;
  const known = context.reuse;
  const report = context.onProgress;
  const progress: UploadProgress = {
    phase: "preparing",
    message: "Compressing your sheet…",
    tabs: [],
    bytesDone: 0,
    bytesTotal: 0,
    blobsDone: 0,
    blobsTotal: 0,
    servers,
  };
  const publish = (patch: Partial<UploadProgress>) => report?.({ ...progress, ...Object.assign(progress, patch) });
  const upload = (frames: Uint8Array[], onBlob?: (blobsDone: number) => void) =>
    uploadFrames(frames, {
      servers,
      signer,
      replicas,
      ...context.env,
      onProgress: (event) => {
        publish({ phase: "uploading", bytesDone: event.bytesDone, bytesTotal: event.bytesTotal, blobsDone: event.blobsDone });
        onBlob?.(event.blobsDone);
      },
    });
  publish({});

  let plans: PlannedChunk[][] = [];
  let uploaded: UploadedBlob[] = [];
  for (let attempt = 1; ; attempt++) {
    const max = chunkLimitFor(servers);
    plans = [];
    for (const sheet of doc.sheets) {
      publish({ message: `Compressing '${sheet.name}'…` });
      plans.push(await planTab(sheet, { max, previous: known?.get(sheet.id), canReuse }));
      // Let the card paint between big tabs.
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    const pending = plans.flat().filter((chunk) => !chunk.reuse);
    const tabOf = new Map<PlannedChunk, number>();
    plans.forEach((chunks, tab) => chunks.forEach((chunk) => tabOf.set(chunk, tab)));
    const tabs = doc.sheets.map((sheet, tab) => ({
      name: sheet.name,
      chunks: plans[tab].length,
      stored: plans[tab].filter((chunk) => chunk.reuse).length,
    }));
    publish({
      message: `Sending ${pending.length} changed file${pending.length === 1 ? "" : "s"} (${plans.flat().length - pending.length} unchanged)…`,
      tabs,
      blobsTotal: pending.length,
      blobsDone: 0,
      bytesDone: 0,
      bytesTotal: 0,
    });
    try {
      uploaded = await upload(
        pending.map((chunk) => chunk.frame as Uint8Array),
        (blobsDone) => {
          // The first `blobsDone` pending chunks are stored.
          const stored = doc.sheets.map((_, tab) => plans[tab].filter((chunk) => chunk.reuse).length);
          pending.slice(0, blobsDone).forEach((chunk) => (stored[tabOf.get(chunk) as number] += 1));
          publish({ tabs: tabs.map((item, tab) => ({ ...item, stored: stored[tab] })) });
        },
      );
      pending.forEach((chunk, i) => {
        chunk.reuse = { ...chunkEntry(chunk, uploaded[i].ref), servers: uploaded[i].servers };
        chunk.frame = null;
      });
      break;
    } catch (error) {
      if (!(error instanceof PayloadTooLargeError) || attempt >= MAX_PLAN_ATTEMPTS) throw error;
      learnServerLimit(error.server, error.size);
      if (chunkLimitFor(servers) <= MIN_CHUNK_BYTES) throw error;
    }
  }

  const memory: TabMemory[] = plans.map((chunks) =>
    chunks.map((chunk) => {
      const held = chunk.reuse as ChunkMemory;
      // Position can move even when the content (and so the blob) did not.
      return { ...held, id: chunk.id, rowStart: chunk.rowStart, rowCount: chunk.rowCount };
    }),
  );
  doc.sheets.forEach((sheet, index) => known?.set(sheet.id, memory[index]));
  const root = await finishBlossomRoot(doc, memory, upload, context.rootMaxBytes);
  publish({ phase: "done", message: "All files are stored and encrypted.", blobsDone: progress.blobsTotal });
  return root;
}

/**
 * Builds the root once every chunk is stored. `memory` is parallel to `doc.sheets` (only the
 * sheets' meta is read). Chunk lists that would push the root over its size limit move into
 * tab-index blobs, which `upload` stores (a second signer approval, only for ~1 GB sheets).
 */
export async function finishBlossomRoot(
  doc: WorkbookDocument,
  memory: TabMemory[],
  upload: (frames: Uint8Array[]) => Promise<UploadedBlob[]>,
  rootMaxBytes?: number,
): Promise<SheetRootV2> {
  const holders = sanitizeServers(memory.flat().flatMap((chunk) => chunk.servers));
  const lists: (ChunkEntry[] | { index: BlobRef })[] = memory.map((tab) =>
    tab.map(({ id, ref, rowStart, rowCount, cells }) =>
      cells === undefined ? { id, ref, rowStart, rowCount } : { id, ref, rowStart, rowCount, cells },
    ),
  );

  // Keep the root under its size limit: move the longest chunk lists into tab-index blobs.
  const placeholder = { index: { x: "0".repeat(64), k: "0".repeat(64), n: 0, p: 0, c: 0, h: "0".repeat(64) } };
  const moved: number[] = [];
  const trial = [...lists];
  const bySize = lists.map((_, i) => i).sort((x, y) => memory[y].length - memory[x].length);
  const rootMax = rootMaxBytes ?? ROOT_MAX_BYTES;
  for (const tab of bySize) {
    // Moving a one-chunk list into an index blob saves nothing.
    if (rootSize(buildBlossomRoot(doc, trial, holders)) <= rootMax || memory[tab].length < 2) break;
    trial[tab] = placeholder;
    moved.push(tab);
  }
  if (moved.length > 0) {
    // A second batch (and signer approval), only for sheets with hundreds of chunks (~1 GB).
    const frames = await Promise.all(
      moved.map(async (tab) => (await prepareTab({ chunks: lists[tab] }, FRAME_TYPE_TAB_INDEX)).frameBytes),
    );
    const indexes = await upload(frames);
    moved.forEach((tab, i) => {
      lists[tab] = { index: indexes[i].ref };
      holders.push(...indexes[i].servers);
    });
  }
  return buildBlossomRoot(doc, lists, sanitizeServers(holders));
}

/**
 * Bytes a sheet takes in storage: the root itself plus every encrypted blob it points at. A tab
 * whose chunk list lives in an index blob counts only that blob (the list is not downloaded).
 */
export function storedBytesOf(root: SheetRootV2): number {
  let bytes = rootSize(root);
  for (const tab of root.tabs) {
    if (tab.blob) bytes += tab.blob.n;
    if (Array.isArray(tab.chunks)) for (const chunk of tab.chunks) bytes += chunk.ref.n;
    else if (tab.chunks) bytes += tab.chunks.index.n;
  }
  return bytes;
}

/**
 * Builds the sheet event as a v2 root. A sheet whose compressed root fits in 32 KB
 * stays inline in the event; a larger one has its tabs encrypted and uploaded to
 * the user's Blossom servers, and the event only points at them. Throws
 * SheetTooLargeError when a sheet needs Blossom but no server/signer is available.
 */
export async function buildSheetEvent(params: {
  sheetId: string;
  document: SheetDocument;
  viewKey: string;
  editKey: string;
  createdAt?: number;
  blossom?: BlossomContext;
  /** Told how many bytes the sheet takes in storage once the root is built. */
  onSize?: (bytes: number) => void;
}): Promise<Event> {
  const createdAt = params.createdAt ?? Math.floor(Date.now() / 1000);
  const inlineRoot = await buildInlineRoot(params.document);
  const inlineBytes = rootSize(inlineRoot);
  const rootValue =
    inlineBytes <= INLINE_MAX_BYTES
      ? inlineRoot
      : await buildBlossomRootFor(params.document, params.blossom, inlineBytes).catch((error: unknown) => {
          // A missing server list is retried by the caller; anything else is a failed upload.
          if (!(error instanceof SheetTooLargeError)) {
            params.blossom?.onProgress?.({
              phase: "error",
              message: "The upload did not finish.",
              error: error instanceof Error ? error.message : String(error),
              tabs: [],
              bytesDone: 0,
              bytesTotal: 0,
              blobsDone: 0,
              blobsTotal: 0,
              servers: params.blossom.servers,
            });
          }
          throw error;
        });
  params.onSize?.(storedBytesOf(rootValue));
  return sheetEventFromRoot({ ...params, createdAt }, rootValue);
}

/** Encrypts a finished root with the view key and signs it with the sheet's edit key (no signer prompt). */
export function sheetEventFromRoot(
  params: { sheetId: string; viewKey: string; editKey: string; createdAt?: number },
  rootValue: SheetRootV2,
): Event {
  const createdAt = params.createdAt ?? Math.floor(Date.now() / 1000);
  const root = serializeRoot(rootValue);
  const template: EventTemplate = {
    kind: KIND_SHEET,
    created_at: createdAt,
    tags: [
      ["d", params.sheetId],
      ["t", "formstr-sheet"],
      ["client", CLIENT_TAG],
      ["encrypted", "nip44"],
      ["v", String(ROOT_VERSION)],
      ["storage", storageOf(rootValue)],
    ],
    content: encryptWithSecret(root, params.viewKey),
  };
  return finalizeEvent(template, hexToBytes(params.editKey));
}

/**
 * Reads a sheet event of any version. v2 roots are expanded; anything else is
 * treated as a v1 event, whose content is the workbook JSON itself (including
 * the older fortune-sheet format that normalizeWorkbook still understands).
 */
export async function decodeSheetEvent(
  event: Pick<Event, "content">,
  viewKey: string,
  fallbackTitle?: string,
  blossom?: Pick<BlossomContext, "servers" | "reuse" | "env">,
  /** Refuse (SheetTooLargeToOpenError) sheets too big for the editor, before downloading them. */
  gate = false,
  /** Told how many bytes the sheet takes in storage (before anything is downloaded). */
  onSize?: (bytes: number) => void,
): Promise<SheetDocument> {
  const parsed: unknown = JSON.parse(decryptWithSecret(event.content, viewKey));
  if (!isSheetRootV2(parsed)) {
    onSize?.(event.content.length);
    return normalizeWorkbook(parsed, fallbackTitle);
  }
  onSize?.(storedBytesOf(parsed));
  if (gate) assertOpenable(parsed);

  // The root names its own servers; the reader's own list is a fallback.
  const doc = await parseRootV2(
    parsed,
    (ref, rootServers, type) =>
      downloadTabBody(ref, [...rootServers, ...(blossom?.servers ?? [])], blossom?.env, type),
    (tabId, entries) => blossom?.reuse?.set(tabId, entries.map((entry) => ({ ...entry, servers: parsed.servers }))),
  );
  return normalizeWorkbook(doc, fallbackTitle);
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
