// Encrypted Blossom storage for worksheet payloads.
//
// Encryption and HTTP are @formstr/drive-sdk's (the NIP-FS segmented AES-GCM
// format Formstr Drive uses), so the blob format is shared across Formstr apps.

import {
  createBlossomAuthorization,
  createFetchBlossomTransport,
  decryptFileBytes,
  encryptFile,
  type FileSigner,
} from "@formstr/drive-sdk";
import { MAX_CHUNK_BYTES, MIN_CHUNK_BYTES } from "./chunks";
import { frame, unframe } from "./codec";
import { sha256Hex } from "./hash";
import { BlobUnavailableError, type BlobRef } from "./root";

/**
 * A blob is copied to this many of the user's servers unless the caller asks for more. The app
 * passes the number of servers the user ticked, so every ticked server gets every blob.
 */
export const MAX_REPLICAS = 1;
/**
 * One signed authorization (one signer prompt) can list many blob hashes (BUD-11 multi-x), the way
 * Drive's single blob needs one prompt however big the file is. The only limit is the HTTP header the
 * signed event travels in: ~100 bytes per hash once base64-encoded, and servers cap a header at
 * different sizes (BUD-11 sets none; nginx's default is 8 KB, Node's 16 KB). So:
 *  - up to SAFE_AUTH_HASHES hashes (~5 KB) are always sent in one authorization;
 *  - above that, each server is probed (below) for the largest header it accepts, and the whole
 *    save is signed in one authorization if every server takes it, else in the biggest batches they do.
 */
export const SAFE_AUTH_HASHES = 50;
/** Largest authorization tried (hashes). 200 is ~20 KB; a 6 MB-chunk sheet of 1.2 GB fits in one. */
export const MAX_AUTH_HASHES = 200;
const AUTH_PROBE_STEPS = [MAX_AUTH_HASHES, 100, SAFE_AUTH_HASHES, 25, 10];
/** Base64 size of an upload authorization event with no hashes, and the size each hash adds. */
const AUTH_BASE_BYTES = 700;
const AUTH_BYTES_PER_HASH = 97;

// Header sizes the servers accepted (as hash counts), learned by probing, per server for the session.
const learnedAuthHashes = new Map<string, number>();

export function forgetAuthLimits(): void {
  learnedAuthHashes.clear();
}

/**
 * The most hashes one authorization may carry for this server. Sends HEAD /upload (BUD-06, which does
 * not store anything) with a throwaway Authorization header of the size an authorization for N hashes
 * would have, largest first. A server whose parser rejects the size answers 431 or 400 (or drops the
 * connection); anything else (401 for the junk token, 200, 404...) means the header size was accepted.
 * No signer is involved, so probing never costs a prompt.
 */
async function authHashLimit(server: string, doFetch: typeof fetch): Promise<number> {
  const known = learnedAuthHashes.get(server);
  if (known !== undefined) return known;
  const probe = async (hashes: number): Promise<boolean> => {
    try {
      const response = await doFetch(`${server}/upload`, {
        method: "HEAD",
        headers: {
          Authorization: `Nostr ${"A".repeat(AUTH_BASE_BYTES + AUTH_BYTES_PER_HASH * hashes)}`,
          "X-Content-Length": "1",
          "X-Content-Type": "application/octet-stream",
          "X-SHA-256": "0".repeat(64),
        },
      });
      return response.status !== 431 && response.status !== 400;
    } catch {
      return false;
    }
  };
  // If even the smallest probe is refused, the HEAD itself is what the server dislikes (not the size),
  // so the probe tells us nothing: stay with the size every common server accepts.
  let limit = SAFE_AUTH_HASHES;
  for (const step of AUTH_PROBE_STEPS) {
    if (await probe(step)) {
      limit = step;
      break;
    }
  }
  learnedAuthHashes.set(server, limit);
  return limit;
}
/** Tries per server for one blob before moving on to the next server. */
export const UPLOAD_ATTEMPTS = 3;
const RETRY_DELAY_MS = 3000;
/** Statuses that will not change on a retry: skip straight to the next server. */
const PERMANENT_STATUSES = new Set([401, 403, 413, 415]);

/** An authorization has to outlive the slowest upload it covers: at least 30 minutes, +2 s per MB (as Drive does). */
export function authTtlSeconds(bytes: number): number {
  return Math.max(1800, Math.ceil(bytes / (1024 * 1024)) * 2);
}

export interface BlossomEnv {
  /** Pass a fetch to override the transport (tests). Without it the browser's XHR is used, which reports upload bytes. */
  fetch?: typeof fetch;
  now?: () => number;
  /** Waits between retries; tests pass an instant one. */
  sleep?: (ms: number) => Promise<void>;
}

/** A server answered 413: this blob is bigger than it accepts. The caller should cut smaller chunks. */
export class PayloadTooLargeError extends Error {
  readonly server: string;
  /** Ciphertext bytes that were refused. */
  readonly size: number;
  constructor(server: string, size: number) {
    super(`${server} refused a ${Math.ceil(size / 1024)} KB upload as too large.`);
    this.name = "PayloadTooLargeError";
    this.server = server;
    this.size = size;
  }
}

// Upload limits learned from 413 answers and BUD-06 preflights, per server, for this session.
const learnedLimits = new Map<string, number>();

/**
 * A server refused `size` bytes. Its real limit is unknown but below that, so remember 75% of it
 * (the next plan then fills to 75% of that). Repeated refusals keep shrinking it.
 */
export function learnServerLimit(server: string, size: number): number {
  const next = Math.floor(size * 0.75);
  const known = learnedLimits.get(server);
  const limit = known === undefined ? next : Math.min(known, next);
  learnedLimits.set(server, limit);
  return limit;
}

/** Largest chunk (compressed bytes) every one of these servers is known to accept. */
export function chunkLimitFor(servers: readonly string[], ceiling = MAX_CHUNK_BYTES): number {
  let max = ceiling;
  for (const server of sanitizeServers(servers)) {
    const known = learnedLimits.get(server);
    if (known !== undefined) max = Math.min(max, known);
  }
  return Math.max(MIN_CHUNK_BYTES, max);
}

export function forgetServerLimits(): void {
  learnedLimits.clear();
}

export interface UploadedBlob {
  ref: BlobRef;
  /** Servers that accepted this blob. */
  servers: string[];
}

/** Only https servers (or http for localhost development) are ever contacted. */
export function sanitizeServers(servers: readonly string[]): string[] {
  const out: string[] = [];
  for (const raw of servers) {
    try {
      const url = new URL(raw);
      const local = url.hostname === "localhost" || url.hostname === "127.0.0.1";
      if (url.protocol === "https:" || (url.protocol === "http:" && local)) {
        const clean = url.origin + url.pathname.replace(/\/+$/, "");
        if (!out.includes(clean)) out.push(clean);
      }
    } catch {
      // not a URL
    }
  }
  return out;
}

/**
 * The list to publish: what is already on the profile, minus what the user removed,
 * plus what they added. Starting from the current list means a server added in another
 * app, or one a slow relay failed to return, is not silently dropped by an edit.
 */
export function mergeServers(current: readonly string[], added: readonly string[], removed: readonly string[]): string[] {
  const dropped = new Set(sanitizeServers(removed));
  return sanitizeServers([...sanitizeServers(current).filter((url) => !dropped.has(url)), ...added]);
}

/** The plaintext frame for a tab body, with its hash, ready to compare or upload. */
export async function prepareTab(body: unknown, type?: number): Promise<{ frameBytes: Uint8Array; hash: string }> {
  const frameBytes = await frame(body, type);
  return { frameBytes, hash: await sha256Hex(frameBytes) };
}

/**
 * Encrypts and uploads frames. Each blob goes to `replicas` servers (in list order, skipping ones
 * that fail) and must land on at least one.
 *
 * - Every frame is encrypted once, up front, and the ciphertext is staged in a `Blob`, which the
 *   browser can page out of the JS heap. Only the blob being sent is read back into memory.
 * - The whole save is signed with ONE authorization when the servers accept a header that big (see
 *   SAFE_AUTH_HASHES), else in the biggest batches they accept. Signed lazily, just before the uploads.
 * - Per server: a BUD-06 preflight with the largest blob, then up to UPLOAD_ATTEMPTS tries with a
 *   pause between, except for definite refusals (401/403/413/415), then the next server.
 *
 * Throws PayloadTooLargeError when a server answers 413, so the caller can cut smaller chunks.
 */
export async function uploadFrames(frames: Uint8Array[], options: UploadOptions): Promise<UploadedBlob[]> {
  const staged: StagedBlob[] = [];
  for (const bytes of frames) staged.push(await encryptFrame(bytes));
  return uploadStaged(staged, options);
}

export interface UploadOptions extends BlossomEnv {
  servers: readonly string[];
  signer: FileSigner;
  replicas?: number;
  /** Called as bytes go out (when the transport reports them) and after each blob is stored. */
  onProgress?: (event: { blobsDone: number; bytesDone: number; bytesTotal: number; server: string }) => void;
}

/** An encrypted frame waiting to be uploaded. The ciphertext sits in a Blob the browser can page out. */
export interface StagedBlob {
  ref: BlobRef;
  blob: Blob;
}

/** Encrypts one frame and stages the ciphertext. The plaintext bytes can be dropped afterwards. */
export async function encryptFrame(bytes: Uint8Array): Promise<StagedBlob> {
  const encrypted = await encryptFile(bytes);
  return {
    ref: {
      x: encrypted.blobHash,
      k: encrypted.encryptionKey,
      n: encrypted.bytes.byteLength,
      p: encrypted.size,
      c: encrypted.chunkSize,
      h: encrypted.unencryptedFileHash,
    },
    blob: new Blob([encrypted.bytes as BlobPart]),
  };
}

/**
 * BUD-06 preflight of a size without uploading anything or signing: learns the limit of every
 * server that answers 413, so chunks can be cut to fit before any work is done.
 */
export async function preflightChunkSize(servers: readonly string[], size: number, env: BlossomEnv = {}): Promise<void> {
  const transport = createFetchBlossomTransport(env.fetch);
  for (const server of sanitizeServers(servers)) {
    try {
      const verdict = await transport.canAccept?.({ server, size, sha256: "0".repeat(64), type: "application/octet-stream" });
      if (verdict && !verdict.ok && verdict.status === 413) learnServerLimit(server, size);
    } catch {
      // No answer is no information; the real upload decides.
    }
  }
}

/** Uploads staged blobs, in order, under as few signed authorizations as the servers allow (usually one). */
export async function uploadStaged(items: StagedBlob[], options: UploadOptions): Promise<UploadedBlob[]> {
  const servers = sanitizeServers(options.servers);
  if (servers.length === 0) throw new BlobUnavailableError("No Blossom server is configured.");
  if (items.length === 0) return [];

  const copies = Math.max(1, options.replicas ?? MAX_REPLICAS);
  const now = options.now ?? (() => Math.floor(Date.now() / 1000));
  const sleep = options.sleep ?? ((ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)));
  const transport = createFetchBlossomTransport(options.fetch);

  const refs = items.map((item) => item.ref);
  const staged = items.map((item) => item.blob);

  // One authorization for everything if the servers take that many hashes in a header.
  let batchSize = refs.length;
  if (refs.length > SAFE_AUTH_HASHES) {
    const doFetch = options.fetch ?? fetch;
    const limits = await Promise.all(servers.map((server) => authHashLimit(server, doFetch)));
    batchSize = Math.max(1, Math.min(refs.length, ...limits));
  }
  const authorizations = new Map<number, Promise<string>>();
  const authorizationFor = (index: number): Promise<string> => {
    const batch = Math.floor(index / batchSize);
    let pending = authorizations.get(batch);
    if (!pending) {
      const members = refs.slice(batch * batchSize, (batch + 1) * batchSize);
      const bytes = members.reduce((sum, ref) => sum + ref.n, 0) * Math.min(copies, servers.length);
      pending = createBlossomAuthorization(
        options.signer,
        "upload",
        members.map((ref) => ref.x),
        "Upload spreadsheet data",
        authTtlSeconds(bytes),
        now,
      );
      authorizations.set(batch, pending);
    }
    return pending;
  };

  // BUD-06 preflight with the largest blob: a 413 here costs one HEAD instead of a failed upload.
  const largestIndex = refs.reduce((best, ref, i) => (ref.n > refs[best].n ? i : best), 0);
  const largest = refs[largestIndex];
  for (const server of servers) {
    // No authorization on the HEAD: it must not cost a signer prompt before the real upload (BUD-06
    // allows it, and a server that wants one answers 401, which we treat as "no information").
    const verdict = await transport.canAccept?.({
      server,
      size: largest.n,
      sha256: largest.x,
      type: "application/octet-stream",
    });
    if (verdict && !verdict.ok && verdict.status === 413) throw new PayloadTooLargeError(server, largest.n);
  }

  const bytesTotal = refs.reduce((sum, ref) => sum + ref.n, 0) * Math.min(copies, servers.length);
  let bytesDone = 0;
  const uploaded: UploadedBlob[] = [];
  for (let index = 0; index < items.length; index++) {
    const ref = refs[index];
    const authorization = await authorizationFor(index);
    const bytes = new Uint8Array(await staged[index].arrayBuffer());
    const accepted: string[] = [];
    const failures: string[] = [];
    for (const server of servers) {
      if (accepted.length >= copies) break;
      let lastError: unknown;
      let stored = false;
      for (let attempt = 1; attempt <= UPLOAD_ATTEMPTS && !stored; attempt++) {
        try {
          await transport.upload({
            server,
            bytes,
            authorization,
            onBytes: (sent: number) =>
              options.onProgress?.({ blobsDone: index, bytesDone: bytesDone + sent, bytesTotal, server }),
          });
          stored = true;
        } catch (error) {
          const status = (error as { status?: unknown }).status;
          if (status === 413) throw new PayloadTooLargeError(server, ref.n);
          lastError = error;
          if (typeof status === "number" && PERMANENT_STATUSES.has(status)) break;
          if (attempt < UPLOAD_ATTEMPTS) await sleep(RETRY_DELAY_MS);
        }
      }
      if (stored) {
        accepted.push(server);
        bytesDone += ref.n;
        options.onProgress?.({ blobsDone: index + (accepted.length >= copies ? 1 : 0), bytesDone, bytesTotal, server });
      } else {
        failures.push(`${server}: ${lastError instanceof Error ? lastError.message : String(lastError)}`);
      }
    }
    if (accepted.length === 0) {
      throw new BlobUnavailableError(`Couldn't upload to any Blossom server. ${failures.join("; ")}`);
    }
    uploaded.push({ servers: accepted, ref });
  }
  return uploaded;
}

/** Downloads, verifies and decrypts one blob, trying each server until one has it. */
export async function downloadTabBody(
  ref: BlobRef,
  servers: readonly string[],
  env: BlossomEnv = {},
  type?: number,
): Promise<unknown> {
  const candidates = sanitizeServers(servers);
  if (candidates.length === 0) throw new BlobUnavailableError("No Blossom server is available to load this sheet from.");
  const transport = createFetchBlossomTransport(env.fetch ?? fetch);
  const failures: string[] = [];
  for (const server of candidates) {
    try {
      const encrypted = await transport.download({ server, hash: ref.x, expectedSize: ref.n });
      const plaintext = await decryptFileBytes(encrypted, {
        size: ref.p,
        chunkSize: ref.c,
        blobHash: ref.x,
        encryptionKey: ref.k,
        unencryptedFileHash: ref.h,
        servers: candidates,
        type: "application/octet-stream",
      });
      return await unframe(plaintext, type);
    } catch (error) {
      failures.push(`${server}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  throw new BlobUnavailableError(`Couldn't load this sheet from any Blossom server. ${failures.join("; ")}`);
}
