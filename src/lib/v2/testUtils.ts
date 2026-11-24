// Test helpers: an in-memory Blossom network and a local signer. Not used by the app.

import { finalizeEvent, generateSecretKey, getPublicKey, type Event, type EventTemplate } from "nostr-tools";
import type { FileSigner } from "@formstr/drive-sdk";

export function localSigner(secret: Uint8Array = generateSecretKey()): FileSigner & { pubkey: string } {
  return {
    pubkey: getPublicKey(secret),
    getPublicKey: async () => getPublicKey(secret),
    signEvent: async (template: EventTemplate): Promise<Event> => finalizeEvent(template, secret),
  };
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as BufferSource);
  return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, "0")).join("");
}

export interface FakeBlossom {
  fetch: typeof fetch;
  /** Blobs stored per server origin. */
  stores: Map<string, Map<string, Uint8Array>>;
  /** Every accepted PUT, in order. */
  uploads: { origin: string; hash: string; authorization: string | null }[];
  /** Origins that currently answer every request with a 500. */
  down: Set<string>;
  /** Per-origin upload size limit in bytes; larger PUTs and BUD-06 HEADs get a 413. */
  limits: Map<string, number>;
  /** Per-origin max Authorization header length in bytes; longer requests get a 431. */
  maxAuthHeader: Map<string, number>;
  /** Per-origin count of upcoming PUTs that answer 503 before the server recovers. */
  flaky: Map<string, number>;
  /** PUT attempts seen, including failed ones. */
  attempts: { origin: string }[];
  /** BUD-06 HEAD /upload preflights seen. */
  preflights: { origin: string; size: number }[];
}

/** A fake Blossom network. Servers are created on first use; each verifies X-SHA-256 on upload. */
export function fakeBlossom(): FakeBlossom {
  const stores = new Map<string, Map<string, Uint8Array>>();
  const uploads: FakeBlossom["uploads"] = [];
  const down = new Set<string>();
  const limits = new Map<string, number>();
  const flaky = new Map<string, number>();
  const maxAuthHeader = new Map<string, number>();
  const attempts: FakeBlossom["attempts"] = [];
  const preflights: FakeBlossom["preflights"] = [];

  const handler = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = new URL(String(input));
    const method = (init?.method ?? "GET").toUpperCase();
    if (down.has(url.origin)) return new Response("down", { status: 500 });
    if (!stores.has(url.origin)) stores.set(url.origin, new Map());
    const store = stores.get(url.origin)!;
    const headers = new Headers(init?.headers);
    const headerLimit = maxAuthHeader.get(url.origin);
    if (headerLimit !== undefined && (headers.get("authorization") ?? "").length > headerLimit) {
      return new Response("header too large", { status: 431 });
    }

    if (url.pathname === "/upload" && method === "HEAD") {
      const size = Number(headers.get("x-content-length"));
      preflights.push({ origin: url.origin, size });
      const limit = limits.get(url.origin);
      return new Response(null, limit !== undefined && size > limit ? { status: 413 } : { status: 200 });
    }
    if (url.pathname === "/upload" && method === "PUT") {
      attempts.push({ origin: url.origin });
      const failing = flaky.get(url.origin) ?? 0;
      if (failing > 0) {
        flaky.set(url.origin, failing - 1);
        return new Response("busy", { status: 503 });
      }
      const bytes = new Uint8Array(init?.body as ArrayBuffer);
      const hash = await sha256Hex(bytes);
      const limit = limits.get(url.origin);
      if (limit !== undefined && bytes.byteLength > limit) return new Response("too large", { status: 413 });
      if (headers.get("x-sha-256") !== hash) return new Response("hash mismatch", { status: 400 });
      store.set(hash, bytes);
      uploads.push({ origin: url.origin, hash, authorization: headers.get("authorization") });
      return new Response("{}", { status: 200 });
    }
    const hash = url.pathname.slice(1);
    if (method === "GET" || method === "HEAD") {
      const bytes = store.get(hash);
      if (!bytes) return new Response("not found", { status: 404 });
      return new Response(method === "HEAD" ? null : (bytes.slice() as BufferSource), { status: 200 });
    }
    return new Response("unsupported", { status: 405 });
  };

  return { fetch: handler as typeof fetch, stores, uploads, down, limits, preflights, flaky, attempts, maxAuthHeader };
}
