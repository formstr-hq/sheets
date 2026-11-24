import { beforeEach, describe, expect, it } from "vitest";
import { verifyEvent, type Event } from "nostr-tools";
import { MAX_AUTH_HASHES, MAX_REPLICAS, SAFE_AUTH_HASHES, forgetAuthLimits, UPLOAD_ATTEMPTS, authTtlSeconds, downloadTabBody, mergeServers, prepareTab, sanitizeServers, uploadFrames } from "./blossom";
import { BlobUnavailableError } from "./root";
import { fakeBlossom, localSigner } from "./testUtils";

const instant = async () => {};
beforeEach(() => forgetAuthLimits());
const A = "https://a.example";
const B = "https://b.example";
const C = "https://c.example";

function decodeAuth(header: string): Event {
  expect(header.startsWith("Nostr ")).toBe(true);
  return JSON.parse(atob(header.slice(6))) as Event;
}

describe("sanitizeServers", () => {
  it("keeps https, allows http only for localhost, drops junk, and dedupes", () => {
    expect(
      sanitizeServers([
        "https://a.example/",
        "https://a.example",
        "http://b.example",
        "http://localhost:3000",
        "javascript:alert(1)",
        "not a url",
        "ftp://c.example",
      ]),
    ).toEqual(["https://a.example", "http://localhost:3000"]);
  });
});

describe("mergeServers", () => {
  it("adds new servers and keeps the ones already on the profile", () => {
    expect(mergeServers([A], [B], [])).toEqual([A, B]);
  });

  it("removes only what the user removed", () => {
    expect(mergeServers([A, B, C], [], [B])).toEqual([A, C]);
  });

  it("keeps a server added elsewhere that this editor never saw", () => {
    // The dialog showed [A] and the user removed A, but the profile now also has C.
    expect(mergeServers([A, C], [B], [A])).toEqual([C, B]);
  });

  it("does not wipe the list when the read came back empty", () => {
    expect(mergeServers([], [B], [])).toEqual([B]);
  });

  it("ignores duplicates, trailing slashes and non-https addresses", () => {
    expect(mergeServers([`${A}/`], [A, "http://evil.example", B], [])).toEqual([A, B]);
  });
});

describe("Blossom upload and download", () => {
  it("round-trips a tab body, and the server only ever holds ciphertext", async () => {
    const net = fakeBlossom();
    const body = { cells: { "0:0": { value: "super secret salary" } }, rowCount: 1, colCount: 1 };
    const { frameBytes } = await prepareTab(body);
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch });

    const stored = net.stores.get(A)!.get(uploaded.ref.x)!;
    expect(new TextDecoder("latin1").decode(stored)).not.toContain("secret");
    expect(uploaded.ref.n).toBe(stored.byteLength);
    expect(await downloadTabBody(uploaded.ref, [A], { fetch: net.fetch })).toEqual(body);
  });

  it("signs one authorization that covers every blob in the save", async () => {
    const net = fakeBlossom();
    const signer = localSigner();
    const frames = [
      (await prepareTab({ a: 1 })).frameBytes,
      (await prepareTab({ b: 2 })).frameBytes,
      (await prepareTab({ c: 3 })).frameBytes,
    ];
    const uploaded = await uploadFrames(frames, { servers: [A], signer, fetch: net.fetch });

    const headers = new Set(net.uploads.map((upload) => upload.authorization));
    expect(headers.size).toBe(1);
    const auth = decodeAuth([...headers][0]!);
    expect(verifyEvent(auth)).toBe(true);
    expect(auth.kind).toBe(24242);
    expect(auth.pubkey).toBe(signer.pubkey);
    expect(auth.tags).toContainEqual(["t", "upload"]);
    expect(auth.tags.filter((tag) => tag[0] === "x").map((tag) => tag[1]).sort()).toEqual(
      uploaded.map((blob) => blob.ref.x).sort(),
    );
    const expiry = Number(auth.tags.find((tag) => tag[0] === "expiration")![1]);
    expect(expiry).toBeGreaterThan(Math.floor(Date.now() / 1000));
  });

  it(`copies each blob to at most ${MAX_REPLICAS} servers and skips one that is down`, async () => {
    const net = fakeBlossom();
    net.down.add(A);
    const { frameBytes } = await prepareTab({ x: 1 });
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A, B, C], signer: localSigner(), fetch: net.fetch, sleep: instant });

    // A is down, so the copies go to the next servers in list order, up to MAX_REPLICAS.
    const expected = [B, C].slice(0, MAX_REPLICAS);
    expect(uploaded.servers).toEqual(expected);
    for (const server of [B, C]) {
      expect(net.stores.get(server)?.has(uploaded.ref.x) ?? false).toBe(expected.includes(server));
    }
  });

  it("copies to exactly the number of servers asked for: 2x, 3x, one", async () => {
    const net = fakeBlossom();
    const { frameBytes } = await prepareTab({ x: 1 });
    const [uploaded] = await uploadFrames([frameBytes], {
      servers: [A, B, C],
      signer: localSigner(),
      fetch: net.fetch,
      replicas: 2,
    });
    expect(uploaded.servers).toEqual([A, B]);
    expect(net.stores.get(C)?.has(uploaded.ref.x) ?? false).toBe(false);

    const all = await uploadFrames([frameBytes], { servers: [A, B, C], signer: localSigner(), fetch: net.fetch, replicas: 3 });
    expect(all[0].servers).toEqual([A, B, C]);
    // Upload bytes scale with the number of servers ticked.
    const net3 = fakeBlossom();
    await uploadFrames([frameBytes], { servers: [A, B, C], signer: localSigner(), fetch: net3.fetch, replicas: 3 });
    expect(net3.uploads).toHaveLength(3);
    const net1 = fakeBlossom();
    await uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net1.fetch, replicas: 1 });
    expect(net1.uploads).toHaveLength(1);
  });

  it("fails clearly when no server accepts the upload", async () => {
    const net = fakeBlossom();
    net.down.add(A);
    const { frameBytes } = await prepareTab({ x: 1 });
    await expect(uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch, sleep: instant })).rejects.toThrow(
      BlobUnavailableError,
    );
  });

  it("refuses to upload with no server configured", async () => {
    const { frameBytes } = await prepareTab({ x: 1 });
    await expect(uploadFrames([frameBytes], { servers: [], signer: localSigner(), fetch: fakeBlossom().fetch })).rejects.toThrow(
      "No Blossom server",
    );
  });

  it("downloads from the next server when the first does not have the blob", async () => {
    const net = fakeBlossom();
    const { frameBytes } = await prepareTab({ hello: "world" });
    const [uploaded] = await uploadFrames([frameBytes], { servers: [B], signer: localSigner(), fetch: net.fetch });
    expect(await downloadTabBody(uploaded.ref, [A, B], { fetch: net.fetch })).toEqual({ hello: "world" });
  });

  it("rejects a blob whose bytes were tampered with", async () => {
    const net = fakeBlossom();
    const { frameBytes } = await prepareTab({ hello: "world" });
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch });
    const stored = net.stores.get(A)!.get(uploaded.ref.x)!;
    stored[10] ^= 0xff;
    await expect(downloadTabBody(uploaded.ref, [A], { fetch: net.fetch })).rejects.toThrow(BlobUnavailableError);
  });

  it("does not contact servers that are not https", async () => {
    const net = fakeBlossom();
    const { frameBytes } = await prepareTab({ x: 1 });
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch });
    await expect(downloadTabBody(uploaded.ref, ["http://evil.example"], { fetch: net.fetch })).rejects.toThrow(
      "No Blossom server",
    );
  });
});

describe("upload hardening", () => {
  it("retries a flaky server before giving up on it, and waits between tries", async () => {
    const net = fakeBlossom();
    net.flaky.set(A, 2); // two 503s, then it recovers
    const waits: number[] = [];
    const { frameBytes } = await prepareTab({ x: 1 });
    const [uploaded] = await uploadFrames([frameBytes], {
      servers: [A, B],
      signer: localSigner(),
      fetch: net.fetch,
      sleep: async (ms) => void waits.push(ms),
    });
    expect(uploaded.servers).toEqual([A]);
    expect(net.attempts.filter((a) => a.origin === A)).toHaveLength(3);
    expect(waits).toEqual([3000, 3000]);
    expect(net.stores.get(B)?.size ?? 0).toBe(0);
  });

  it(`falls back to the next server after ${UPLOAD_ATTEMPTS} failed tries`, async () => {
    const net = fakeBlossom();
    net.flaky.set(A, 99);
    const { frameBytes } = await prepareTab({ x: 1 });
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A, B], signer: localSigner(), fetch: net.fetch, sleep: instant });
    expect(net.attempts.filter((a) => a.origin === A)).toHaveLength(UPLOAD_ATTEMPTS);
    expect(uploaded.servers).toEqual([B]);
  });

  it("makes authorizations last at least 30 minutes, and longer for big uploads", async () => {
    expect(authTtlSeconds(1024)).toBe(1800);
    expect(authTtlSeconds(2000 * 1024 * 1024)).toBe(4000);
    const net = fakeBlossom();
    const { frameBytes } = await prepareTab({ x: 1 });
    await uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch });
    const auth = decodeAuth(net.uploads[0].authorization!);
    const expiry = Number(auth.tags.find((tag) => tag[0] === "expiration")![1]);
    expect(expiry - auth.created_at).toBeGreaterThanOrEqual(1800);
  });

  it("reports progress as blobs finish", async () => {
    const net = fakeBlossom();
    const frames = await Promise.all([1, 2, 3].map(async (i) => (await prepareTab({ i })).frameBytes));
    const events: { blobsDone: number; bytesDone: number; bytesTotal: number }[] = [];
    await uploadFrames(frames, { servers: [A], signer: localSigner(), fetch: net.fetch, onProgress: (e) => events.push(e) });
    const last = events[events.length - 1];
    expect(last.blobsDone).toBe(3);
    expect(last.bytesDone).toBe(last.bytesTotal);
    for (let i = 1; i < events.length; i++) expect(events[i].bytesDone).toBeGreaterThanOrEqual(events[i - 1].bytesDone);
  });

  async function promptsFor(count: number, net = fakeBlossom()) {
    const signer = localSigner();
    let signed = 0;
    const counting = { ...signer, signEvent: async (t: Parameters<typeof signer.signEvent>[0]) => (signed++, signer.signEvent(t)) };
    const frames = await Promise.all(Array.from({ length: count }, async (_, i) => (await prepareTab({ i })).frameBytes));
    await uploadFrames(frames, { servers: [A], signer: counting, fetch: net.fetch });
    return { signed, net };
  }

  it("needs ONE signer prompt for the whole save when the server takes a big header, up to MAX_AUTH_HASHES chunks", { timeout: 60_000 }, async () => {
    for (const [count, prompts] of [[1, 1], [SAFE_AUTH_HASHES, 1], [SAFE_AUTH_HASHES + 1, 1], [MAX_AUTH_HASHES, 1], [MAX_AUTH_HASHES + 1, 2]] as const) {
      expect((await promptsFor(count)).signed, `${count} chunks`).toBe(prompts);
    }
  });

  it("probes a server's header limit without the signer and falls back to the biggest batch it accepts", async () => {
    // A server that takes ~16 KB headers: 100 hashes fit, 200 do not -> 120 chunks need 2 prompts.
    const roomy = fakeBlossom();
    roomy.maxAuthHeader.set(A, 16 * 1024);
    expect((await promptsFor(120, roomy)).signed).toBe(2);
    // A server with nginx's 8 KB default: only SAFE_AUTH_HASHES fit -> 120 chunks need 3 prompts.
    const strict = fakeBlossom();
    strict.maxAuthHeader.set(A, 8 * 1024);
    forgetAuthLimits();
    const { signed, net } = await promptsFor(120, strict);
    expect(signed).toBe(3);
    // Every real authorization was accepted, so the probe's size estimate is not an under-estimate.
    expect(net.uploads).toHaveLength(120);
    for (const upload of net.uploads) expect((upload.authorization ?? "").length).toBeLessThanOrEqual(8 * 1024);
  });

  it("uses the smallest limit among the servers, and probing costs no signatures", async () => {
    const net = fakeBlossom();
    net.maxAuthHeader.set(B, 8 * 1024);
    const signer = localSigner();
    let signedAtProbe = -1;
    let signed = 0;
    const counting = { ...signer, signEvent: async (t: Parameters<typeof signer.signEvent>[0]) => (signed++, signer.signEvent(t)) };
    const frames = await Promise.all(Array.from({ length: 80 }, async (_, i) => (await prepareTab({ i })).frameBytes));
    const spy: typeof fetch = async (input, init) => {
      if ((init?.method ?? "GET") === "HEAD" && String(input).includes("/upload")) signedAtProbe = signed;
      return net.fetch(input, init);
    };
    await uploadFrames(frames, { servers: [A, B], signer: counting, replicas: 2, fetch: spy });
    expect(signedAtProbe).toBe(0); // the last HEAD happened before any signing
    expect(signed).toBe(2); // B only takes 50 hashes -> 50 + 30
  });

  it("keeps the safe batch size when the probe itself is refused at every size", async () => {
    const net = fakeBlossom();
    const picky: typeof fetch = async (input, init) =>
      (init?.method ?? "GET") === "HEAD" && String(input).endsWith("/upload") ? new Response(null, { status: 400 }) : net.fetch(input, init);
    const signer = localSigner();
    let signed = 0;
    const counting = { ...signer, signEvent: async (t: Parameters<typeof signer.signEvent>[0]) => (signed++, signer.signEvent(t)) };
    const frames = await Promise.all(Array.from({ length: 120 }, async (_, i) => (await prepareTab({ i })).frameBytes));
    await uploadFrames(frames, { servers: [A], signer: counting, fetch: picky });
    expect(signed).toBe(3); // 50 + 50 + 20
  });

  it("does not ask the signer for the size preflight", async () => {
    const net = fakeBlossom();
    const signer = localSigner();
    let signedBeforeFirstPut = -1;
    let signed = 0;
    const counting = { ...signer, signEvent: async (t: Parameters<typeof signer.signEvent>[0]) => (signed++, signer.signEvent(t)) };
    const { frameBytes } = await prepareTab({ x: 1 });
    const spy: typeof fetch = async (input, init) => {
      if ((init?.method ?? "GET") === "PUT" && signedBeforeFirstPut < 0) signedBeforeFirstPut = signed;
      return net.fetch(input, init);
    };
    await uploadFrames([frameBytes], { servers: [A], signer: counting, fetch: spy });
    expect(net.preflights.length).toBe(1);
    expect(net.preflights.length === 1 && signedBeforeFirstPut).toBe(1); // signed once, only when the upload began
    expect(signed).toBe(1);
  });
});
