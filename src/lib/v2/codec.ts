// FSB2 framing for v2 sheet payloads.
//
//   magic "FSB2" (4) | type (1) | codec (1) | payload
//
// The payload is deflate-raw compressed JSON. Compression happens before
// encryption, because ciphertext does not compress.

export const FSB2_MAGIC = [0x46, 0x53, 0x42, 0x32] as const; // "FSB2"
export const FSB2_HEADER_BYTES = 6;

export const FSB2_TYPE_CHUNK = 0x01;

export const FSB2_CODEC_RAW = 0x00;
export const FSB2_CODEC_DEFLATE_RAW = 0x01;

async function pipe(bytes: Uint8Array, stream: CompressionStream | DecompressionStream): Promise<Uint8Array> {
  const writer = stream.writable.getWriter();
  // Do not await the write: the readable side must be drained at the same time.
  void writer.write(bytes as BufferSource).catch(() => {});
  void writer.close().catch(() => {});
  return new Uint8Array(await new Response(stream.readable).arrayBuffer());
}

export function deflate(bytes: Uint8Array): Promise<Uint8Array> {
  return pipe(bytes, new CompressionStream("deflate-raw"));
}

export function inflate(bytes: Uint8Array): Promise<Uint8Array> {
  return pipe(bytes, new DecompressionStream("deflate-raw"));
}

/**
 * JSON with object keys sorted, so equal content always produces equal bytes
 * regardless of the order properties were inserted in. The frame's hash is how we
 * tell an unchanged tab from a changed one, so it must follow content, not key order.
 */
export function canonicalJson(value: unknown): string {
  return JSON.stringify(value, (_key, item: unknown) => {
    if (item && typeof item === "object" && !Array.isArray(item)) {
      const source = item as Record<string, unknown>;
      return Object.fromEntries(Object.keys(source).sort().map((key) => [key, source[key]]));
    }
    return item;
  });
}

export async function frame(json: unknown, type = FSB2_TYPE_CHUNK): Promise<Uint8Array> {
  const body = await deflate(new TextEncoder().encode(canonicalJson(json)));
  const out = new Uint8Array(FSB2_HEADER_BYTES + body.length);
  out.set(FSB2_MAGIC, 0);
  out[4] = type;
  out[5] = FSB2_CODEC_DEFLATE_RAW;
  out.set(body, FSB2_HEADER_BYTES);
  return out;
}

export async function unframe(bytes: Uint8Array, expectedType = FSB2_TYPE_CHUNK): Promise<unknown> {
  if (bytes.length < FSB2_HEADER_BYTES || FSB2_MAGIC.some((byte, i) => bytes[i] !== byte)) {
    throw new Error("Not an FSB2 payload.");
  }
  if (bytes[4] !== expectedType) throw new Error("Unexpected FSB2 payload type.");
  const body = bytes.subarray(FSB2_HEADER_BYTES);
  const codec = bytes[5];
  let raw: Uint8Array;
  if (codec === FSB2_CODEC_RAW) raw = body;
  else if (codec === FSB2_CODEC_DEFLATE_RAW) raw = await inflate(body);
  else throw new Error(`Unsupported FSB2 codec ${codec}.`);
  return JSON.parse(new TextDecoder().decode(raw));
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const step = 0x8000;
  for (let i = 0; i < bytes.length; i += step) {
    binary += String.fromCharCode(...bytes.subarray(i, i + step));
  }
  return btoa(binary);
}

export function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}
