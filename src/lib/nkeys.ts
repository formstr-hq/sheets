import { bech32 } from "bech32";

const utf8Encoder = new TextEncoder();
const utf8Decoder = new TextDecoder();

function concatBytes(chunks: Uint8Array[]): Uint8Array {
  const length = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

function encodeTlv(fields: Record<number, Uint8Array[]>): Uint8Array {
  const parts: Uint8Array[] = [];
  for (const [typeString, values] of Object.entries(fields)) {
    const type = Number(typeString);
    for (const value of values) {
      parts.push(new Uint8Array([type, value.length]));
      parts.push(value);
    }
  }
  return concatBytes(parts);
}

function decodeTlv(data: Uint8Array): Record<number, Uint8Array[]> {
  const result: Record<number, Uint8Array[]> = {};
  let index = 0;
  while (index < data.length) {
    const type = data[index];
    const length = data[index + 1];
    const value = data.slice(index + 2, index + 2 + length);
    if (!result[type]) result[type] = [];
    result[type].push(value);
    index += 2 + length;
  }
  return result;
}

export function encodeNKeys(values: Record<string, string>): string {
  const fields: Record<number, Uint8Array[]> = { 0: [], 1: [] };
  for (const [key, value] of Object.entries(values)) {
    fields[0].push(utf8Encoder.encode(key));
    fields[1].push(utf8Encoder.encode(value));
  }
  const data = encodeTlv(fields);
  return bech32.encode("nkeys", bech32.toWords(data), 2048);
}

export function decodeNKeys(value: string): Record<string, string> {
  const { words } = bech32.decode(value, 2048);
  const data = new Uint8Array(bech32.fromWords(words));
  const tlv = decodeTlv(data);
  const result: Record<string, string> = {};
  const keys = tlv[0] || [];
  const values = tlv[1] || [];
  keys.forEach((keyBytes, index) => {
    const key = utf8Decoder.decode(keyBytes);
    const valueBytes = values[index];
    result[key] = valueBytes ? utf8Decoder.decode(valueBytes) : "";
  });
  return result;
}
