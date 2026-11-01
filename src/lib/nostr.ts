import { dataLayer, type ObserveHandle } from "@formstr/local-relay";
import type { Event, Filter } from "nostr-tools";
import { KIND_SHEET } from "./constants";
import { mergeRelays } from "./utils";

export type { Event, Filter } from "nostr-tools";

export function coordinateFor(editPubkey: string, sheetId: string): string {
  return `${KIND_SHEET}:${editPubkey}:${sheetId}`;
}

export function parseCoordinate(coordinate: string): {
  kind: number;
  pubkey: string;
  id: string;
} | null {
  const [kindText, pubkey, ...rest] = coordinate.split(":");
  const kind = Number(kindText);
  const id = rest.join(":");
  if (!Number.isFinite(kind) || !pubkey || !id) return null;
  return { kind, pubkey, id };
}

function collectEvents(
  relays: string[],
  filters: Filter[],
  timeoutMs: number,
): Promise<Event[]> {
  return new Promise((resolve) => {
    const found = new Map<string, Event>();
    let settled = false;
    let handle: ObserveHandle | null = null;
    let quietTimer: ReturnType<typeof setTimeout> | null = null;
    let repollTimer: ReturnType<typeof setTimeout> | null = null;
    let deadline: ReturnType<typeof setTimeout> | null = null;

    const clearTimers = () => {
      if (quietTimer) clearTimeout(quietTimer);
      if (repollTimer) clearTimeout(repollTimer);
      if (deadline) clearTimeout(deadline);
      quietTimer = null;
      repollTimer = null;
      deadline = null;
    };

    const settle = () => {
      if (settled) return;
      settled = true;
      clearTimers();
      handle?.unobserve();
      resolve(Array.from(found.values()));
    };

    const bumpQuiet = () => {
      if (quietTimer) clearTimeout(quietTimer);
      quietTimer = setTimeout(settle, 900);
    };

    const scheduleRepoll = () => {
      if (settled || repollTimer || found.size > 0) return;
      repollTimer = setTimeout(() => {
        repollTimer = null;
        if (settled || found.size > 0) return;
        handle?.update(filters);
        scheduleRepoll();
      }, 500);
    };

    handle = dataLayer.observe(
      filters,
      {
        onEvent(event) {
          found.set(event.id, event as Event);
          if (repollTimer) {
            clearTimeout(repollTimer);
            repollTimer = null;
          }
          bumpQuiet();
        },
        onEose() {
          if (found.size > 0) bumpQuiet();
          else scheduleRepoll();
        },
      },
      { relays },
    );

    deadline = setTimeout(settle, timeoutMs);
  });
}

export function fetchEvents(
  relays: string[],
  filter: Filter,
  timeoutMs = 7000,
): Promise<Event[]> {
  return collectEvents(relays, [filter], timeoutMs);
}

export async function fetchLatestEvent(
  relays: string[],
  filter: Filter,
  timeoutMs = 7000,
): Promise<Event | null> {
  const events = await fetchEvents(relays, filter, timeoutMs);
  events.sort((a, b) => b.created_at - a.created_at);
  return events[0] ?? null;
}

export async function publishEvent(
  event: Event,
  relays: string[],
  _timeoutMs = 10_000,
): Promise<{ accepted: number; relays: string[] }> {
  const result = await dataLayer.publishEvent(event, {
    relays: mergeRelays(relays),
  });
  const accepted = result.relayResults
    .filter((relay) => relay.status === "accepted")
    .map((relay) => relay.relay);

  if (accepted.length === 0) {
    throw new Error("Couldn't reach any relay. Your work is saved on this device.");
  }

  return { accepted: accepted.length, relays: Array.from(new Set(accepted)) };
}

export function subscribeSheet(
  relays: string[],
  editPubkey: string,
  sheetId: string,
  onEvent: (event: Event) => void,
): () => void {
  const handle = dataLayer.observe(
    [{ kinds: [KIND_SHEET], authors: [editPubkey], "#d": [sheetId] }],
    { onEvent: (event) => onEvent(event as Event) },
    { relays },
  );
  return () => handle.unobserve();
}

export async function fetchUserRelays(
  relays: string[],
  pubkey: string,
): Promise<string[]> {
  const event = await fetchLatestEvent(
    relays,
    { kinds: [10002], authors: [pubkey], limit: 1 },
    4500,
  );
  if (!event) return [];
  return event.tags
    .filter((tag) => tag[0] === "r" && tag[1]?.startsWith("wss://"))
    .map((tag) => tag[1]);
}
