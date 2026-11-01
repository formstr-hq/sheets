import {
  DataLayer,
  getDataLayer,
  LocalRelayClient,
  setDataLayer,
  workerChannel,
  type Channel,
} from "@formstr/local-relay";
import type { Event, EventTemplate } from "nostr-tools";
import { DEFAULT_RELAYS } from "./constants";
import { signerInstance } from "./signerInstance";

let started = false;
let localRelayClient: LocalRelayClient | null = null;

async function signWithActiveSigner(template: EventTemplate): Promise<Event> {
  const signer = signerInstance.getActiveSigner();
  if (!signer) throw new Error("No signer available.");
  return (await signer.signEvent(template)) as Event;
}

export function bootstrapLocalRelay(): DataLayer {
  if (started) return getDataLayer();
  started = true;

  const worker = new Worker(new URL("../worker/relay.worker.ts", import.meta.url), {
    type: "module",
  });
  const baseChannel: Channel = workerChannel(worker);
  const channel: Channel = {
    post: (message) => baseChannel.post(message),
    close: () => baseChannel.close(),
    onMessage: (handler) => baseChannel.onMessage(handler),
  };

  const client = new LocalRelayClient(channel, {
    onSignRequest: async (template) => {
      try {
        return (await signWithActiveSigner(template)) as never;
      } catch {
        return null;
      }
    },
  });
  localRelayClient = client;
  client.setUserRelays(DEFAULT_RELAYS);
  client.setActiveAccount(signerInstance.getActiveAccount()?.pubkey ?? null);
  signerInstance.onChange(() => {
    client.setActiveAccount(signerInstance.getActiveAccount()?.pubkey ?? null);
  });

  if (typeof document !== "undefined" && document.addEventListener) {
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "hidden") client.pause();
      else client.resume();
    });
  }

  const dataLayer = new DataLayer({
    client,
    sign: signWithActiveSigner,
  });
  setDataLayer(dataLayer);
  return dataLayer;
}

export function getLocalRelayClient(): LocalRelayClient | null {
  return localRelayClient;
}

export function setLocalRelayRelays(relays: string[]): void {
  localRelayClient?.setUserRelays(relays);
  if (started) {
    try {
      getDataLayer().setUserRelays(relays);
    } catch {
      // not bootstrapped yet
    }
  }
}
