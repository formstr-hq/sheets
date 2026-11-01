import {
  defaultPrunePolicy,
  IndexedDBStorage,
  RelayService,
  selfChannel,
} from "@formstr/local-relay";
import {
  KIND_SHEET,
  KIND_SHEET_GRANT,
  KIND_SHEET_INDEX,
} from "../lib/constants";

const channel = selfChannel(
  self as unknown as {
    postMessage: (message: unknown) => void;
    onmessage: ((event: MessageEvent) => void) | null;
  },
);

const prunePolicy = defaultPrunePolicy();
prunePolicy.protectedKinds.add(KIND_SHEET);
prunePolicy.protectedKinds.add(KIND_SHEET_GRANT);
prunePolicy.protectedKinds.add(KIND_SHEET_INDEX);

const service = new RelayService({
  channel,
  storage: new IndexedDBStorage("formstr-sheets"),
  persistence: { prunePolicy },
});

void service.start().then(() => {
  channel.post({ kind: "hydrated" });
});

export {};
