// Web Worker entry: forwards messages to the import logic in workerCore.ts.
import type { MainToWorker, WorkerToMain } from "./protocol";
import { createWorkerHandler } from "./workerCore";

interface WorkerScope {
  onmessage: ((event: MessageEvent<MainToWorker>) => void) | null;
  postMessage(message: WorkerToMain, transfer?: Transferable[]): void;
}

const scope = self as unknown as WorkerScope;
const handle = createWorkerHandler((message, transfer) => scope.postMessage(message, transfer));
scope.onmessage = (event) => handle(event.data);
