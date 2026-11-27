// Progress of a Blossom upload, for the upload card. A tiny store (no provider needed): the save
// path reports into it, the card reads it with useSyncExternalStore.

export interface UploadTabProgress {
  name: string;
  /** Chunks in this worksheet. */
  chunks: number;
  /** Chunks already stored (unchanged ones count as soon as the upload starts). */
  stored: number;
}

export interface UploadProgress {
  phase: "preparing" | "uploading" | "done" | "error";
  /** What is happening right now, in words. */
  message: string;
  tabs: UploadTabProgress[];
  /** Encrypted bytes sent so far / in total, counting every server copy. */
  bytesDone: number;
  bytesTotal: number;
  /** Blobs finished / to send in this save (unchanged chunks are not counted). */
  blobsDone: number;
  blobsTotal: number;
  servers: string[];
  error?: string;
}

type Listener = () => void;

let current: UploadProgress | null = null;
const listeners = new Set<Listener>();

function emit(): void {
  for (const listener of listeners) listener();
}

export const uploadProgress = {
  subscribe(listener: Listener): () => void {
    listeners.add(listener);
    return () => listeners.delete(listener);
  },
  get(): UploadProgress | null {
    return current;
  },
  report(next: UploadProgress): void {
    current = next;
    emit();
  },
  dismiss(): void {
    current = null;
    emit();
  },
};
