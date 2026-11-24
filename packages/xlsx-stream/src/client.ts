// Main-thread side: starts the bundled worker and asks it about a workbook.
export interface XlsxProbe {
  sheets: { name: string; hidden: boolean }[];
  date1904: boolean;
  sharedStrings: number;
}

export interface ProbeOptions {
  /**
   * Creates the worker. The default finds `worker.js` next to this file through `import.meta.url`,
   * which exists in the ESM build only. CommonJS consumers pass their own factory, for example
   * `() => new Worker(new URL("@formstr/xlsx-stream/worker", ...))` as their bundler requires.
   */
  createWorker?: () => Worker;
}

export function probeXlsx(file: Blob, options: ProbeOptions = {}): Promise<XlsxProbe> {
  const worker = options.createWorker
    ? options.createWorker()
    : new Worker(new URL("./worker.js", import.meta.url), { type: "module" });
  return new Promise<XlsxProbe>((resolve, reject) => {
    worker.onmessage = (event: MessageEvent<{ ok: true; probe: XlsxProbe } | { ok: false; error: string }>) => {
      worker.terminate();
      if (event.data.ok) resolve(event.data.probe);
      else reject(new Error(event.data.error));
    };
    worker.onerror = (event) => {
      worker.terminate();
      reject(new Error(event.message || "worker failed"));
    };
    worker.postMessage({ file });
  });
}
