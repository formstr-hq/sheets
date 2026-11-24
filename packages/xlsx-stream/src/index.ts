// Spike A: the public API is only the worker client plus the error type. The reading code lives in
// ./worker, which ships as its own bundled file (`@formstr/xlsx-stream/worker`).
export { ImportError, type ImportErrorCode } from "./errors";
export { probeXlsx, type XlsxProbe } from "./client";
