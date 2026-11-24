// Errors the streaming import reports. `code` is what the worker protocol carries.
import { ZipError } from "./zip";

export type ImportErrorCode =
  | "UNSUPPORTED_ZIP"
  | "NOT_XLSX"
  | "STRING_TABLE_TOO_LARGE"
  /** Raised by whoever uploads a chunk when a single row is over the server limit (not by the engine). */
  | "ROW_TOO_LARGE"
  /** The file is a zip/xml that is damaged or not laid out as the xlsx format requires. */
  | "CORRUPT"
  | "CANCELLED"
  | "UNKNOWN";

export class ImportError extends Error {
  readonly code: ImportErrorCode;
  /** Technical text (parser message) for logs; `message` is for the user. */
  readonly detail?: string;
  constructor(code: ImportErrorCode, message: string, detail?: string) {
    super(message);
    this.name = "ImportError";
    this.code = code;
    this.detail = detail;
  }
}

/** Maps any thrown value to an ImportError (zip problems become NOT_XLSX / UNSUPPORTED_ZIP). */
export function toImportError(error: unknown): ImportError {
  if (error instanceof ImportError) return error;
  if (error instanceof ZipError) {
    if (error.code === "CANCELLED") return new ImportError("CANCELLED", error.message);
    if (error.code === "UNSUPPORTED_ZIP") return new ImportError("UNSUPPORTED_ZIP", error.message);
    if (error.code === "CORRUPT_ZIP") {
      return new ImportError("CORRUPT", "This file looks damaged and could not be read. Try opening and re-saving it in a spreadsheet app.", error.message);
    }
    return new ImportError("NOT_XLSX", "This does not look like a valid .xlsx file.", error.message);
  }
  const message = error instanceof Error ? error.message : String(error);
  // saxes reports "line:column: problem" for malformed XML.
  if (/^\d+:\d+: /.test(message)) {
    return new ImportError("CORRUPT", "This file looks damaged: part of it is not valid XML. Try opening and re-saving it in a spreadsheet app.", message);
  }
  return new ImportError("UNKNOWN", message);
}
