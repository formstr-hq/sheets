// Spike worker: opens the zip, reads the workbook parts with ExcelJS's parsers, answers with a summary.
import { loadSharedStrings } from "./strings";
import { readWorkbookMeta } from "./workbookMeta";
import { openZip } from "./zip";
import type { XlsxProbe } from "./client";

self.onmessage = async (event: MessageEvent<{ file: Blob }>) => {
  try {
    const zip = await openZip(event.data.file, {});
    const meta = await readWorkbookMeta(zip);
    const strings = await loadSharedStrings(zip, { budgetBytes: 64 * 1024 * 1024 });
    const probe: XlsxProbe = {
      sheets: meta.sheets.map((sheet) => ({ name: sheet.name, hidden: sheet.hidden })),
      date1904: meta.date1904,
      sharedStrings: strings?.length ?? 0,
    };
    (self as unknown as Worker).postMessage({ ok: true, probe });
  } catch (error) {
    (self as unknown as Worker).postMessage({ ok: false, error: error instanceof Error ? error.message : String(error) });
  }
};
