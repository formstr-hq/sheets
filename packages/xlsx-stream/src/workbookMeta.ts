// Reads the workbook-level parts of an xlsx: sheet list, relationships, date system, styles.
import { createStyleManager, parseRelationships, parseWorkbookXml, type StyleManager } from "./excelInternals";
import { ImportError } from "./errors";
import type { ZipArchive } from "./zip";

export interface SheetRef {
  name: string;
  hidden: boolean;
  /** Path of the worksheet part inside the zip, e.g. "xl/worksheets/sheet1.xml". */
  path: string;
  /** Position among the importable worksheets. */
  order: number;
}

export interface WorkbookMeta {
  sheets: SheetRef[];
  date1904: boolean;
  styles: StyleManager;
  /** Sheets that could not be imported, in words for the user. */
  notes: string[];
}

const WORKSHEET_TYPE = /\/worksheet$/;

/** Resolves a relationship target against the folder of the part that owns it ("xl/"). */
export function resolveTarget(target: string, baseDir = "xl/"): string {
  const trimmed = target.trim().replace(/\\/g, "/");
  const parts = (trimmed.startsWith("/") ? trimmed.slice(1) : baseDir + trimmed).split("/");
  const out: string[] = [];
  for (const part of parts) {
    if (part === "..") out.pop();
    else if (part !== "." && part !== "") out.push(part);
  }
  return out.join("/");
}

export async function readWorkbookMeta(zip: ZipArchive, signal?: AbortSignal): Promise<WorkbookMeta> {
  if (!zip.has("xl/workbook.xml")) {
    throw new ImportError("NOT_XLSX", "This does not look like an .xlsx file (xl/workbook.xml is missing).");
  }
  const workbook = await parseWorkbookXml(zip.text("xl/workbook.xml", { signal }));
  const rels = zip.has("xl/_rels/workbook.xml.rels")
    ? await parseRelationships(zip.text("xl/_rels/workbook.xml.rels", { signal }))
    : [];
  const byId = new Map(rels.map((rel) => [rel.Id, rel]));

  const lowerNames = new Map<string, string>();
  for (const key of zip.entries.keys()) lowerNames.set(key.toLowerCase(), key);

  const sheets: SheetRef[] = [];
  const notes: string[] = [];
  for (const sheet of workbook?.sheets ?? []) {
    const rel = sheet.rId ? byId.get(sheet.rId) : undefined;
    if (!rel || rel.TargetMode === "External") {
      notes.push(`Sheet "${sheet.name}" was skipped: the file has no data for it.`);
      continue;
    }
    if (!WORKSHEET_TYPE.test(rel.Type)) {
      notes.push(`Sheet "${sheet.name}" was skipped: it is not a normal worksheet (for example a chart sheet).`);
      continue;
    }
    const wanted = resolveTarget(rel.Target);
    const path = zip.has(wanted) ? wanted : lowerNames.get(wanted.toLowerCase());
    if (!path) {
      notes.push(`Sheet "${sheet.name}" was skipped: its data part (${wanted}) is missing from the file.`);
      continue;
    }
    sheets.push({
      name: sheet.name,
      hidden: sheet.state === "hidden" || sheet.state === "veryHidden",
      path,
      order: sheets.length,
    });
  }

  const styles = createStyleManager();
  if (zip.has("xl/styles.xml")) await styles.parseStream(zip.text("xl/styles.xml", { signal }));

  return { sheets, date1904: workbook?.properties?.date1904 === true, styles, notes };
}
