import { useEffect, useRef, useState, type CSSProperties } from "react";
import { useNavigate } from "react-router-dom";
import { ArrowRight, FileSpreadsheet, Loader2, Upload } from "lucide-react";
import { Dialog } from "./Dialog";
import { TEMPLATES } from "../lib/templates";
import { useLibrary } from "../providers/LibraryProvider";
import { useToast } from "../providers/ToastProvider";
import { sheetPath } from "../lib/routes";
import { cx } from "../lib/utils";
import { parseSpreadsheetFile } from "../lib/xlsx";
import { isStreamImportSupported } from "../lib/xlsxStream/client";
import { STREAM_IMPORT_MIN_BYTES } from "../lib/v2/streamImport";

/** Above this the in-memory importer may exhaust the browser tab's memory (5 MB was fine, 40 MB hung). */
const LARGE_IMPORT_BYTES = 20 * 1024 * 1024;

/** Big .xlsx files are read in a worker and uploaded chunk by chunk instead of loaded whole. */
function shouldStream(file: File): boolean {
  return /\.xlsx$/i.test(file.name) && file.size > STREAM_IMPORT_MIN_BYTES && isStreamImportSupported();
}

export function NewSheetDialog({
  open,
  onClose,
  initialTemplate = "blank",
}: {
  open: boolean;
  onClose: () => void;
  initialTemplate?: string;
}) {
  const library = useLibrary();
  const navigate = useNavigate();
  const { toast } = useToast();
  const [templateId, setTemplateId] = useState(initialTemplate);
  const [title, setTitle] = useState("");
  const [titleTouched, setTitleTouched] = useState(false);
  const [busy, setBusy] = useState<"create" | "import" | null>(null);
  const [importStage, setImportStage] = useState("");
  const [error, setError] = useState("");
  const fileRef = useRef<HTMLInputElement>(null);
  const abortRef = useRef<AbortController | null>(null);

  useEffect(() => {
    if (!open) return;
    setTemplateId(initialTemplate);
    setTitle("");
    setTitleTouched(false);
    setBusy(null);
    setError("");
  }, [initialTemplate, open]);

  const selected = TEMPLATES.find((template) => template.id === templateId) ?? TEMPLATES[0];
  const finalTitle = title.trim() || (titleTouched ? "Untitled spreadsheet" : selected.name);

  const create = async () => {
    setBusy("create");
    setError("");
    try {
      const entry = await library.createSheet(templateId, finalTitle);
      onClose();
      navigate(sheetPath(entry));
      toast("Spreadsheet created.", "success");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't create the spreadsheet.");
    } finally {
      setBusy(null);
    }
  };

  const importBig = async (file: File) => {
    setBusy("import");
    setError("");
    setImportStage(`Reading ${(file.size / 1024 / 1024).toFixed(0)} MB…`);
    const controller = new AbortController();
    abortRef.current = controller;
    try {
      const name = title.trim() || file.name.replace(/\.[^.]+$/, "") || finalTitle;
      const { entry, dropped } = await library.importStreamed(name, file, controller.signal);
      onClose();
      navigate(sheetPath(entry));
      toast(dropped.length ? `Imported ${name}. Not carried over: ${dropped.join("; ")}` : `Imported ${name}.`, "success");
    } catch (err) {
      setError(
        controller.signal.aborted
          ? "Import cancelled. Nothing was added to your library."
          : err instanceof Error
            ? err.message
            : "Couldn't import that file.",
      );
    } finally {
      abortRef.current = null;
      setBusy(null);
      setImportStage("");
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  const importFile = async (file: File) => {
    if (shouldStream(file)) return importBig(file);
    // Measured: a 40 MB xlsx needs ~3.7 GB of memory to import, and a 100 MB one needs far more than
    // a browser tab can have, so it freezes the page with nothing saved. Say so before starting.
    if (file.size > LARGE_IMPORT_BYTES) {
      const mb = (file.size / 1024 / 1024).toFixed(0);
      const go = window.confirm(
        `This file is ${mb} MB. Importing reads the whole workbook into browser memory, and files this size usually freeze or crash the tab before anything is saved.

Try anyway?`,
      );
      if (!go) {
        if (fileRef.current) fileRef.current.value = "";
        return;
      }
    }
    setBusy("import");
    setError("");
    // Big files keep the page busy for a while: say what is happening, and let the label paint first.
    const stage = async (text: string) => {
      setImportStage(text);
      await new Promise((resolve) => setTimeout(resolve, 30));
    };
    try {
      await stage(`Reading ${(file.size / 1024 / 1024).toFixed(1)} MB…`);
      const imported = await parseSpreadsheetFile(file, finalTitle);
      if (imported.sheets.length === 0) throw new Error("This file has no worksheets.");
      await stage("Encrypting and saving…");
      const entry = await library.importDocument(imported.title || finalTitle, imported.sheets);
      onClose();
      navigate(sheetPath(entry));
      toast(`Imported ${imported.title}.`, "success");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't import that file.");
    } finally {
      setBusy(null);
      setImportStage("");
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  return (
    <Dialog
      open={open}
      onClose={busy ? () => {} : onClose}
      title="New spreadsheet"
      subtitle="Start blank, borrow a template, or bring an Excel file."
      size="lg"
      footer={
        <div className="dialog-footer-row">
          <button
            type="button"
            className="btn btn-ghost"
            disabled={!!busy}
            onClick={() => fileRef.current?.click()}
          >
            {busy === "import" ? <Loader2 size={15} className="spin" /> : <Upload size={15} />}
            {busy === "import" && importStage ? importStage : "Import file"}
          </button>
          <div className="dialog-footer-actions">
            {busy === "import" && abortRef.current ? (
              <button type="button" className="btn btn-ghost" onClick={() => abortRef.current?.abort()}>
                Stop import
              </button>
            ) : (
              <button type="button" className="btn btn-ghost" disabled={!!busy} onClick={onClose}>
                Cancel
              </button>
            )}
            <button type="button" className="btn btn-primary" disabled={!!busy} onClick={create}>
              {busy === "create" ? <Loader2 size={16} className="spin" /> : <FileSpreadsheet size={16} />}
              Create sheet <ArrowRight size={15} />
            </button>
          </div>
        </div>
      }
    >
      <div className="new-sheet-panel">
        <label className="field">
          <span className="field-label">Name</span>
          <input
            value={title}
            autoFocus
            placeholder={selected.name}
            onChange={(event) => {
              setTitle(event.target.value);
              setTitleTouched(true);
            }}
            onKeyDown={(event) => {
              if (event.key === "Enter" && !busy) void create();
            }}
          />
        </label>

        <div className="field">
          <span className="field-label">Template</span>
          <div className="template-grid">
            {TEMPLATES.map((template) => (
              <button
                key={template.id}
                type="button"
                className={cx("template-card", templateId === template.id && "is-selected")}
                onClick={() => setTemplateId(template.id)}
              >
                <span className="template-preview" style={{ "--template-color": template.color } as CSSProperties}>
                  <i />
                  <i />
                  <i />
                  <i />
                  <i />
                  <i />
                  <i />
                  <i />
                  <b />
                </span>
                <span className="template-copy">
                  <strong>{template.name}</strong>
                  <small>{template.description}</small>
                </span>
              </button>
            ))}
          </div>
        </div>

        <input
          ref={fileRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          hidden
          onChange={(event) => {
            const file = event.target.files?.[0];
            if (file) void importFile(file);
          }}
        />

        {error && <div className="form-error">{error}</div>}
      </div>
    </Dialog>
  );
}
