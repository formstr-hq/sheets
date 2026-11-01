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
  const [error, setError] = useState("");
  const fileRef = useRef<HTMLInputElement>(null);

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

  const importFile = async (file: File) => {
    setBusy("import");
    setError("");
    try {
      const imported = await parseSpreadsheetFile(file, finalTitle);
      if (imported.sheets.length === 0) throw new Error("This file has no worksheets.");
      const entry = await library.importDocument(imported.title || finalTitle, imported.sheets);
      onClose();
      navigate(sheetPath(entry));
      toast(`Imported ${imported.title}.`, "success");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't import that file.");
    } finally {
      setBusy(null);
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
            Import file
          </button>
          <div className="dialog-footer-actions">
            <button type="button" className="btn btn-ghost" disabled={!!busy} onClick={onClose}>
              Cancel
            </button>
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
