// Data validation dialog for big sheets: choose what the selected cells accept.
import { useState } from "react";
import type { WorkbookValidation } from "../lib/workbook";
import { buildRule, OPERATORS, RULE_TYPES, type Operator } from "../lib/v2/bigValidation";

export function BigValidationDialog({
  current,
  rangeLabel,
  onSave,
  onRemove,
  onClose,
}: {
  current?: WorkbookValidation;
  rangeLabel: string;
  onSave: (rule: WorkbookValidation) => void;
  onRemove: () => void;
  onClose: () => void;
}) {
  const [type, setType] = useState(current?.type ?? "list");
  const [list, setList] = useState((current?.options ?? []).join(", "));
  const [op, setOp] = useState<Operator>((current?.params?.op as Operator | undefined) ?? "between");
  const [a, setA] = useState(String(current?.params?.a ?? ""));
  const [b, setB] = useState(String(current?.params?.b ?? ""));
  const [error, setError] = useState("");
  const range = op === "between" || op === "notBetween";

  const save = () => {
    const rule = buildRule(type, { list, op, a, b });
    if (typeof rule === "string") setError(rule);
    else onSave(rule);
  };

  return (
    <div className="big-dialog-backdrop" onPointerDown={onClose}>
      <div
        className="big-dialog"
        role="dialog"
        aria-label="Data validation"
        onPointerDown={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === "Escape") onClose();
          if (event.key === "Enter" && (event.target as HTMLElement).tagName !== "TEXTAREA") save();
        }}
      >
        <h3>Data validation</h3>
        <p className="big-dialog-sub">Cells {rangeLabel} will only accept:</p>
        <label className="big-dialog-field">
          <span>Rule</span>
          <select value={type} onChange={(event) => setType(event.target.value)}>
            {RULE_TYPES.map((item) => (
              <option key={item.value} value={item.value}>
                {item.label}
              </option>
            ))}
          </select>
        </label>
        {type === "checkbox" ? (
          <p className="big-dialog-sub">The cells show a box that ticks and unticks when clicked (the value is TRUE or FALSE).</p>
        ) : type === "textContains" ? (
          <label className="big-dialog-field">
            <span>The text must contain</span>
            <input value={a} onChange={(event) => setA(event.target.value)} autoFocus />
          </label>
        ) : type === "list" ? (
          <label className="big-dialog-field">
            <span>Options (separated by commas)</span>
            <textarea rows={3} value={list} onChange={(event) => setList(event.target.value)} autoFocus />
          </label>
        ) : (
          <div className="big-dialog-row">
            <select value={op} onChange={(event) => setOp(event.target.value as Operator)}>
              {OPERATORS.map((item) => (
                <option key={item.value} value={item.value}>
                  {item.label}
                </option>
              ))}
            </select>
            <input value={a} onChange={(event) => setA(event.target.value)} inputMode={type === "date" ? "text" : "decimal"} placeholder={type === "date" ? (range ? "from yyyy-mm-dd" : "yyyy-mm-dd") : range ? "from" : "number"} autoFocus />
            {range && <input value={b} onChange={(event) => setB(event.target.value)} inputMode={type === "date" ? "text" : "decimal"} placeholder={type === "date" ? "to yyyy-mm-dd" : "to"} />}
          </div>
        )}
        {error && <p className="big-dialog-error">{error}</p>}
        <p className="big-dialog-sub">Anything else typed into these cells is rejected. Empty cells are always fine.</p>
        <div className="big-dialog-actions">
          {current && (
            <button type="button" className="btn btn-sm" onClick={onRemove}>
              Remove rule
            </button>
          )}
          <span style={{ flex: 1 }} />
          <button type="button" className="btn btn-sm" onClick={onClose}>
            Cancel
          </button>
          <button type="button" className="btn btn-primary btn-sm" onClick={save}>
            Save
          </button>
        </div>
      </div>
    </div>
  );
}
