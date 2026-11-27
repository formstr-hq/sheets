// Conditional formatting dialog for big sheets: colour the selected cells when their value matches.
import { useState } from "react";
import type { WorkbookConditionalRule } from "../lib/workbook";
import { CONDITIONS, SCALE_TYPES } from "../lib/v2/bigConditional";

export function BigConditionalDialog({
  rangeLabel,
  hasRules,
  onSave,
  onClearHere,
  onClose,
}: {
  rangeLabel: string;
  /** Is there already a rule on part of the selection (so "Remove rules" does something)? */
  hasRules: boolean;
  onSave: (condition: WorkbookConditionalRule["condition"], format: WorkbookConditionalRule["format"]) => void;
  onClearHere: () => void;
  onClose: () => void;
}) {
  const [type, setType] = useState<string>("gt");
  const [first, setFirst] = useState("");
  const [second, setSecond] = useState("");
  const [fill, setFill] = useState("#fde2e2");
  const [useFill, setUseFill] = useState(true);
  const [text, setText] = useState("#b42318");
  const [useText, setUseText] = useState(true);
  const [toFill, setToFill] = useState("#63be7b");
  const [error, setError] = useState("");
  const scale = SCALE_TYPES.includes(type);
  const wants = CONDITIONS.find((item) => item.value === type)?.values ?? 1;

  const save = () => {
    if (wants >= 1 && first.trim() === "") return setError("Enter a value to compare with.");
    if (wants === 2 && second.trim() === "") return setError("Enter both values.");
    if (scale) {
      onSave({ type, values: [] }, type === "dataBar" ? { cellColor: fill } : { cellColor: fill, toColor: toFill });
      return;
    }
    if (!useFill && !useText) return setError("Pick a fill colour or a text colour.");
    onSave(
      { type, values: wants === 0 ? [] : wants === 1 ? [first.trim()] : [first.trim(), second.trim()] },
      { ...(useFill ? { cellColor: fill } : {}), ...(useText ? { textColor: text } : {}) },
    );
  };

  return (
    <div className="big-dialog-backdrop" onPointerDown={onClose}>
      <div
        className="big-dialog"
        role="dialog"
        aria-label="Conditional formatting"
        onPointerDown={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === "Escape") onClose();
          if (event.key === "Enter") save();
        }}
      >
        <h3>Conditional formatting</h3>
        <p className="big-dialog-sub">Format cells {rangeLabel} when the value:</p>
        <div className="big-dialog-row">
          <select
            value={type}
            onChange={(event) => {
              setType(event.target.value);
              // Sensible starting colours for bars and scales.
              if (event.target.value === "dataBar") setFill("#8ab4f8");
              else if (event.target.value === "colorScale") setFill("#fde2e2");
            }}
          >
            {CONDITIONS.map((item) => (
              <option key={item.value} value={item.value}>
                {item.label}
              </option>
            ))}
          </select>
          {wants >= 1 && <input value={first} onChange={(event) => setFirst(event.target.value)} placeholder="value" autoFocus />}
          {wants === 2 && <input value={second} onChange={(event) => setSecond(event.target.value)} placeholder="and" />}
        </div>
        {scale ? (
          <div className="big-dialog-row">
            <label className="big-dialog-check">
              {type === "dataBar" ? "Bar colour" : "Lowest"}
              <input type="color" value={fill} onChange={(event) => setFill(event.target.value)} />
            </label>
            {type === "colorScale" && (
              <label className="big-dialog-check">
                Highest
                <input type="color" value={toFill} onChange={(event) => setToFill(event.target.value)} />
              </label>
            )}
            <span className="big-dialog-sub">Compared with the lowest and highest numbers in the selection now.</span>
          </div>
        ) : null}
        <div className="big-dialog-row" hidden={scale}>
          <label className="big-dialog-check">
            <input type="checkbox" checked={useFill} onChange={(event) => setUseFill(event.target.checked)} /> Fill
            <input type="color" value={fill} onChange={(event) => setFill(event.target.value)} disabled={!useFill} />
          </label>
          <label className="big-dialog-check">
            <input type="checkbox" checked={useText} onChange={(event) => setUseText(event.target.checked)} /> Text
            <input type="color" value={text} onChange={(event) => setText(event.target.value)} disabled={!useText} />
          </label>
          <span className="big-dialog-preview" style={{ background: useFill ? fill : undefined, color: useText ? text : undefined }}>
            Sample
          </span>
        </div>
        {error && <p className="big-dialog-error">{error}</p>}
        <div className="big-dialog-actions">
          {hasRules && (
            <button type="button" className="btn btn-sm" onClick={onClearHere}>
              Remove rules here
            </button>
          )}
          <span style={{ flex: 1 }} />
          <button type="button" className="btn btn-sm" onClick={onClose}>
            Cancel
          </button>
          <button type="button" className="btn btn-primary btn-sm" onClick={save}>
            Add rule
          </button>
        </div>
      </div>
    </div>
  );
}
