// Sort dialog for big sheets: pick the column and direction, and say whether the first row is a header.
import { useState } from "react";
import { columnName } from "@formstr/sheets-grid";

export function BigSortDialog({
  firstCol,
  lastCol,
  activeCol,
  rows,
  onSort,
  onClose,
}: {
  firstCol: number;
  lastCol: number;
  activeCol: number;
  rows: number;
  onSort: (options: { col: number; descending: boolean; header: boolean }) => void;
  onClose: () => void;
}) {
  const [col, setCol] = useState(Math.min(Math.max(activeCol, firstCol), lastCol));
  const [descending, setDescending] = useState(false);
  const [header, setHeader] = useState(false);
  const columns = Array.from({ length: Math.min(lastCol - firstCol + 1, 200) }, (_, i) => firstCol + i);

  return (
    <div className="big-dialog-backdrop" onPointerDown={onClose}>
      <div
        className="big-dialog"
        role="dialog"
        aria-label="Sort range"
        onPointerDown={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === "Escape") onClose();
          if (event.key === "Enter") onSort({ col, descending, header });
        }}
      >
        <h3>Sort range</h3>
        <p className="big-dialog-sub">{rows.toLocaleString()} rows are sorted; each row moves as a whole.</p>
        <label className="big-dialog-check">
          <input type="checkbox" checked={header} onChange={(event) => setHeader(event.target.checked)} /> The first row is a header (stays on top)
        </label>
        <div className="big-dialog-row">
          <select value={col} onChange={(event) => setCol(Number(event.target.value))} aria-label="Sort by column">
            {columns.map((c) => (
              <option key={c} value={c}>
                Sort by column {columnName(c)}
              </option>
            ))}
          </select>
          <select value={descending ? "desc" : "asc"} onChange={(event) => setDescending(event.target.value === "desc")} aria-label="Order">
            <option value="asc">A → Z (smallest first)</option>
            <option value="desc">Z → A (largest first)</option>
          </select>
        </div>
        <div className="big-dialog-actions">
          <span style={{ flex: 1 }} />
          <button type="button" className="btn btn-sm" onClick={onClose}>
            Cancel
          </button>
          <button type="button" className="btn btn-primary btn-sm" onClick={() => onSort({ col, descending, header })}>
            Sort
          </button>
        </div>
      </div>
    </div>
  );
}
