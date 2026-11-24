// Find (Ctrl+F) and replace (Ctrl+H) for big sheets: searches every part of the tab, one after another.
import { useEffect, useRef, useState } from "react";
import { ChevronDown, ChevronUp, Loader2, X } from "lucide-react";
import type { BigSheetStore } from "../lib/v2/bigSheet";
import { findCells, replaceCells, type FindOptions, type Match } from "../lib/v2/bigFind";

export function BigFindBar({
  store,
  tab,
  canEdit,
  withReplace,
  onGo,
  onReplaced,
  onClose,
}: {
  store: BigSheetStore;
  tab: number;
  canEdit: boolean;
  withReplace: boolean;
  /** Moves the grid's selection to a match. */
  onGo: (match: Match) => void;
  /** Cells whose text was replaced (so formulas that read them can be refreshed). */
  onReplaced: (changed: Match[]) => void;
  onClose: () => void;
}) {
  const [query, setQuery] = useState("");
  const [replacement, setReplacement] = useState("");
  const [options, setOptions] = useState<FindOptions>({});
  const [matches, setMatches] = useState<Match[]>([]);
  const [current, setCurrent] = useState(-1);
  const [truncated, setTruncated] = useState(false);
  const [progress, setProgress] = useState<string | null>(null);
  const [message, setMessage] = useState("");
  const input = useRef<HTMLInputElement>(null);
  /** Bumped to cancel a search that is still running (new search, new tab, closing). */
  const run = useRef(0);

  useEffect(() => {
    input.current?.focus();
    input.current?.select();
    return () => {
      run.current++;
    };
  }, []);
  useEffect(() => {
    run.current++;
    setMatches([]);
    setCurrent(-1);
    setProgress(null);
  }, [tab]);

  const search = async (): Promise<Match[]> => {
    const mine = ++run.current;
    setMessage("");
    setProgress("Searching…");
    try {
      const result = await findCells(
        store,
        tab,
        query,
        options,
        (done, total) => run.current === mine && setProgress(`Searching… ${done} of ${total} parts`),
        () => run.current !== mine,
      );
      if (run.current !== mine) return [];
      setMatches(result.matches);
      setTruncated(result.truncated);
      setCurrent(result.matches.length > 0 ? 0 : -1);
      setMessage(result.matches.length === 0 ? "No matches" : "");
      if (result.matches.length > 0) onGo(result.matches[0]);
      return result.matches;
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Couldn't search.");
      return [];
    } finally {
      if (run.current === mine) setProgress(null);
    }
  };

  const step = (delta: number) => {
    if (matches.length === 0) return void search();
    const next = (current + delta + matches.length) % matches.length;
    setCurrent(next);
    onGo(matches[next]);
  };

  const replaceNow = async (all: boolean) => {
    if (!canEdit || query === "") return;
    const list = matches.length > 0 ? matches : await search();
    const targets = all ? list : list.slice(current < 0 ? 0 : current, (current < 0 ? 0 : current) + 1);
    if (targets.length === 0) return;
    setProgress("Replacing…");
    try {
      const changed = await replaceCells(store, tab, targets, query, replacement, options);
      onReplaced(changed);
      const gone = new Set(changed.map((m) => `${m.row}:${m.col}`));
      const left = matches.filter((m) => !gone.has(`${m.row}:${m.col}`));
      setMatches(left);
      setCurrent(left.length > 0 ? Math.min(Math.max(current, 0), left.length - 1) : -1);
      setMessage(`Replaced ${changed.length.toLocaleString()} cell${changed.length === 1 ? "" : "s"}`);
      if (!all && left.length > 0) onGo(left[Math.min(Math.max(current, 0), left.length - 1)]);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Couldn't replace.");
    } finally {
      setProgress(null);
    }
  };

  return (
    <div
      className="big-find"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.stopPropagation();
          onClose();
        }
      }}
    >
      <div className="big-find-row">
        <input
          ref={input}
          className="big-find-input"
          placeholder="Find in this sheet"
          value={query}
          onChange={(event) => {
            setQuery(event.target.value);
            setMatches([]);
            setCurrent(-1);
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              if (matches.length > 0) step(event.shiftKey ? -1 : 1);
              else void search();
            }
          }}
        />
        <span className="big-find-count">
          {progress ? (
            <>
              <Loader2 size={12} className="spin" /> {progress}
            </>
          ) : matches.length > 0 ? (
            `${current + 1} of ${matches.length.toLocaleString()}${truncated ? "+" : ""}`
          ) : (
            message
          )}
        </span>
        <button type="button" className="big-find-btn" title="Previous (Shift+Enter)" onClick={() => step(-1)}>
          <ChevronUp size={14} />
        </button>
        <button type="button" className="big-find-btn" title="Next (Enter)" onClick={() => step(1)}>
          <ChevronDown size={14} />
        </button>
        <label className="big-find-opt">
          <input
            type="checkbox"
            checked={!!options.matchCase}
            onChange={(event) => {
              setOptions({ ...options, matchCase: event.target.checked });
              setMatches([]);
            }}
          />
          Match case
        </label>
        <label className="big-find-opt">
          <input
            type="checkbox"
            checked={!!options.wholeCell}
            onChange={(event) => {
              setOptions({ ...options, wholeCell: event.target.checked });
              setMatches([]);
            }}
          />
          Whole cell
        </label>
        <button type="button" className="big-find-btn" title="Close (Esc)" onClick={onClose}>
          <X size={14} />
        </button>
      </div>
      {withReplace && (
        <div className="big-find-row">
          <input
            className="big-find-input"
            placeholder="Replace with"
            value={replacement}
            disabled={!canEdit}
            onChange={(event) => setReplacement(event.target.value)}
          />
          <button type="button" className="btn btn-sm" disabled={!canEdit || query === "" || !!progress} onClick={() => void replaceNow(false)}>
            Replace
          </button>
          <button type="button" className="btn btn-sm" disabled={!canEdit || query === "" || !!progress} onClick={() => void replaceNow(true)}>
            Replace all
          </button>
        </div>
      )}
    </div>
  );
}
