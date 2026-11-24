import { useEffect, useState } from "react";
import { Loader2 } from "lucide-react";
import { Dialog } from "./Dialog";
import { useLibrary } from "../providers/LibraryProvider";
import { measureWorkbook, type FileInfo } from "../lib/v2/info";
import { formatBytes, formatRelativeTime } from "../lib/utils";
import type { SheetDocument } from "../lib/types";

/** Roughly where relays start refusing events (relay.formstr.app: 100 KB of content). */
const RELAY_EVENT_LIMIT_BYTES = 100 * 1024;

export function FileInfoDialog({
  open,
  onClose,
  doc,
  lastSavedAt,
}: {
  open: boolean;
  onClose: () => void;
  doc: SheetDocument | null;
  lastSavedAt?: number;
}) {
  const library = useLibrary();
  const [info, setInfo] = useState<FileInfo | null>(null);
  const [failed, setFailed] = useState(false);

  useEffect(() => {
    if (!open || !doc) return;
    let cancelled = false;
    setInfo(null);
    setFailed(false);
    measureWorkbook(doc)
      .then((result) => !cancelled && setInfo(result))
      .catch(() => !cancelled && setFailed(true));
    return () => {
      cancelled = true;
    };
  }, [open, doc]);

  const stored = info ? (info.storage === "blossom" ? info.blobBytes : info.eventBytes) : 0;
  const ratio = info && stored > 0 ? info.rawBytes / stored : 0;
  const oldEventBytes = info ? Math.ceil((info.rawBytes * 4) / 3) : 0;

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="File info"
      subtitle={doc?.title}
      size="md"
      footer={
        <div className="dialog-footer-actions dialog-footer-end">
          <button type="button" className="btn btn-primary" onClick={onClose}>
            Done
          </button>
        </div>
      }
    >
      {!doc || failed ? (
        <div className="server-empty">Couldn't measure this sheet.</div>
      ) : !info ? (
        <div className="server-empty">
          <Loader2 size={14} className="spin" /> Measuring…
        </div>
      ) : (
        <div className="new-sheet-panel">
          <dl className="info-grid">
            <dt>Size when saved</dt>
            <dd>
              <strong>{formatBytes(stored)}</strong>
              {ratio >= 1.5 && <small> · {ratio.toFixed(1)}× smaller than the plain data ({formatBytes(info.rawBytes)})</small>}
            </dd>

            <dt>Stored</dt>
            <dd>
              {info.storage === "blossom"
                ? `Encrypted on Blossom: ${info.blobCount} ${info.blobCount === 1 ? "file" : "files"}, plus a ${formatBytes(info.eventBytes)} pointer in the Nostr event`
                : "Inside the Nostr event (small enough, so no storage server is needed)"}
            </dd>

            {info.storage === "blossom" && (
              <>
                <dt>Servers</dt>
                <dd>{library.blossomServers.length > 0 ? library.blossomServers.join(", ") : "None set — add one under Storage servers"}</dd>
              </>
            )}

            <dt>Old format</dt>
            <dd>
              Would be about {formatBytes(oldEventBytes)} in one event
              {oldEventBytes > RELAY_EVENT_LIMIT_BYTES ? ", over what relays accept." : ", which fits."}
            </dd>

            <dt>Worksheets</dt>
            <dd>{info.tabs.length}</dd>

            <dt>Filled cells</dt>
            <dd>{info.cells.toLocaleString()}</dd>

            {lastSavedAt && (
              <>
                <dt>Last saved</dt>
                <dd>{formatRelativeTime(lastSavedAt)}</dd>
              </>
            )}
          </dl>

          <div className="field">
            <span className="field-label">Per worksheet</span>
            <ul className="server-list">
              {info.tabs.map((tab) => (
                <li key={tab.id}>
                  <span>
                    {tab.name} <small>· {tab.cells.toLocaleString()} cells</small>
                  </span>
                  <strong>{formatBytes(tab.storedBytes)}</strong>
                </li>
              ))}
            </ul>
          </div>

          <p className="server-note">
            Sizes are measured from what is open now, as it would be saved: compressed, then encrypted.
          </p>
        </div>
      )}
    </Dialog>
  );
}
