import { useEffect, useSyncExternalStore } from "react";
import { AlertCircle, CheckCircle2, Loader2, UploadCloud, X } from "lucide-react";
import { uploadProgress } from "../lib/v2/progress";
import { formatBytes } from "../lib/utils";

/** Floating card, like Drive's upload list: shows a save's chunks going to the Blossom servers. */
export function UploadProgressCard() {
  const progress = useSyncExternalStore(uploadProgress.subscribe, uploadProgress.get);

  // A finished upload clears itself; an error stays until it is dismissed.
  useEffect(() => {
    if (progress?.phase !== "done") return;
    const timer = window.setTimeout(() => uploadProgress.dismiss(), 4000);
    return () => window.clearTimeout(timer);
  }, [progress]);

  if (!progress) return null;
  const { phase, tabs, bytesDone, bytesTotal, blobsDone, blobsTotal, servers } = progress;
  const percent = phase === "done" ? 100 : bytesTotal > 0 ? Math.min(99, Math.floor((bytesDone / bytesTotal) * 100)) : 0;
  const busy = phase === "preparing" || phase === "uploading";

  return (
    <div className="upload-card" role="status" aria-live="polite">
      <div className="upload-card-head">
        <span className="upload-card-icon">
          {phase === "done" ? <CheckCircle2 size={17} /> : phase === "error" ? <AlertCircle size={17} /> : <UploadCloud size={17} />}
        </span>
        <strong>
          {phase === "done" ? "Upload complete" : phase === "error" ? "Upload failed" : phase === "preparing" ? "Preparing upload" : "Uploading to Blossom"}
        </strong>
        {!busy && (
          <button type="button" className="icon-btn" aria-label="Dismiss" onClick={() => uploadProgress.dismiss()}>
            <X size={15} />
          </button>
        )}
      </div>

      <div className="upload-card-summary">
        {phase === "error" ? (
          <span className="upload-card-error">{progress.error ?? progress.message}</span>
        ) : (
          <>
            {busy && <Loader2 size={13} className="spin" />}
            <span>{progress.message}</span>
          </>
        )}
      </div>

      <div className="upload-bar" aria-hidden="true">
        <span className={phase === "error" ? "is-error" : ""} style={{ width: `${phase === "error" ? 100 : percent}%` }} />
      </div>
      {(phase !== "error" || blobsTotal > 0) && (
      <div className="upload-card-meta">
        <span>
          {blobsTotal > 0 ? `${blobsDone} of ${blobsTotal} files` : "Nothing to send"}
          {bytesTotal > 0 ? ` · ${formatBytes(bytesDone)} of ${formatBytes(bytesTotal)}` : ""}
        </span>
        <span>{percent}%</span>
      </div>
      )}

      {tabs.length > 0 && (
        <ul className="upload-card-tabs">
          {tabs.map((tab, index) => (
            <li key={index}>
              <span className="upload-card-tab-name">{tab.name}</span>
              <span>
                {tab.stored >= tab.chunks ? <CheckCircle2 size={13} className="upload-ok" /> : `${tab.stored}/${tab.chunks}`}
              </span>
            </li>
          ))}
        </ul>
      )}
      {servers.length > 0 && (
        <div className="upload-card-servers">
          {servers.length === 1 ? "Server" : `${servers.length} servers (${servers.length}x copies)`}: {servers.join(", ")}
        </div>
      )}
    </div>
  );
}
