import { useEffect, useState } from "react";
import { HardDrive, Loader2, Plus, X } from "lucide-react";
import { Dialog } from "./Dialog";
import { useLibrary } from "../providers/LibraryProvider";
import { useToast } from "../providers/ToastProvider";
import { sanitizeServers } from "../lib/v2/blossom";

/**
 * Where large sheets are stored. The list is the user's Blossom server list (a Nostr
 * kind 10063 event), shared with other Nostr apps such as Formstr Drive. Which servers
 * are ticked is kept on this device. Every ticked server gets every upload (replication).
 */
export function StorageServersDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  const library = useLibrary();
  const { toast } = useToast();
  const [loaded, setLoaded] = useState<string[] | null>(null);
  const [servers, setServers] = useState<string[]>([]);
  const [unticked, setUnticked] = useState<string[]>([]);
  const [draft, setDraft] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    if (!open) return;
    let cancelled = false;
    setLoaded(null);
    setDraft("");
    setError("");
    setUnticked(library.uploadPrefs.off);
    library
      .refreshBlossomServers()
      .then((list) => {
        if (cancelled) return;
        setLoaded(list);
        setServers(list);
      })
      .catch(() => {
        if (cancelled) return;
        setLoaded([]);
        setServers([]);
      });
    return () => {
      cancelled = true;
    };
    // Reload only when the dialog opens.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const ticked = servers.filter((url) => !unticked.includes(url));

  const add = () => {
    const [clean] = sanitizeServers([draft.trim()]);
    if (!clean) {
      setError("Enter a server address that starts with https://");
      return;
    }
    setError("");
    if (!servers.includes(clean)) setServers([...servers, clean]);
    setUnticked(unticked.filter((url) => url !== clean));
    setDraft("");
  };

  const toggle = (url: string) =>
    setUnticked(unticked.includes(url) ? unticked.filter((item) => item !== url) : [...unticked, url]);

  const save = async () => {
    // A pending, valid address in the box counts as added.
    const pending = sanitizeServers([draft.trim()])[0];
    const next = pending && !servers.includes(pending) ? [...servers, pending] : servers;
    const before = loaded ?? [];
    const removed = before.filter((url) => !next.includes(url));
    const added = next.filter((url) => !before.includes(url));
    const nextOff = unticked.filter((url) => next.includes(url));
    setSaving(true);
    setError("");
    try {
      // Only the list itself is a Nostr event; skip the signer when it did not change.
      if (removed.length > 0 || added.length > 0) await library.saveBlossomServers(next, removed);
      library.saveUploadPrefs({ off: nextOff });
      toast("Storage servers saved.", "success");
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't save your servers.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Dialog
      open={open}
      onClose={saving ? () => {} : onClose}
      title="Storage servers"
      subtitle="Large spreadsheets are encrypted and stored on Blossom servers you choose."
      size="md"
      footer={
        <div className="dialog-footer-row">
          <span />
          <div className="dialog-footer-actions">
            <button type="button" className="btn btn-ghost" disabled={saving} onClick={onClose}>
              Cancel
            </button>
            <button type="button" className="btn btn-primary" disabled={saving || loaded === null} onClick={save}>
              {saving ? <Loader2 size={16} className="spin" /> : <HardDrive size={16} />}
              Save servers
            </button>
          </div>
        </div>
      }
    >
      <div className="new-sheet-panel">
        <div className="field">
          <span className="field-label">Replicate my data to these servers</span>
          {loaded === null ? (
            <div className="server-empty">
              <Loader2 size={14} className="spin" /> Looking up your server list…
            </div>
          ) : servers.length === 0 ? (
            <div className="server-empty">No servers yet. Add one below to save large sheets.</div>
          ) : (
            <ul className="server-list">
              {servers.map((url) => (
                <li key={url}>
                  <label className="server-pick">
                    <input type="checkbox" checked={!unticked.includes(url)} onChange={() => toggle(url)} />
                    <span>{url}</span>
                  </label>
                  <button
                    type="button"
                    className="icon-btn"
                    title="Remove from my list"
                    aria-label={`Remove ${url}`}
                    onClick={() => setServers(servers.filter((item) => item !== url))}
                  >
                    <X size={15} />
                  </button>
                </li>
              ))}
            </ul>
          )}
          {servers.length > 0 && ticked.length === 0 && (
            <div className="form-error">Tick at least one server, or sheets too big for a Nostr event can't be saved.</div>
          )}
        </div>

        {ticked.length >= 2 && (
          <div className="server-warning" role="alert">
            <strong>
              Your data is copied to all {ticked.length} ticked servers ({ticked.length}x).
            </strong>
            <small>
              Upload bandwidth and storage used are multiplied by {ticked.length}, and big sheets save slower. Tick a
              single server to store one copy and turn replication off.
            </small>
          </div>
        )}

        <label className="field">
          <span className="field-label">Add a server</span>
          <div className="server-add">
            <input
              value={draft}
              placeholder="https://blossom.example.com"
              onChange={(event) => setDraft(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  add();
                }
              }}
            />
            <button type="button" className="btn btn-outline" onClick={add} disabled={!draft.trim()}>
              <Plus size={15} /> Add
            </button>
          </div>
        </label>

        <p className="server-note">
          The list is saved to your Nostr profile and shared with other apps that use Blossom; changing it asks your
          signer to approve one event. Unticked servers are still used to open sheets, just not to upload to. The ticks
          stay on this device.
        </p>

        {error && <div className="form-error">{error}</div>}
      </div>
    </Dialog>
  );
}
