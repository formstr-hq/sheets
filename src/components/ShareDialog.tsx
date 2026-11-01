import { useEffect, useMemo, useState } from "react";
import {
  AlertCircle,
  Check,
  Copy,
  Link2,
  Loader2,
  LockKeyhole,
  ShieldCheck,
  Trash2,
  UserPlus,
} from "lucide-react";
import { Dialog } from "./Dialog";
import { Avatar } from "./Avatar";
import { Select } from "./Select";
import { useAuth } from "../providers/AuthProvider";
import { useLibrary } from "../providers/LibraryProvider";
import { useProfiles } from "../providers/ProfileProvider";
import { useToast } from "../providers/ToastProvider";
import { parsePubkey, shortNpub, toNpub } from "../lib/utils";
import type { Collaborator, LibraryEntry } from "../lib/types";

export function ShareDialog({
  open,
  onClose,
  entry,
}: {
  open: boolean;
  onClose: () => void;
  entry: LibraryEntry;
}) {
  const auth = useAuth();
  const library = useLibrary();
  const profiles = useProfiles();
  const { toast } = useToast();
  const [npub, setNpub] = useState("");
  const [role, setRole] = useState<"viewer" | "editor">("viewer");
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState("");
  const [confirmRemove, setConfirmRemove] = useState<string | null>(null);
  const [copied, setCopied] = useState("");

  const isOwner = entry.role === "owner";
  const ownerProfile = entry.ownerPubkey ? profiles.profiles.get(entry.ownerPubkey) : null;
  const ownerName =
    ownerProfile?.displayName ||
    ownerProfile?.name ||
    entry.ownerName ||
    (entry.ownerPubkey ? shortNpub(entry.ownerPubkey) : "Sheet owner");

  const people = useMemo(() => entry.sharedWith, [entry.sharedWith]);

  useEffect(() => {
    if (!open) return;
    setNpub("");
    setError("");
    setBusy(null);
    setConfirmRemove(null);
    setCopied("");
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const keys = [entry.ownerPubkey, ...people.map((person) => person.pubkey)].filter(Boolean);
    if (keys.length > 0) void profiles.ensureProfiles(keys);
  }, [entry.ownerPubkey, open, people, profiles.ensureProfiles]);

  const copy = async (value: string, key: string) => {
    try {
      await navigator.clipboard.writeText(value);
      setCopied(key);
      toast("Link copied.", "success");
      window.setTimeout(() => setCopied(""), 1600);
    } catch {
      toast("Couldn't copy the link.", "error");
    }
  };

  const addPerson = async () => {
    setBusy("add");
    setError("");
    try {
      const pubkey = parsePubkey(npub);
      if (pubkey === auth.pubkey) throw new Error("That's you. You already have access.");
      if (pubkey === entry.ownerPubkey) throw new Error("The owner already has access.");
      await profiles.ensureProfiles([pubkey]);
      const profile = profiles.profileFor(pubkey);
      const collaborator: Collaborator = {
        pubkey,
        npub: toNpub(pubkey),
        name: profile?.displayName || profile?.name,
        picture: profile?.picture,
        role,
        grantedAt: Date.now(),
      };
      await library.shareSheet(entry, collaborator, role);
      setNpub("");
      toast(`${profile?.displayName || profile?.name || shortNpub(pubkey)} can now ${role === "editor" ? "edit" : "view"}.`, "success");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't grant access.");
    } finally {
      setBusy(null);
    }
  };

  const updateRole = async (person: Collaborator, nextRole: "viewer" | "editor") => {
    if (person.role === nextRole) return;
    setBusy(`role-${person.pubkey}`);
    setError("");
    try {
      await library.shareSheet(entry, { ...person, role: nextRole }, nextRole);
      toast(`Role updated to ${nextRole}.`, "success");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't update access.");
    } finally {
      setBusy(null);
    }
  };

  const removePerson = async (pubkey: string) => {
    setBusy(`remove-${pubkey}`);
    setError("");
    try {
      await library.revokeAccess(entry, pubkey);
      setConfirmRemove(null);
      toast("Access removed.", "success");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't remove access.");
    } finally {
      setBusy(null);
    }
  };

  return (
    <Dialog
      open={open}
      onClose={busy ? () => {} : onClose}
      title="Share this sheet"
      subtitle="Access is encrypted to each person's Nostr public key."
      size="md"
      footer={
        <div className="dialog-footer-row">
          <span className="share-footer-note">
            <LockKeyhole size={13} /> Keys are never posted in plaintext.
          </span>
          <button type="button" className="btn btn-ghost" onClick={onClose} disabled={!!busy}>
            Done
          </button>
        </div>
      }
    >
      <div className="share-dialog">
        <section className="share-section">
          <div className="share-section-head">
            <h3>People with access</h3>
            <span>{people.length + 1} total</span>
          </div>

          <div className="access-list">
            <div className="access-row">
              <Avatar size={34} name={ownerName} picture={ownerProfile?.picture || entry.ownerPicture} />
              <div className="access-person">
                <strong>{ownerName}</strong>
                <small>{entry.ownerPubkey ? shortNpub(entry.ownerPubkey) : "Workspace owner"}</small>
              </div>
              <span className="role-chip role-owner">Owner</span>
            </div>

            {people.map((person) => {
              const profile = profiles.profiles.get(person.pubkey);
              const name = profile?.displayName || profile?.name || person.name || shortNpub(person.pubkey);
              const removing = confirmRemove === person.pubkey;
              return (
                <div className="access-row" key={person.pubkey}>
                  <Avatar size={34} name={name} picture={profile?.picture || person.picture} />
                  <div className="access-person">
                    <strong>{name}</strong>
                    <small>{shortNpub(person.pubkey)}</small>
                  </div>
                  {isOwner ? (
                    removing ? (
                      <div className="access-confirm">
                        <span>Remove?</span>
                        <button
                          type="button"
                          className="btn btn-danger btn-xs"
                          disabled={!!busy}
                          onClick={() => void removePerson(person.pubkey)}
                        >
                          {busy === `remove-${person.pubkey}` ? (
                            <Loader2 size={12} className="spin" />
                          ) : (
                            "Remove"
                          )}
                        </button>
                        <button
                          type="button"
                          className="btn btn-ghost btn-xs"
                          onClick={() => setConfirmRemove(null)}
                        >
                          Cancel
                        </button>
                      </div>
                    ) : (
                      <div className="access-controls">
                        <Select
                          value={person.role}
                          disabled={!!busy}
                          onValueChange={(value) =>
                            void updateRole(person, value as "viewer" | "editor")
                          }
                          options={[
                            { value: "viewer", label: "Viewer" },
                            { value: "editor", label: "Editor" },
                          ]}
                          ariaLabel={`Role for ${name}`}
                          className="role-select-trigger"
                        />
                        <button
                          type="button"
                          className="icon-btn icon-btn-danger"
                          aria-label={`Remove ${name}`}
                          onClick={() => setConfirmRemove(person.pubkey)}
                        >
                          <Trash2 size={15} />
                        </button>
                      </div>
                    )
                  ) : (
                    <span className={`role-chip role-${person.role}`}>
                      {person.role === "editor" ? "Editor" : "Viewer"}
                    </span>
                  )}
                </div>
              );
            })}
          </div>

          {isOwner ? (
            <div className="add-person">
              <div className="add-person-input">
                <UserPlus size={16} />
                <input
                  value={npub}
                  placeholder="npub1… or hex public key"
                  onChange={(event) => setNpub(event.target.value)}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" && npub.trim()) void addPerson();
                  }}
                />
              </div>
              <Select
                value={role}
                onValueChange={(value) => setRole(value as "viewer" | "editor")}
                options={[
                  { value: "viewer", label: "Can view" },
                  { value: "editor", label: "Can edit" },
                ]}
                ariaLabel="Access level"
                className="role-select-trigger"
              />
              <button
                type="button"
                className="btn btn-primary"
                disabled={!npub.trim() || !!busy}
                onClick={() => void addPerson()}
              >
                {busy === "add" ? <Loader2 size={15} className="spin" /> : "Invite"}
              </button>
            </div>
          ) : (
            <div className="callout callout-info">
              <ShieldCheck size={16} />
              <p>Only the owner can invite people. You can still copy a share link below.</p>
            </div>
          )}

          {error && (
            <div className="callout callout-error">
              <AlertCircle size={16} />
              <p>{error}</p>
            </div>
          )}
        </section>

        <section className="share-section">
          <div className="share-section-head">
            <h3>Link access</h3>
            <span>Keys live in the URL fragment</span>
          </div>
          <div className="link-actions">
            <button
              type="button"
              className="link-action"
              onClick={() => void copy(library.buildShareLink(entry, "view"), "view")}
            >
              <span className="link-action-icon">
                <Link2 size={17} />
              </span>
              <span>
                <strong>Copy view link</strong>
                <small>Anyone with the link can read this sheet.</small>
              </span>
              {copied === "view" ? <Check size={16} /> : <Copy size={16} />}
            </button>
            {entry.editKey && (
              <button
                type="button"
                className="link-action"
                onClick={() => void copy(library.buildShareLink(entry, "edit"), "edit")}
              >
                <span className="link-action-icon">
                  <ShieldCheck size={17} />
                </span>
                <span>
                  <strong>Copy edit link</strong>
                  <small>Anyone with the link can edit this sheet.</small>
                </span>
                {copied === "edit" ? <Check size={16} /> : <Copy size={16} />}
              </button>
            )}
          </div>
        </section>

        <div className="callout callout-warning">
          <AlertCircle size={16} />
          <p>
            Removing access prevents new grants, but anyone who already opened this
            sheet may have a local copy. Treat access removal as future-facing, not
            a retroactive wipe.
          </p>
        </div>
      </div>
    </Dialog>
  );
}
