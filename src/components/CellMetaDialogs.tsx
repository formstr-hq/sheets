import { useEffect, useMemo, useState } from "react";
import { AtSign, Loader2, Search, Tag } from "lucide-react";
import { Dialog } from "./Dialog";
import { Avatar } from "./Avatar";
import { useProfiles } from "../providers/ProfileProvider";
import { colorFromString, parsePubkey, shortNpub, toNpub } from "../lib/utils";
import type { UserRef } from "../lib/types";

export interface PersonChoice extends UserRef {
  role?: "owner" | "editor" | "viewer";
}

export function PersonDialog({
  open,
  onClose,
  onPick,
  people,
}: {
  open: boolean;
  onClose: () => void;
  onPick: (user: UserRef) => void;
  people: PersonChoice[];
}) {
  const profiles = useProfiles();
  const [query, setQuery] = useState("");
  const [manual, setManual] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [resolved, setResolved] = useState<UserRef | null>(null);

  useEffect(() => {
    if (!open) return;
    setQuery("");
    setManual("");
    setBusy(false);
    setError("");
    setResolved(null);
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const keys = people.map((person) => person.pubkey);
    if (keys.length > 0) void profiles.ensureProfiles(keys);
  }, [open, people, profiles.ensureProfiles]);

  const filtered = useMemo(() => {
    const term = query.trim().toLowerCase();
    if (!term) return people;
    return people.filter((person) => {
      const profile = profiles.profiles.get(person.pubkey);
      return (
        (profile?.displayName ?? person.name ?? "").toLowerCase().includes(term) ||
        person.npub.toLowerCase().includes(term) ||
        person.pubkey.includes(term)
      );
    });
  }, [people, profiles.profiles, query]);

  const resolveManual = async () => {
    setBusy(true);
    setError("");
    setResolved(null);
    try {
      const pubkey = parsePubkey(manual);
      await profiles.ensureProfiles([pubkey]);
      const profile = profiles.profileFor(pubkey);
      setResolved({
        pubkey,
        npub: toNpub(pubkey),
        name: profile?.displayName || profile?.name,
        picture: profile?.picture,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Invalid public key.");
    } finally {
      setBusy(false);
    }
  };

  const pick = (user: UserRef) => {
    onPick(user);
    onClose();
  };

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="Tag a person"
      subtitle="Insert a Nostr profile pill into the selected cells."
      size="sm"
    >
      <div className="person-dialog">
        <label className="search-field">
          <Search size={15} />
          <input
            value={query}
            autoFocus
            placeholder="Search collaborators"
            onChange={(event) => setQuery(event.target.value)}
          />
        </label>

        {filtered.length > 0 && (
          <div className="person-list">
            {filtered.map((person) => {
              const profile = profiles.profiles.get(person.pubkey);
              const name = profile?.displayName || profile?.name || person.name || shortNpub(person.pubkey);
              return (
                <button
                  type="button"
                  className="person-row"
                  key={person.pubkey}
                  onClick={() =>
                    pick({
                      pubkey: person.pubkey,
                      npub: person.npub,
                      name,
                      picture: profile?.picture || person.picture,
                    })
                  }
                >
                  <Avatar size={30} name={name} picture={profile?.picture || person.picture} />
                  <span>
                    <strong>{name}</strong>
                    <small>{shortNpub(person.pubkey)}</small>
                  </span>
                  {person.role && (
                    <span className={`role-chip role-${person.role}`}>
                      {person.role === "owner" ? "Owner" : person.role === "editor" ? "Editor" : "Viewer"}
                    </span>
                  )}
                </button>
              );
            })}
          </div>
        )}

        <div className="person-manual">
          <div className="field-label">
            <AtSign size={13} /> Or use any npub
          </div>
          <div className="inline-field">
            <input
              value={manual}
              placeholder="npub1… or hex key"
              onChange={(event) => {
                setManual(event.target.value);
                setResolved(null);
                setError("");
              }}
              onKeyDown={(event) => {
                if (event.key === "Enter") void resolveManual();
              }}
            />
            <button type="button" className="btn btn-outline" disabled={busy} onClick={resolveManual}>
              {busy ? <Loader2 size={14} className="spin" /> : "Look up"}
            </button>
          </div>
          {error && <div className="form-error">{error}</div>}
          {resolved && (
            <button type="button" className="person-row is-resolved" onClick={() => pick(resolved)}>
              <Avatar size={30} name={resolved.name || "?"} picture={resolved.picture} />
              <span>
                <strong>{resolved.name || shortNpub(resolved.pubkey)}</strong>
                <small>{shortNpub(resolved.pubkey)}</small>
              </span>
              <span className="role-chip role-viewer">Use</span>
            </button>
          )}
        </div>
      </div>
    </Dialog>
  );
}

const TAG_PRESETS = ["High", "Medium", "Low", "Done", "Blocked", "Idea", "Draft", "Series", "One-off"];

export function TagDialog({
  open,
  onClose,
  onPick,
}: {
  open: boolean;
  onClose: () => void;
  onPick: (tag: string) => void;
}) {
  const [value, setValue] = useState("");

  useEffect(() => {
    if (open) setValue("");
  }, [open]);

  const preview = value.trim() || "Tag";
  const colors = colorFromString(preview);

  const apply = () => {
    const tag = value.trim();
    if (!tag) return;
    onPick(tag);
    onClose();
  };

  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="Add a tag pill"
      subtitle="A short, colorful label for the selected cells."
      size="sm"
      footer={
        <div className="dialog-footer-actions dialog-footer-end">
          <button type="button" className="btn btn-ghost" onClick={onClose}>
            Cancel
          </button>
          <button type="button" className="btn btn-primary" disabled={!value.trim()} onClick={apply}>
            <Tag size={15} /> Add tag
          </button>
        </div>
      }
    >
      <div className="tag-dialog">
        <label className="field">
          <span className="field-label">Tag text</span>
          <input
            value={value}
            autoFocus
            maxLength={40}
            placeholder="e.g. High priority"
            onChange={(event) => setValue(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") apply();
            }}
          />
        </label>
        <div className="tag-presets">
          {TAG_PRESETS.map((tag) => (
            <button type="button" key={tag} className="tag-suggestion" onClick={() => setValue(tag)}>
              {tag}
            </button>
          ))}
        </div>
        <div className="tag-preview">
          <span className="field-label">Preview</span>
          <span className="pill" style={{ background: colors.bg, color: colors.fg }}>
            {preview}
          </span>
        </div>
      </div>
    </Dialog>
  );
}
