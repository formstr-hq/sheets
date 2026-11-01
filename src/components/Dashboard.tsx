import { useEffect, useMemo, useRef, useState, type CSSProperties } from "react";
import { useNavigate } from "react-router-dom";
import {
  ChevronDown,
  Copy,
  Download,
  FileSpreadsheet,
  Grid2X2,
  KeyRound,
  Loader2,
  LockKeyhole,
  LogOut,
  Moon,
  MoreHorizontal,
  Plus,
  RefreshCw,
  Search,
  Sun,
  Trash2,
  Users,
} from "lucide-react";
import { Logo, LogoMark } from "./Logo";
import { Avatar } from "./Avatar";
import { Menu, MenuItem, MenuSeparator } from "./Menu";
import { Dialog } from "./Dialog";
import { NewSheetDialog } from "./NewSheetDialog";
import { useAuth } from "../providers/AuthProvider";
import { useLibrary } from "../providers/LibraryProvider";
import { useProfiles } from "../providers/ProfileProvider";
import { useToast } from "../providers/ToastProvider";
import { getTheme, setTheme } from "../lib/local";
import { downloadText, formatRelativeTime, shortNpub } from "../lib/utils";
import { sheetPath } from "../lib/routes";
import type { LibraryEntry } from "../lib/types";

type Filter = "all" | "owned" | "shared";
type Sort = "recent" | "name";

function SheetPreview({ color, seed }: { color: string; seed: string }) {
  const bars = useMemo(() => {
    let hash = 0;
    for (let i = 0; i < seed.length; i++) hash = (hash * 31 + seed.charCodeAt(i)) | 0;
    return Array.from({ length: 9 }, (_, index) => {
      const value = Math.abs(hash + index * 37) % 100;
      return 18 + (value % 52);
    });
  }, [seed]);
  return (
    <div className="sheet-preview" style={{ "--sheet-color": color } as CSSProperties}>
      <div className="sheet-preview-bar">
        <span />
        <span />
        <span />
      </div>
      <div className="sheet-preview-grid">
        {bars.map((width, index) => (
          <i key={index} style={{ width: `${width}%` }} />
        ))}
      </div>
      <div className="sheet-preview-lock">
        <LockKeyhole size={11} />
      </div>
    </div>
  );
}

export function Dashboard() {
  const auth = useAuth();
  const library = useLibrary();
  const profiles = useProfiles();
  const navigate = useNavigate();
  const { toast } = useToast();
  const [filter, setFilter] = useState<Filter>("all");
  const [sort, setSort] = useState<Sort>("recent");
  const [search, setSearch] = useState("");
  const [newOpen, setNewOpen] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<LibraryEntry | null>(null);
  const [deleting, setDeleting] = useState(false);
  const [theme, setThemeState] = useState<"light" | "dark">(getTheme);
  const searchRef = useRef<HTMLInputElement>(null);

  const userPubkey = auth.pubkey ?? "";
  const profile = profiles.profiles.get(userPubkey);
  const displayName = profile?.displayName || profile?.name || shortNpub(userPubkey);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "/" && !(event.target instanceof HTMLInputElement) && !(event.target instanceof HTMLTextAreaElement)) {
        event.preventDefault();
        searchRef.current?.focus();
      }
    };
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, []);

  useEffect(() => {
    const pubkeys = new Set<string>();
    for (const entry of library.entries) {
      if (entry.ownerPubkey) pubkeys.add(entry.ownerPubkey);
      for (const person of entry.sharedWith) pubkeys.add(person.pubkey);
    }
    if (pubkeys.size > 0) void profiles.ensureProfiles(Array.from(pubkeys));
  }, [library.entries, profiles]);

  const filtered = useMemo(() => {
    const term = search.trim().toLowerCase();
    const list = library.entries.filter((entry) => {
      if (filter === "owned" && entry.role !== "owner") return false;
      if (filter === "shared" && entry.role === "owner") return false;
      if (!term) return true;
      return (
        entry.title.toLowerCase().includes(term) ||
        (entry.ownerName ?? "").toLowerCase().includes(term) ||
        (entry.ownerPubkey ? profiles.profiles.get(entry.ownerPubkey)?.name ?? "" : "")
          .toLowerCase()
          .includes(term)
      );
    });
    return list.sort((a, b) => {
      if (sort === "name") return a.title.localeCompare(b.title);
      return (b.lastOpenedAt ?? b.updatedAt) - (a.lastOpenedAt ?? a.updatedAt);
    });
  }, [filter, library.entries, profiles.profiles, search, sort]);

  const openEntry = (entry: LibraryEntry) => {
    navigate(sheetPath(entry));
  };

  const copyLink = async (entry: LibraryEntry, mode: "view" | "edit") => {
    try {
      await navigator.clipboard.writeText(library.buildShareLink(entry, mode));
      toast(`${mode === "edit" ? "Edit" : "View"} link copied.`, "success");
    } catch {
      toast("Couldn't copy the link.", "error");
    }
  };

  const toggleTheme = () => {
    const next = theme === "light" ? "dark" : "light";
    setTheme(next);
    setThemeState(next);
  };

  const confirmDelete = async () => {
    if (!deleteTarget) return;
    setDeleting(true);
    try {
      await library.deleteSheet(deleteTarget);
      toast(
        deleteTarget.role === "owner" ? "Spreadsheet deleted." : "Removed from your sheets.",
        "success",
      );
      setDeleteTarget(null);
    } catch (err) {
      toast(err instanceof Error ? err.message : "Couldn't delete the sheet.", "error");
    } finally {
      setDeleting(false);
    }
  };

  const ownerLabel = (entry: LibraryEntry) => {
    if (entry.role === "owner") return "Private";
    if (!entry.ownerPubkey) return "Shared link";
    const owner = profiles.profiles.get(entry.ownerPubkey);
    return `Shared by ${owner?.displayName || owner?.name || entry.ownerName || shortNpub(entry.ownerPubkey)}`;
  };

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="sidebar-top">
          <Logo />
          <button type="button" className="btn btn-primary sidebar-new" onClick={() => setNewOpen(true)}>
            <Plus size={17} /> New sheet
          </button>
          <nav className="sidebar-nav">
            <button
              type="button"
              className={filter === "all" ? "is-active" : ""}
              onClick={() => setFilter("all")}
            >
              <Grid2X2 size={17} />
              All sheets
              <span>{library.entries.length}</span>
            </button>
            <button
              type="button"
              className={filter === "owned" ? "is-active" : ""}
              onClick={() => setFilter("owned")}
            >
              <LockKeyhole size={17} />
              My sheets
              <span>{library.entries.filter((entry) => entry.role === "owner").length}</span>
            </button>
            <button
              type="button"
              className={filter === "shared" ? "is-active" : ""}
              onClick={() => setFilter("shared")}
            >
              <Users size={17} />
              Shared with me
              <span>{library.entries.filter((entry) => entry.role !== "owner").length}</span>
            </button>
          </nav>
        </div>

        <div className="sidebar-bottom">
          <div className="relay-card">
            <span className="relay-pulse" />
            <div>
              <strong>{auth.relays.length} relays</strong>
              <span>End-to-end encrypted</span>
            </div>
          </div>

          <Menu
            align="left"
            width={230}
            trigger={
              <button type="button" className="user-chip">
                <Avatar
                  size={32}
                  name={profile?.displayName || profile?.name || displayName}
                  picture={profile?.picture}
                />
                <span className="user-chip-copy">
                  <strong>{displayName}</strong>
                  <small>{shortNpub(userPubkey)}</small>
                </span>
                <ChevronDown size={15} />
              </button>
            }
          >
            {(close) => (
              <>
                <div className="menu-user-card">
                  <Avatar
                    size={38}
                    name={profile?.displayName || profile?.name || displayName}
                    picture={profile?.picture}
                  />
                  <div>
                    <strong>{displayName}</strong>
                    <small>{shortNpub(userPubkey)}</small>
                  </div>
                </div>
                <MenuSeparator />
                <MenuItem
                  icon={<Copy size={15} />}
                  label="Copy public key"
                  onClick={() => {
                    void navigator.clipboard.writeText(auth.npub ?? userPubkey);
                    toast("Public key copied.", "success");
                    close();
                  }}
                />
                {auth.account?.ncryptsec && (
                  <MenuItem
                    icon={<Download size={15} />}
                    label="Download recovery file"
                    onClick={() => {
                      downloadText(
                        `formstr-sheets-recovery-${(auth.npub ?? userPubkey).slice(0, 12)}.txt`,
                        [
                          "Formstr Sheets — account recovery",
                          "",
                          `npub: ${auth.npub ?? userPubkey}`,
                          `ncryptsec: ${auth.account?.ncryptsec}`,
                          "",
                          "Keep this file offline. You need it and your passphrase to sign in on another device.",
                        ].join("\n"),
                      );
                      toast("Recovery file downloaded.", "success");
                      close();
                    }}
                  />
                )}
                <MenuItem
                  icon={<RefreshCw size={15} />}
                  label="Refresh library"
                  onClick={() => {
                    void library.refresh();
                    close();
                  }}
                />
                <MenuItem
                  icon={theme === "light" ? <Moon size={15} /> : <Sun size={15} />}
                  label={theme === "light" ? "Dark mode" : "Light mode"}
                  onClick={() => {
                    toggleTheme();
                    close();
                  }}
                />
                <MenuSeparator />
                <MenuItem
                  icon={<LogOut size={15} />}
                  label="Sign out"
                  danger
                  onClick={() => {
                    close();
                    void auth.logout();
                  }}
                />
              </>
            )}
          </Menu>
        </div>
      </aside>

      <main className="dashboard">
        <header className="dashboard-header">
          <div>
            <span className="eyebrow">Your library</span>
            <h1>Sheets</h1>
            <p>Private grids, encrypted on the relays you choose.</p>
          </div>
          <div className="dashboard-actions">
            <label className="search-field">
              <Search size={16} />
              <input
                ref={searchRef}
                value={search}
                placeholder="Search sheets"
                onChange={(event) => setSearch(event.target.value)}
              />
              <kbd>/</kbd>
            </label>
            <Menu
              width={170}
              trigger={
                <button type="button" className="btn btn-outline">
                  {sort === "recent" ? "Recent" : "Name"} <ChevronDown size={14} />
                </button>
              }
            >
              {(close) => (
                <>
                  <MenuItem
                    label="Recent"
                    active={sort === "recent"}
                    onClick={() => {
                      setSort("recent");
                      close();
                    }}
                  />
                  <MenuItem
                    label="Name"
                    active={sort === "name"}
                    onClick={() => {
                      setSort("name");
                      close();
                    }}
                  />
                </>
              )}
            </Menu>
            <button type="button" className="btn btn-primary" onClick={() => setNewOpen(true)}>
              <Plus size={16} /> New
            </button>
          </div>
        </header>

        {library.error && (
          <div className="banner banner-warning">
            <LockKeyhole size={16} />
            <span>{library.error}</span>
            <button type="button" className="btn btn-ghost btn-sm" onClick={() => void library.refresh()}>
              <RefreshCw size={14} /> Retry
            </button>
          </div>
        )}

        {library.loading && library.entries.length === 0 ? (
          <div className="sheet-grid">
            {Array.from({ length: 6 }, (_, index) => (
              <div className="sheet-card sheet-card-skeleton" key={index}>
                <div className="skeleton skeleton-preview" />
                <div className="skeleton skeleton-line" />
                <div className="skeleton skeleton-line skeleton-line-short" />
              </div>
            ))}
          </div>
        ) : filtered.length === 0 ? (
          <div className="empty-state">
            <div className="empty-art">
              <div className="empty-art-grid">
                {Array.from({ length: 12 }, (_, index) => (
                  <i key={index} className={index === 5 ? "is-accent" : ""} />
                ))}
              </div>
              <div className="empty-art-lock">
                <LockKeyhole size={18} />
              </div>
            </div>
            <h2>
              {search
                ? "No sheets match that search"
                : filter === "shared"
                  ? "Nothing shared with you yet"
                  : filter === "owned"
                    ? "No private sheets yet"
                    : "Your first sheet is a blank canvas"}
            </h2>
            <p>
              {search
                ? "Try a different title or clear the search."
                : "Start from scratch or borrow a template. Every cell is encrypted before it touches a relay."}
            </p>
            <div className="empty-actions">
              <button type="button" className="btn btn-primary" onClick={() => setNewOpen(true)}>
                <FileSpreadsheet size={16} /> Create a sheet
              </button>
              <button
                type="button"
                className="btn btn-outline"
                onClick={() => {
                  setFilter("all");
                  library.refresh();
                }}
              >
                <RefreshCw size={15} /> Refresh
              </button>
            </div>
          </div>
        ) : (
          <div className="sheet-grid">
            {filtered.map((entry) => {
              const ownerProfile = entry.ownerPubkey ? profiles.profiles.get(entry.ownerPubkey) : null;
              return (
                <article
                  key={entry.coordinate}
                  className="sheet-card"
                  tabIndex={0}
                  role="button"
                  onClick={() => openEntry(entry)}
                  onKeyDown={(event) => {
                    if (event.key === "Enter") openEntry(entry);
                  }}
                >
                  <SheetPreview color={entry.color || "#FF5A36"} seed={entry.coordinate} />
                  <div className="sheet-card-body">
                    <div className="sheet-card-title-row">
                      <div>
                        <h3>{entry.title}</h3>
                        <p>{ownerLabel(entry)}</p>
                      </div>
                      <div onClick={(event) => event.stopPropagation()}>
                        <Menu
                          width={200}
                          trigger={
                            <button type="button" className="icon-btn" aria-label="Sheet actions">
                              <MoreHorizontal size={17} />
                            </button>
                          }
                        >
                          {(close) => (
                            <>
                              <MenuItem
                                icon={<FileSpreadsheet size={15} />}
                                label="Open"
                                onClick={() => {
                                  close();
                                  openEntry(entry);
                                }}
                              />
                              <MenuItem
                                icon={<Copy size={15} />}
                                label="Copy view link"
                                onClick={() => {
                                  close();
                                  void copyLink(entry, "view");
                                }}
                              />
                              {entry.editKey && (
                                <MenuItem
                                  icon={<KeyRound size={15} />}
                                  label="Copy edit link"
                                  onClick={() => {
                                    close();
                                    void copyLink(entry, "edit");
                                  }}
                                />
                              )}
                              <MenuSeparator />
                              <MenuItem
                                icon={<Trash2 size={15} />}
                                label={entry.role === "owner" ? "Delete sheet" : "Remove from my sheets"}
                                danger
                                onClick={() => {
                                  close();
                                  setDeleteTarget(entry);
                                }}
                              />
                            </>
                          )}
                        </Menu>
                      </div>
                    </div>
                    <div className="sheet-card-meta">
                      <span className={`role-chip role-${entry.role}`}>
                        {entry.role === "owner" ? "Owner" : entry.role === "editor" ? "Editor" : "Viewer"}
                      </span>
                      <span>{formatRelativeTime(entry.lastOpenedAt ?? entry.updatedAt)}</span>
                      {entry.role !== "owner" && ownerProfile?.picture && (
                        <Avatar size={18} picture={ownerProfile.picture} name={ownerProfile.displayName} />
                      )}
                    </div>
                  </div>
                </article>
              );
            })}
            <button type="button" className="sheet-card sheet-card-new" onClick={() => setNewOpen(true)}>
              <span className="sheet-new-icon">
                <Plus size={20} />
              </span>
              <span>New spreadsheet</span>
            </button>
          </div>
        )}

        <footer className="dashboard-footer">
          <LogoMark size={18} />
          <span>Encrypted with NIP-44 · Keys stay on your device</span>
          <button type="button" className="link-btn" onClick={toggleTheme}>
            {theme === "light" ? "Switch to dark" : "Switch to light"}
          </button>
        </footer>
      </main>

      <NewSheetDialog open={newOpen} onClose={() => setNewOpen(false)} />

      <Dialog
        open={!!deleteTarget}
        onClose={() => (deleting ? undefined : setDeleteTarget(null))}
        title={deleteTarget?.role === "owner" ? "Delete this spreadsheet?" : "Remove from your sheets?"}
        subtitle={
          deleteTarget?.role === "owner"
            ? "A deletion request is published to relays. People who already opened it may still hold copies."
            : "This only removes it from your library. The owner's access grants are unchanged."
        }
        size="sm"
        footer={
          <div className="dialog-footer-actions dialog-footer-end">
            <button type="button" className="btn btn-ghost" disabled={deleting} onClick={() => setDeleteTarget(null)}>
              Cancel
            </button>
            <button type="button" className="btn btn-danger" disabled={deleting} onClick={confirmDelete}>
              {deleting ? <Loader2 size={15} className="spin" /> : <Trash2 size={15} />}
              {deleteTarget?.role === "owner" ? "Delete" : "Remove"}
            </button>
          </div>
        }
      >
        <p className="dialog-copy">
          {deleteTarget?.title
            ? `“${deleteTarget.title}”${deleteTarget.role === "owner" ? " will be removed from your library." : " will no longer appear in Shared with me."}`
            : ""}
        </p>
      </Dialog>
    </div>
  );
}
