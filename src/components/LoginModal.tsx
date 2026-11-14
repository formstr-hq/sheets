import { useEffect, useState } from "react";
import {
  AlertCircle,
  Check,
  Globe,
  Copy,
  Download,
  Eye,
  EyeOff,
  KeyRound,
  Link2,
  Loader2,
  ShieldCheck,
} from "lucide-react";
import { Dialog } from "./Dialog";
import { LogoMark } from "./Logo";
import { useAuth } from "../providers/AuthProvider";
import { useToast } from "../providers/ToastProvider";
import { cx, downloadText, shortNpub } from "../lib/utils";
import { hasNostrExtension, watchNostrExtension } from "../lib/extension";

type Mode = "extension" | "create" | "existing" | "bunker" | "backup";

export function LoginModal({ open, onClose }: { open: boolean; onClose: () => void }) {
  const auth = useAuth();
  const { toast } = useToast();
  const [mode, setMode] = useState<Mode>(() =>
    hasNostrExtension()
      ? "extension"
      : "create",
  );
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [passphrase, setPassphrase] = useState("");
  const [confirmPassphrase, setConfirmPassphrase] = useState("");
  const [nsec, setNsec] = useState("");
  const [bunkerUri, setBunkerUri] = useState("");
  const [showSecret, setShowSecret] = useState(false);
  const [backup, setBackup] = useState<{ npub: string; ncryptsec: string } | null>(null);
  const [copied, setCopied] = useState("");

  const [hasExtension, setHasExtension] = useState(hasNostrExtension);

  useEffect(() => {
    if (!open) return;
    return watchNostrExtension(setHasExtension);
  }, [open]);

  const close = () => {
    if (mode === "backup") return;
    setError("");
    onClose();
  };

  async function run(fn: () => Promise<void>) {
    setBusy(true);
    setError("");
    try {
      await fn();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Something went wrong.");
    } finally {
      setBusy(false);
    }
  }

  const loginExtension = () =>
    run(async () => {
      if (!hasNostrExtension()) {
        throw new Error("Your extension isn't available on this page. Enable its access to this site, then retry. If needed, reload the page after enabling it.");
      }
      await auth.loginWithExtension();
      toast("Signed in with your browser extension.", "success");
      onClose();
    });

  const createKey = () =>
    run(async () => {
      if (passphrase.length < 8) throw new Error("Use at least 8 characters for your passphrase.");
      if (passphrase !== confirmPassphrase) throw new Error("Passphrases don't match.");
      const result = await auth.createAccount(passphrase);
      setBackup(result);
      setMode("backup");
      setPassphrase("");
      setConfirmPassphrase("");
    });

  const importKey = () =>
    run(async () => {
      if (passphrase.length < 8) throw new Error("Use at least 8 characters for your passphrase.");
      await auth.importNsec(nsec, passphrase);
      toast("Signed in with your existing key.", "success");
      onClose();
    });

  const loginBunker = () =>
    run(async () => {
      if (!bunkerUri.trim().startsWith("bunker://")) {
        throw new Error("Enter a bunker:// connection string.");
      }
      await auth.loginWithBunker(bunkerUri.trim());
      toast("Connected to your remote signer.", "success");
      onClose();
    });

  const copy = async (value: string, key: string) => {
    try {
      await navigator.clipboard.writeText(value);
      setCopied(key);
      toast("Copied to clipboard.", "success");
      window.setTimeout(() => setCopied(""), 1600);
    } catch {
      toast("Couldn't copy. Select the text and copy it manually.", "error");
    }
  };

  const downloadRecovery = () => {
    if (!backup) return;
    downloadText(
      `formstr-sheets-recovery-${backup.npub.slice(0, 12)}.txt`,
      [
        "Formstr Sheets — account recovery",
        "",
        `npub: ${backup.npub}`,
        `ncryptsec: ${backup.ncryptsec}`,
        "",
        "This file contains your encrypted private key. Keep it offline and safe.",
        "You will need this ncryptsec (and your passphrase) to sign in on another device.",
        "Never share it with anyone. Formstr cannot recover it for you.",
      ].join("\n"),
    );
  };

  const closeLabel = mode === "backup" ? "I saved my recovery key" : null;

  return (
    <Dialog
      open={open}
      onClose={close}
      title={mode === "backup" ? "Save your recovery key" : "Open Formstr Sheets"}
      subtitle={
        mode === "backup"
          ? "This is the only way to get back into your account on a new device."
          : "Sign in with Nostr. Your keys stay on your device."
      }
      size="md"
      closeOnBackdrop={mode !== "backup"}
      hideClose={mode === "backup"}
      className={mode === "backup" ? "login-dialog login-backup" : "login-dialog"}
      footer={
        mode === "backup" ? (
          <div className="dialog-footer-row">
            <button type="button" className="btn btn-ghost" onClick={downloadRecovery}>
              <Download size={15} /> Download file
            </button>
            <button
              type="button"
              className="btn btn-primary"
              onClick={() => {
                setBackup(null);
                onClose();
              }}
            >
              {closeLabel} <Check size={16} />
            </button>
          </div>
        ) : undefined
      }
    >
      {mode === "backup" && backup ? (
        <div className="backup-panel">
          <div className="backup-hero">
            <div className="backup-icon">
              <ShieldCheck size={26} />
            </div>
            <div>
              <strong>Your account is ready</strong>
              <p>
                Formstr can't reset or recover this. Save it somewhere safe before
                you continue.
              </p>
            </div>
          </div>
          <label className="field">
            <span className="field-label">Public key</span>
            <div className="copy-field">
              <code>{backup.npub}</code>
              <button type="button" className="icon-btn" onClick={() => copy(backup.npub, "npub")}>
                {copied === "npub" ? <Check size={15} /> : <Copy size={15} />}
              </button>
            </div>
          </label>
          <label className="field">
            <span className="field-label">Encrypted recovery key (ncryptsec)</span>
            <div className="copy-field">
              <code className="mono-wrap">{backup.ncryptsec}</code>
              <button
                type="button"
                className="icon-btn"
                onClick={() => copy(backup.ncryptsec, "ncryptsec")}
              >
                {copied === "ncryptsec" ? <Check size={15} /> : <Copy size={15} />}
              </button>
            </div>
          </label>
          <div className="callout callout-warning">
            <AlertCircle size={16} />
            <p>
              Your passphrase is not saved anywhere. If you lose both the recovery
              key and the passphrase, this account is gone.
            </p>
          </div>
        </div>
      ) : (
        <div className="login-panel">
          <button
            type="button"
            className={cx("login-method", mode === "extension" && "is-active")}
            onClick={() => setMode("extension")}
          >
            <span className="login-method-icon">
              <Globe size={19} />
            </span>
            <span>
              <strong>Browser extension</strong>
              <small>NIP-07 · Alby, nos2x, Keys</small>
            </span>
          </button>

          {mode === "extension" && (
            <div className="login-form">
              {hasExtension ? (
                <p className="form-note">
                  Your extension will ask for permission to sign and encrypt on
                  your behalf.
                </p>
              ) : (
                <div className="callout callout-info">
                  <AlertCircle size={16} />
                  <p>
                    No NIP-07 extension found. Install one, or use a local key
                    below.
                  </p>
                </div>
              )}
              <button
                type="button"
                className="btn btn-primary btn-full"
                disabled={busy}
                onClick={loginExtension}
              >
                {busy ? <Loader2 size={16} className="spin" /> : <Globe size={16} />}
                {busy ? "Connecting…" : "Connect extension"}
              </button>
            </div>
          )}

          <div className="login-divider">
            <span>or use a local key</span>
          </div>

          <div className="login-tabs" role="tablist">
            <button
              type="button"
              role="tab"
              aria-selected={mode === "create"}
              className={cx("login-tab", mode === "create" && "is-active")}
              onClick={() => setMode("create")}
            >
              <KeyRound size={14} /> Create
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={mode === "existing"}
              className={cx("login-tab", mode === "existing" && "is-active")}
              onClick={() => setMode("existing")}
            >
              <Download size={14} /> Import
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={mode === "bunker"}
              className={cx("login-tab", mode === "bunker" && "is-active")}
              onClick={() => setMode("bunker")}
            >
              <Link2 size={14} /> Bunker
            </button>
          </div>

          {mode === "create" && (
            <div className="login-form">
              <p className="form-note">
                Generates a new Nostr identity and stores it encrypted on this
                device with your passphrase.
              </p>
              <label className="field">
                <span className="field-label">Passphrase</span>
                <div className="input-wrap">
                  <input
                    type={showSecret ? "text" : "password"}
                    value={passphrase}
                    autoComplete="new-password"
                    placeholder="At least 8 characters"
                    onChange={(event) => setPassphrase(event.target.value)}
                  />
                  <button
                    type="button"
                    className="input-affix"
                    onClick={() => setShowSecret((value) => !value)}
                    aria-label={showSecret ? "Hide passphrase" : "Show passphrase"}
                  >
                    {showSecret ? <EyeOff size={16} /> : <Eye size={16} />}
                  </button>
                </div>
              </label>
              <label className="field">
                <span className="field-label">Confirm passphrase</span>
                <input
                  type={showSecret ? "text" : "password"}
                  value={confirmPassphrase}
                  autoComplete="new-password"
                  placeholder="Type it again"
                  onChange={(event) => setConfirmPassphrase(event.target.value)}
                />
              </label>
              <button
                type="button"
                className="btn btn-primary btn-full"
                disabled={busy}
                onClick={createKey}
              >
                {busy ? <Loader2 size={16} className="spin" /> : <KeyRound size={16} />}
                Create my account
              </button>
            </div>
          )}

          {mode === "existing" && (
            <div className="login-form">
              <p className="form-note">
                Import an nsec. It is encrypted with your passphrase and saved on
                this device — never sent anywhere.
              </p>
              <label className="field">
                <span className="field-label">Private key (nsec)</span>
                <div className="input-wrap">
                  <input
                    type={showSecret ? "text" : "password"}
                    value={nsec}
                    autoComplete="off"
                    placeholder="nsec1…"
                    onChange={(event) => setNsec(event.target.value)}
                  />
                  <button
                    type="button"
                    className="input-affix"
                    onClick={() => setShowSecret((value) => !value)}
                    aria-label={showSecret ? "Hide key" : "Show key"}
                  >
                    {showSecret ? <EyeOff size={16} /> : <Eye size={16} />}
                  </button>
                </div>
              </label>
              <label className="field">
                <span className="field-label">Passphrase for this device</span>
                <input
                  type={showSecret ? "text" : "password"}
                  value={passphrase}
                  autoComplete="new-password"
                  placeholder="At least 8 characters"
                  onChange={(event) => setPassphrase(event.target.value)}
                />
              </label>
              <button
                type="button"
                className="btn btn-primary btn-full"
                disabled={busy}
                onClick={importKey}
              >
                {busy ? <Loader2 size={16} className="spin" /> : <Download size={16} />}
                Import key
              </button>
            </div>
          )}

          {mode === "bunker" && (
            <div className="login-form">
              <p className="form-note">
                Paste a bunker:// connection from your remote signer (Amber,
                nsec.app, etc.).
              </p>
              <label className="field">
                <span className="field-label">Bunker URI</span>
                <textarea
                  rows={3}
                  value={bunkerUri}
                  placeholder="bunker://…"
                  onChange={(event) => setBunkerUri(event.target.value)}
                />
              </label>
              <button
                type="button"
                className="btn btn-primary btn-full"
                disabled={busy}
                onClick={loginBunker}
              >
                {busy ? <Loader2 size={16} className="spin" /> : <Link2 size={16} />}
                Connect remote signer
              </button>
            </div>
          )}

          {error && (
            <div className="callout callout-error">
              <AlertCircle size={16} />
              <p>{error}</p>
            </div>
          )}

          <div className="login-security">
            <LogoMark size={22} />
            <p>
              Formstr never sees your private key. Sheets are encrypted with
              NIP-44 and stored on relays you choose.
            </p>
          </div>
        </div>
      )}
    </Dialog>
  );
}

export function UnlockScreen() {
  const auth = useAuth();
  const { toast } = useToast();
  const [passphrase, setPassphrase] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [show, setShow] = useState(false);

  const unlock = async () => {
    setBusy(true);
    setError("");
    try {
      if (auth.account?.method === "ncryptsec") {
        await auth.unlockWithPassphrase(passphrase);
      } else if (auth.account?.method === "nip46") {
        await auth.reconnectRemote();
      } else {
        await auth.loginWithExtension();
      }
      toast("Unlocked.", "success");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't unlock this account.");
    } finally {
      setBusy(false);
    }
  };

  const needsPassphrase = auth.account?.method === "ncryptsec";

  return (
    <div className="unlock-page">
      <div className="unlock-card">
        <LogoMark size={44} />
        <span className="eyebrow">Welcome back</span>
        <h1>Unlock your sheets</h1>
        <p className="unlock-npub">{auth.npub ? shortNpub(auth.npub) : ""}</p>
        <p className="muted">
          {needsPassphrase
            ? "Enter the passphrase for this device. It never leaves your browser."
            : auth.account?.method === "nip46"
              ? "Reconnect to your remote signer to continue."
              : "Your extension will re-authorize this session."}
        </p>
        {needsPassphrase && (
          <label className="field">
            <span className="field-label">Passphrase</span>
            <div className="input-wrap">
              <input
                type={show ? "text" : "password"}
                value={passphrase}
                autoFocus
                autoComplete="current-password"
                placeholder="Your passphrase"
                onChange={(event) => setPassphrase(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") void unlock();
                }}
              />
              <button
                type="button"
                className="input-affix"
                onClick={() => setShow((value) => !value)}
                aria-label={show ? "Hide passphrase" : "Show passphrase"}
              >
                {show ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </div>
          </label>
        )}
        {error && (
          <div className="callout callout-error">
            <AlertCircle size={16} />
            <p>{error}</p>
          </div>
        )}
        <button
          type="button"
          className="btn btn-primary btn-full"
          disabled={busy || (needsPassphrase && !passphrase)}
          onClick={unlock}
        >
          {busy ? <Loader2 size={16} className="spin" /> : <KeyRound size={16} />}
          {needsPassphrase ? "Unlock" : "Reconnect"}
        </button>
        <button type="button" className="btn btn-ghost btn-full" onClick={() => void auth.logout()}>
          Use another account
        </button>
      </div>
    </div>
  );
}
