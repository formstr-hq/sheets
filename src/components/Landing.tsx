import { ArrowRight, LockKeyhole, Sparkles, Users, Zap } from "lucide-react";
import { Logo } from "./Logo";

export function Landing({ onOpen }: { onOpen: () => void }) {
  return (
    <div className="landing">
      <header className="landing-nav">
        <Logo />
        <div className="landing-nav-actions">
          <span className="landing-nav-note">Private sheets on Nostr</span>
          <button type="button" className="btn btn-ghost" onClick={onOpen}>
            Sign in
          </button>
          <button type="button" className="btn btn-primary" onClick={onOpen}>
            Open Sheets <ArrowRight size={16} />
          </button>
        </div>
      </header>

      <main>
        <section className="hero">
          <div className="hero-copy">
            <div className="eyebrow">
              <Sparkles size={14} />
              Sheets, without the platform
            </div>
            <h1>
              Spreadsheets
              <br />
              that belong
              <br />
              to <em>you</em>.
            </h1>
            <p className="hero-lede">
              Formula-rich grids, dropdowns, pills and collaboration — saved as
              end-to-end encrypted Nostr events. Share a sheet with a person,
              not a company.
            </p>
            <div className="hero-actions">
              <button type="button" className="btn btn-primary btn-lg" onClick={onOpen}>
                Create your first sheet <ArrowRight size={17} />
              </button>
              <button type="button" className="btn btn-outline btn-lg" onClick={onOpen}>
                I already have a key
              </button>
            </div>
            <div className="hero-trust">
              <span>
                <LockKeyhole size={14} /> NIP-44 encrypted
              </span>
              <span>
                <Zap size={14} /> Relay agnostic
              </span>
              <span>
                <Users size={14} /> User-granted access
              </span>
            </div>
          </div>

          <div className="hero-visual" aria-hidden="true">
            <div className="mock-window">
              <div className="mock-topbar">
                <div className="mock-dots">
                  <i />
                  <i />
                  <i />
                </div>
                <div className="mock-title">Q4 launch plan</div>
                <div className="mock-avatars">
                  <span>R</span>
                  <span>A</span>
                  <span className="mock-avatar-more">+2</span>
                </div>
              </div>
              <div className="mock-formula">
                <span className="mock-fx">fx</span>
                <code>=SUMIF(C3:C8,"Done",D3:D8)</code>
                <span className="mock-saved">Encrypted</span>
              </div>
              <div className="mock-grid">
                <div className="mock-row mock-head">
                  <span />
                  <span>Task</span>
                  <span>Owner</span>
                  <span>Status</span>
                  <span>Budget</span>
                </div>
                {[
                  ["Launch page", "Ravi", "Done", "$4,200"],
                  ["Invite flow", "Ama", "In progress", "$2,900"],
                  ["Beta emails", "Ravi", "Review", "$800"],
                  ["Docs refresh", "Len", "Done", "$1,400"],
                  ["Launch event", "Ama", "Blocked", "$6,100"],
                ].map(([task, owner, status, budget]) => (
                  <div className="mock-row" key={task}>
                    <span className="mock-cell-check" />
                    <span>{task}</span>
                    <span className="mock-person">
                      <i>{owner[0]}</i>
                      {owner}
                    </span>
                    <span>
                      <b
                        className={
                          status === "Done"
                            ? "pill pill-green"
                            : status === "In progress"
                              ? "pill pill-amber"
                              : status === "Blocked"
                                ? "pill pill-red"
                                : "pill pill-blue"
                        }
                      >
                        {status}
                      </b>
                    </span>
                    <span className="mock-number">{budget}</span>
                  </div>
                ))}
              </div>
              <div className="mock-tabs">
                <span className="is-active">Plan</span>
                <span>Budget</span>
                <span>Notes</span>
                <span className="mock-add">+</span>
              </div>
            </div>
            <div className="mock-sticker mock-sticker-lock">
              <LockKeyhole size={16} />
              <div>
                <strong>Only key holders</strong>
                <span>Relays see ciphertext</span>
              </div>
            </div>
            <div className="mock-sticker mock-sticker-share">
              <div className="mock-avatar-stack">
                <span>R</span>
                <span>A</span>
              </div>
              <div>
                <strong>Shared with 2 people</strong>
                <span>Viewer · Editor</span>
              </div>
            </div>
          </div>
        </section>

        <section className="feature-strip">
          <article className="feature">
            <div className="feature-icon">
              <LockKeyhole size={19} />
            </div>
            <h3>Encrypted before it leaves</h3>
            <p>
              Every cell is encrypted with a random sheet key. Relays store
              ciphertext, never your formulas or data.
            </p>
          </article>
          <article className="feature">
            <div className="feature-icon">
              <Users size={19} />
            </div>
            <h3>Access you grant, person by person</h3>
            <p>
              Invite collaborators by npub. Their key unlocks the sheet; remove
              it and the grant disappears from the network.
            </p>
          </article>
          <article className="feature">
            <div className="feature-icon">
              <Sparkles size={19} />
            </div>
            <h3>A real spreadsheet</h3>
            <p>
              Formulas, dropdowns, comments, conditional formatting, frozen rows,
              tag pills and Excel import/export.
            </p>
          </article>
        </section>

        <section className="how-strip">
          <div>
            <span className="eyebrow">Under the hood</span>
            <h2>Private by construction</h2>
          </div>
          <ol className="how-steps">
            <li>
              <span>01</span>
              <div>
                <strong>Create a sheet</strong>
                <p>We generate a fresh view key and edit key on your device.</p>
              </div>
            </li>
            <li>
              <span>02</span>
              <div>
                <strong>Save encrypted events</strong>
                <p>The sheet is NIP-44 encrypted and published to your relays.</p>
              </div>
            </li>
            <li>
              <span>03</span>
              <div>
                <strong>Grant access</strong>
                <p>Keys are encrypted to each collaborator's Nostr public key.</p>
              </div>
            </li>
            <li>
              <span>04</span>
              <div>
                <strong>Edit together</strong>
                <p>Editors sign new snapshots with the sheet's edit key.</p>
              </div>
            </li>
          </ol>
        </section>
      </main>

      <footer className="landing-footer">
        <Logo compact />
        <p>
          A Formstr project. Built on Nostr. No accounts, no analytics, no
          tracking.
        </p>
      </footer>
    </div>
  );
}
