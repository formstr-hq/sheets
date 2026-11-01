import { cx } from "../lib/utils";

export function LogoMark({ size = 34, className }: { size?: number; className?: string }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 64 64"
      role="img"
      aria-label="Formstr Sheets"
      className={className}
    >
      <rect width="64" height="64" rx="15" fill="#141719" />
      <rect x="11" y="13" width="42" height="38" rx="7" fill="#FAF8F5" />
      <path
        d="M11 24h42M11 35h42M25 13v38M39 13v38"
        stroke="#141719"
        strokeOpacity=".14"
        strokeWidth="2"
      />
      <path d="M39 13h7a7 7 0 0 1 7 7v4H39z" fill="#FF5A36" />
      <path
        d="M19.5 29.2v5.6M16.7 32h5.6M17.7 30l3.9 3.9M21.6 30l-3.9 3.9"
        stroke="#141719"
        strokeWidth="1.7"
        strokeLinecap="round"
      />
    </svg>
  );
}

export function Logo({ compact = false }: { compact?: boolean }) {
  return (
    <div className={cx("logo", compact && "logo-compact")}>
      <LogoMark size={compact ? 28 : 34} />
      <div className="logo-type">
        <span className="logo-word">Formstr</span>
        <span className="logo-pill">Sheets</span>
      </div>
    </div>
  );
}
