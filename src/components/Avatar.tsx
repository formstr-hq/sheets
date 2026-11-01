import { useState } from "react";
import { colorFromString, cx, initials } from "../lib/utils";

interface AvatarProps {
  name?: string;
  picture?: string;
  size?: number;
  className?: string;
  ring?: boolean;
}

export function Avatar({ name, picture, size = 30, className, ring }: AvatarProps) {
  const [broken, setBroken] = useState(false);
  const label = name?.trim() || "?";
  const colors = colorFromString(label);
  return (
    <span
      className={cx("avatar", ring && "avatar-ring", className)}
      style={{
        width: size,
        height: size,
        background: picture && !broken ? undefined : `linear-gradient(135deg, ${colors.bg}, ${colors.fg}22)`,
        color: colors.fg,
        fontSize: Math.max(9, Math.round(size * 0.36)),
      }}
      title={label}
    >
      {picture && !broken ? (
        <img src={picture} alt="" onError={() => setBroken(true)} referrerPolicy="no-referrer" />
      ) : (
        initials(label)
      )}
    </span>
  );
}
