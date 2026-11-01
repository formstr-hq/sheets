import type { Cell } from "@fortune-sheet/core";
import type { Profile } from "./types";
import { colorFromString, initials } from "./utils";
import { cellMeta } from "./sheetData";

interface CellInfo {
  startX: number;
  startY: number;
  endX: number;
  endY: number;
}

function roundedRect(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  radius: number,
): void {
  const r = Math.min(radius, width / 2, height / 2);
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + width, y, x + width, y + height, r);
  ctx.arcTo(x + width, y + height, x, y + height, r);
  ctx.arcTo(x, y + height, x, y, r);
  ctx.arcTo(x, y, x + width, y, r);
  ctx.closePath();
}

function fitText(
  ctx: CanvasRenderingContext2D,
  text: string,
  maxWidth: number,
): string {
  if (ctx.measureText(text).width <= maxWidth) return text;
  let value = text;
  while (value.length > 1 && ctx.measureText(`${value}…`).width > maxWidth) {
    value = value.slice(0, -1);
  }
  return `${value}…`;
}

export function drawCellPill(
  cell: Cell | null | undefined,
  info: CellInfo,
  ctx: CanvasRenderingContext2D,
  profileFor: (pubkey: string) => Profile | null,
): void {
  const meta = cellMeta(cell);
  if (!meta) return;
  const profile = meta.fs === "user" && meta.pubkey ? profileFor(meta.pubkey) : null;
  const raw =
    meta.fs === "user"
      ? profile?.displayName || profile?.name || meta.name || String(cell?.m ?? cell?.v ?? "")
      : String(cell?.m ?? cell?.v ?? "");
  const label = raw.trim();
  if (!label) return;

  const { startX, startY, endX, endY } = info;
  const width = endX - startX - 2;
  const height = endY - startY - 2;
  if (width < 10 || height < 8) return;

  ctx.save();
  ctx.beginPath();
  ctx.rect(startX + 1, startY + 1, width, height);
  ctx.clip();

  ctx.fillStyle = cell?.bg || "#FFFFFF";
  ctx.fillRect(startX + 1, startY + 1, width, height);

  const colors = colorFromString(label);
  const font = "500 11px Inter, 'Segoe UI', sans-serif";
  const paddingX = 7;
  const avatarSize = meta.fs === "user" ? 14 : 0;
  const avatarGap = meta.fs === "user" ? 5 : 0;
  ctx.font = font;
  const maxTextWidth = Math.max(8, width - paddingX * 2 - avatarSize - avatarGap - 4);
  const text = fitText(ctx, label, maxTextWidth);
  const textWidth = ctx.measureText(text).width;
  const pillHeight = Math.min(height - 4, 20);
  const pillWidth = Math.min(
    width - 4,
    Math.ceil(textWidth + paddingX * 2 + avatarSize + avatarGap),
  );

  let x = startX + 2;
  if (cell?.ht === 0) x = startX + 2 + (width - pillWidth) / 2;
  else if (cell?.ht === 2) x = endX - 2 - pillWidth;
  const y = startY + 2 + (height - pillHeight) / 2;

  roundedRect(ctx, x, y, pillWidth, pillHeight, pillHeight / 2);
  ctx.fillStyle = colors.bg;
  ctx.fill();
  ctx.strokeStyle = `${colors.fg}22`;
  ctx.lineWidth = 1;
  ctx.stroke();

  let textX = x + paddingX;
  ctx.textBaseline = "middle";
  if (meta.fs === "user") {
    const centerX = textX + avatarSize / 2;
    const centerY = y + pillHeight / 2;
    ctx.beginPath();
    ctx.arc(centerX, centerY, avatarSize / 2, 0, Math.PI * 2);
    ctx.fillStyle = colors.fg;
    ctx.fill();
    ctx.fillStyle = colors.bg;
    ctx.font = "700 7px Inter, 'Segoe UI', sans-serif";
    ctx.textAlign = "center";
    ctx.fillText(initials(label).slice(0, 1), centerX, centerY + 0.4);
    ctx.textAlign = "left";
    textX += avatarSize + avatarGap;
  }

  ctx.fillStyle = colors.fg;
  ctx.font = font;
  ctx.fillText(text, textX, y + pillHeight / 2 + 0.4);
  ctx.restore();
}
