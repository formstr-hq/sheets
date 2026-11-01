import { useState, type ReactNode } from "react";
import * as DropdownMenu from "@radix-ui/react-dropdown-menu";
import { cx } from "../lib/utils";

interface MenuProps {
  trigger: ReactNode;
  children: (close: () => void) => ReactNode;
  align?: "left" | "right";
  width?: number;
  className?: string;
  menuClassName?: string;
}

export function Menu({
  trigger,
  children,
  align = "right",
  width,
  className,
  menuClassName,
}: MenuProps) {
  const [open, setOpen] = useState(false);

  return (
    <DropdownMenu.Root open={open} onOpenChange={setOpen}>
      <DropdownMenu.Trigger asChild>{trigger}</DropdownMenu.Trigger>
      <DropdownMenu.Portal>
        <DropdownMenu.Content
          className={cx("menu", className, menuClassName)}
          align={align === "right" ? "end" : "start"}
          sideOffset={6}
          collisionPadding={12}
          style={width ? { width } : undefined}
        >
          {children(() => setOpen(false))}
        </DropdownMenu.Content>
      </DropdownMenu.Portal>
    </DropdownMenu.Root>
  );
}

export function MenuItem({
  icon,
  label,
  shortcut,
  onClick,
  danger,
  disabled,
  active,
}: {
  icon?: ReactNode;
  label: ReactNode;
  shortcut?: string;
  onClick?: () => void;
  danger?: boolean;
  disabled?: boolean;
  active?: boolean;
}) {
  return (
    <DropdownMenu.Item
      className={cx("menu-item", danger && "menu-item-danger", active && "is-active")}
      disabled={disabled}
      onSelect={() => onClick?.()}
    >
      {icon && <span className="menu-item-icon">{icon}</span>}
      <span className="menu-item-label">{label}</span>
      {shortcut && <span className="menu-item-shortcut">{shortcut}</span>}
    </DropdownMenu.Item>
  );
}

export function MenuSeparator() {
  return <DropdownMenu.Separator className="menu-separator" />;
}

export function MenuLabel({ children }: { children: ReactNode }) {
  return <DropdownMenu.Label className="menu-label">{children}</DropdownMenu.Label>;
}
