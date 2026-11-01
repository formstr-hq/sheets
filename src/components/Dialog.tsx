import { useId, type ReactNode } from "react";
import * as DialogPrimitive from "@radix-ui/react-dialog";
import { X } from "lucide-react";
import { cx } from "../lib/utils";

interface DialogProps {
  open: boolean;
  onClose: () => void;
  title?: ReactNode;
  subtitle?: ReactNode;
  children: ReactNode;
  footer?: ReactNode;
  size?: "sm" | "md" | "lg";
  className?: string;
  closeOnBackdrop?: boolean;
  hideHeader?: boolean;
  hideClose?: boolean;
}

export function Dialog({
  open,
  onClose,
  title,
  subtitle,
  children,
  footer,
  size = "md",
  className,
  closeOnBackdrop = true,
  hideHeader,
  hideClose,
}: DialogProps) {
  const titleId = useId();

  return (
    <DialogPrimitive.Root
      open={open}
      onOpenChange={(nextOpen) => {
        if (!nextOpen) onClose();
      }}
    >
      <DialogPrimitive.Portal>
        <DialogPrimitive.Overlay className="dialog-backdrop" />
        <DialogPrimitive.Content
          className={cx("dialog", `dialog-${size}`, className)}
          aria-labelledby={title ? titleId : undefined}
          onPointerDownOutside={(event) => {
            if (!closeOnBackdrop) event.preventDefault();
          }}
          onEscapeKeyDown={(event) => {
            if (!closeOnBackdrop) event.preventDefault();
          }}
        >
          {hideHeader ? (
            <DialogPrimitive.Title className="sr-only">
              {title || "Dialog"}
            </DialogPrimitive.Title>
          ) : (
            <div className="dialog-header">
              <div>
                <DialogPrimitive.Title className="dialog-title" id={titleId}>
                  {title || "Dialog"}
                </DialogPrimitive.Title>
                <DialogPrimitive.Description className={cx("dialog-subtitle", !subtitle && "sr-only")}>
                  {subtitle || "Dialog"}
                </DialogPrimitive.Description>
              </div>
              {!hideClose && (
                <DialogPrimitive.Close asChild>
                  <button type="button" className="icon-btn dialog-close" aria-label="Close">
                    <X size={18} />
                  </button>
                </DialogPrimitive.Close>
              )}
            </div>
          )}
          <div className="dialog-body">{children}</div>
          {footer && <div className="dialog-footer">{footer}</div>}
        </DialogPrimitive.Content>
      </DialogPrimitive.Portal>
    </DialogPrimitive.Root>
  );
}
