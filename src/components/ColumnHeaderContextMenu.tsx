import { useCallback, useEffect, useLayoutEffect, useRef } from "react";
import { useClampedMenuPosition } from "./useClampedMenuPosition";

const MENU_ITEM_STYLE = { minHeight: 40, minWidth: 40 } as const;

export type ColumnHeaderContextMenuProps = Readonly<{
  label: string;
  x: number;
  y: number;
  opener: HTMLElement;
  onClose(): void;
  onInsertLeft(): void;
  onInsertRight(): void;
  onDelete(): void;
}>;

export function ColumnHeaderContextMenu({
  label,
  x,
  y,
  opener,
  onClose,
  onInsertLeft,
  onInsertRight,
  onDelete
}: ColumnHeaderContextMenuProps) {
  const { menuRef, position } = useClampedMenuPosition(x, y);
  const firstItemRef = useRef<HTMLButtonElement>(null);
  const closedRef = useRef(false);

  useLayoutEffect(() => {
    firstItemRef.current?.focus({ preventScroll: true });
  }, []);

  const closeMenu = useCallback((restoreFocus = true) => {
    if (closedRef.current) return;
    closedRef.current = true;
    onClose();
    if (restoreFocus && opener.isConnected) {
      opener.focus({ preventScroll: true });
    }
  }, [onClose, opener]);

  useEffect(() => {
    function closeOnOutsidePointer(event: PointerEvent) {
      if (!menuRef.current?.contains(event.target as Node)) {
        closeMenu();
      }
    }

    window.addEventListener("pointerdown", closeOnOutsidePointer);
    return () => {
      window.removeEventListener("pointerdown", closeOnOutsidePointer);
    };
  }, [closeMenu]);

  function run(action: () => void) {
    action();
    closeMenu();
  }

  return (
    <div
      ref={menuRef}
      className="cell-context-menu"
      role="menu"
      aria-label={`Column ${label} context menu`}
      style={{ left: position.left, top: position.top }}
      onPointerDown={(event) => event.stopPropagation()}
      onContextMenu={(event) => event.preventDefault()}
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          event.stopPropagation();
          closeMenu();
        }
      }}
      onBlur={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
          closeMenu(false);
        }
      }}
    >
      <div className="cell-context-menu-heading">Column {label}</div>
      <button
        ref={firstItemRef}
        type="button"
        role="menuitem"
        style={MENU_ITEM_STYLE}
        onClick={() => run(onInsertLeft)}
      >
        Insert column left
      </button>
      <button type="button" role="menuitem" style={MENU_ITEM_STYLE} onClick={() => run(onInsertRight)}>
        Insert column right
      </button>
      <button type="button" role="menuitem" style={MENU_ITEM_STYLE} onClick={() => run(onDelete)}>
        Delete column
      </button>
    </div>
  );
}
