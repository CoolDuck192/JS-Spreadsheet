import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";

const MENU_VIEWPORT_MARGIN = 8;
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
  const menuRef = useRef<HTMLDivElement>(null);
  const firstItemRef = useRef<HTMLButtonElement>(null);
  const [position, setPosition] = useState({ left: x, top: y });

  useLayoutEffect(() => {
    const menu = menuRef.current;
    if (!menu) {
      setPosition({ left: x, top: y });
      return;
    }
    const menuElement = menu;

    function updatePosition() {
      const rect = menuElement.getBoundingClientRect();
      const maxLeft = Math.max(MENU_VIEWPORT_MARGIN, window.innerWidth - rect.width - MENU_VIEWPORT_MARGIN);
      const maxTop = Math.max(MENU_VIEWPORT_MARGIN, window.innerHeight - rect.height - MENU_VIEWPORT_MARGIN);
      const nextPosition = {
        left: Math.min(Math.max(MENU_VIEWPORT_MARGIN, x), maxLeft),
        top: Math.min(Math.max(MENU_VIEWPORT_MARGIN, y), maxTop)
      };
      setPosition((currentPosition) =>
        currentPosition.left === nextPosition.left && currentPosition.top === nextPosition.top
          ? currentPosition
          : nextPosition
      );
    }

    updatePosition();
    window.addEventListener("resize", updatePosition);
    return () => window.removeEventListener("resize", updatePosition);
  }, [x, y]);

  useLayoutEffect(() => {
    firstItemRef.current?.focus({ preventScroll: true });
  }, []);

  const closeAndRestoreFocus = useCallback(() => {
    onClose();
    if (opener.isConnected) {
      opener.focus({ preventScroll: true });
    }
  }, [onClose, opener]);

  useEffect(() => {
    function closeOnOutsidePointer(event: PointerEvent) {
      if (!menuRef.current?.contains(event.target as Node)) {
        closeAndRestoreFocus();
      }
    }

    window.addEventListener("pointerdown", closeOnOutsidePointer);
    return () => {
      window.removeEventListener("pointerdown", closeOnOutsidePointer);
    };
  }, [closeAndRestoreFocus]);

  function run(action: () => void) {
    action();
    closeAndRestoreFocus();
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
          closeAndRestoreFocus();
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
