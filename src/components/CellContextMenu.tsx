import type { ReactElement } from "react";
import { useClampedMenuPosition } from "./useClampedMenuPosition";

type CellContextMenuProps = {
  address: string;
  x: number;
  y: number;
  canPasteSpecial: boolean;
  isWrapped: boolean;
  onClose: () => void;
  onCut: () => void;
  onCopy: () => void;
  onPaste: () => void;
  onPasteValues: () => void;
  onPasteFormats: () => void;
  onClearAll: () => void;
  onClearContents: () => void;
  onClearFormats: () => void;
  onClearConditionalFormats: () => void;
  onClearHyperlinks: () => void;
  onClearValidation: () => void;
  onClearComments: () => void;
  onToggleWrapText: () => void;
  onInsertRow: () => void;
  onDeleteRow: () => void;
  onInsertColumnLeft: () => void;
  onInsertColumnRight: () => void;
  onDeleteColumn: () => void;
  onComment: () => void;
  onLink: () => void;
};

export function CellContextMenu({
  address,
  x,
  y,
  canPasteSpecial,
  isWrapped,
  onClose,
  onCut,
  onCopy,
  onPaste,
  onPasteValues,
  onPasteFormats,
  onClearAll,
  onClearContents,
  onClearFormats,
  onClearConditionalFormats,
  onClearHyperlinks,
  onClearValidation,
  onClearComments,
  onToggleWrapText,
  onInsertRow,
  onDeleteRow,
  onInsertColumnLeft,
  onInsertColumnRight,
  onDeleteColumn,
  onComment,
  onLink
}: CellContextMenuProps): ReactElement {
  const { menuRef, position } = useClampedMenuPosition(x, y);

  function run(action: () => void) {
    action();
    onClose();
  }

  return (
    <div
      ref={menuRef}
      className="cell-context-menu"
      role="menu"
      aria-label="Cell context menu"
      style={{ left: position.left, top: position.top }}
      onPointerDown={(event) => event.stopPropagation()}
      onContextMenu={(event) => event.preventDefault()}
    >
      <div className="cell-context-menu-heading">{address}</div>
      <button type="button" role="menuitem" onClick={() => run(onCut)}>
        Cut
      </button>
      <button type="button" role="menuitem" onClick={() => run(onCopy)}>
        Copy
      </button>
      <button type="button" role="menuitem" disabled={!canPasteSpecial} onClick={() => run(onPaste)}>
        Paste
      </button>
      <button type="button" role="menuitem" disabled={!canPasteSpecial} onClick={() => run(onPasteValues)}>
        Paste values
      </button>
      <button type="button" role="menuitem" disabled={!canPasteSpecial} onClick={() => run(onPasteFormats)}>
        Paste formats
      </button>
      <span className="cell-context-menu-divider" aria-hidden="true" />
      <button type="button" role="menuitem" onClick={() => run(onClearAll)}>
        Clear all
      </button>
      <button type="button" role="menuitem" onClick={() => run(onClearContents)}>
        Clear contents
      </button>
      <button type="button" role="menuitem" onClick={() => run(onClearFormats)}>
        Clear formats
      </button>
      <button type="button" role="menuitem" onClick={() => run(onClearConditionalFormats)}>
        Clear conditional formats
      </button>
      <button type="button" role="menuitem" onClick={() => run(onClearHyperlinks)}>
        Clear hyperlinks
      </button>
      <button type="button" role="menuitem" onClick={() => run(onClearValidation)}>
        Clear validation
      </button>
      <button type="button" role="menuitem" onClick={() => run(onToggleWrapText)}>
        {isWrapped ? "Unwrap text" : "Wrap text"}
      </button>
      <span className="cell-context-menu-divider" aria-hidden="true" />
      <button type="button" role="menuitem" onClick={() => run(onInsertRow)}>
        Insert row above
      </button>
      <button type="button" role="menuitem" onClick={() => run(onDeleteRow)}>
        Delete row
      </button>
      <button type="button" role="menuitem" onClick={() => run(onInsertColumnLeft)}>
        Insert column left
      </button>
      <button type="button" role="menuitem" onClick={() => run(onInsertColumnRight)}>
        Insert column right
      </button>
      <button type="button" role="menuitem" onClick={() => run(onDeleteColumn)}>
        Delete column
      </button>
      <span className="cell-context-menu-divider" aria-hidden="true" />
      <button type="button" role="menuitem" onClick={() => run(onComment)}>
        Comment
      </button>
      <button type="button" role="menuitem" onClick={() => run(onClearComments)}>
        Clear comments
      </button>
      <button type="button" role="menuitem" onClick={() => run(onLink)}>
        Link
      </button>
    </div>
  );
}
