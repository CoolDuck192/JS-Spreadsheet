import { useLayoutEffect, useRef, type CSSProperties } from "react";
import type { SheetModel } from "../types";

type SheetTabsProps = {
  sheets: SheetModel[];
  activeSheetId: string;
  onSelect: (sheetId: string) => void;
  onAdd: () => void;
};

export function SheetTabs({ sheets, activeSheetId, onSelect, onAdd }: SheetTabsProps) {
  const visibleSheets = sheets.filter((sheet) => sheet.isHidden !== true);
  const tabRefs = useRef(new Map<string, HTMLButtonElement>());

  useLayoutEffect(() => {
    tabRefs.current.get(activeSheetId)?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
  }, [activeSheetId]);

  return (
    <div className="js-spreadsheet-sheet-tabs" role="tablist" aria-label="Sheet tabs">
      {visibleSheets.map((sheet) => {
        const tabStyle: CSSProperties | undefined = sheet.tabColor
          ? { borderTopColor: sheet.tabColor }
          : undefined;

        return (
          <button
            key={sheet.id}
            type="button"
            role="tab"
            aria-selected={sheet.id === activeSheetId}
            className={sheet.id === activeSheetId ? "js-spreadsheet-sheet-tab js-spreadsheet-active-sheet-tab" : "js-spreadsheet-sheet-tab"}
            style={tabStyle}
            ref={(element) => {
              if (element) {
                tabRefs.current.set(sheet.id, element);
              } else {
                tabRefs.current.delete(sheet.id);
              }
            }}
            onClick={() => onSelect(sheet.id)}
          >
            {sheet.name}
          </button>
        );
      })}
      <button type="button" className="js-spreadsheet-add-sheet-tab" aria-label="Add sheet tab" title="Add sheet" onClick={onAdd}>
        +
      </button>
    </div>
  );
}
