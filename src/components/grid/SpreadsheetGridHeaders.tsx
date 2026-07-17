import type { MouseEvent as ReactMouseEvent, ReactNode } from "react";

export type SpreadsheetColumnHeaderProps = {
  column: number;
  label: string;
  onStartResize(column: number, event: ReactMouseEvent<HTMLButtonElement>): void;
  onAutoFit?(column: number): void;
};

export function SpreadsheetColumnHeader({
  column,
  label,
  onStartResize,
  onAutoFit
}: SpreadsheetColumnHeaderProps): ReactNode {
  return (
    <>
      {label}
      <button
        type="button"
        className="js-spreadsheet-column-resize-handle"
        aria-label={`Resize column ${label}`}
        onClick={(event) => {
          event.preventDefault();
          event.stopPropagation();
        }}
        onDoubleClick={(event) => {
          event.preventDefault();
          event.stopPropagation();
          onAutoFit?.(column);
        }}
        onMouseDown={(event) => {
          event.preventDefault();
          event.stopPropagation();
          onStartResize(column, event);
        }}
      />
    </>
  );
}

export type SpreadsheetRowHeaderProps = {
  row: number;
  label: string;
  onStartResize(row: number, event: ReactMouseEvent<HTMLButtonElement>): void;
  onAutoFit?(row: number): void;
};

export function SpreadsheetRowHeader({
  row,
  label,
  onStartResize,
  onAutoFit
}: SpreadsheetRowHeaderProps): ReactNode {
  return (
    <>
      {label}
      <button
        type="button"
        className="js-spreadsheet-row-resize-handle"
        aria-label={`Resize row ${label}`}
        onClick={(event) => {
          event.preventDefault();
          event.stopPropagation();
        }}
        onDoubleClick={(event) => {
          event.preventDefault();
          event.stopPropagation();
          onAutoFit?.(row);
        }}
        onMouseDown={(event) => {
          event.preventDefault();
          event.stopPropagation();
          onStartResize(row, event);
        }}
      />
    </>
  );
}
