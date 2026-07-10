export type CellContent = string | number | boolean | null;

export type CellCoord = {
  row: number;
  column: number;
};

export type CellRange = {
  start: CellCoord;
  end: CellCoord;
};

export type CellBorderSide = {
  style: "thin";
  color: string;
};

export type CellBorders = {
  top?: CellBorderSide;
  right?: CellBorderSide;
  bottom?: CellBorderSide;
  left?: CellBorderSide;
};

export type BorderPreset = "none" | "all" | "outer" | "top" | "right" | "bottom" | "left";

export type CellFormat = {
  bold?: boolean;
  italic?: boolean;
  fontFamily?: string;
  /** In points, as in Excel; the grid renders it with the CSS `pt` unit. */
  fontSize?: number;
  textColor?: string;
  backgroundColor?: string;
  numberFormat?: "general" | "number" | "currency" | "percent" | "date" | "dateTime";
  horizontalAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  wrapText?: boolean;
  borders?: CellBorders;
};

/** A format field summarized over a selection: concrete when every cell agrees. */
export type MixedFormatValue<T> = T | "mixed" | undefined;

/** Excel-style ribbon state for the whole selection, with "mixed" indicators. */
export type SelectionFormatSummary = {
  bold: boolean | "mixed";
  italic: boolean | "mixed";
  wrapText: boolean | "mixed";
  fontFamily: MixedFormatValue<string>;
  fontSize: MixedFormatValue<number>;
  numberFormat: MixedFormatValue<NonNullable<CellFormat["numberFormat"]>>;
  horizontalAlign: MixedFormatValue<NonNullable<CellFormat["horizontalAlign"]>>;
  verticalAlign: MixedFormatValue<NonNullable<CellFormat["verticalAlign"]>>;
};

export type DataValidationRule =
  | {
      type: "list";
      values: readonly string[];
      allowBlank?: boolean;
    }
  | {
      type: "number";
      min?: number;
      max?: number;
      allowBlank?: boolean;
    }
  | {
      type: "textLength";
      min?: number;
      max?: number;
      allowBlank?: boolean;
    };

export type ConditionalFormatCondition =
  | {
      type: "greaterThan" | "lessThan" | "equalTo" | "textContains";
      value: string;
    }
  | {
      type: "between";
      value: string;
      secondValue: string;
    }
  | {
      type: "blank" | "notBlank" | "duplicate" | "unique";
    }
  | {
      type: "top" | "bottom";
      count: number;
    }
  | {
      type: "dataBar";
      color: string;
    }
  | {
      type: "colorScale";
      minColor: string;
      maxColor: string;
    };

export type ConditionalFormatRule = {
  id: string;
  range: CellRange;
  condition: ConditionalFormatCondition;
  format: CellFormat;
};

export type FilterOperator = "contains" | "equals" | "greaterThan" | "lessThan";

export type SheetFilter = {
  id: string;
  range: CellRange;
  column: number;
  operator: FilterOperator;
  value: string;
  values?: string[];
  hasHeader?: boolean;
};

export type SheetChartType = "bar" | "line" | "pie";

export type SheetChart = {
  id: string;
  title: string;
  type: SheetChartType;
  range: CellRange;
  anchor: CellCoord;
};

export type SheetMerge = {
  id: string;
  range: CellRange;
};

export type SheetProtection = {
  isProtected: boolean;
  lockedCells: Record<string, boolean>;
  unlockedCells: Record<string, boolean>;
};

export type NamedRange = {
  name: string;
  sheetId: string;
  range: CellRange;
};

export type SheetModel = {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  isHidden?: boolean;
  tabColor?: string;
  cells: Record<string, CellContent>;
  formats: Record<string, CellFormat>;
  columnWidths: Record<string, number>;
  rowHeights: Record<string, number>;
  hiddenColumns?: Record<string, boolean>;
  hiddenRows?: Record<string, boolean>;
  freezeTopRow?: boolean;
  freezeFirstColumn?: boolean;
  comments: Record<string, string>;
  hyperlinks: Record<string, string>;
  validations: Record<string, DataValidationRule>;
  conditionalFormats: ConditionalFormatRule[];
  autoFilterRange?: CellRange;
  filters: SheetFilter[];
  charts: SheetChart[];
  merges: SheetMerge[];
  protection: SheetProtection;
};

export type WorkbookModel = {
  version: 1;
  activeSheetId: string;
  sheets: SheetModel[];
  namedRanges: NamedRange[];
};

export type Selection = CellRange;

export type WorkbookHistoryLimits = Readonly<{
  maxEntries: number;
  maxWeight: number;
}>;

export type WorkbookHistory = {
  past: WorkbookModel[];
  present: WorkbookModel;
  future: WorkbookModel[];
  /** Absent on histories created by releases before weighted retention. */
  limits?: WorkbookHistoryLimits;
};

/** @deprecated Import `WorkbookHistory` from `core/workbook/history` for new code. */
export type HistoryState = WorkbookHistory;
