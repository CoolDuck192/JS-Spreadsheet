import {
  createContext,
  useContext,
  useEffect,
  useId,
  useRef,
  useState,
  type KeyboardEvent,
  type ReactNode
} from "react";
import type { TableIssue } from "../core/commands/types";
import { normalizeRange, parseRangeAddress } from "../lib/addressing";
import { formatRangeAddress } from "../lib/autoSum";
import type { FilterExpression } from "../table/core/query";
import type { CellRange, StructuredTable, TableAggregate, TableStyle } from "../types";
import { CalculatedColumnPanel } from "./CalculatedColumnPanel";
import { StructuredTableFilterPanel } from "./StructuredTableFilterPanel";

const DEFAULT_EXPORT_REASON = "Export is not supported by this table";
const SpreadsheetTableExportReasonContext = createContext<string | undefined>(DEFAULT_EXPORT_REASON);

const TABLE_STYLES = [
  { value: "TableStyleLight1", label: "Light 1" },
  { value: "TableStyleLight9", label: "Light 9" },
  { value: "TableStyleMedium2", label: "Medium 2" },
  { value: "TableStyleMedium9", label: "Medium 9" },
  { value: "TableStyleDark1", label: "Dark 1" }
] as const;

const TOTALS_FUNCTIONS: Array<{ value: TableAggregate; label: string }> = [
  { value: "none", label: "None" },
  { value: "sum", label: "Sum" },
  { value: "average", label: "Average" },
  { value: "count", label: "Count" },
  { value: "countNumbers", label: "Count numbers" },
  { value: "min", label: "Minimum" },
  { value: "max", label: "Maximum" },
  { value: "standardDeviation", label: "Standard deviation" },
  { value: "variance", label: "Variance" }
];

export type SpreadsheetTableTabProps = {
  table: StructuredTable;
  activeColumnId?: string;
  issues: readonly TableIssue[];
  onRename(name: string): void;
  onResize(range: CellRange): void;
  onHeaderRow(enabled: boolean): void;
  onTotalsRow(enabled: boolean): void;
  onTotalsFunction(columnId: string, aggregate: TableAggregate): void;
  onStyle(style: TableStyle): void;
  onKeyColumn(columnId?: string): void;
  onCalculatedColumn(columnId: string, formula?: string): void;
  onFilter(filter?: FilterExpression): void;
  onExport(): void;
  onConvertToRange(): void;
  onOpenTableView(): void;
};

export function SpreadsheetTableExportReasonProvider({ reason, children }: { reason?: string; children: ReactNode }) {
  return (
    <SpreadsheetTableExportReasonContext.Provider value={reason}>
      {children}
    </SpreadsheetTableExportReasonContext.Provider>
  );
}

export function SpreadsheetTableTab({
  table,
  activeColumnId,
  issues,
  onRename,
  onResize,
  onHeaderRow,
  onTotalsRow,
  onTotalsFunction,
  onStyle,
  onKeyColumn,
  onCalculatedColumn,
  onFilter,
  onExport,
  onConvertToRange,
  onOpenTableView
}: SpreadsheetTableTabProps) {
  const exportReason = useContext(SpreadsheetTableExportReasonContext);
  const [nameDraft, setNameDraft] = useState(table.name);
  const [rangeDraft, setRangeDraft] = useState(formatRangeAddress(table.range));
  const [localNameError, setLocalNameError] = useState("");
  const [localRangeError, setLocalRangeError] = useState("");
  const [totalsColumnId, setTotalsColumnId] = useState(activeColumnId ?? table.columns[0]?.id ?? "");
  const [openPanel, setOpenPanel] = useState<"calculated" | "filter" | null>(null);
  const nameRef = useRef<HTMLInputElement>(null);
  const rangeRef = useRef<HTMLInputElement>(null);
  const id = useId();
  const nameErrorId = `table-name-error-${id}`;
  const rangeErrorId = `table-range-error-${id}`;
  const exportReasonId = `structured-table-export-reason-${id}`;

  useEffect(() => {
    setNameDraft(table.name);
    setLocalNameError("");
  }, [table.id, table.name]);

  useEffect(() => {
    setRangeDraft(formatRangeAddress(table.range));
    setLocalRangeError("");
  }, [table.id, table.range.end.column, table.range.end.row, table.range.start.column, table.range.start.row]);

  useEffect(() => {
    setTotalsColumnId(
      activeColumnId && table.columns.some((column) => column.id === activeColumnId)
        ? activeColumnId
        : table.columns[0]?.id ?? ""
    );
    setOpenPanel(null);
  }, [activeColumnId, table.columns, table.id]);

  const nameIssue = issues.find((issue) => issue.code.startsWith("TABLE_NAME_"))?.message;
  const rangeIssue = issues.find((issue) => isRangeIssue(issue.code))?.message;
  const calculatedIssue = issues.find((issue) =>
    issue.code.includes("FORMULA") || issue.code.includes("CALCULATED_COLUMN")
  )?.message;
  const generalIssues = issues.filter(
    (issue) => !issue.code.startsWith("TABLE_NAME_") && !isRangeIssue(issue.code)
  );
  const nameError = localNameError || nameIssue;
  const rangeError = localRangeError || rangeIssue;
  const selectedTotalsColumn = table.columns.find((column) => column.id === totalsColumnId);
  const style = table.style ?? { theme: "TableStyleLight1", showRowStripes: true };

  function commitName() {
    const nextName = nameDraft.trim();
    if (!nextName) {
      setLocalNameError("Enter a table name");
      nameRef.current?.focus();
      return;
    }
    setLocalNameError("");
    if (nextName !== table.name) {
      onRename(nextName);
    }
  }

  function commitRange() {
    try {
      const nextRange = normalizeRange(parseRangeAddress(rangeDraft.trim()));
      setLocalRangeError("");
      if (formatRangeAddress(nextRange) !== formatRangeAddress(table.range)) {
        onResize(nextRange);
      }
    } catch {
      setLocalRangeError("Enter a valid cell range, for example A1:D12");
      rangeRef.current?.focus();
    }
  }

  function commitOnEnter(event: KeyboardEvent<HTMLInputElement>, commit: () => void) {
    if (event.key === "Enter") {
      event.preventDefault();
      commit();
    }
  }

  function updateStyle(patch: Partial<TableStyle>) {
    onStyle({ ...style, ...patch });
  }

  return (
    <div className="js-spreadsheet-spreadsheet-table-tab" role="group" aria-label={`Table ${table.name}`}>
      <div className="js-spreadsheet-structured-table-tab-fields">
        <label className="js-spreadsheet-structured-table-field">
          <span>Table name</span>
          <input
            ref={nameRef}
            aria-label="Table name"
            aria-describedby={nameError ? nameErrorId : undefined}
            value={nameDraft}
            onChange={(event) => {
              setNameDraft(event.currentTarget.value);
              setLocalNameError("");
            }}
            onKeyDown={(event) => commitOnEnter(event, commitName)}
          />
        </label>
        <label className="js-spreadsheet-structured-table-field">
          <span>Table range</span>
          <input
            ref={rangeRef}
            aria-label="Table range"
            aria-describedby={rangeError ? rangeErrorId : undefined}
            value={rangeDraft}
            onChange={(event) => {
              setRangeDraft(event.currentTarget.value);
              setLocalRangeError("");
            }}
            onKeyDown={(event) => commitOnEnter(event, commitRange)}
          />
        </label>
        <label className="js-spreadsheet-structured-table-check">
          <input type="checkbox" aria-label="Header row" checked={table.headerRow} onChange={(event) => onHeaderRow(event.currentTarget.checked)} />
          <span>Header row</span>
        </label>
        <label className="js-spreadsheet-structured-table-check">
          <input type="checkbox" aria-label="Totals row" checked={table.totalsRow} onChange={(event) => onTotalsRow(event.currentTarget.checked)} />
          <span>Totals row</span>
        </label>
      </div>

      <div className="js-spreadsheet-structured-table-tab-fields">
        <label className="js-spreadsheet-structured-table-field">
          <span>Totals column</span>
          <select aria-label="Totals column" value={totalsColumnId} onChange={(event) => setTotalsColumnId(event.currentTarget.value)}>
            {table.columns.map((column) => (
              <option key={column.id} value={column.id}>{column.name}</option>
            ))}
          </select>
        </label>
        <label className="js-spreadsheet-structured-table-field">
          <span>Totals function</span>
          <select
            aria-label="Totals function"
            value={selectedTotalsColumn?.totalsFunction ?? "none"}
            disabled={!totalsColumnId}
            onChange={(event) => onTotalsFunction(totalsColumnId, event.currentTarget.value as TableAggregate)}
          >
            {TOTALS_FUNCTIONS.map((aggregate) => (
              <option key={aggregate.value} value={aggregate.value}>{aggregate.label}</option>
            ))}
          </select>
        </label>
        <label className="js-spreadsheet-structured-table-field">
          <span>Table style</span>
          <select aria-label="Table style" value={style.theme} onChange={(event) => updateStyle({ theme: event.currentTarget.value })}>
            {TABLE_STYLES.map((option) => (
              <option key={option.value} value={option.value}>{option.label}</option>
            ))}
          </select>
        </label>
        <label className="js-spreadsheet-structured-table-field">
          <span>Key column</span>
          <select aria-label="Key column" value={table.keyColumnId ?? ""} onChange={(event) => onKeyColumn(event.currentTarget.value || undefined)}>
            <option value="">No key column</option>
            {table.columns.map((column) => (
              <option key={column.id} value={column.id}>{column.name}</option>
            ))}
          </select>
        </label>
      </div>

      <div className="js-spreadsheet-structured-table-style-options" role="group" aria-label="Table style options">
        <StyleToggle label="First column" checked={Boolean(style.showFirstColumn)} onChange={(checked) => updateStyle({ showFirstColumn: checked })} />
        <StyleToggle label="Last column" checked={Boolean(style.showLastColumn)} onChange={(checked) => updateStyle({ showLastColumn: checked })} />
        <StyleToggle label="Row stripes" checked={style.showRowStripes !== false} onChange={(checked) => updateStyle({ showRowStripes: checked })} />
        <StyleToggle label="Column stripes" checked={Boolean(style.showColumnStripes)} onChange={(checked) => updateStyle({ showColumnStripes: checked })} />
      </div>

      <div className="js-spreadsheet-structured-table-actions" role="group" aria-label="Table actions">
        <button type="button" aria-expanded={openPanel === "calculated"} onClick={() => setOpenPanel((current) => current === "calculated" ? null : "calculated")}>
          Calculated column
        </button>
        <button type="button" aria-expanded={openPanel === "filter"} onClick={() => setOpenPanel((current) => current === "filter" ? null : "filter")}>
          Filter table
        </button>
        <button
          type="button"
          aria-label="Export table"
          aria-describedby={exportReason ? exportReasonId : undefined}
          onClick={onExport}
          disabled={Boolean(exportReason)}
        >
          Export table
        </button>
        {exportReason ? <span id={exportReasonId} className="js-spreadsheet-visually-hidden">{exportReason}</span> : null}
        <button type="button" onClick={onConvertToRange}>Convert to range</button>
        <button type="button" onClick={onOpenTableView}>Open table view</button>
      </div>

      {openPanel === "calculated" ? (
        <CalculatedColumnPanel
          columns={table.columns}
          activeColumnId={activeColumnId}
          issue={calculatedIssue}
          onApply={onCalculatedColumn}
          onClose={() => setOpenPanel(null)}
        />
      ) : null}
      {openPanel === "filter" ? (
        <StructuredTableFilterPanel
          columns={table.columns}
          activeColumnId={activeColumnId}
          filter={table.filter}
          onApply={onFilter}
          onClose={() => setOpenPanel(null)}
        />
      ) : null}

      {nameError || rangeError || generalIssues.length > 0 ? (
        <div className="js-spreadsheet-structured-table-issue-summary" role="alert" aria-live="assertive">
          {nameError ? <p id={nameErrorId}>{nameError}</p> : null}
          {rangeError ? <p id={rangeErrorId}>{rangeError}</p> : null}
          {generalIssues.map((issue, index) => <p key={`${issue.code}-${index}`}>{issue.message}</p>)}
        </div>
      ) : null}
    </div>
  );
}

function StyleToggle({ label, checked, onChange }: { label: string; checked: boolean; onChange(checked: boolean): void }) {
  return (
    <label className="js-spreadsheet-structured-table-check">
      <input type="checkbox" aria-label={label} checked={checked} onChange={(event) => onChange(event.currentTarget.checked)} />
      <span>{label}</span>
    </label>
  );
}

function isRangeIssue(code: string): boolean {
  return ["TABLE_RANGE_BLOCKED", "TABLE_RANGE_OVERLAP", "TABLE_MERGE_CONFLICT", "TABLE_PROTECTED"].includes(code);
}
