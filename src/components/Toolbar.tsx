import {
  ArrowDown,
  ArrowDownAZ,
  ArrowDownZA,
  ArrowLeft,
  ArrowRight,
  Bold,
  ChartColumn,
  MessageSquare,
  Columns3,
  Copy,
  Download,
  Edit3,
  Eraser,
  Eye,
  EyeOff,
  Filter,
  FilePlus,
  GitBranch,
  Grid2X2,
  Italic,
  Link,
  ListChecks,
  ListX,
  Lock,
  MapPin,
  PaintBucket,
  Paintbrush,
  Palette,
  PanelBottom,
  PanelLeft,
  PanelTop,
  Plus,
  Printer,
  Redo2,
  RotateCcw,
  Rows3,
  Search,
  Shield,
  Sigma,
  SquareFunction,
  TableCellsMerge,
  TableCellsSplit,
  TableProperties,
  Trash2,
  Undo2,
  Unlink,
  Unlock,
  Upload,
  WrapText
} from "lucide-react";
import { useState, type KeyboardEvent, type ReactNode } from "react";
import type { BorderPreset, CellFormat } from "../types";
import type { AutoFunctionName } from "../lib/autoSum";

const NUMBER_FORMAT_OPTIONS = [
  { value: "general", label: "General" },
  { value: "number", label: "Number" },
  { value: "currency", label: "Currency" },
  { value: "percent", label: "Percent" },
  { value: "date", label: "Date" }
] as const;

const ALIGN_OPTIONS = [
  { value: "left", label: "Left" },
  { value: "center", label: "Center" },
  { value: "right", label: "Right" }
] as const;

const VERTICAL_ALIGN_OPTIONS = [
  { value: "top", label: "Top" },
  { value: "middle", label: "Middle" },
  { value: "bottom", label: "Bottom" }
] as const;

const BORDER_OPTIONS = [
  { value: "", label: "Borders" },
  { value: "all", label: "All borders" },
  { value: "outer", label: "Outer border" },
  { value: "top", label: "Top border" },
  { value: "right", label: "Right border" },
  { value: "bottom", label: "Bottom border" },
  { value: "left", label: "Left border" },
  { value: "none", label: "Clear borders" }
] as const;

const AUTO_FUNCTION_OPTIONS = [
  { value: "", label: "Auto" },
  { value: "SUM", label: "Sum" },
  { value: "AVERAGE", label: "Average" },
  { value: "COUNT", label: "Count" },
  { value: "MAX", label: "Max" },
  { value: "MIN", label: "Min" }
] as const;

const RIBBON_TABS = [
  { id: "file", label: "File" },
  { id: "home", label: "Home" },
  { id: "insert", label: "Insert" },
  { id: "formulas", label: "Formulas" },
  { id: "data", label: "Data" },
  { id: "review", label: "Review" },
  { id: "view", label: "View" }
] as const;

type RibbonTabId = (typeof RIBBON_TABS)[number]["id"];

type ToolbarProps = {
  canUndo: boolean;
  canRedo: boolean;
  activeFormat: CellFormat;
  onNew: () => void;
  onImport: () => void;
  onExport: () => void;
  onImportXlsx: () => void;
  onExportXlsx: () => void;
  onPrint: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onClear: () => void;
  canPasteSpecial: boolean;
  onPasteValues: () => void;
  onPasteFormats: () => void;
  formatPainterActive: boolean;
  onFormatPainter: () => void;
  onTransposePaste: () => void;
  onAutoSum: (functionName?: AutoFunctionName) => void;
  onFunctionLibrary: () => void;
  formulaAuditOpen: boolean;
  onFormulaAudit: () => void;
  onNamedRanges: () => void;
  onGoTo: () => void;
  onComment: () => void;
  onLink: () => void;
  onUnlink: () => void;
  sheetProtected: boolean;
  onLockCells: () => void;
  onUnlockCells: () => void;
  onToggleProtectSheet: () => void;
  onMergeCells: () => void;
  onUnmergeCells: () => void;
  onFillDown: () => void;
  onFillRight: () => void;
  onSortAsc: () => void;
  onSortDesc: () => void;
  onRemoveDuplicates: () => void;
  onInsertRows: () => void;
  onDeleteRows: () => void;
  onInsertColumns: () => void;
  onDeleteColumns: () => void;
  onAutoFitRows: () => void;
  onAutoFitColumns: () => void;
  onHideRows: () => void;
  onHideColumns: () => void;
  onUnhideAll: () => void;
  freezeTopRow: boolean;
  freezeFirstColumn: boolean;
  onToggleFreezeTopRow: () => void;
  onToggleFreezeFirstColumn: () => void;
  onFindReplace: () => void;
  onFilter: () => void;
  onDataValidation: () => void;
  onConditionalFormatting: () => void;
  onAddSheet: () => void;
  onRenameSheet: () => void;
  onDuplicateSheet: () => void;
  onDeleteSheet: () => void;
  onHideSheet: () => void;
  onUnhideSheets: () => void;
  onMoveSheetLeft: () => void;
  onMoveSheetRight: () => void;
  activeSheetTabColor: string;
  onSheetTabColor: (color: string) => void;
  showGridlines: boolean;
  onToggleGridlines: () => void;
  showHeaders: boolean;
  onToggleHeaders: () => void;
  showFormulaBar: boolean;
  onToggleFormulaBar: () => void;
  showFormulas: boolean;
  onToggleShowFormulas: () => void;
  showSheetTabs: boolean;
  onToggleSheetTabs: () => void;
  onResetView: () => void;
  onPivot: () => void;
  onChart: () => void;
  onBold: () => void;
  onItalic: () => void;
  onWrapText: () => void;
  onNumberFormat: (format: NonNullable<CellFormat["numberFormat"]>) => void;
  onHorizontalAlign: (align: NonNullable<CellFormat["horizontalAlign"]>) => void;
  onVerticalAlign: (align: NonNullable<CellFormat["verticalAlign"]>) => void;
  onBorders: (preset: BorderPreset) => void;
  onTextColor: (color: string) => void;
  onFillColor: (color: string) => void;
};

export function Toolbar(props: ToolbarProps) {
  const [activeTab, setActiveTab] = useState<RibbonTabId>("home");
  const activeTabIndex = RIBBON_TABS.findIndex((tab) => tab.id === activeTab);
  const activeTabLabel = RIBBON_TABS[activeTabIndex]?.label ?? "Home";

  function handleTabKeyDown(event: KeyboardEvent<HTMLButtonElement>, index: number) {
    if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) {
      return;
    }

    event.preventDefault();
    const nextIndex =
      event.key === "Home"
        ? 0
        : event.key === "End"
          ? RIBBON_TABS.length - 1
          : event.key === "ArrowRight"
            ? (index + 1) % RIBBON_TABS.length
            : (index - 1 + RIBBON_TABS.length) % RIBBON_TABS.length;
    setActiveTab(RIBBON_TABS[nextIndex].id);
    document.getElementById(ribbonTabId(RIBBON_TABS[nextIndex].id))?.focus();
  }

  return (
    <div className="toolbar" role="toolbar" aria-label="Toolbar">
      <div className="ribbon-header">
        <div className="brand">
          <span className="brand-mark">JS</span>
          <span>JavaScript Spreadsheet</span>
        </div>
        <div className="ribbon-tabs" role="tablist" aria-label="Ribbon tabs">
          {RIBBON_TABS.map((tab, index) => {
            const selected = tab.id === activeTab;

            return (
              <button
                key={tab.id}
                id={ribbonTabId(tab.id)}
                className={selected ? "ribbon-tab active-ribbon-tab" : "ribbon-tab"}
                type="button"
                role="tab"
                aria-selected={selected}
                aria-controls={ribbonPanelId(tab.id)}
                tabIndex={selected ? 0 : -1}
                onClick={() => setActiveTab(tab.id)}
                onKeyDown={(event) => handleTabKeyDown(event, index)}
              >
                {tab.label}
              </button>
            );
          })}
        </div>
      </div>
      <div
        id={ribbonPanelId(activeTab)}
        className="ribbon-panel"
        role="tabpanel"
        aria-label={`${activeTabLabel} ribbon`}
        aria-labelledby={ribbonTabId(activeTab)}
      >
        {renderRibbonTab(activeTab, props)}
      </div>
    </div>
  );
}

function renderRibbonTab(activeTab: RibbonTabId, props: ToolbarProps) {
  if (activeTab === "file") {
    return <WorkbookGroup {...props} />;
  }

  if (activeTab === "insert") {
    return <InsertGroup {...props} />;
  }

  if (activeTab === "formulas") {
    return <FormulaGroup {...props} />;
  }

  if (activeTab === "data") {
    return <DataToolsGroup {...props} />;
  }

  if (activeTab === "review") {
    return <ReviewGroup {...props} />;
  }

  if (activeTab === "view") {
    return <ViewGroup {...props} />;
  }

  return (
    <>
      <ClipboardGroup {...props} />
      <FontGroup {...props} />
      <NumberGroup {...props} />
      <AlignmentGroup {...props} />
      <StylesGroup {...props} />
      <CellsGroup {...props} />
      <SheetGroup {...props} />
      <EditingGroup {...props} />
    </>
  );
}

function WorkbookGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Workbook">
      <ToolbarButton label="New workbook" onClick={props.onNew} icon={<FilePlus />} />
      <ToolbarButton label="Import CSV" onClick={props.onImport} icon={<Upload />} />
      <ToolbarButton label="Export CSV" onClick={props.onExport} icon={<Download />} />
      <ToolbarButton label="Import XLSX" onClick={props.onImportXlsx} icon={<Upload />} />
      <ToolbarButton label="Export XLSX" onClick={props.onExportXlsx} icon={<Download />} />
      <ToolbarButton label="Print workbook" onClick={props.onPrint} icon={<Printer />} />
    </ToolbarGroup>
  );
}

function ClipboardGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Clipboard">
        <ToolbarButton label="Undo" onClick={props.onUndo} icon={<Undo2 />} disabled={!props.canUndo} compact />
        <ToolbarButton label="Redo" onClick={props.onRedo} icon={<Redo2 />} disabled={!props.canRedo} compact />
        <ToolbarButton label="Clear selection" onClick={props.onClear} icon={<Eraser />} compact />
        <ToolbarButton label="Paste values" onClick={props.onPasteValues} icon={<Copy />} disabled={!props.canPasteSpecial} compact />
        <ToolbarButton label="Paste formats" onClick={props.onPasteFormats} icon={<Palette />} disabled={!props.canPasteSpecial} compact />
        <ToolbarButton
          label="Format painter"
          onClick={props.onFormatPainter}
          icon={<Paintbrush />}
          active={props.formatPainterActive}
          compact
        />
        <ToolbarButton
          label="Transpose paste"
          onClick={props.onTransposePaste}
          icon={<TableProperties />}
          disabled={!props.canPasteSpecial}
          compact
        />
      </ToolbarGroup>
  );
}

function FormulaGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Function Library">
        <ToolbarButton label="AutoSum" onClick={() => props.onAutoSum()} icon={<Sigma />} compact />
        <SelectControl
          label="Auto function"
          value=""
          options={AUTO_FUNCTION_OPTIONS}
          compact
          onChange={(value) => {
            if (value) {
              props.onAutoSum(value as AutoFunctionName);
            }
          }}
        />
        <ToolbarButton label="Function library" onClick={props.onFunctionLibrary} icon={<SquareFunction />} compact />
        <ToolbarButton
          label="Audit formulas"
          onClick={props.onFormulaAudit}
          icon={<GitBranch />}
          active={props.formulaAuditOpen}
          compact
        />
        <ToolbarButton label="Named ranges" onClick={props.onNamedRanges} icon={<TableProperties />} compact />
        <ToolbarButton label="Go to" onClick={props.onGoTo} icon={<MapPin />} compact />
      </ToolbarGroup>
  );
}

function ReviewGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Review">
        <ToolbarButton label="Comment" onClick={props.onComment} icon={<MessageSquare />} compact />
        <ToolbarButton label="Link" onClick={props.onLink} icon={<Link />} compact />
        <ToolbarButton label="Unlink" onClick={props.onUnlink} icon={<Unlink />} compact />
        <ToolbarButton label="Lock cells" onClick={props.onLockCells} icon={<Lock />} compact />
        <ToolbarButton label="Unlock cells" onClick={props.onUnlockCells} icon={<Unlock />} compact />
        <ToolbarButton
          label="Protect sheet"
          onClick={props.onToggleProtectSheet}
          icon={<Shield />}
          active={props.sheetProtected}
          compact
        />
      </ToolbarGroup>
  );
}

function DataToolsGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Data Tools">
        <ToolbarButton label="Fill down" onClick={props.onFillDown} icon={<ArrowDown />} compact />
        <ToolbarButton label="Fill right" onClick={props.onFillRight} icon={<ArrowRight />} compact />
        <ToolbarButton label="Sort A to Z" onClick={props.onSortAsc} icon={<ArrowDownAZ />} compact />
        <ToolbarButton label="Sort Z to A" onClick={props.onSortDesc} icon={<ArrowDownZA />} compact />
        <ToolbarButton label="Remove duplicates" onClick={props.onRemoveDuplicates} icon={<ListX />} compact />
        <ToolbarButton label="Filter" onClick={props.onFilter} icon={<Filter />} compact />
        <ToolbarButton label="Find and replace" onClick={props.onFindReplace} icon={<Search />} compact />
        <ToolbarButton label="Data validation" onClick={props.onDataValidation} icon={<ListChecks />} compact />
        <ToolbarButton label="Conditional formatting" onClick={props.onConditionalFormatting} icon={<Palette />} compact />
      </ToolbarGroup>
  );
}

function InsertGroup(props: ToolbarProps) {
  return (
    <>
      <ToolbarGroup label="Tables">
        <ToolbarButton label="Pivot table" onClick={props.onPivot} icon={<TableProperties />} compact />
      </ToolbarGroup>
      <ToolbarGroup label="Charts">
        <ToolbarButton label="Chart" onClick={props.onChart} icon={<ChartColumn />} compact />
      </ToolbarGroup>
      <ToolbarGroup label="Links">
        <ToolbarButton label="Link" onClick={props.onLink} icon={<Link />} compact />
        <ToolbarButton label="Comment" onClick={props.onComment} icon={<MessageSquare />} compact />
      </ToolbarGroup>
    </>
  );
}

function CellsGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Cells">
        <ToolbarButton label="Merge cells" onClick={props.onMergeCells} icon={<TableCellsMerge />} compact />
        <ToolbarButton label="Unmerge cells" onClick={props.onUnmergeCells} icon={<TableCellsSplit />} compact />
        <ToolbarButton label="Insert row above" onClick={props.onInsertRows} icon={<Rows3 />} compact />
        <ToolbarButton label="Delete row" onClick={props.onDeleteRows} icon={<Trash2 />} compact />
        <ToolbarButton label="Insert column left" onClick={props.onInsertColumns} icon={<Columns3 />} compact />
        <ToolbarButton label="Delete column" onClick={props.onDeleteColumns} icon={<Trash2 />} compact />
        <ToolbarButton label="Auto-fit rows" onClick={props.onAutoFitRows} icon={<Rows3 />} compact />
        <ToolbarButton label="Auto-fit columns" onClick={props.onAutoFitColumns} icon={<Columns3 />} compact />
        <ToolbarButton label="Hide rows" onClick={props.onHideRows} icon={<EyeOff />} compact />
        <ToolbarButton label="Hide columns" onClick={props.onHideColumns} icon={<EyeOff />} compact />
        <ToolbarButton label="Unhide all" onClick={props.onUnhideAll} icon={<Eye />} compact />
        <ToolbarButton
          label="Freeze top row"
          onClick={props.onToggleFreezeTopRow}
          icon={<PanelTop />}
          active={props.freezeTopRow}
          compact
        />
        <ToolbarButton
          label="Freeze first column"
          onClick={props.onToggleFreezeFirstColumn}
          icon={<PanelLeft />}
          active={props.freezeFirstColumn}
          compact
        />
      </ToolbarGroup>
  );
}

function FontGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Font">
        <ToolbarButton label="Bold" onClick={props.onBold} icon={<Bold />} active={Boolean(props.activeFormat.bold)} compact />
        <ToolbarButton
          label="Italic"
          onClick={props.onItalic}
          icon={<Italic />}
          active={Boolean(props.activeFormat.italic)}
          compact
        />
        <ToolbarButton
          label="Wrap text"
          onClick={props.onWrapText}
          icon={<WrapText />}
          active={Boolean(props.activeFormat.wrapText)}
          compact
        />
        <ColorControl
          label="Text color"
          icon={<Palette />}
          value={props.activeFormat.textColor ?? "#1f2937"}
          onChange={props.onTextColor}
          compact
        />
        <ColorControl
          label="Fill color"
          icon={<PaintBucket />}
          value={props.activeFormat.backgroundColor ?? "#ffffff"}
          onChange={props.onFillColor}
          compact
        />
      </ToolbarGroup>
  );
}

function NumberGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Number">
        <SelectControl
          label="Number format"
          value={props.activeFormat.numberFormat ?? "general"}
          options={NUMBER_FORMAT_OPTIONS}
          onChange={(value) => props.onNumberFormat(value as NonNullable<CellFormat["numberFormat"]>)}
        />
      </ToolbarGroup>
  );
}

function AlignmentGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Alignment">
        <SelectControl
          label="Horizontal align"
          value={props.activeFormat.horizontalAlign ?? "left"}
          options={ALIGN_OPTIONS}
          onChange={(value) => props.onHorizontalAlign(value as NonNullable<CellFormat["horizontalAlign"]>)}
        />
        <SelectControl
          label="Vertical align"
          value={props.activeFormat.verticalAlign ?? "middle"}
          options={VERTICAL_ALIGN_OPTIONS}
          onChange={(value) => props.onVerticalAlign(value as NonNullable<CellFormat["verticalAlign"]>)}
        />
        <SelectControl
          label="Borders"
          value=""
          options={BORDER_OPTIONS}
          onChange={(value) => {
            if (value) {
              props.onBorders(value as BorderPreset);
            }
          }}
        />
      </ToolbarGroup>
  );
}

function StylesGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Styles">
      <ToolbarButton label="Conditional formatting" onClick={props.onConditionalFormatting} icon={<Palette />} compact />
      <ToolbarButton label="Data validation" onClick={props.onDataValidation} icon={<ListChecks />} compact />
    </ToolbarGroup>
  );
}

function SheetGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Sheet">
        <ToolbarButton label="Add sheet" onClick={props.onAddSheet} icon={<Plus />} compact />
        <ToolbarButton label="Rename active sheet" onClick={props.onRenameSheet} icon={<Edit3 />} compact />
        <ToolbarButton label="Duplicate active sheet" onClick={props.onDuplicateSheet} icon={<Copy />} compact />
        <ToolbarButton label="Delete active sheet" onClick={props.onDeleteSheet} icon={<Trash2 />} compact />
        <ToolbarButton label="Hide sheet" onClick={props.onHideSheet} icon={<EyeOff />} compact />
        <ToolbarButton label="Unhide sheets" onClick={props.onUnhideSheets} icon={<Eye />} compact />
        <ToolbarButton label="Move sheet left" onClick={props.onMoveSheetLeft} icon={<ArrowLeft />} compact />
        <ToolbarButton label="Move sheet right" onClick={props.onMoveSheetRight} icon={<ArrowRight />} compact />
        <ColorControl
          label="Sheet tab color"
          icon={<Palette />}
          value={props.activeSheetTabColor}
          onChange={props.onSheetTabColor}
          compact
        />
      </ToolbarGroup>
  );
}

function EditingGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Editing">
      <ToolbarButton label="AutoSum" onClick={() => props.onAutoSum()} icon={<Sigma />} compact />
      <ToolbarButton label="Fill down" onClick={props.onFillDown} icon={<ArrowDown />} compact />
      <ToolbarButton label="Fill right" onClick={props.onFillRight} icon={<ArrowRight />} compact />
      <ToolbarButton label="Sort A to Z" onClick={props.onSortAsc} icon={<ArrowDownAZ />} compact />
      <ToolbarButton label="Sort Z to A" onClick={props.onSortDesc} icon={<ArrowDownZA />} compact />
      <ToolbarButton label="Remove duplicates" onClick={props.onRemoveDuplicates} icon={<ListX />} compact />
      <ToolbarButton label="Filter" onClick={props.onFilter} icon={<Filter />} compact />
      <ToolbarButton label="Find and replace" onClick={props.onFindReplace} icon={<Search />} compact />
      <ToolbarButton label="Go to" onClick={props.onGoTo} icon={<MapPin />} compact />
    </ToolbarGroup>
  );
}

function ViewGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Workbook Views">
        <ToolbarButton
          label={props.showGridlines ? "Hide gridlines" : "Show gridlines"}
          onClick={props.onToggleGridlines}
          icon={<Grid2X2 />}
          active={!props.showGridlines}
          compact
        />
        <ToolbarButton
          label={props.showHeaders ? "Hide headers" : "Show headers"}
          onClick={props.onToggleHeaders}
          icon={<PanelTop />}
          active={!props.showHeaders}
          compact
        />
        <ToolbarButton
          label={props.showFormulaBar ? "Hide formula bar" : "Show formula bar"}
          onClick={props.onToggleFormulaBar}
          icon={<SquareFunction />}
          active={!props.showFormulaBar}
          compact
        />
        <ToolbarButton
          label={props.showFormulas ? "Show formula results" : "Show formulas"}
          onClick={props.onToggleShowFormulas}
          icon={<Sigma />}
          active={props.showFormulas}
          compact
        />
        <ToolbarButton
          label={props.showSheetTabs ? "Hide sheet tabs" : "Show sheet tabs"}
          onClick={props.onToggleSheetTabs}
          icon={<PanelBottom />}
          active={!props.showSheetTabs}
          compact
        />
        <ToolbarButton
          label="Freeze top row"
          onClick={props.onToggleFreezeTopRow}
          icon={<PanelTop />}
          active={props.freezeTopRow}
          compact
        />
        <ToolbarButton
          label="Freeze first column"
          onClick={props.onToggleFreezeFirstColumn}
          icon={<PanelLeft />}
          active={props.freezeFirstColumn}
          compact
        />
        <ToolbarButton label="Reset view" onClick={props.onResetView} icon={<RotateCcw />} compact />
      </ToolbarGroup>
  );
}

function ribbonTabId(tab: RibbonTabId): string {
  return `ribbon-tab-${tab}`;
}

function ribbonPanelId(tab: RibbonTabId): string {
  return `ribbon-panel-${tab}`;
}

function ToolbarGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="toolbar-group" role="group" aria-label={label}>
      <span className="toolbar-group-label">{label}</span>
      <div className="toolbar-group-controls">{children}</div>
    </div>
  );
}

function ToolbarButton({
  label,
  onClick,
  icon,
  disabled = false,
  active = false,
  compact = false
}: {
  label: string;
  onClick: () => void;
  icon: ReactNode;
  disabled?: boolean;
  active?: boolean;
  compact?: boolean;
}) {
  const className = [
    "toolbar-button",
    active ? "active-toolbar-button" : "",
    compact ? "compact-toolbar-button" : ""
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <button
      type="button"
      className={className}
      aria-label={label}
      title={label}
      onClick={onClick}
      disabled={disabled}
    >
      {icon}
      <span>{label}</span>
    </button>
  );
}

function SelectControl({
  label,
  value,
  options,
  onChange,
  compact = false
}: {
  label: string;
  value: string;
  options: readonly { value: string; label: string }[];
  onChange: (value: string) => void;
  compact?: boolean;
}) {
  return (
    <label className={compact ? "select-control compact-select-control" : "select-control"} title={label}>
      <span>{label}</span>
      <select aria-label={label} value={value} onChange={(event) => onChange(event.currentTarget.value)}>
        {options.map((option) => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
    </label>
  );
}

function ColorControl({
  label,
  icon,
  value,
  onChange,
  compact = false
}: {
  label: string;
  icon: ReactNode;
  value: string;
  onChange: (value: string) => void;
  compact?: boolean;
}) {
  return (
    <label className={compact ? "color-control compact-color-control" : "color-control"} title={label}>
      {icon}
      <span>{label}</span>
      <input
        type="color"
        aria-label={label}
        value={value}
        onInput={(event) => onChange(event.currentTarget.value)}
      />
    </label>
  );
}
