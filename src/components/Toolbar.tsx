import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDown,
  ArrowDownAZ,
  ArrowDownToLine,
  ArrowDownZA,
  ArrowLeft,
  ArrowRight,
  ArrowUpToLine,
  Bold,
  ChartColumn,
  ChevronDown,
  Cloud,
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
  FoldVertical,
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
  Sparkles,
  SquareFunction,
  TableCellsMerge,
  TableCellsSplit,
  TableProperties,
  Trash2,
  Type,
  Undo2,
  Unlink,
  Unlock,
  Upload,
  WrapText
} from "lucide-react";
import { useEffect, useRef, useState, type KeyboardEvent, type ReactNode } from "react";
import type { BorderPreset, CellFormat, SelectionFormatSummary } from "../types";
import type { AutoFunctionName } from "../lib/autoSum";

const MIXED_SELECT_VALUE = "__mixed__";

const NUMBER_FORMAT_OPTIONS = [
  { value: "general", label: "General" },
  { value: "number", label: "Number" },
  { value: "currency", label: "Currency" },
  { value: "percent", label: "Percent" },
  { value: "date", label: "Date" }
] as const;

const FONT_FAMILY_OPTIONS = [
  { value: "", label: "Default font" },
  { value: "Arial", label: "Arial" },
  { value: "Calibri", label: "Calibri" },
  { value: "Courier New", label: "Courier New" },
  { value: "Georgia", label: "Georgia" },
  { value: "Helvetica", label: "Helvetica" },
  { value: "Times New Roman", label: "Times New Roman" },
  { value: "Verdana", label: "Verdana" }
] as const;

const FONT_SIZE_OPTIONS = [
  { value: "", label: "Size" },
  ...[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36].map((size) => ({ value: String(size), label: String(size) }))
] as const;

const HORIZONTAL_ALIGN_BUTTONS = [
  { value: "left", label: "Align left", icon: <AlignLeft /> },
  { value: "center", label: "Align center", icon: <AlignCenter /> },
  { value: "right", label: "Align right", icon: <AlignRight /> }
] as const;

const VERTICAL_ALIGN_BUTTONS = [
  { value: "top", label: "Align top", icon: <ArrowUpToLine /> },
  { value: "middle", label: "Align middle", icon: <FoldVertical /> },
  { value: "bottom", label: "Align bottom", icon: <ArrowDownToLine /> }
] as const;

const BORDER_MENU_ITEMS: Array<{ value: BorderPreset; label: string }> = [
  { value: "all", label: "All borders" },
  { value: "outer", label: "Outer border" },
  { value: "top", label: "Top border" },
  { value: "right", label: "Right border" },
  { value: "bottom", label: "Bottom border" },
  { value: "left", label: "Left border" },
  { value: "none", label: "Clear borders" }
];

const AUTO_FUNCTION_MENU_ITEMS: Array<{ value: AutoFunctionName; label: string }> = [
  { value: "SUM", label: "Sum" },
  { value: "AVERAGE", label: "Average" },
  { value: "COUNT", label: "Count" },
  { value: "MAX", label: "Max" },
  { value: "MIN", label: "Min" }
];

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
  selectionFormat: SelectionFormatSummary;
  findPanelOpen: boolean;
  filterPanelOpen: boolean;
  validationPanelOpen: boolean;
  conditionalPanelOpen: boolean;
  pivotPanelOpen: boolean;
  chartPanelOpen: boolean;
  functionLibraryOpen: boolean;
  namedRangesOpen: boolean;
  goToPanelOpen: boolean;
  onNew: () => void;
  onImport: () => void;
  onExport: () => void;
  onImportXlsx: () => void;
  onImportGoogleSheet: () => void;
  onExportXlsx: () => void;
  onPrint: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onClear: () => void;
  canPasteSpecial: boolean;
  onPaste: () => void;
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
  onFontFamily: (fontFamily: string) => void;
  onFontSize: (fontSize: number | null) => void;
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

  // Excel's Home order: Clipboard, Font, Alignment, Number, Styles, Cells, Editing.
  return (
    <>
      <ClipboardGroup {...props} />
      <FontGroup {...props} />
      <AlignmentGroup {...props} />
      <NumberGroup {...props} />
      <StylesGroup {...props} />
      <CellsGroup {...props} />
      <EditingGroup {...props} />
      <SheetGroup {...props} />
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
      <ToolbarButton label="Link Google Sheet" onClick={props.onImportGoogleSheet} icon={<Cloud />} />
      <ToolbarButton label="Print workbook" onClick={props.onPrint} icon={<Printer />} />
    </ToolbarGroup>
  );
}

function ClipboardGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Clipboard">
      <ToolbarButton label="Undo" onClick={props.onUndo} icon={<Undo2 />} disabled={!props.canUndo} shortcut="Control+Z" compact />
      <ToolbarButton label="Redo" onClick={props.onRedo} icon={<Redo2 />} disabled={!props.canRedo} shortcut="Control+Y" compact />
      <SplitButton
        label="Paste"
        icon={<Copy />}
        disabled={!props.canPasteSpecial}
        onPrimary={props.onPaste}
        items={[
          { label: "Paste values", onSelect: props.onPasteValues },
          { label: "Paste formats", onSelect: props.onPasteFormats },
          { label: "Transpose paste", onSelect: props.onTransposePaste }
        ]}
      />
      <ToolbarButton
        label="Format painter"
        onClick={props.onFormatPainter}
        icon={<Paintbrush />}
        pressed={props.formatPainterActive}
        compact
      />
      <ToolbarButton label="Clear selection" onClick={props.onClear} icon={<Eraser />} compact />
    </ToolbarGroup>
  );
}

function FormulaGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Function Library">
      <SplitButton
        label="AutoSum"
        icon={<Sigma />}
        onPrimary={() => props.onAutoSum()}
        items={AUTO_FUNCTION_MENU_ITEMS.map((item) => ({
          label: item.label,
          onSelect: () => props.onAutoSum(item.value)
        }))}
      />
      <ToolbarButton
        label="Function library"
        onClick={props.onFunctionLibrary}
        icon={<SquareFunction />}
        expanded={props.functionLibraryOpen}
        compact
      />
      <ToolbarButton
        label="Audit formulas"
        onClick={props.onFormulaAudit}
        icon={<GitBranch />}
        expanded={props.formulaAuditOpen}
        compact
      />
      <ToolbarButton
        label="Named ranges"
        onClick={props.onNamedRanges}
        icon={<TableProperties />}
        expanded={props.namedRangesOpen}
        compact
      />
      <ToolbarButton label="Go to" onClick={props.onGoTo} icon={<MapPin />} expanded={props.goToPanelOpen} compact />
    </ToolbarGroup>
  );
}

function ReviewGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Review">
      <ToolbarButton label="Comment" onClick={props.onComment} icon={<MessageSquare />} compact />
      <ToolbarButton label="Link" onClick={props.onLink} icon={<Link />} shortcut="Control+K" compact />
      <ToolbarButton label="Unlink" onClick={props.onUnlink} icon={<Unlink />} compact />
      <ToolbarButton label="Lock cells" onClick={props.onLockCells} icon={<Lock />} compact />
      <ToolbarButton label="Unlock cells" onClick={props.onUnlockCells} icon={<Unlock />} compact />
      <ToolbarButton
        label="Protect sheet"
        onClick={props.onToggleProtectSheet}
        icon={<Shield />}
        pressed={props.sheetProtected}
        compact
      />
    </ToolbarGroup>
  );
}

function DataToolsGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Data Tools">
      <ToolbarButton label="Fill down" onClick={props.onFillDown} icon={<ArrowDown />} shortcut="Control+D" compact />
      <ToolbarButton label="Fill right" onClick={props.onFillRight} icon={<ArrowRight />} shortcut="Control+R" compact />
      <ToolbarButton label="Sort A to Z" onClick={props.onSortAsc} icon={<ArrowDownAZ />} compact />
      <ToolbarButton label="Sort Z to A" onClick={props.onSortDesc} icon={<ArrowDownZA />} compact />
      <ToolbarButton label="Remove duplicates" onClick={props.onRemoveDuplicates} icon={<ListX />} compact />
      <ToolbarButton label="Filter" onClick={props.onFilter} icon={<Filter />} expanded={props.filterPanelOpen} shortcut="Control+Shift+L" compact />
      <ToolbarButton
        label="Find and replace"
        onClick={props.onFindReplace}
        icon={<Search />}
        expanded={props.findPanelOpen}
        shortcut="Control+F"
        compact
      />
      <ToolbarButton
        label="Data validation"
        onClick={props.onDataValidation}
        icon={<ListChecks />}
        expanded={props.validationPanelOpen}
        compact
      />
      <ToolbarButton
        label="Conditional formatting"
        onClick={props.onConditionalFormatting}
        icon={<Sparkles />}
        expanded={props.conditionalPanelOpen}
        compact
      />
    </ToolbarGroup>
  );
}

function InsertGroup(props: ToolbarProps) {
  return (
    <>
      <ToolbarGroup label="Tables">
        <ToolbarButton label="Pivot table" onClick={props.onPivot} icon={<TableProperties />} expanded={props.pivotPanelOpen} compact />
      </ToolbarGroup>
      <ToolbarGroup label="Charts">
        <ToolbarButton label="Chart" onClick={props.onChart} icon={<ChartColumn />} expanded={props.chartPanelOpen} compact />
      </ToolbarGroup>
      <ToolbarGroup label="Links">
        <ToolbarButton label="Link" onClick={props.onLink} icon={<Link />} shortcut="Control+K" compact />
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
    </ToolbarGroup>
  );
}

function FontGroup(props: ToolbarProps) {
  const fontFamilyMixed = props.selectionFormat.fontFamily === "mixed";
  const fontSizeMixed = props.selectionFormat.fontSize === "mixed";
  const currentFamily = fontFamilyMixed ? undefined : (props.selectionFormat.fontFamily as string | undefined);
  const currentSize = fontSizeMixed ? undefined : (props.selectionFormat.fontSize as number | undefined);
  // Imported files can carry fonts outside our list; surface them instead of
  // showing a misleading "Default font" placeholder.
  const familyOptions =
    currentFamily && !FONT_FAMILY_OPTIONS.some((option) => option.value === currentFamily)
      ? [...FONT_FAMILY_OPTIONS, { value: currentFamily, label: currentFamily }]
      : FONT_FAMILY_OPTIONS;
  const sizeOptions =
    currentSize !== undefined && !FONT_SIZE_OPTIONS.some((option) => option.value === String(currentSize))
      ? [...FONT_SIZE_OPTIONS, { value: String(currentSize), label: String(currentSize) }]
      : FONT_SIZE_OPTIONS;
  return (
    <ToolbarGroup label="Font">
      <SelectControl
        label="Font family"
        value={fontFamilyMixed ? MIXED_SELECT_VALUE : (currentFamily ?? "")}
        mixed={fontFamilyMixed}
        options={familyOptions}
        compact
        className="font-family-select"
        onChange={(value) => props.onFontFamily(value)}
      />
      <SelectControl
        label="Font size"
        value={fontSizeMixed ? MIXED_SELECT_VALUE : String(currentSize ?? "")}
        mixed={fontSizeMixed}
        options={sizeOptions}
        compact
        onChange={(value) => props.onFontSize(value === "" ? null : Number(value))}
      />
      <ToolbarButton
        label="Bold"
        onClick={props.onBold}
        icon={<Bold />}
        pressed={props.selectionFormat.bold}
        shortcut="Control+B"
        compact
      />
      <ToolbarButton
        label="Italic"
        onClick={props.onItalic}
        icon={<Italic />}
        pressed={props.selectionFormat.italic}
        shortcut="Control+I"
        compact
      />
      <SplitButton
        label="All borders"
        icon={<Grid2X2 />}
        onPrimary={() => props.onBorders("all")}
        items={BORDER_MENU_ITEMS.map((item) => ({ label: item.label, onSelect: () => props.onBorders(item.value) }))}
      />
      <ColorControl
        label="Text color"
        icon={<Type />}
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
  const numberFormatMixed = props.selectionFormat.numberFormat === "mixed";
  return (
    <ToolbarGroup label="Number">
      <ToolbarButton
        label="Currency format"
        onClick={() => props.onNumberFormat("currency")}
        icon={<span className="glyph-icon">$</span>}
        pressed={props.selectionFormat.numberFormat === "currency"}
        compact
      />
      <ToolbarButton
        label="Percent format"
        onClick={() => props.onNumberFormat("percent")}
        icon={<span className="glyph-icon">%</span>}
        pressed={props.selectionFormat.numberFormat === "percent"}
        compact
      />
      <ToolbarButton
        label="Comma style"
        onClick={() => props.onNumberFormat("number")}
        icon={<span className="glyph-icon">,</span>}
        pressed={props.selectionFormat.numberFormat === "number"}
        compact
      />
      <SelectControl
        label="Number format"
        value={
          numberFormatMixed
            ? MIXED_SELECT_VALUE
            : ((props.selectionFormat.numberFormat as string | undefined) ?? "general")
        }
        mixed={numberFormatMixed}
        options={NUMBER_FORMAT_OPTIONS}
        onChange={(value) => props.onNumberFormat(value as NonNullable<CellFormat["numberFormat"]>)}
      />
    </ToolbarGroup>
  );
}

function AlignmentGroup(props: ToolbarProps) {
  const horizontal = props.selectionFormat.horizontalAlign === "mixed" ? null : (props.selectionFormat.horizontalAlign ?? "left");
  const vertical = props.selectionFormat.verticalAlign === "mixed" ? null : (props.selectionFormat.verticalAlign ?? "middle");
  return (
    <ToolbarGroup label="Alignment">
      {HORIZONTAL_ALIGN_BUTTONS.map((option) => (
        <ToolbarButton
          key={option.value}
          label={option.label}
          onClick={() => props.onHorizontalAlign(option.value)}
          icon={option.icon}
          pressed={horizontal === option.value}
          compact
        />
      ))}
      {VERTICAL_ALIGN_BUTTONS.map((option) => (
        <ToolbarButton
          key={option.value}
          label={option.label}
          onClick={() => props.onVerticalAlign(option.value)}
          icon={option.icon}
          pressed={vertical === option.value}
          compact
        />
      ))}
      <ToolbarButton
        label="Wrap text"
        onClick={props.onWrapText}
        icon={<WrapText />}
        pressed={props.selectionFormat.wrapText}
        compact
      />
    </ToolbarGroup>
  );
}

function StylesGroup(props: ToolbarProps) {
  return (
    <ToolbarGroup label="Styles">
      <ToolbarButton
        label="Conditional formatting"
        onClick={props.onConditionalFormatting}
        icon={<Sparkles />}
        expanded={props.conditionalPanelOpen}
        compact
      />
      <ToolbarButton
        label="Data validation"
        onClick={props.onDataValidation}
        icon={<ListChecks />}
        expanded={props.validationPanelOpen}
        compact
      />
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
      <SplitButton
        label="AutoSum"
        icon={<Sigma />}
        onPrimary={() => props.onAutoSum()}
        items={AUTO_FUNCTION_MENU_ITEMS.map((item) => ({
          label: item.label,
          onSelect: () => props.onAutoSum(item.value)
        }))}
      />
      <ToolbarButton label="Fill down" onClick={props.onFillDown} icon={<ArrowDown />} shortcut="Control+D" compact />
      <ToolbarButton label="Fill right" onClick={props.onFillRight} icon={<ArrowRight />} shortcut="Control+R" compact />
      <ToolbarButton label="Sort A to Z" onClick={props.onSortAsc} icon={<ArrowDownAZ />} compact />
      <ToolbarButton label="Sort Z to A" onClick={props.onSortDesc} icon={<ArrowDownZA />} compact />
      <ToolbarButton label="Remove duplicates" onClick={props.onRemoveDuplicates} icon={<ListX />} compact />
      <ToolbarButton label="Filter" onClick={props.onFilter} icon={<Filter />} expanded={props.filterPanelOpen} shortcut="Control+Shift+L" compact />
      <ToolbarButton
        label="Find and replace"
        onClick={props.onFindReplace}
        icon={<Search />}
        expanded={props.findPanelOpen}
        shortcut="Control+F"
        compact
      />
      <ToolbarButton label="Go to" onClick={props.onGoTo} icon={<MapPin />} expanded={props.goToPanelOpen} compact />
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
        pressed={!props.showGridlines}
        compact
      />
      <ToolbarButton
        label={props.showHeaders ? "Hide headers" : "Show headers"}
        onClick={props.onToggleHeaders}
        icon={<PanelTop />}
        pressed={!props.showHeaders}
        compact
      />
      <ToolbarButton
        label={props.showFormulaBar ? "Hide formula bar" : "Show formula bar"}
        onClick={props.onToggleFormulaBar}
        icon={<SquareFunction />}
        pressed={!props.showFormulaBar}
        compact
      />
      <ToolbarButton
        label={props.showFormulas ? "Show formula results" : "Show formulas"}
        onClick={props.onToggleShowFormulas}
        icon={<Sigma />}
        pressed={props.showFormulas}
        compact
      />
      <ToolbarButton
        label={props.showSheetTabs ? "Hide sheet tabs" : "Show sheet tabs"}
        onClick={props.onToggleSheetTabs}
        icon={<PanelBottom />}
        pressed={!props.showSheetTabs}
        compact
      />
      <ToolbarButton
        label="Freeze top row"
        onClick={props.onToggleFreezeTopRow}
        icon={<PanelTop />}
        pressed={props.freezeTopRow}
        compact
      />
      <ToolbarButton
        label="Freeze first column"
        onClick={props.onToggleFreezeFirstColumn}
        icon={<PanelLeft />}
        pressed={props.freezeFirstColumn}
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
  pressed,
  expanded,
  shortcut,
  compact = false
}: {
  label: string;
  onClick: () => void;
  icon: ReactNode;
  disabled?: boolean;
  /** Toggle state; "mixed" when the selection disagrees. Renders aria-pressed. */
  pressed?: boolean | "mixed";
  /** Panel-launcher state. Renders aria-expanded and the active style while open. */
  expanded?: boolean;
  /** Shown in the tooltip and exposed via aria-keyshortcuts. */
  shortcut?: string;
  compact?: boolean;
}) {
  const isActive = pressed === true || expanded === true;
  const isMixed = pressed === "mixed";
  const className = [
    "toolbar-button",
    isActive ? "active-toolbar-button" : "",
    isMixed ? "mixed-toolbar-button" : "",
    compact ? "compact-toolbar-button" : ""
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <button
      type="button"
      className={className}
      aria-label={label}
      aria-pressed={pressed === undefined ? undefined : isMixed ? "mixed" : pressed}
      aria-expanded={expanded === undefined ? undefined : expanded}
      aria-keyshortcuts={shortcut}
      data-shortcut={shortcut ? shortcutHint(shortcut) : undefined}
      // Compact buttons get an instant CSS tooltip from aria-label; keeping the
      // native title too would show a second, delayed tooltip on top of it.
      title={compact ? undefined : label}
      onClick={onClick}
      disabled={disabled}
    >
      {icon}
      <span>{label}</span>
    </button>
  );
}

function shortcutHint(shortcut: string): string {
  return shortcut.replace(/Control/g, "Ctrl").replace(/\+/g, "+");
}

function SplitButton({
  label,
  icon,
  onPrimary,
  items,
  disabled = false
}: {
  label: string;
  icon: ReactNode;
  onPrimary: () => void;
  items: Array<{ label: string; onSelect: () => void }>;
  disabled?: boolean;
}) {
  const [open, setOpen] = useState(false);
  // The menu is position:fixed so the ribbon's overflow scroll container
  // (mobile) cannot clip it; coordinates are captured when it opens.
  const [menuPosition, setMenuPosition] = useState<{ top: number; left: number } | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const toggleRef = useRef<HTMLButtonElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) {
      return undefined;
    }
    function handleDocumentMouseDown(event: MouseEvent) {
      if (!containerRef.current?.contains(event.target as Node)) {
        setOpen(false);
      }
    }
    // A fixed-position menu would float away from its anchor on scroll/resize.
    function handleViewportChange() {
      setOpen(false);
    }
    document.addEventListener("mousedown", handleDocumentMouseDown);
    window.addEventListener("scroll", handleViewportChange, true);
    window.addEventListener("resize", handleViewportChange);
    return () => {
      document.removeEventListener("mousedown", handleDocumentMouseDown);
      window.removeEventListener("scroll", handleViewportChange, true);
      window.removeEventListener("resize", handleViewportChange);
    };
  }, [open]);

  useEffect(() => {
    if (open) {
      // Keyboard users land inside the menu; arrows then cycle the items
      // instead of falling through to grid navigation.
      menuRef.current?.querySelector("button")?.focus();
    }
  }, [open]);

  function openMenu() {
    const rect = containerRef.current?.getBoundingClientRect();
    setMenuPosition(rect ? { top: rect.bottom + 4, left: rect.left } : null);
    setOpen(true);
  }

  function closeMenu(refocusToggle: boolean) {
    setOpen(false);
    if (refocusToggle) {
      toggleRef.current?.focus();
    }
  }

  function handleMenuKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      closeMenu(true);
      return;
    }
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      const buttons = Array.from(menuRef.current?.querySelectorAll("button") ?? []);
      if (buttons.length === 0) {
        return;
      }
      const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
      const nextIndex =
        event.key === "ArrowDown"
          ? (index + 1) % buttons.length
          : (index - 1 + buttons.length) % buttons.length;
      buttons[nextIndex]?.focus();
      return;
    }
    if (event.key === "Tab") {
      setOpen(false);
    }
  }

  return (
    <div className="split-button" ref={containerRef}>
      <button
        type="button"
        className="toolbar-button compact-toolbar-button split-button-primary"
        aria-label={label}
        onClick={() => {
          setOpen(false);
          onPrimary();
        }}
        disabled={disabled}
      >
        {icon}
        <span>{label}</span>
      </button>
      <button
        type="button"
        ref={toggleRef}
        className="split-button-toggle"
        aria-label={`${label} options`}
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => (open ? closeMenu(false) : openMenu())}
        onKeyDown={(event) => {
          if (event.key === "ArrowDown" && !open) {
            event.preventDefault();
            openMenu();
          }
          if (event.key === "Escape" && open) {
            event.preventDefault();
            event.stopPropagation();
            closeMenu(true);
          }
        }}
        disabled={disabled}
      >
        <ChevronDown aria-hidden="true" />
      </button>
      {open ? (
        <div
          className="split-button-menu"
          role="menu"
          aria-label={`${label} menu`}
          ref={menuRef}
          style={menuPosition ? { top: menuPosition.top, left: menuPosition.left } : undefined}
          onKeyDown={handleMenuKeyDown}
        >
          {items.map((item) => (
            <button
              key={item.label}
              type="button"
              role="menuitem"
              onClick={() => {
                closeMenu(true);
                item.onSelect();
              }}
            >
              {item.label}
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function SelectControl({
  label,
  value,
  options,
  onChange,
  mixed = false,
  compact = false,
  className
}: {
  label: string;
  value: string;
  options: readonly { value: string; label: string }[];
  onChange: (value: string) => void;
  /** When the selection disagrees, show a disabled "Mixed" entry as the value. */
  mixed?: boolean;
  compact?: boolean;
  className?: string;
}) {
  return (
    <label
      className={["select-control", compact ? "compact-select-control" : "", className ?? ""].filter(Boolean).join(" ")}
      title={label}
    >
      <span>{label}</span>
      <select
        aria-label={label}
        value={value}
        onChange={(event) => {
          if (event.currentTarget.value !== MIXED_SELECT_VALUE) {
            onChange(event.currentTarget.value);
          }
        }}
      >
        {mixed ? (
          <option value={MIXED_SELECT_VALUE} disabled>
            Mixed
          </option>
        ) : null}
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
