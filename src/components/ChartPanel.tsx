import { useEffect, useState } from "react";
import { X } from "lucide-react";
import type { SheetChartType } from "../types";

type ChartPanelProps = {
  isOpen: boolean;
  onClose: () => void;
  onCreate: (config: { type: SheetChartType; title: string }) => void;
};

export function ChartPanel({ isOpen, onClose, onCreate }: ChartPanelProps) {
  const [type, setType] = useState<SheetChartType>("bar");
  const [title, setTitle] = useState("");

  useEffect(() => {
    if (isOpen) {
      setType("bar");
      setTitle("");
    }
  }, [isOpen]);

  if (!isOpen) {
    return null;
  }

  return (
    <aside className="js-spreadsheet-chart-panel" aria-label="Chart builder">
      <div className="js-spreadsheet-chart-panel-header">
        <strong>Chart</strong>
        <button type="button" aria-label="Close chart builder" onClick={onClose}>
          <X />
        </button>
      </div>
      <label>
        Type
        <select
          aria-label="Chart type"
          value={type}
          onChange={(event) => setType(event.currentTarget.value as SheetChartType)}
        >
          <option value="bar">Bar</option>
          <option value="line">Line</option>
          <option value="pie">Pie</option>
        </select>
      </label>
      <label>
        Title
        <input
          aria-label="Chart title"
          value={title}
          onChange={(event) => setTitle(event.currentTarget.value)}
          placeholder="Auto"
        />
      </label>
      <div className="js-spreadsheet-chart-panel-actions">
        <button type="button" onClick={onClose}>
          Cancel
        </button>
        <button type="button" className="js-spreadsheet-primary-panel-button" onClick={() => onCreate({ type, title: title.trim() })}>
          Create chart
        </button>
      </div>
    </aside>
  );
}
