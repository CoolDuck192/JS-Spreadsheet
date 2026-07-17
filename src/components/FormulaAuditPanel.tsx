import type { ReactElement } from "react";
import { X } from "lucide-react";
import type { CellRange } from "../types";

export type FormulaAuditReference = {
  label: string;
  range: CellRange;
  formula?: string;
};

type FormulaAuditPanelProps = {
  selectionLabel: string;
  activeAddress: string;
  formula: string;
  precedents: FormulaAuditReference[];
  dependents: FormulaAuditReference[];
  onSelectReference: (range: CellRange) => void;
  onClose: () => void;
};

export function FormulaAuditPanel({
  selectionLabel,
  activeAddress,
  formula,
  precedents,
  dependents,
  onSelectReference,
  onClose
}: FormulaAuditPanelProps): ReactElement {
  return (
    <aside className="js-spreadsheet-formula-audit-panel" aria-label="Formula audit">
      <div className="js-spreadsheet-formula-audit-panel-header">
        <strong>Formula Audit</strong>
        <button type="button" aria-label="Close formula audit" onClick={onClose}>
          <X />
        </button>
      </div>
      <div className="js-spreadsheet-formula-audit-target">
        <span>Selection</span>
        <strong>{selectionLabel}</strong>
      </div>
      <section className="js-spreadsheet-formula-audit-section" aria-label="Active formula">
        <h3>{activeAddress}</h3>
        {formula ? <code>{formula}</code> : <p>No formula in active cell</p>}
      </section>
      <FormulaReferenceList title="Precedents" emptyLabel="No precedents" references={precedents} onSelectReference={onSelectReference} />
      <FormulaReferenceList title="Dependents" emptyLabel="No dependents" references={dependents} onSelectReference={onSelectReference} />
    </aside>
  );
}

function FormulaReferenceList({
  title,
  emptyLabel,
  references,
  onSelectReference
}: {
  title: string;
  emptyLabel: string;
  references: FormulaAuditReference[];
  onSelectReference: (range: CellRange) => void;
}) {
  return (
    <section className="js-spreadsheet-formula-audit-section" aria-label={title}>
      <h3>{title}</h3>
      {references.length > 0 ? (
        <div className="js-spreadsheet-formula-audit-list" role="list" aria-label={`${title} list`}>
          {references.map((reference) => (
            <article key={reference.label} className="js-spreadsheet-formula-audit-item" role="listitem">
              <div>
                <strong>{reference.label}</strong>
                {reference.formula ? <code>{reference.formula}</code> : null}
              </div>
              <button type="button" aria-label={`Go to ${reference.label}`} onClick={() => onSelectReference(reference.range)}>
                Go
              </button>
            </article>
          ))}
        </div>
      ) : (
        <p className="js-spreadsheet-formula-audit-empty">{emptyLabel}</p>
      )}
    </section>
  );
}
