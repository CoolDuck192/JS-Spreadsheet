import { forwardRef, useEffect, useMemo, useRef, useState } from "react";
import type { FilterExpression, QueryScalar } from "../table/core/query";
import type { StructuredTableColumn } from "../types";

type ScalarType = QueryScalar["type"];
type ComparisonOperator = Extract<FilterExpression, { kind: "comparison" }>["operator"];
type SetOperator = Extract<FilterExpression, { kind: "set" }>["operator"];
type RangeOperator = Extract<FilterExpression, { kind: "range" }>["operator"];
type BlankOperator = Extract<FilterExpression, { kind: "blank" }>["operator"];
type LogicalOperator = Extract<FilterExpression, { kind: "logical" }>["operator"];

export type StructuredTableFilterPanelProps = {
  columns: readonly StructuredTableColumn[];
  activeColumnId?: string;
  filter?: FilterExpression;
  onApply(filter?: FilterExpression): void;
  onClose?(): void;
};

export function StructuredTableFilterPanel({
  columns,
  activeColumnId,
  filter,
  onApply,
  onClose
}: StructuredTableFilterPanelProps) {
  const initialColumnId = activeColumnId && columns.some((column) => column.id === activeColumnId)
    ? activeColumnId
    : firstLeafColumnId(filter) ?? columns[0]?.id ?? "";
  const [kind, setKind] = useState<FilterExpression["kind"]>(filter?.kind ?? "comparison");
  const [columnId, setColumnId] = useState(initialColumnId);
  const [scalarType, setScalarType] = useState<ScalarType>("string");
  const [comparisonOperator, setComparisonOperator] = useState<ComparisonOperator>("eq");
  const [setOperator, setSetOperator] = useState<SetOperator>("in");
  const [rangeOperator, setRangeOperator] = useState<RangeOperator>("between");
  const [blankOperator, setBlankOperator] = useState<BlankOperator>("isBlank");
  const [logicalOperator, setLogicalOperator] = useState<LogicalOperator>("and");
  const [value, setValue] = useState("");
  const [secondValue, setSecondValue] = useState("");
  const [setValues, setSetValues] = useState("");
  const [error, setError] = useState("");
  const firstValueRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setColumnId(initialColumnId);
  }, [initialColumnId]);

  const selectedColumnExists = useMemo(
    () => columns.some((column) => column.id === columnId),
    [columnId, columns]
  );

  function applyFilter() {
    if (!selectedColumnExists) {
      setError("Choose a table column");
      return;
    }

    try {
      const nextFilter = buildFilter({
        kind,
        columnId,
        scalarType,
        comparisonOperator,
        setOperator,
        rangeOperator,
        blankOperator,
        logicalOperator,
        value,
        secondValue,
        setValues
      });
      setError("");
      onApply(nextFilter);
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : "Filter values are invalid");
      firstValueRef.current?.focus();
    }
  }

  return (
    <section className="structured-table-popover" aria-label="Table filter">
      <label className="structured-table-field">
        <span>Filter type</span>
        <select
          aria-label="Filter type"
          value={kind}
          onChange={(event) => {
            setKind(event.currentTarget.value as FilterExpression["kind"]);
            setError("");
          }}
        >
          <option value="comparison">Comparison</option>
          <option value="set">Set</option>
          <option value="range">Range</option>
          <option value="blank">Blank</option>
          <option value="logical">Logical</option>
          <option value="not">Not</option>
        </select>
      </label>
      <label className="structured-table-field">
        <span>Column</span>
        <select aria-label="Filter column" value={columnId} onChange={(event) => setColumnId(event.currentTarget.value)}>
          {columns.map((column) => (
            <option key={column.id} value={column.id}>
              {column.name}
            </option>
          ))}
        </select>
      </label>
      {kind !== "blank" ? (
        <label className="structured-table-field">
          <span>Value type</span>
          <select
            aria-label="Value type"
            value={scalarType}
            onChange={(event) => setScalarType(event.currentTarget.value as ScalarType)}
          >
            <option value="string">Text</option>
            <option value="number">Number</option>
            <option value="boolean">Boolean</option>
            <option value="date">Date</option>
            <option value="datetime">Date and time</option>
            <option value="error">Error</option>
            <option value="null">Null</option>
          </select>
        </label>
      ) : null}
      {kind === "comparison" ? (
        <>
          <OperatorSelect
            label="Comparison operator"
            value={comparisonOperator}
            options={["eq", "neq", "contains", "startsWith", "endsWith", "gt", "gte", "lt", "lte"]}
            onChange={(next) => setComparisonOperator(next as ComparisonOperator)}
          />
          <ValueInput ref={firstValueRef} label="Filter value" value={value} disabled={scalarType === "null"} onChange={setValue} />
        </>
      ) : null}
      {kind === "set" ? (
        <>
          <OperatorSelect label="Set operator" value={setOperator} options={["in", "notIn"]} onChange={(next) => setSetOperator(next as SetOperator)} />
          <ValueInput ref={firstValueRef} label="Set values" value={setValues} disabled={scalarType === "null"} onChange={setSetValues} />
        </>
      ) : null}
      {kind === "range" ? (
        <>
          <OperatorSelect
            label="Range operator"
            value={rangeOperator}
            options={["between", "notBetween"]}
            onChange={(next) => setRangeOperator(next as RangeOperator)}
          />
          <ValueInput ref={firstValueRef} label="Lower value" value={value} disabled={scalarType === "null"} onChange={setValue} />
          <ValueInput label="Upper value" value={secondValue} disabled={scalarType === "null"} onChange={setSecondValue} />
        </>
      ) : null}
      {kind === "blank" ? (
        <OperatorSelect
          label="Blank operator"
          value={blankOperator}
          options={["isNull", "isNotNull", "isEmpty", "isNotEmpty", "isBlank", "isNotBlank"]}
          onChange={(next) => setBlankOperator(next as BlankOperator)}
        />
      ) : null}
      {kind === "logical" ? (
        <>
          <OperatorSelect label="Logical operator" value={logicalOperator} options={["and", "or"]} onChange={(next) => setLogicalOperator(next as LogicalOperator)} />
          <ValueInput ref={firstValueRef} label="First comparison value" value={value} disabled={scalarType === "null"} onChange={setValue} />
          <ValueInput label="Second comparison value" value={secondValue} disabled={scalarType === "null"} onChange={setSecondValue} />
        </>
      ) : null}
      {kind === "not" ? (
        <ValueInput ref={firstValueRef} label="Comparison value" value={value} disabled={scalarType === "null"} onChange={setValue} />
      ) : null}
      {error ? (
        <p className="structured-table-field-error" role="alert">
          {error}
        </p>
      ) : null}
      <div className="structured-table-actions">
        <button type="button" onClick={applyFilter} disabled={!selectedColumnExists}>
          Apply table filter
        </button>
        <button
          type="button"
          onClick={() => {
            setError("");
            onApply(undefined);
          }}
        >
          Clear table filter
        </button>
        {onClose ? (
          <button type="button" onClick={onClose}>
            Close table filter
          </button>
        ) : null}
      </div>
    </section>
  );
}

function OperatorSelect({
  label,
  value,
  options,
  onChange
}: {
  label: string;
  value: string;
  options: readonly string[];
  onChange(value: string): void;
}) {
  return (
    <label className="structured-table-field">
      <span>{label}</span>
      <select aria-label={label} value={value} onChange={(event) => onChange(event.currentTarget.value)}>
        {options.map((option) => (
          <option key={option} value={option}>
            {operatorLabel(option)}
          </option>
        ))}
      </select>
    </label>
  );
}

const ValueInput = forwardRef<HTMLInputElement, {
  label: string;
  value: string;
  disabled: boolean;
  onChange(value: string): void;
}>(function ValueInput({ label, value, disabled, onChange }, ref) {
  return (
    <label className="structured-table-field">
      <span>{label}</span>
      <input
        ref={ref}
        aria-label={label}
        value={value}
        disabled={disabled}
        onChange={(event) => onChange(event.currentTarget.value)}
      />
    </label>
  );
});

function buildFilter(input: {
  kind: FilterExpression["kind"];
  columnId: string;
  scalarType: ScalarType;
  comparisonOperator: ComparisonOperator;
  setOperator: SetOperator;
  rangeOperator: RangeOperator;
  blankOperator: BlankOperator;
  logicalOperator: LogicalOperator;
  value: string;
  secondValue: string;
  setValues: string;
}): FilterExpression {
  const comparison = (rawValue: string): Extract<FilterExpression, { kind: "comparison" }> => ({
    kind: "comparison",
    columnId: input.columnId,
    operator: input.comparisonOperator,
    value: parseScalar(input.scalarType, rawValue)
  });

  switch (input.kind) {
    case "comparison":
      return comparison(input.value);
    case "set": {
      const rawValues = input.scalarType === "null"
        ? [""]
        : input.setValues.split(",").map((value) => value.trim()).filter(Boolean);
      if (rawValues.length === 0) {
        throw new Error("Enter at least one set value");
      }
      return {
        kind: "set",
        columnId: input.columnId,
        operator: input.setOperator,
        values: rawValues.map((value) => parseScalar(input.scalarType, value))
      };
    }
    case "range":
      return {
        kind: "range",
        columnId: input.columnId,
        operator: input.rangeOperator,
        lower: parseScalar(input.scalarType, input.value),
        upper: parseScalar(input.scalarType, input.secondValue)
      };
    case "blank":
      return { kind: "blank", columnId: input.columnId, operator: input.blankOperator };
    case "logical":
      return {
        kind: "logical",
        operator: input.logicalOperator,
        operands: [comparison(input.value), comparison(input.secondValue)]
      };
    case "not":
      return { kind: "not", operand: comparison(input.value) };
  }
}

function parseScalar(type: ScalarType, rawValue: string): QueryScalar {
  switch (type) {
    case "number": {
      const value = Number(rawValue);
      if (!rawValue.trim() || !Number.isFinite(value)) {
        throw new Error("Enter a valid number");
      }
      return { type, value };
    }
    case "boolean":
      if (!/^(true|false)$/i.test(rawValue.trim())) {
        throw new Error("Enter true or false");
      }
      return { type, value: rawValue.trim().toLowerCase() === "true" };
    case "date":
    case "datetime":
    case "string":
    case "error":
      return { type, value: rawValue.trim() };
    case "null":
      return { type };
  }
}

function firstLeafColumnId(filter?: FilterExpression): string | undefined {
  if (!filter) return undefined;
  if (filter.kind === "logical") return firstLeafColumnId(filter.operands[0]);
  if (filter.kind === "not") return firstLeafColumnId(filter.operand);
  return filter.columnId;
}

function operatorLabel(operator: string): string {
  return operator.replace(/([A-Z])/g, " $1").replace(/^./, (character) => character.toUpperCase());
}
