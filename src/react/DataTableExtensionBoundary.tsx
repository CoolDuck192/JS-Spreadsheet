import { Component, type ErrorInfo, type ReactNode } from "react";
import type { TableDiagnosticEvent } from "../table/core/types";

type DataTableExtensionBoundaryProps = {
  children: ReactNode;
  columnId: string;
  slot: "cell" | "header" | "editor";
  onDiagnostic?(event: TableDiagnosticEvent): void;
};

type DataTableExtensionBoundaryState = { failed: boolean };

export class DataTableExtensionBoundary extends Component<
  DataTableExtensionBoundaryProps,
  DataTableExtensionBoundaryState
> {
  state: DataTableExtensionBoundaryState = { failed: false };

  static getDerivedStateFromError(): DataTableExtensionBoundaryState {
    return { failed: true };
  }

  componentDidCatch(_error: Error, _info: ErrorInfo): void {
    try {
      this.props.onDiagnostic?.({
        category: "extension",
        metadata: { columnId: this.props.columnId, slot: this.props.slot }
      });
    } catch {
      // Host diagnostics never take down the table surface.
    }
  }

  render() {
    return this.state.failed
      ? <span role="status">Cell renderer unavailable</span>
      : this.props.children;
  }
}
