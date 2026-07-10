import type { QueryRequest } from "../core/query";
import type { RemoteMutationController } from "./RemoteMutationController";
import type { RemoteQueryController } from "./RemoteQueryController";
import type { RemoteSourceEvent, RemoteTableSource } from "./types";

export type RemoteSubscriptionControllerOptions<TRow> = {
  source: RemoteTableSource<TRow>;
  queryController: RemoteQueryController<TRow>;
  mutationController: RemoteMutationController<TRow>;
  getActiveQuery(): QueryRequest;
  createOperationId(): string;
};

export class RemoteSubscriptionController<TRow> {
  private readonly source: RemoteTableSource<TRow>;
  private readonly queryController: RemoteQueryController<TRow>;
  private readonly mutationController: RemoteMutationController<TRow>;
  private readonly getActiveQuery: () => QueryRequest;
  private readonly createOperationId: () => string;
  private unsubscribe: (() => void) | null = null;
  private destroyed = false;

  constructor(options: RemoteSubscriptionControllerOptions<TRow>) {
    this.source = options.source;
    this.queryController = options.queryController;
    this.mutationController = options.mutationController;
    this.getActiveQuery = options.getActiveQuery;
    this.createOperationId = options.createOperationId;
  }

  start(): void {
    if (this.destroyed || this.unsubscribe || !this.source.subscribe) return;
    this.unsubscribe = this.source.subscribe((event) => this.accept(event));
  }

  stop(): void {
    this.unsubscribe?.();
    this.unsubscribe = null;
  }

  accept(event: RemoteSourceEvent<TRow>): void {
    if (this.destroyed) return;
    const currentRevision = this.queryController.getCurrentRevision();
    if (event.kind === "invalidate" && event.revision === undefined) {
      this.invalidateAndRefresh();
      return;
    }
    const revision = event.revision;
    if (!revision || !currentRevision) {
      this.invalidateAndRefresh();
      return;
    }
    const comparison = this.source.compareRevisions(revision, currentRevision);
    if (comparison === "older" || comparison === "equal") return;
    if (comparison === "unknown") {
      this.invalidateAndRefresh();
      return;
    }

    if (event.kind === "invalidate") {
      this.queryController.noteCurrentRevision(revision, false);
      this.invalidateAndRefresh();
      return;
    }

    const projected = queryIsProjected(this.getActiveQuery());
    if (event.kind === "rows-upserted") {
      const insertsNewRow = event.rows.some((row) =>
        this.queryController.getCanonicalRow(this.source.getRowId(row)) === undefined
      );
      this.mutationController.acceptAuthoritativeRows(event.rows, revision);
      if (projected || insertsNewRow) {
        this.invalidateAndRefresh();
      } else {
        this.queryController.publishCanonicalState();
      }
      return;
    }

    this.mutationController.acceptAuthoritativeDeletions(event.rowIds, revision);
    if (projected) {
      this.invalidateAndRefresh();
    } else {
      this.queryController.publishCanonicalState();
    }
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.stop();
  }

  private invalidateAndRefresh(): void {
    if (this.destroyed) return;
    this.queryController.invalidate();
    void this.queryController.refresh(this.createOperationId()).catch(() => {});
  }
}

function queryIsProjected(query: QueryRequest): boolean {
  return query.sorting.length > 0
    || query.filter !== null
    || query.grouping.length > 0
    || query.aggregates.length > 0
    || query.pagination.kind !== "none";
}
