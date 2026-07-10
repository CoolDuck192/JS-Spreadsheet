export type CommandEnvelope<TIntent> = {
  id: string;
  intent: TIntent;
  transactionId?: string;
  expectedRevision?: string;
};

export type TableIssue = {
  code: string;
  message: string;
  sheetId?: string;
  address?: string;
};

export type CommandResult<TCurrent = unknown> =
  | { status: "committed"; revision: string; changed?: boolean }
  | { status: "pending"; operationId: string }
  | {
      status: "rejected";
      reason: "validation" | "permission" | "unsupported";
      issues?: readonly TableIssue[];
    }
  | { status: "conflict"; revision: string; current: TCurrent };
