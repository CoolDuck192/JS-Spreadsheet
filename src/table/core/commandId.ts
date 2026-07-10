export type CommandIdFactory = () => string;

function defaultSessionId(): string {
  if (typeof globalThis.crypto?.randomUUID !== "function") {
    throw new Error("crypto.randomUUID is required to create table command ids");
  }
  return globalThis.crypto.randomUUID();
}

export function createCommandIdFactory(createSessionId: () => string = defaultSessionId): CommandIdFactory {
  const sessionId = createSessionId();
  let sequence = 0;
  return () => `${sessionId}:${++sequence}`;
}
