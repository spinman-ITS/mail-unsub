export type UnsubscribeAuditEntry = {
  id: string;
  timestamp: string;
  method: "GET" | "POST";
  target: string;
  outcome: "accepted" | "blocked" | "failed";
  status?: number;
  message: string;
  redirectTarget?: string;
};

const maxEntries = 50;
const entries: UnsubscribeAuditEntry[] = [];

export function addAuditEntry(entry: Omit<UnsubscribeAuditEntry, "id" | "timestamp">): UnsubscribeAuditEntry {
  const saved = {
    id: crypto.randomUUID(),
    timestamp: new Date().toISOString(),
    ...entry
  };

  entries.unshift(saved);
  entries.splice(maxEntries);

  console.log(
    `[unsubscribe] ${saved.outcome} ${saved.method} ${saved.target} status=${saved.status ?? "n/a"} message="${saved.message}"${
      saved.redirectTarget ? ` redirect=${saved.redirectTarget}` : ""
    }`
  );

  return saved;
}

export function getAuditEntries(): UnsubscribeAuditEntry[] {
  return entries;
}

export function redactUrl(value: string): string {
  try {
    const url = new URL(value);
    return `${url.protocol}//${url.host}${url.pathname}${url.search ? "?..." : ""}`;
  } catch {
    return "invalid-url";
  }
}
