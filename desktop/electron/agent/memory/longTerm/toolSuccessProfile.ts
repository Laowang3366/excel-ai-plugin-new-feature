export type ToolProfileOperation =
  | "open"
  | "create"
  | "inspect"
  | "read"
  | "write"
  | "replace"
  | "format"
  | "style"
  | "chart"
  | "validate"
  | "script"
  | "unknown";

export interface ToolProfileKeyInput {
  app: "excel" | "word" | "powerpoint" | "office";
  operation: ToolProfileOperation;
  toolFamily: "openxml" | "com" | "script" | "shell" | "python" | "office_action" | "other";
}

export interface ToolProfileStats {
  app: ToolProfileKeyInput["app"];
  operation: ToolProfileOperation;
  toolFamily: ToolProfileKeyInput["toolFamily"];
  successCount: number;
  failureCount: number;
  lastUpdatedAt: number;
}

const TOOL_PROFILE_OPERATIONS = new Set<ToolProfileOperation>([
  "open",
  "create",
  "inspect",
  "read",
  "write",
  "replace",
  "format",
  "style",
  "chart",
  "validate",
  "script",
  "unknown",
]);

const TOOL_PROFILE_OPERATION_ALIASES: Record<string, ToolProfileOperation> = {
  edit: "write",
  applystyle: "style",
  apply_style: "style",
  insertchart: "chart",
  insert_chart: "chart",
  readtext: "read",
};

export function normalizeToolProfileOperation(value: unknown): ToolProfileOperation {
  if (typeof value !== "string") {
    return "unknown";
  }

  const operation = value.trim().toLowerCase();
  if (TOOL_PROFILE_OPERATIONS.has(operation as ToolProfileOperation)) {
    return operation as ToolProfileOperation;
  }

  return TOOL_PROFILE_OPERATION_ALIASES[operation] ?? "unknown";
}

export function buildToolProfileKey(input: ToolProfileKeyInput): string {
  return `${input.app}:${normalizeToolProfileOperation(input.operation)}:${input.toolFamily}`;
}

export function shouldPromoteToolProfile(
  stats: Pick<ToolProfileStats, "successCount" | "failureCount">
): boolean {
  return stats.successCount + stats.failureCount >= 3;
}

export function updateToolProfileStats(
  current: ToolProfileStats | undefined,
  event: ToolProfileKeyInput & { success: boolean }
): ToolProfileStats {
  return {
    app: event.app,
    operation: normalizeToolProfileOperation(event.operation),
    toolFamily: event.toolFamily,
    successCount: (current?.successCount ?? 0) + (event.success ? 1 : 0),
    failureCount: (current?.failureCount ?? 0) + (event.success ? 0 : 1),
    lastUpdatedAt: Date.now(),
  };
}
