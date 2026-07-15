import type { RolloutItem, ThreadMetadata } from "../shared/types";
import {
  fieldAad,
  unprotectFieldValue,
  unprotectRequiredField,
} from "../../main-modules/localDataProtection/fieldCrypto";
import type {
  RuntimeGoalRecord,
  RuntimeLongTermMemoryRecord,
  RuntimeMemoryRecord,
  RuntimeToolExecutionLogRecord,
} from "./stateRuntimeTypes";

/**
 * StateRuntime 数据行映射。
 *
 * 关联模块：
 * - stateRuntimeStore.ts: 查询 SQLite 行后调用本模块恢复领域对象。
 */
export function getRolloutTurnId(item: RolloutItem): string | null {
  return "turnId" in item ? item.turnId : null;
}

export function mapGoal(row: Record<string, any>): RuntimeGoalRecord {
  const goalId = String(row.goal_id);
  const payloadJson = unprotectFieldValue(
    row.payload_json,
    fieldAad("goals", "goals", goalId, "payload_json"),
  );
  return {
    goalId: row.goal_id,
    threadId: row.thread_id ?? undefined,
    objective: unprotectRequiredField(
      row.objective,
      fieldAad("goals", "goals", goalId, "objective"),
    ),
    status: row.status,
    tokenBudget: row.token_budget ?? undefined,
    tokenUsage: row.token_usage ?? undefined,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    completedAt: row.completed_at ?? undefined,
    payload: payloadJson ? JSON.parse(payloadJson) : undefined,
  };
}

export function mapMemory(row: Record<string, any>): RuntimeMemoryRecord {
  const memoryId = String(row.memory_id);
  const metadataJson = unprotectFieldValue(
    row.metadata_json,
    fieldAad("memories", "memories", memoryId, "metadata_json"),
  );
  return {
    memoryId: row.memory_id,
    namespace: row.namespace,
    content: unprotectRequiredField(
      row.content,
      fieldAad("memories", "memories", memoryId, "content"),
    ),
    metadata: metadataJson ? JSON.parse(metadataJson) : undefined,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}

export function mapLongTermMemory(row: Record<string, any>): RuntimeLongTermMemoryRecord {
  const memoryId = String(row.memory_id);
  const metadataJson = unprotectFieldValue(
    row.metadata_json,
    fieldAad("memories", "long_term_memories", memoryId, "metadata_json"),
  );
  const citationsJson = unprotectFieldValue(
    row.citations_json,
    fieldAad("memories", "long_term_memories", memoryId, "citations_json"),
  );
  const summary = unprotectFieldValue(
    row.summary,
    fieldAad("memories", "long_term_memories", memoryId, "summary"),
  );
  return {
    memoryId: row.memory_id,
    namespace: row.namespace,
    kind: row.kind,
    visibility: row.visibility,
    status: row.status,
    content: unprotectRequiredField(
      row.content,
      fieldAad("memories", "long_term_memories", memoryId, "content"),
    ),
    summary: summary ?? undefined,
    confidence: row.confidence ?? undefined,
    sourceThreadId: row.source_thread_id ?? undefined,
    sourceEventId: row.source_event_id ?? undefined,
    workspaceFingerprint: row.workspace_fingerprint ?? undefined,
    expiresAt: row.expires_at ?? undefined,
    metadata: metadataJson ? JSON.parse(metadataJson) : undefined,
    citations: citationsJson ? JSON.parse(citationsJson) : undefined,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}

export function mapThreadSnapshot(row: Record<string, any>): ThreadMetadata {
  const threadId = String(row.thread_id);
  const compacted = unprotectFieldValue(
    row.compacted_history,
    fieldAad("state", "thread_snapshots", threadId, "compacted_history"),
  );
  return {
    threadId: row.thread_id,
    preview: unprotectRequiredField(
      row.preview ?? "",
      fieldAad("state", "thread_snapshots", threadId, "preview"),
    ),
    name: row.name
      ? unprotectRequiredField(row.name, fieldAad("state", "thread_snapshots", threadId, "name"))
      : undefined,
    modelProvider: row.model_provider,
    model: row.model ?? undefined,
    contextWindowSize: row.context_window_size ?? undefined,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    activeTurnId: row.active_turn_id ?? undefined,
    lastTurnStatus: row.last_turn_status ?? undefined,
    totalTokenUsage: row.total_token_usage ? JSON.parse(row.total_token_usage) : undefined,
    archivedAt: row.archived_at ?? undefined,
    folderId: row.folder_id ?? undefined,
    compactedHistory: compacted ? JSON.parse(compacted) : undefined,
  };
}

export function mapToolExecutionLog(row: Record<string, any>): RuntimeToolExecutionLogRecord {
  const rowId = String(row.id);
  const metadataJson = unprotectFieldValue(
    row.metadata_json,
    fieldAad("logs", "tool_execution_logs", rowId, "metadata_json"),
  );
  return {
    id: row.id,
    threadId: row.thread_id,
    turnId: row.turn_id,
    toolCallId: row.tool_call_id,
    toolName: row.tool_name,
    status: row.status,
    durationMs: row.duration_ms,
    timestamp: row.timestamp,
    argumentsSummary: unprotectRequiredField(
      row.arguments_summary,
      fieldAad("logs", "tool_execution_logs", rowId, "arguments_summary"),
    ),
    resultSummary: unprotectRequiredField(
      row.result_summary,
      fieldAad("logs", "tool_execution_logs", rowId, "result_summary"),
    ),
    error:
      unprotectFieldValue(row.error, fieldAad("logs", "tool_execution_logs", rowId, "error")) ??
      undefined,
    metadata: metadataJson ? JSON.parse(metadataJson) : undefined,
  };
}

export function buildRolloutFtsQuery(query: string): string {
  const terms = query.split(/\s+/).filter(Boolean);
  if (terms.length === 0) return quoteFtsTerm(query);
  return terms.map(quoteFtsTerm).join(" ");
}

export function clampMemoryListOffset(offset: number | undefined): number {
  if (!Number.isFinite(offset)) return 0;
  return Math.max(0, Math.floor(offset as number));
}

function quoteFtsTerm(term: string): string {
  return `"${term.replace(/"/g, '""')}"`;
}
