import type { RolloutItem, ThreadMetadata } from "../shared/types";
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
  return {
    goalId: row.goal_id,
    threadId: row.thread_id ?? undefined,
    objective: row.objective,
    status: row.status,
    tokenBudget: row.token_budget ?? undefined,
    tokenUsage: row.token_usage ?? undefined,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    completedAt: row.completed_at ?? undefined,
    payload: row.payload_json ? JSON.parse(row.payload_json) : undefined,
  };
}

export function mapMemory(row: Record<string, any>): RuntimeMemoryRecord {
  return {
    memoryId: row.memory_id,
    namespace: row.namespace,
    content: row.content,
    metadata: row.metadata_json ? JSON.parse(row.metadata_json) : undefined,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}

export function mapLongTermMemory(row: Record<string, any>): RuntimeLongTermMemoryRecord {
  return {
    memoryId: row.memory_id,
    namespace: row.namespace,
    kind: row.kind,
    visibility: row.visibility,
    status: row.status,
    content: row.content,
    summary: row.summary ?? undefined,
    confidence: row.confidence ?? undefined,
    sourceThreadId: row.source_thread_id ?? undefined,
    sourceEventId: row.source_event_id ?? undefined,
    workspaceFingerprint: row.workspace_fingerprint ?? undefined,
    expiresAt: row.expires_at ?? undefined,
    metadata: row.metadata_json ? JSON.parse(row.metadata_json) : undefined,
    citations: row.citations_json ? JSON.parse(row.citations_json) : undefined,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}

export function mapThreadSnapshot(row: Record<string, any>): ThreadMetadata {
  return {
    threadId: row.thread_id,
    preview: row.preview,
    name: row.name ?? undefined,
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
    compactedHistory: row.compacted_history ? JSON.parse(row.compacted_history) : undefined,
  };
}

export function mapToolExecutionLog(row: Record<string, any>): RuntimeToolExecutionLogRecord {
  return {
    id: row.id,
    threadId: row.thread_id,
    turnId: row.turn_id,
    toolCallId: row.tool_call_id,
    toolName: row.tool_name,
    status: row.status,
    durationMs: row.duration_ms,
    timestamp: row.timestamp,
    argumentsSummary: row.arguments_summary,
    resultSummary: row.result_summary,
    error: row.error ?? undefined,
    metadata: row.metadata_json ? JSON.parse(row.metadata_json) : undefined,
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
