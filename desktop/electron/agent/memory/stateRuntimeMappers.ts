import type { RolloutItem } from "../shared/types";
import type {
  RuntimeGoalRecord,
  RuntimeLongTermMemoryRecord,
  RuntimeMemoryRecord,
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
