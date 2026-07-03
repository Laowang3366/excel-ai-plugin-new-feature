import type { SqliteDatabase } from "../storage/nodeSqlite";

import type { RolloutItem, ThreadId } from "../shared/types";

export type RuntimeDbName = "state" | "logs" | "goals" | "memories";

export interface RuntimeDatabasePaths {
  state: string;
  logs: string;
  goals: string;
  memories: string;
}

export interface ResolvedRuntimePaths {
  dbPaths: RuntimeDatabasePaths;
  legacyStateDbPath?: string;
}

export interface RuntimeConnections {
  state: SqliteDatabase;
  logs: SqliteDatabase;
  goals: SqliteDatabase;
  memories: SqliteDatabase;
}

export interface RuntimeRecoveryReport {
  dbName: RuntimeDbName;
  dbPath: string;
  reason: string;
  backupPaths: string[];
  recoveredAt: number;
}

export interface RuntimeMigration {
  id: string;
  sql: string;
}

export interface RuntimeGoalRecord {
  goalId: string;
  threadId?: ThreadId;
  objective: string;
  status: "active" | "complete" | "blocked";
  tokenBudget?: number;
  tokenUsage?: number;
  createdAt: number;
  updatedAt: number;
  completedAt?: number;
  payload?: Record<string, unknown>;
}

export type RuntimeMemoryKind =
  | "preference"
  | "constraint"
  | "correction"
  | "style_preference"
  | "operation_preference"
  | "file_impression"
  | "tool_success_profile"
  | "project_fact"
  | "workflow";

export type RuntimeMemoryVisibility = "user" | "internal";
export type RuntimeMemoryStatus = "active" | "stale" | "archived";

export interface RuntimeMemoryCitation {
  threadId: ThreadId;
  eventId?: number;
  turnId?: string;
}

export interface RuntimeMemoryRecord {
  memoryId: string;
  namespace: string;
  content: string;
  metadata?: Record<string, unknown>;
  createdAt: number;
  updatedAt: number;
}

export interface RuntimeLongTermMemoryRecord {
  memoryId: string;
  namespace: string;
  kind: RuntimeMemoryKind;
  visibility: RuntimeMemoryVisibility;
  status: RuntimeMemoryStatus;
  content: string;
  summary?: string;
  confidence?: number;
  sourceThreadId?: ThreadId;
  sourceEventId?: number;
  workspaceFingerprint?: string;
  expiresAt?: number;
  metadata?: Record<string, unknown>;
  citations?: RuntimeMemoryCitation[];
  createdAt: number;
  updatedAt: number;
}

export interface RuntimeMemoryListOptions {
  namespace?: string;
  kind?: RuntimeMemoryKind;
  visibility?: RuntimeMemoryVisibility;
  status?: RuntimeMemoryStatus;
  limit?: number;
  offset?: number;
}

export interface RuntimeRolloutEvent {
  id: number;
  threadId: ThreadId;
  turnId?: string;
  itemType: string;
  timestamp: string;
  item: RolloutItem;
}

export interface RuntimeRolloutSearchMatch extends RuntimeRolloutEvent {
  snippet: string;
}

export type RuntimeToolExecutionStatus = "success" | "error" | "cancelled" | "blocked";

export interface RuntimeToolExecutionLogRecord {
  id?: number;
  threadId: ThreadId;
  turnId: string;
  toolCallId: string;
  toolName: string;
  status: RuntimeToolExecutionStatus;
  durationMs: number;
  timestamp: number;
  argumentsSummary: string;
  resultSummary: string;
  error?: string;
  metadata?: Record<string, unknown>;
}
