import type { SqliteDatabase } from "../storage/nodeSqlite";
import { clampNumber } from "../shared/numberLimits";
import {
  fieldAad,
  protectFieldValue,
  protectRequiredField,
} from "../../main-modules/localDataProtection/fieldCrypto";
import { clampMemoryListOffset, mapLongTermMemory, mapMemory } from "./stateRuntimeMappers";
import type {
  RuntimeLongTermMemoryRecord,
  RuntimeMemoryListOptions,
  RuntimeMemoryRecord,
} from "./stateRuntimeTypes";

export function upsertMemoryInDb(memoriesDb: SqliteDatabase, memory: RuntimeMemoryRecord): void {
  const content = protectRequiredField(
    memory.content,
    fieldAad("memories", "memories", memory.memoryId, "content"),
  );
  const metadataJson = memory.metadata
    ? protectRequiredField(
        JSON.stringify(memory.metadata),
        fieldAad("memories", "memories", memory.memoryId, "metadata_json"),
      )
    : null;
  memoriesDb
    .prepare(
      `INSERT INTO memories (
      memory_id, namespace, content, metadata_json, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT(memory_id) DO UPDATE SET
      namespace = excluded.namespace,
      content = excluded.content,
      metadata_json = excluded.metadata_json,
      updated_at = excluded.updated_at`,
    )
    .run(
      memory.memoryId,
      memory.namespace,
      content,
      metadataJson,
      memory.createdAt,
      memory.updatedAt,
    );
}

export function listMemoriesFromDb(
  memoriesDb: SqliteDatabase,
  namespace?: string,
): RuntimeMemoryRecord[] {
  const rows = namespace
    ? memoriesDb
        .prepare(`SELECT * FROM memories WHERE namespace = ? ORDER BY updated_at DESC`)
        .all(namespace)
    : memoriesDb.prepare(`SELECT * FROM memories ORDER BY updated_at DESC`).all();

  return (rows as Record<string, any>[]).map(mapMemory);
}

export function upsertLongTermMemoryInDb(
  memoriesDb: SqliteDatabase,
  memory: RuntimeLongTermMemoryRecord,
): void {
  memoriesDb
    .prepare(
      `INSERT INTO long_term_memories (
      memory_id, namespace, kind, visibility, status, content, summary, confidence,
      source_thread_id, source_event_id, workspace_fingerprint, expires_at,
      metadata_json, citations_json, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(memory_id) DO UPDATE SET
      namespace = excluded.namespace,
      kind = excluded.kind,
      visibility = excluded.visibility,
      status = excluded.status,
      content = excluded.content,
      summary = excluded.summary,
      confidence = excluded.confidence,
      source_thread_id = excluded.source_thread_id,
      source_event_id = excluded.source_event_id,
      workspace_fingerprint = excluded.workspace_fingerprint,
      expires_at = excluded.expires_at,
      metadata_json = excluded.metadata_json,
      citations_json = excluded.citations_json,
      updated_at = excluded.updated_at`,
    )
    .run(
      memory.memoryId,
      memory.namespace,
      memory.kind,
      memory.visibility,
      memory.status,
      protectRequiredField(
        memory.content,
        fieldAad("memories", "long_term_memories", memory.memoryId, "content"),
      ),
      memory.summary
        ? protectRequiredField(
            memory.summary,
            fieldAad("memories", "long_term_memories", memory.memoryId, "summary"),
          )
        : null,
      memory.confidence ?? null,
      memory.sourceThreadId ?? null,
      memory.sourceEventId ?? null,
      memory.workspaceFingerprint ?? null,
      memory.expiresAt ?? null,
      memory.metadata
        ? protectRequiredField(
            JSON.stringify(memory.metadata),
            fieldAad("memories", "long_term_memories", memory.memoryId, "metadata_json"),
          )
        : null,
      memory.citations
        ? protectRequiredField(
            JSON.stringify(memory.citations),
            fieldAad("memories", "long_term_memories", memory.memoryId, "citations_json"),
          )
        : null,
      memory.createdAt,
      memory.updatedAt,
    );
}

export function getLongTermMemoryFromDb(
  memoriesDb: SqliteDatabase,
  memoryId: string,
): RuntimeLongTermMemoryRecord | null {
  const row = memoriesDb
    .prepare(`SELECT * FROM long_term_memories WHERE memory_id = ?`)
    .get(memoryId) as Record<string, any> | undefined;

  return row ? mapLongTermMemory(row) : null;
}

export function archiveLongTermMemoryInDb(
  memoriesDb: SqliteDatabase,
  memoryId: string,
  updatedAt = Date.now(),
): RuntimeLongTermMemoryRecord | null {
  const existing = getLongTermMemoryFromDb(memoriesDb, memoryId);
  if (!existing) return null;

  memoriesDb
    .prepare(
      `UPDATE long_term_memories
     SET status = 'archived', updated_at = ?
     WHERE memory_id = ?`,
    )
    .run(updatedAt, memoryId);

  return getLongTermMemoryFromDb(memoriesDb, memoryId);
}

export function listLongTermMemoriesFromDb(
  memoriesDb: SqliteDatabase,
  options: RuntimeMemoryListOptions = {},
): RuntimeLongTermMemoryRecord[] {
  const filters: string[] = [];
  const params: Array<string | number> = [];

  if (options.namespace) {
    filters.push("namespace = ?");
    params.push(options.namespace);
  }
  if (options.kind) {
    filters.push("kind = ?");
    params.push(options.kind);
  }
  if (options.visibility) {
    filters.push("visibility = ?");
    params.push(options.visibility);
  }
  if (options.status) {
    filters.push("status = ?");
    params.push(options.status);
  }

  const where = filters.length > 0 ? `WHERE ${filters.join(" AND ")}` : "";
  const limit = clampNumber(options.limit, { fallback: 20, min: 1, max: 100 });
  const offset = clampMemoryListOffset(options.offset);
  const rows = memoriesDb
    .prepare(`SELECT * FROM long_term_memories ${where} ORDER BY updated_at DESC LIMIT ? OFFSET ?`)
    .all(...params, limit, offset) as Record<string, any>[];

  return rows.map(mapLongTermMemory);
}

export function getMemoryPipelineCursorFromDb(
  memoriesDb: SqliteDatabase,
  pipelineId: string,
): number {
  const row = memoriesDb
    .prepare(`SELECT last_event_id FROM memory_pipeline_state WHERE pipeline_id = ?`)
    .get(pipelineId) as { last_event_id: number } | undefined;

  return row?.last_event_id ?? 0;
}

export function setMemoryPipelineCursorInDb(
  memoriesDb: SqliteDatabase,
  pipelineId: string,
  lastEventId: number,
): void {
  memoriesDb
    .prepare(
      `INSERT INTO memory_pipeline_state (pipeline_id, last_event_id, updated_at)
     VALUES (?, ?, ?)
     ON CONFLICT(pipeline_id) DO UPDATE SET
       last_event_id = excluded.last_event_id,
       updated_at = excluded.updated_at`,
    )
    .run(pipelineId, lastEventId, Date.now());
}
