import type {
  RuntimeLongTermMemoryRecord,
  RuntimeMemoryKind,
} from "../stateRuntimeTypes";
import { StateRuntimeStore } from "../stateRuntimeStore";
import {
  isToolWritableMemoryKind,
  normalizeMemoryWriteInput,
  type MemoryWriteInput,
} from "./memoryTypes";
import { clampNumber } from "../../shared/numberLimits";

export interface MemorySearchOptions {
  query?: string;
  namespace?: string;
  kind?: RuntimeMemoryKind;
  includeInternal?: boolean;
  limit?: number;
}

export class LongTermMemoryStore {
  constructor(private runtime: StateRuntimeStore) {}

  updateRuntime(runtime: StateRuntimeStore): void {
    this.runtime = runtime;
  }

  async write(input: MemoryWriteInput): Promise<RuntimeLongTermMemoryRecord> {
    const now = Date.now();
    const normalized = normalizeMemoryWriteInput(input);
    const record: RuntimeLongTermMemoryRecord = {
      ...normalized,
      memoryId: createMemoryId(normalized.kind),
      createdAt: now,
      updatedAt: now,
    };
    await this.runtime.upsertLongTermMemory(record);
    return record;
  }

  async search(
    options: MemorySearchOptions = {},
  ): Promise<RuntimeLongTermMemoryRecord[]> {
    const query = options.query?.trim().toLowerCase();
    const publicLimit = clampNumber(options.limit, { fallback: 20, min: 1, max: 100 });
    const baseOptions = {
      namespace: options.namespace,
      kind: options.kind,
      visibility: options.includeInternal ? undefined : "user",
      status: "active",
    } as const;

    if (!query && options.includeInternal) {
      return this.runtime.listLongTermMemories({
        ...baseOptions,
        limit: publicLimit,
      });
    }

    const pageSize = 100;
    const matches: RuntimeLongTermMemoryRecord[] = [];
    for (let offset = 0; ; offset += pageSize) {
      const rows = await this.runtime.listLongTermMemories({
        ...baseOptions,
        limit: pageSize,
        offset,
      });
      for (const memory of rows) {
        if (!options.includeInternal && !isToolWritableMemoryKind(memory.kind)) {
          continue;
        }
        if (!query || matchesQuery(memory, query)) {
          matches.push(memory);
          if (matches.length >= publicLimit) return matches;
        }
      }
      if (rows.length < pageSize) return matches;
    }
  }

  async list(namespace?: string): Promise<RuntimeLongTermMemoryRecord[]> {
    return this.search({
      namespace,
      limit: 50,
    });
  }

  async delete(memoryId: string): Promise<RuntimeLongTermMemoryRecord | null> {
    const normalizedId = memoryId.trim();
    if (!normalizedId) throw new Error("记忆 ID 不能为空");

    const existing = await this.runtime.getLongTermMemory(normalizedId);
    if (!existing || existing.status !== "active") return null;
    if (existing.visibility !== "user" || !isToolWritableMemoryKind(existing.kind)) {
      throw new Error("只能删除用户可见的长期记忆");
    }

    return this.runtime.archiveLongTermMemory(normalizedId);
  }
}

function createMemoryId(kind: RuntimeMemoryKind): string {
  return `${kind}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function matchesQuery(
  memory: RuntimeLongTermMemoryRecord,
  query: string,
): boolean {
  const haystack = `${memory.content}\n${memory.summary ?? ""}`.toLowerCase();
  return haystack.includes(query);
}
