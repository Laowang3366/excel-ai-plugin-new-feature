# Long Term Office Memory Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a database-backed long-term memory system for the Office assistant that remembers user-visible preferences, constraints, corrections, file impressions, and internal tool success profiles without leaking internal tool strategy into normal prompts.

**Architecture:** Extend the existing four-database `StateRuntimeStore` with structured long-term memory schema and a focused `memory/longTerm` module. User-visible memories flow into `buildStreamParams` and `memory.*` tools; internal `tool_success_profile` records are consumed only by tool routing/execution code.

**Tech Stack:** TypeScript, Electron agent runtime, SQLite via Node/Electron built-in `node:sqlite`, Vitest, existing `StateRuntimeStore`, existing tool registry/executor pattern.

---

## File Structure

- Create `desktop/electron/agent/memory/longTerm/memoryTypes.ts`
  - Owns memory kind/status/source types, user-visible kind checks, tool profile metadata, and validation helpers.
- Create `desktop/electron/agent/memory/longTerm/memoryStore.ts`
  - Thin long-term memory API over `StateRuntimeStore`; no AI calls.
- Create `desktop/electron/agent/memory/longTerm/memoryExtraction.ts`
  - Converts rollout events into candidate extraction input and parses structured candidate output.
- Create `desktop/electron/agent/memory/longTerm/memoryConsolidation.ts`
  - Merges candidates with existing memory records and enforces visibility/tool-profile rules.
- Create `desktop/electron/agent/memory/longTerm/memoryPruning.ts`
  - Applies stale/archive rules without physical deletion.
- Create `desktop/electron/agent/memory/longTerm/memoryStartupTask.ts`
  - Runs lightweight background extraction/consolidation/prune using a persisted cursor.
- Create `desktop/electron/agent/prompts/templates/memory/stage_one_system.zh-CN.md`
  - Chinese extraction instructions that filter one-off file content and internal tool names from user-visible memory.
- Create `desktop/electron/agent/prompts/templates/memory/consolidation.zh-CN.md`
  - Chinese consolidation instructions for merge/update/archive decisions.
- Create `desktop/electron/agent/prompts/memoryPrompt.ts`
  - Loads memory prompt templates.
- Create `desktop/electron/agent/tools/registry/memory.ts`
  - Defines `memory.write/read/search/list`.
- Create `desktop/electron/agent/tools/executors/memoryExecutors.ts`
  - Implements memory tools over `memoryStore`.
- Modify `desktop/electron/agent/memory/stateRuntimeTypes.ts`
  - Expand `RuntimeMemoryRecord` and add query/input types.
- Modify `desktop/electron/agent/memory/stateRuntimeSchema.ts`
  - Add `002_long_term_memories` migration for structured columns and indexes.
- Modify `desktop/electron/agent/memory/stateRuntimeMappers.ts`
  - Map structured memory rows.
- Modify `desktop/electron/agent/memory/stateRuntimeStore.ts`
  - Add structured memory CRUD/search/cursor APIs while preserving current `upsertMemory/listMemories` compatibility.
- Modify `desktop/electron/agent/core/agentLoop/buildStreamParams.ts`
  - Inject only user-visible memory snippets.
- Modify `desktop/electron/agent/tools/registry/toolDefinitions.ts`
  - Include memory tool definitions.
- Modify `desktop/electron/agent/tools/executors/createToolExecutors.ts`
  - Wire memory executor dependencies.
- Modify `desktop/electron/agent/runtime/agentRuntime.ts`
  - Initialize `MemoryRuntime` or pass `StateRuntimeStore` to memory executors.

## Task 1: Structured Memory Schema And Store API

**Files:**
- Modify: `desktop/electron/agent/memory/stateRuntimeTypes.ts`
- Modify: `desktop/electron/agent/memory/stateRuntimeSchema.ts`
- Modify: `desktop/electron/agent/memory/stateRuntimeMappers.ts`
- Modify: `desktop/electron/agent/memory/stateRuntimeStore.ts`
- Test: `desktop/electron/agent/memory/stateRuntimeStore.test.ts`

- [ ] **Step 1: Write failing tests for migration and visibility-aware memory records**

Add these tests to `desktop/electron/agent/memory/stateRuntimeStore.test.ts`:

```ts
it("persists structured long-term memories with visibility fields", async () => {
  const store = new StateRuntimeStore(":memory:");
  await store.init();

  await store.upsertLongTermMemory({
    memoryId: "mem-1",
    namespace: "global",
    kind: "preference",
    visibility: "user",
    status: "active",
    content: "回复先给结论",
    summary: "回复风格偏好",
    confidence: 0.9,
    citations: [{ threadId: "thread-1", eventId: 1 }],
    metadata: { source: "tool" },
    createdAt: 100,
    updatedAt: 100,
  });

  expect(await store.listLongTermMemories({ visibility: "user" })).toMatchObject([
    {
      memoryId: "mem-1",
      kind: "preference",
      visibility: "user",
      status: "active",
      content: "回复先给结论",
    },
  ]);
  expect(await store.listLongTermMemories({ visibility: "internal" })).toEqual([]);
  await store.close();
});

it("tracks memory pipeline cursor in memories database", async () => {
  const store = new StateRuntimeStore(":memory:");
  await store.init();

  expect(await store.getMemoryPipelineCursor("default")).toBe(0);
  await store.setMemoryPipelineCursor("default", 42);
  expect(await store.getMemoryPipelineCursor("default")).toBe(42);
  await store.close();
});
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/stateRuntimeStore.test.ts
```

Expected: failures because `upsertLongTermMemory`, `listLongTermMemories`, `getMemoryPipelineCursor`, and `setMemoryPipelineCursor` do not exist.

- [ ] **Step 3: Extend runtime types**

In `desktop/electron/agent/memory/stateRuntimeTypes.ts`, replace `RuntimeMemoryRecord` with:

```ts
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
}
```

- [ ] **Step 4: Add memory migration**

In `desktop/electron/agent/memory/stateRuntimeSchema.ts`, append this migration to `memories` after `001_memories`:

```ts
{
  id: "002_long_term_memories",
  sql: `
    CREATE TABLE IF NOT EXISTS long_term_memories (
      memory_id TEXT PRIMARY KEY,
      namespace TEXT NOT NULL,
      kind TEXT NOT NULL,
      visibility TEXT NOT NULL,
      status TEXT NOT NULL,
      content TEXT NOT NULL,
      summary TEXT,
      confidence REAL,
      source_thread_id TEXT,
      source_event_id INTEGER,
      workspace_fingerprint TEXT,
      expires_at INTEGER,
      metadata_json TEXT,
      citations_json TEXT,
      created_at INTEGER NOT NULL,
      updated_at INTEGER NOT NULL
    );

    CREATE TABLE IF NOT EXISTS memory_pipeline_state (
      pipeline_id TEXT PRIMARY KEY,
      last_event_id INTEGER NOT NULL,
      updated_at INTEGER NOT NULL
    );

    CREATE INDEX IF NOT EXISTS idx_long_term_memories_visibility_status_updated_at
      ON long_term_memories(visibility, status, updated_at);
    CREATE INDEX IF NOT EXISTS idx_long_term_memories_namespace_kind_status
      ON long_term_memories(namespace, kind, status);
    CREATE INDEX IF NOT EXISTS idx_long_term_memories_source_thread
      ON long_term_memories(source_thread_id);
  `,
}
```

- [ ] **Step 5: Add row mapper**

In `desktop/electron/agent/memory/stateRuntimeMappers.ts`, add:

```ts
import type { RuntimeMemoryRecord } from "./stateRuntimeTypes";

export function mapLongTermMemory(row: Record<string, any>): RuntimeMemoryRecord {
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
```

If `stateRuntimeMappers.ts` already exports `mapMemory`, keep it and add `mapLongTermMemory` alongside it.

- [ ] **Step 6: Add store methods**

In `desktop/electron/agent/memory/stateRuntimeStore.ts`, import `mapLongTermMemory` and add methods:

```ts
async upsertLongTermMemory(memory: RuntimeMemoryRecord): Promise<void> {
  this.getDbs().memories.prepare(
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
      updated_at = excluded.updated_at`
  ).run(
    memory.memoryId,
    memory.namespace,
    memory.kind,
    memory.visibility,
    memory.status,
    memory.content,
    memory.summary ?? null,
    memory.confidence ?? null,
    memory.sourceThreadId ?? null,
    memory.sourceEventId ?? null,
    memory.workspaceFingerprint ?? null,
    memory.expiresAt ?? null,
    memory.metadata ? JSON.stringify(memory.metadata) : null,
    memory.citations ? JSON.stringify(memory.citations) : null,
    memory.createdAt,
    memory.updatedAt
  );
}

async listLongTermMemories(options: RuntimeMemoryListOptions = {}): Promise<RuntimeMemoryRecord[]> {
  const clauses: string[] = [];
  const values: unknown[] = [];
  if (options.namespace) {
    clauses.push("namespace = ?");
    values.push(options.namespace);
  }
  if (options.kind) {
    clauses.push("kind = ?");
    values.push(options.kind);
  }
  if (options.visibility) {
    clauses.push("visibility = ?");
    values.push(options.visibility);
  }
  if (options.status) {
    clauses.push("status = ?");
    values.push(options.status);
  }
  const limit = Math.min(100, Math.max(1, Math.floor(options.limit ?? 20)));
  const where = clauses.length ? `WHERE ${clauses.join(" AND ")}` : "";
  const rows = this.getDbs().memories.prepare(
    `SELECT * FROM long_term_memories ${where} ORDER BY updated_at DESC LIMIT ?`
  ).all(...values, limit) as Record<string, any>[];
  return rows.map(mapLongTermMemory);
}

async getMemoryPipelineCursor(pipelineId: string): Promise<number> {
  const row = this.getDbs().memories.prepare(
    `SELECT last_event_id FROM memory_pipeline_state WHERE pipeline_id = ?`
  ).get(pipelineId) as { last_event_id: number } | undefined;
  return row?.last_event_id ?? 0;
}

async setMemoryPipelineCursor(pipelineId: string, lastEventId: number): Promise<void> {
  this.getDbs().memories.prepare(
    `INSERT INTO memory_pipeline_state (pipeline_id, last_event_id, updated_at)
     VALUES (?, ?, ?)
     ON CONFLICT(pipeline_id) DO UPDATE SET
       last_event_id = excluded.last_event_id,
       updated_at = excluded.updated_at`
  ).run(pipelineId, Math.max(0, Math.floor(lastEventId)), Date.now());
}
```

- [ ] **Step 7: Run tests and commit**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/stateRuntimeStore.test.ts
npm run typecheck
git add electron/agent/memory/stateRuntimeTypes.ts electron/agent/memory/stateRuntimeSchema.ts electron/agent/memory/stateRuntimeMappers.ts electron/agent/memory/stateRuntimeStore.ts electron/agent/memory/stateRuntimeStore.test.ts
git commit -m "feat: add structured long term memory store"
```

Expected: tests pass and typecheck passes.

## Task 2: Memory Types And Visibility Rules

**Files:**
- Create: `desktop/electron/agent/memory/longTerm/memoryTypes.ts`
- Test: `desktop/electron/agent/memory/longTerm/memoryTypes.test.ts`

- [ ] **Step 1: Write failing tests**

Create `desktop/electron/agent/memory/longTerm/memoryTypes.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import {
  getMemoryVisibility,
  isUserVisibleMemoryKind,
  normalizeMemoryWriteInput,
} from "./memoryTypes";

describe("long term memory types", () => {
  it("keeps tool success profiles internal", () => {
    expect(getMemoryVisibility("tool_success_profile")).toBe("internal");
    expect(isUserVisibleMemoryKind("tool_success_profile")).toBe(false);
  });

  it("keeps user preference kinds visible", () => {
    expect(getMemoryVisibility("preference")).toBe("user");
    expect(getMemoryVisibility("operation_preference")).toBe("user");
    expect(isUserVisibleMemoryKind("file_impression")).toBe(true);
  });

  it("rejects internal tool profiles from ordinary tool writes", () => {
    expect(() => normalizeMemoryWriteInput({
      kind: "tool_success_profile",
      namespace: "global",
      content: "内部工具统计",
      source: "tool",
    })).toThrow("tool_success_profile 只能由内部遥测写入");
  });
});
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/longTerm/memoryTypes.test.ts
```

Expected: module not found.

- [ ] **Step 3: Implement memory type helpers**

Create `desktop/electron/agent/memory/longTerm/memoryTypes.ts`:

```ts
import type {
  RuntimeMemoryKind,
  RuntimeMemoryRecord,
  RuntimeMemoryVisibility,
} from "../stateRuntimeTypes";

export const USER_VISIBLE_MEMORY_KINDS: RuntimeMemoryKind[] = [
  "preference",
  "constraint",
  "correction",
  "style_preference",
  "operation_preference",
  "file_impression",
];

export const INTERNAL_MEMORY_KINDS: RuntimeMemoryKind[] = [
  "tool_success_profile",
];

export interface MemoryWriteInput {
  kind: RuntimeMemoryKind;
  namespace?: string;
  content: string;
  summary?: string;
  confidence?: number;
  source?: "tool" | "telemetry" | "extraction";
  metadata?: Record<string, unknown>;
}

export function isUserVisibleMemoryKind(kind: RuntimeMemoryKind): boolean {
  return USER_VISIBLE_MEMORY_KINDS.includes(kind);
}

export function getMemoryVisibility(kind: RuntimeMemoryKind): RuntimeMemoryVisibility {
  return isUserVisibleMemoryKind(kind) ? "user" : "internal";
}

export function normalizeMemoryWriteInput(input: MemoryWriteInput): Omit<RuntimeMemoryRecord, "memoryId" | "createdAt" | "updatedAt"> {
  const content = input.content.trim();
  if (!content) throw new Error("记忆内容不能为空");
  if (content.length > 1000) throw new Error("记忆内容不能超过 1000 字");
  if (input.kind === "tool_success_profile" && input.source !== "telemetry") {
    throw new Error("tool_success_profile 只能由内部遥测写入");
  }

  return {
    namespace: input.namespace?.trim() || "global",
    kind: input.kind,
    visibility: getMemoryVisibility(input.kind),
    status: "active",
    content,
    summary: input.summary?.trim() || undefined,
    confidence: clampConfidence(input.confidence),
    metadata: {
      ...(input.metadata ?? {}),
      source: input.source ?? "tool",
    },
  };
}

function clampConfidence(value: number | undefined): number | undefined {
  if (value === undefined) return undefined;
  if (!Number.isFinite(value)) return undefined;
  return Math.min(1, Math.max(0, value));
}
```

- [ ] **Step 4: Run tests and commit**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/longTerm/memoryTypes.test.ts
npm run typecheck
git add electron/agent/memory/longTerm/memoryTypes.ts electron/agent/memory/longTerm/memoryTypes.test.ts
git commit -m "feat: define long term memory visibility rules"
```

Expected: tests pass and typecheck passes.

## Task 3: Long Term Memory Store Wrapper

**Files:**
- Create: `desktop/electron/agent/memory/longTerm/memoryStore.ts`
- Test: `desktop/electron/agent/memory/longTerm/memoryStore.test.ts`

- [ ] **Step 1: Write failing tests**

Create `desktop/electron/agent/memory/longTerm/memoryStore.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import { StateRuntimeStore } from "../stateRuntimeStore";
import { LongTermMemoryStore } from "./memoryStore";

describe("LongTermMemoryStore", () => {
  it("writes user-visible memories and hides internal profiles by default", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const store = new LongTermMemoryStore(runtime);

    await store.write({
      kind: "preference",
      namespace: "global",
      content: "回复先给结论",
      source: "tool",
    });
    await store.write({
      kind: "tool_success_profile",
      namespace: "global",
      content: "PPT 文件级编辑成功率更高",
      source: "telemetry",
      metadata: { successCount: 3, failureCount: 1 },
    });

    expect((await store.search({ query: "PPT" })).map((m) => m.kind)).not.toContain("tool_success_profile");
    expect((await store.search({ query: "PPT", includeInternal: true })).map((m) => m.kind)).toContain("tool_success_profile");
    await runtime.close();
  });
});
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/longTerm/memoryStore.test.ts
```

Expected: module not found.

- [ ] **Step 3: Implement store wrapper**

Create `desktop/electron/agent/memory/longTerm/memoryStore.ts`:

```ts
import type { RuntimeMemoryKind, RuntimeMemoryRecord } from "../stateRuntimeTypes";
import { StateRuntimeStore } from "../stateRuntimeStore";
import { normalizeMemoryWriteInput, type MemoryWriteInput } from "./memoryTypes";

export interface MemorySearchOptions {
  query?: string;
  namespace?: string;
  kind?: RuntimeMemoryKind;
  includeInternal?: boolean;
  limit?: number;
}

export class LongTermMemoryStore {
  constructor(private readonly runtime: StateRuntimeStore) {}

  async write(input: MemoryWriteInput): Promise<RuntimeMemoryRecord> {
    const now = Date.now();
    const normalized = normalizeMemoryWriteInput(input);
    const record: RuntimeMemoryRecord = {
      ...normalized,
      memoryId: createMemoryId(normalized.kind),
      createdAt: now,
      updatedAt: now,
    };
    await this.runtime.upsertLongTermMemory(record);
    return record;
  }

  async search(options: MemorySearchOptions = {}): Promise<RuntimeMemoryRecord[]> {
    const rows = await this.runtime.listLongTermMemories({
      namespace: options.namespace,
      kind: options.kind,
      visibility: options.includeInternal ? undefined : "user",
      status: "active",
      limit: options.limit ?? 20,
    });
    const query = options.query?.trim().toLowerCase();
    if (!query) return rows;
    return rows.filter((memory) => {
      const haystack = `${memory.content}\n${memory.summary ?? ""}`.toLowerCase();
      return haystack.includes(query);
    });
  }

  async list(namespace?: string): Promise<RuntimeMemoryRecord[]> {
    return this.runtime.listLongTermMemories({
      namespace,
      visibility: "user",
      status: "active",
      limit: 50,
    });
  }
}

function createMemoryId(kind: RuntimeMemoryKind): string {
  return `${kind}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}
```

- [ ] **Step 4: Run tests and commit**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/longTerm/memoryStore.test.ts
npm run typecheck
git add electron/agent/memory/longTerm/memoryStore.ts electron/agent/memory/longTerm/memoryStore.test.ts
git commit -m "feat: add long term memory store wrapper"
```

Expected: tests pass and typecheck passes.

## Task 4: Memory Tools

**Files:**
- Create: `desktop/electron/agent/tools/registry/memory.ts`
- Create: `desktop/electron/agent/tools/executors/memoryExecutors.ts`
- Modify: `desktop/electron/agent/tools/registry/toolDefinitions.ts`
- Modify: `desktop/electron/agent/tools/executors/createToolExecutors.ts`
- Test: `desktop/electron/agent/tools/executors/memoryExecutors.test.ts`

- [ ] **Step 1: Write failing executor tests**

Create `desktop/electron/agent/tools/executors/memoryExecutors.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import { StateRuntimeStore } from "../../memory/stateRuntimeStore";
import { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";
import { addMemoryExecutors } from "./memoryExecutors";

describe("memory executors", () => {
  it("writes and searches user-visible memory", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const memoryStore = new LongTermMemoryStore(runtime);
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore });

    const write = await executors.get("memory.write").execute({
      kind: "preference",
      content: "回复先给结论",
    });
    expect(write.success).toBe(true);

    const search = await executors.get("memory.search").execute({ query: "结论" });
    expect(search.success).toBe(true);
    expect(search.data[0].content).toBe("回复先给结论");
    await runtime.close();
  });

  it("rejects direct writes to internal tool profiles", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();
    const executors = new Map();
    addMemoryExecutors(executors, { memoryStore: new LongTermMemoryStore(runtime) });

    const result = await executors.get("memory.write").execute({
      kind: "tool_success_profile",
      content: "内部统计",
    });
    expect(result.success).toBe(false);
    expect(result.error).toContain("tool_success_profile");
    await runtime.close();
  });
});
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```powershell
cd desktop
npm test -- electron/agent/tools/executors/memoryExecutors.test.ts
```

Expected: module not found.

- [ ] **Step 3: Add tool definitions**

Create `desktop/electron/agent/tools/registry/memory.ts`:

```ts
import type { ToolDefinition } from "../../shared/types";

export const MEMORY_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "memory.write",
    description: "写入用户长期记忆。仅用于用户明确偏好、长期约束、纠正、文档风格偏好、操作方式偏好和低敏文件印象；不要写入文件正文、表格明细、临时路径或内部工具统计。",
    parameters: {
      type: "object",
      properties: {
        kind: {
          type: "string",
          enum: ["preference", "constraint", "correction", "style_preference", "operation_preference", "file_impression"],
        },
        namespace: { type: "string", default: "global" },
        content: { type: "string" },
        summary: { type: "string" },
        confidence: { type: "number" },
      },
      required: ["kind", "content"],
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "memory.search",
    description: "搜索用户可见长期记忆，用于了解用户偏好、长期约束、纠正、文档风格偏好、操作方式偏好和过往文件印象。",
    parameters: {
      type: "object",
      properties: {
        query: { type: "string" },
        namespace: { type: "string" },
        kind: { type: "string" },
        limit: { type: "number", default: 10 },
      },
      required: ["query"],
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "memory.list",
    description: "列出用户可见长期记忆摘要。",
    parameters: {
      type: "object",
      properties: {
        namespace: { type: "string" },
      },
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
];
```

Modify `desktop/electron/agent/tools/registry/toolDefinitions.ts`:

```ts
import { MEMORY_TOOL_DEFINITIONS } from "./memory";

export const ALL_TOOL_DEFINITIONS: ToolDefinition[] = [
  // keep existing order
  ...KNOWLEDGE_TOOL_DEFINITIONS,
  ...MEMORY_TOOL_DEFINITIONS,
  ...OFFICE_TOOL_DEFINITIONS,
];
```

Keep the existing entries before and after; only insert `...MEMORY_TOOL_DEFINITIONS` near knowledge tools.

- [ ] **Step 4: Add executors**

Create `desktop/electron/agent/tools/executors/memoryExecutors.ts`:

```ts
import type { ToolExecutor } from "../../shared/types";
import type { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";
import { validateArgs } from "./validation";

export interface MemoryExecutorDeps {
  memoryStore?: LongTermMemoryStore;
}

export function addMemoryExecutors(target: Map<string, ToolExecutor>, deps: MemoryExecutorDeps): void {
  target.set("memory.write", {
    name: "memory.write",
    execute: async (args: Record<string, unknown>) => {
      if (!deps.memoryStore) return { success: false, error: "长期记忆尚未初始化" };
      const err = validateArgs(args, { kind: "string", content: "string" });
      if (err) return { success: false, error: err };
      try {
        const record = await deps.memoryStore.write({
          kind: args.kind as any,
          namespace: typeof args.namespace === "string" ? args.namespace : undefined,
          content: args.content as string,
          summary: typeof args.summary === "string" ? args.summary : undefined,
          confidence: typeof args.confidence === "number" ? args.confidence : undefined,
          source: "tool",
        });
        return { success: true, data: record };
      } catch (error: any) {
        return { success: false, error: error.message };
      }
    },
  });

  target.set("memory.search", {
    name: "memory.search",
    execute: async (args: Record<string, unknown>) => {
      if (!deps.memoryStore) return { success: false, error: "长期记忆尚未初始化" };
      const err = validateArgs(args, { query: "string" });
      if (err) return { success: false, error: err };
      const data = await deps.memoryStore.search({
        query: args.query as string,
        namespace: typeof args.namespace === "string" ? args.namespace : undefined,
        kind: typeof args.kind === "string" ? args.kind as any : undefined,
        limit: typeof args.limit === "number" ? args.limit : 10,
      });
      return { success: true, data };
    },
  });

  target.set("memory.list", {
    name: "memory.list",
    execute: async (args: Record<string, unknown>) => {
      if (!deps.memoryStore) return { success: false, error: "长期记忆尚未初始化" };
      const data = await deps.memoryStore.list(typeof args.namespace === "string" ? args.namespace : undefined);
      return { success: true, data };
    },
  });
}
```

- [ ] **Step 5: Wire executor creation**

Modify `desktop/electron/agent/tools/executors/createToolExecutors.ts`:

```ts
import type { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";
import { addMemoryExecutors } from "./memoryExecutors";

export function createToolExecutors(
  workbookBridge: ExcelWorkbookBridge,
  vbaBridge: ExcelVbaBridge,
  scriptBridge: ExcelScriptBridge,
  uiBridge: ExcelUiBridge,
  sessionFolderPath?: string,
  knowledgeRetriever?: Retriever,
  wordBridge?: WordDocumentBridge,
  presentationBridge?: PresentationBridge,
  officeScriptBridge?: OfficeScriptBridge,
  officeActionBridge?: OfficeActionBridge,
  memoryStore?: LongTermMemoryStore
): Map<string, ToolExecutor> {
  const executors = new Map<string, ToolExecutor>();
  addMemoryExecutors(executors, { memoryStore });
  return executors;
}
```

Preserve all existing `add*Executors` calls; add `addMemoryExecutors` after `addKnowledgeExecutors`.

- [ ] **Step 6: Run tests and commit**

Run:

```powershell
cd desktop
npm test -- electron/agent/tools/executors/memoryExecutors.test.ts
npm run typecheck
git add electron/agent/tools/registry/memory.ts electron/agent/tools/registry/toolDefinitions.ts electron/agent/tools/executors/memoryExecutors.ts electron/agent/tools/executors/createToolExecutors.ts electron/agent/tools/executors/memoryExecutors.test.ts
git commit -m "feat: add memory tools"
```

Expected: tests pass and typecheck passes.

## Task 5: User-Visible Prompt Injection Boundary

**Files:**
- Modify: `desktop/electron/agent/core/agentLoop/buildStreamParams.ts`
- Test: `desktop/electron/agent/core/agentLoop/buildStreamParams.test.ts`

- [ ] **Step 1: Write failing tests**

Create or extend `desktop/electron/agent/core/agentLoop/buildStreamParams.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import { appendLongTermMemoryContext } from "./buildStreamParams";

describe("appendLongTermMemoryContext", () => {
  it("injects only user-visible memory", () => {
    const prompt = appendLongTermMemoryContext("base", [
      {
        memoryId: "mem-user",
        namespace: "global",
        kind: "preference",
        visibility: "user",
        status: "active",
        content: "回复先给结论",
        createdAt: 1,
        updatedAt: 1,
      },
      {
        memoryId: "mem-internal",
        namespace: "global",
        kind: "tool_success_profile",
        visibility: "internal",
        status: "active",
        content: "内部工具统计",
        createdAt: 1,
        updatedAt: 1,
      },
    ]);

    expect(prompt).toContain("回复先给结论");
    expect(prompt).not.toContain("内部工具统计");
    expect(prompt).not.toContain("tool_success_profile");
  });
});
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```powershell
cd desktop
npm test -- electron/agent/core/agentLoop/buildStreamParams.test.ts
```

Expected: `appendLongTermMemoryContext` is not exported.

- [ ] **Step 3: Add prompt append helper**

In `desktop/electron/agent/core/agentLoop/buildStreamParams.ts`, add:

```ts
import type { RuntimeMemoryRecord } from "../../memory/stateRuntimeTypes";

export function appendLongTermMemoryContext(
  prompt: string,
  memories: RuntimeMemoryRecord[]
): string {
  const visible = memories.filter((memory) =>
    memory.visibility === "user" &&
    memory.status === "active" &&
    memory.kind !== "tool_success_profile"
  );
  if (visible.length === 0) return prompt;

  const lines = visible.slice(0, 8).map((memory) =>
    `- [${memory.kind}] ${memory.summary || memory.content}`
  );
  return `${prompt}\n\n## 用户长期记忆\n${lines.join("\n")}`;
}
```

- [ ] **Step 4: Keep runtime wiring out of this task**

Do not fetch memories in `buildEffectiveSystemPrompt` yet unless `MemoryRuntime` is already available in the call path. This task only establishes and tests the injection boundary.

- [ ] **Step 5: Run tests and commit**

Run:

```powershell
cd desktop
npm test -- electron/agent/core/agentLoop/buildStreamParams.test.ts
npm run typecheck
git add electron/agent/core/agentLoop/buildStreamParams.ts electron/agent/core/agentLoop/buildStreamParams.test.ts
git commit -m "feat: isolate user visible memory prompt context"
```

Expected: tests pass and typecheck passes.

## Task 6: Tool Success Profile Internal Route

**Files:**
- Create: `desktop/electron/agent/memory/longTerm/toolSuccessProfile.ts`
- Test: `desktop/electron/agent/memory/longTerm/toolSuccessProfile.test.ts`

- [ ] **Step 1: Write failing tests**

Create `desktop/electron/agent/memory/longTerm/toolSuccessProfile.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import {
  buildToolProfileKey,
  shouldPromoteToolProfile,
  updateToolProfileStats,
} from "./toolSuccessProfile";

describe("tool success profile", () => {
  it("requires multiple samples before promotion", () => {
    expect(shouldPromoteToolProfile({ successCount: 1, failureCount: 0 })).toBe(false);
    expect(shouldPromoteToolProfile({ successCount: 3, failureCount: 1 })).toBe(true);
  });

  it("updates counts without storing user prompt text", () => {
    const updated = updateToolProfileStats(undefined, {
      app: "powerpoint",
      operation: "create",
      toolFamily: "openxml",
      success: true,
    });
    expect(updated.successCount).toBe(1);
    expect(updated.failureCount).toBe(0);
    expect(JSON.stringify(updated)).not.toContain("用户");
  });

  it("builds stable app operation family key", () => {
    expect(buildToolProfileKey({
      app: "word",
      operation: "format",
      toolFamily: "com",
    })).toBe("word:format:com");
  });
});
```

- [ ] **Step 2: Run tests and verify they fail**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/longTerm/toolSuccessProfile.test.ts
```

Expected: module not found.

- [ ] **Step 3: Implement profile helpers**

Create `desktop/electron/agent/memory/longTerm/toolSuccessProfile.ts`:

```ts
export interface ToolProfileKeyInput {
  app: "excel" | "word" | "powerpoint" | "office";
  operation: string;
  toolFamily: "openxml" | "com" | "script" | "shell" | "python" | "office_action" | "other";
}

export interface ToolProfileStats {
  app: ToolProfileKeyInput["app"];
  operation: string;
  toolFamily: ToolProfileKeyInput["toolFamily"];
  successCount: number;
  failureCount: number;
  lastUpdatedAt: number;
}

export function buildToolProfileKey(input: ToolProfileKeyInput): string {
  return `${input.app}:${input.operation}:${input.toolFamily}`;
}

export function shouldPromoteToolProfile(stats: Pick<ToolProfileStats, "successCount" | "failureCount">): boolean {
  return stats.successCount + stats.failureCount >= 3;
}

export function updateToolProfileStats(
  current: ToolProfileStats | undefined,
  event: ToolProfileKeyInput & { success: boolean }
): ToolProfileStats {
  return {
    app: event.app,
    operation: event.operation,
    toolFamily: event.toolFamily,
    successCount: (current?.successCount ?? 0) + (event.success ? 1 : 0),
    failureCount: (current?.failureCount ?? 0) + (event.success ? 0 : 1),
    lastUpdatedAt: Date.now(),
  };
}
```

- [ ] **Step 4: Run tests and commit**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/longTerm/toolSuccessProfile.test.ts
npm run typecheck
git add electron/agent/memory/longTerm/toolSuccessProfile.ts electron/agent/memory/longTerm/toolSuccessProfile.test.ts
git commit -m "feat: add internal tool success profiles"
```

Expected: tests pass and typecheck passes.

## Task 7: Extraction And Consolidation Skeleton

**Files:**
- Create: `desktop/electron/agent/memory/longTerm/memoryExtraction.ts`
- Create: `desktop/electron/agent/memory/longTerm/memoryConsolidation.ts`
- Test: `desktop/electron/agent/memory/longTerm/memoryExtraction.test.ts`
- Test: `desktop/electron/agent/memory/longTerm/memoryConsolidation.test.ts`

- [ ] **Step 1: Write failing extraction tests**

Create `desktop/electron/agent/memory/longTerm/memoryExtraction.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import { parseStageOneOutput, shouldIgnoreCandidateContent } from "./memoryExtraction";

describe("memory extraction", () => {
  it("drops candidates that contain temp paths", () => {
    expect(shouldIgnoreCandidateContent("临时路径 C:\\Users\\wfq\\AppData\\Local\\Temp\\make.py")).toBe(true);
  });

  it("parses valid user-visible candidate output", () => {
    const output = parseStageOneOutput(JSON.stringify({
      memories: [{
        kind: "operation_preference",
        namespace: "global",
        content: "优先使用稳定的文件级编辑",
        confidence: 0.8,
        citations: [{ threadId: "thread-1", eventId: 1 }],
      }],
    }));
    expect(output[0]).toMatchObject({
      kind: "operation_preference",
      visibility: "user",
    });
  });
});
```

- [ ] **Step 2: Write failing consolidation tests**

Create `desktop/electron/agent/memory/longTerm/memoryConsolidation.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import { chooseConsolidationAction } from "./memoryConsolidation";

describe("memory consolidation", () => {
  it("ignores single-sample tool profiles", () => {
    expect(chooseConsolidationAction({
      kind: "tool_success_profile",
      visibility: "internal",
      metadata: { successCount: 1, failureCount: 0 },
    } as any)).toBe("ignore");
  });

  it("adds user corrections", () => {
    expect(chooseConsolidationAction({
      kind: "correction",
      visibility: "user",
      content: "不要反复尝试打开应用",
    } as any)).toBe("add");
  });
});
```

- [ ] **Step 3: Run tests and verify they fail**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/longTerm/memoryExtraction.test.ts electron/agent/memory/longTerm/memoryConsolidation.test.ts
```

Expected: modules not found.

- [ ] **Step 4: Implement extraction skeleton**

Create `desktop/electron/agent/memory/longTerm/memoryExtraction.ts`:

```ts
import type { RuntimeMemoryRecord } from "../stateRuntimeTypes";
import { getMemoryVisibility } from "./memoryTypes";

export function shouldIgnoreCandidateContent(content: string): boolean {
  return /AppData\\Local\\Temp|\\Temp\\|临时路径|完整正文|表格明细/.test(content);
}

export function parseStageOneOutput(raw: string): Array<Omit<RuntimeMemoryRecord, "memoryId" | "createdAt" | "updatedAt" | "status">> {
  const parsed = JSON.parse(raw) as { memories?: any[] };
  const memories = Array.isArray(parsed.memories) ? parsed.memories : [];
  return memories
    .filter((item) => typeof item.content === "string" && !shouldIgnoreCandidateContent(item.content))
    .map((item) => ({
      namespace: typeof item.namespace === "string" ? item.namespace : "global",
      kind: item.kind,
      visibility: getMemoryVisibility(item.kind),
      content: item.content.trim(),
      summary: typeof item.summary === "string" ? item.summary : undefined,
      confidence: typeof item.confidence === "number" ? item.confidence : undefined,
      citations: Array.isArray(item.citations) ? item.citations : undefined,
      metadata: item.file || item.toolProfile ? { file: item.file, toolProfile: item.toolProfile } : undefined,
    }));
}
```

- [ ] **Step 5: Implement consolidation skeleton**

Create `desktop/electron/agent/memory/longTerm/memoryConsolidation.ts`:

```ts
import type { RuntimeMemoryRecord } from "../stateRuntimeTypes";

export type ConsolidationAction = "add" | "merge" | "update" | "ignore";

export function chooseConsolidationAction(memory: RuntimeMemoryRecord): ConsolidationAction {
  if (memory.kind === "tool_success_profile") {
    const stats = memory.metadata?.toolProfile ?? memory.metadata;
    const successCount = Number((stats as any)?.successCount ?? 0);
    const failureCount = Number((stats as any)?.failureCount ?? 0);
    return successCount + failureCount >= 3 ? "add" : "ignore";
  }
  if (!memory.content.trim()) return "ignore";
  return "add";
}
```

- [ ] **Step 6: Run tests and commit**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/longTerm/memoryExtraction.test.ts electron/agent/memory/longTerm/memoryConsolidation.test.ts
npm run typecheck
git add electron/agent/memory/longTerm/memoryExtraction.ts electron/agent/memory/longTerm/memoryConsolidation.ts electron/agent/memory/longTerm/memoryExtraction.test.ts electron/agent/memory/longTerm/memoryConsolidation.test.ts
git commit -m "feat: add long term memory extraction skeleton"
```

Expected: tests pass and typecheck passes.

## Task 8: Prompt Templates

**Files:**
- Create: `desktop/electron/agent/prompts/templates/memory/stage_one_system.zh-CN.md`
- Create: `desktop/electron/agent/prompts/templates/memory/consolidation.zh-CN.md`
- Create: `desktop/electron/agent/prompts/templates/memory/instructions.zh-CN.md`
- Create: `desktop/electron/agent/prompts/memoryPrompt.ts`
- Test: `desktop/electron/agent/prompts/memoryPrompt.test.ts`

- [ ] **Step 1: Write failing loader test**

Create `desktop/electron/agent/prompts/memoryPrompt.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import { loadMemoryPromptTemplate } from "./memoryPrompt";

describe("memory prompt templates", () => {
  it("loads Chinese stage one prompt", () => {
    const text = loadMemoryPromptTemplate("stage_one_system");
    expect(text).toContain("用户可见记忆");
    expect(text).toContain("tool_success_profile");
    expect(text).toContain("不要进入普通对话提示词");
  });
});
```

- [ ] **Step 2: Add templates**

Create `desktop/electron/agent/prompts/templates/memory/stage_one_system.zh-CN.md`:

```md
# 长期记忆候选提取

你只提取跨任务稳定复用的信息。

用户可见记忆包括 preference、constraint、correction、style_preference、operation_preference、file_impression。

系统内部策略记忆只包括 tool_success_profile。它只能描述工具执行统计，不要进入普通对话提示词，不要写成“用户偏好”。

忽略文件正文、表格明细、完整 PPT 文案、临时路径、临时脚本、单次偶然成功或失败。
```

Create `desktop/electron/agent/prompts/templates/memory/consolidation.zh-CN.md`:

```md
# 长期记忆清洗合并

对候选记忆做新增、合并、更新或忽略。

同一用户纠正多次出现时合并为更强规则。
同一文件印象只更新最近操作和高层摘要。
tool_success_profile 必须有多次同类结果，不能因单次成功或失败写入长期策略。
```

Create `desktop/electron/agent/prompts/templates/memory/instructions.zh-CN.md`:

```md
# 长期记忆使用说明

普通对话上下文只允许注入用户可见记忆。
内部工具策略记忆只允许工具路由器和执行器读取。
不要把内部工具名、执行路线或成功率统计当成用户偏好展示。
```

- [ ] **Step 3: Add loader**

Create `desktop/electron/agent/prompts/memoryPrompt.ts`:

```ts
import * as fs from "fs";
import * as path from "path";

export type MemoryPromptTemplateName =
  | "stage_one_system"
  | "consolidation"
  | "instructions";

export function loadMemoryPromptTemplate(name: MemoryPromptTemplateName): string {
  const filePath = path.join(__dirname, "templates", "memory", `${name}.zh-CN.md`);
  return fs.readFileSync(filePath, "utf8");
}
```

- [ ] **Step 4: Run tests and commit**

Run:

```powershell
cd desktop
npm test -- electron/agent/prompts/memoryPrompt.test.ts
npm run typecheck
git add electron/agent/prompts/templates/memory/stage_one_system.zh-CN.md electron/agent/prompts/templates/memory/consolidation.zh-CN.md electron/agent/prompts/templates/memory/instructions.zh-CN.md electron/agent/prompts/memoryPrompt.ts electron/agent/prompts/memoryPrompt.test.ts
git commit -m "feat: add long term memory prompts"
```

Expected: tests pass and typecheck passes.

## Task 9: Startup Task And Pruning

**Files:**
- Create: `desktop/electron/agent/memory/longTerm/memoryPruning.ts`
- Create: `desktop/electron/agent/memory/longTerm/memoryStartupTask.ts`
- Test: `desktop/electron/agent/memory/longTerm/memoryPruning.test.ts`
- Test: `desktop/electron/agent/memory/longTerm/memoryStartupTask.test.ts`

- [ ] **Step 1: Write failing pruning test**

Create `desktop/electron/agent/memory/longTerm/memoryPruning.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import { shouldArchiveMemory } from "./memoryPruning";

describe("memory pruning", () => {
  it("archives expired file impressions", () => {
    expect(shouldArchiveMemory({
      kind: "file_impression",
      status: "active",
      expiresAt: 100,
    } as any, 101)).toBe(true);
  });

  it("keeps non-expired user constraints", () => {
    expect(shouldArchiveMemory({
      kind: "constraint",
      status: "active",
    } as any, 101)).toBe(false);
  });
});
```

- [ ] **Step 2: Write failing startup task test**

Create `desktop/electron/agent/memory/longTerm/memoryStartupTask.test.ts`:

```ts
import { describe, expect, it } from "vitest";

import { StateRuntimeStore } from "../stateRuntimeStore";
import { runMemoryStartupTask } from "./memoryStartupTask";

describe("memory startup task", () => {
  it("does not throw when there are no new rollout events", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();

    await expect(runMemoryStartupTask({ runtime, pipelineId: "test" })).resolves.toMatchObject({
      processed: 0,
    });
    await runtime.close();
  });
});
```

- [ ] **Step 3: Implement pruning helper**

Create `desktop/electron/agent/memory/longTerm/memoryPruning.ts`:

```ts
import type { RuntimeMemoryRecord } from "../stateRuntimeTypes";

export function shouldArchiveMemory(memory: RuntimeMemoryRecord, now = Date.now()): boolean {
  if (memory.status === "archived") return false;
  if (memory.expiresAt !== undefined && memory.expiresAt <= now) return true;
  return false;
}
```

- [ ] **Step 4: Implement startup task skeleton**

Create `desktop/electron/agent/memory/longTerm/memoryStartupTask.ts`:

```ts
import { StateRuntimeStore } from "../stateRuntimeStore";

export interface MemoryStartupTaskOptions {
  runtime: StateRuntimeStore;
  pipelineId?: string;
}

export interface MemoryStartupTaskResult {
  processed: number;
}

export async function runMemoryStartupTask(options: MemoryStartupTaskOptions): Promise<MemoryStartupTaskResult> {
  const pipelineId = options.pipelineId ?? "default";
  const cursor = await options.runtime.getMemoryPipelineCursor(pipelineId);
  await options.runtime.setMemoryPipelineCursor(pipelineId, cursor);
  return { processed: 0 };
}
```

- [ ] **Step 5: Run tests and commit**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/longTerm/memoryPruning.test.ts electron/agent/memory/longTerm/memoryStartupTask.test.ts
npm run typecheck
git add electron/agent/memory/longTerm/memoryPruning.ts electron/agent/memory/longTerm/memoryStartupTask.ts electron/agent/memory/longTerm/memoryPruning.test.ts electron/agent/memory/longTerm/memoryStartupTask.test.ts
git commit -m "feat: add long term memory startup task"
```

Expected: tests pass and typecheck passes.

## Task 10: Final Integration Review

**Files:**
- Review: `desktop/electron/agent/memory/longTerm/*.ts`
- Review: `desktop/electron/agent/memory/stateRuntimeStore.ts`
- Review: `desktop/electron/agent/core/agentLoop/buildStreamParams.ts`
- Review: `desktop/electron/agent/tools/registry/memory.ts`
- Review: `desktop/electron/agent/tools/executors/memoryExecutors.ts`

- [ ] **Step 1: Run focused test suite**

Run:

```powershell
cd desktop
npm test -- electron/agent/memory/stateRuntimeStore.test.ts electron/agent/memory/longTerm electron/agent/tools/executors/memoryExecutors.test.ts electron/agent/core/agentLoop/buildStreamParams.test.ts electron/agent/prompts/memoryPrompt.test.ts
```

Expected: all focused tests pass.

- [ ] **Step 2: Run full verification**

Run:

```powershell
cd desktop
npm test
npm run typecheck
npm run build
```

Expected: all tests pass, typecheck passes, build passes with only existing Vite chunk-size warnings if they already exist.

- [ ] **Step 3: Manual review checklist**

Review these exact invariants:

```text
1. buildStreamParams never injects tool_success_profile.
2. memory.search and memory.list default to user-visible memory only.
3. memory.write cannot create tool_success_profile from ordinary model calls.
4. tool_success_profile contains counts/routes only, not user file text or prompt text.
5. Existing legacy upsertMemory/listMemories callers still compile.
6. No temporary test files remain outside committed test files.
```

- [ ] **Step 4: Commit final integration fixes**

If review requires fixes, commit them:

```powershell
git add electron/agent
git commit -m "fix: finalize long term memory integration"
```

If no fixes are required, do not create an empty commit.

## Self-Review

- Spec coverage: storage, schema, user-visible memory, correction memory, file impressions, internal tool success profiles, prompt injection boundary, memory tools, pruning, startup skeleton, and tests are covered.
- Scope control: workspace diff, remote compaction, vector search, and UI memory manager are not included in this plan because the approved spec marks them as non-core follow-up work.
- Internal strategy boundary: `tool_success_profile` is excluded from `buildStreamParams`, `memory.search`, and `memory.list` by default, and ordinary `memory.write` cannot create it.
- Red-flag scan: no unresolved implementation markers or unspecified deferred steps remain.
