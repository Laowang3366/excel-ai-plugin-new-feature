import type {
  RuntimeLongTermMemoryRecord,
  RuntimeMemoryCitation,
  RuntimeMemoryKind,
  RuntimeMemoryVisibility,
} from "../stateRuntimeTypes";
import type { ThreadId } from "../../shared/types";

export const TOOL_WRITABLE_MEMORY_KINDS = [
  "preference",
  "constraint",
  "correction",
  "style_preference",
  "operation_preference",
  "file_impression",
] as const;

export type ToolWritableMemoryKind = (typeof TOOL_WRITABLE_MEMORY_KINDS)[number];

export const MEMORY_VISIBILITY_BY_KIND: Record<RuntimeMemoryKind, RuntimeMemoryVisibility> = {
  preference: "user",
  constraint: "user",
  correction: "user",
  style_preference: "user",
  operation_preference: "user",
  file_impression: "user",
  tool_success_profile: "internal",
};

export interface MemoryWriteInput {
  kind: RuntimeMemoryKind;
  namespace?: string;
  content: string;
  summary?: string;
  confidence?: number;
  source?: "tool" | "telemetry" | "extraction";
  sourceThreadId?: ThreadId;
  sourceEventId?: number;
  workspaceFingerprint?: string;
  expiresAt?: number;
  metadata?: Record<string, unknown>;
  citations?: RuntimeMemoryCitation[];
}

export function isToolWritableMemoryKind(kind: string): kind is ToolWritableMemoryKind {
  return (TOOL_WRITABLE_MEMORY_KINDS as readonly string[]).includes(kind);
}

export function getMemoryVisibility(kind: RuntimeMemoryKind): RuntimeMemoryVisibility {
  return MEMORY_VISIBILITY_BY_KIND[kind];
}

export function normalizeMemoryWriteInput(
  input: MemoryWriteInput,
): Omit<RuntimeLongTermMemoryRecord, "memoryId" | "createdAt" | "updatedAt"> {
  const content = input.content.trim();
  if (!content) throw new Error("记忆内容不能为空");
  if (content.length > 1000) throw new Error("记忆内容不能超过 1000 字");
  const namespace = normalizeNamespace(input.namespace);
  if (input.kind === "tool_success_profile" && input.source !== "telemetry") {
    throw new Error("tool_success_profile 只能由内部遥测写入");
  }

  return {
    namespace,
    kind: input.kind,
    visibility: getMemoryVisibility(input.kind),
    status: "active",
    content,
    summary: input.summary?.trim() || undefined,
    confidence: clampConfidence(input.confidence),
    sourceThreadId: input.sourceThreadId,
    sourceEventId: input.sourceEventId,
    workspaceFingerprint: input.workspaceFingerprint,
    expiresAt: input.expiresAt,
    metadata: {
      ...(input.metadata ?? {}),
      source: input.source ?? "tool",
    },
    citations: input.citations,
  };
}

function normalizeNamespace(value: string | undefined): string {
  if (value === undefined) return "global";
  const namespace = value.trim();
  if (!namespace) throw new Error("记忆命名空间不能为空");
  return namespace;
}

function clampConfidence(value: number | undefined): number | undefined {
  if (value === undefined) return undefined;
  if (!Number.isFinite(value)) return undefined;
  return Math.min(1, Math.max(0, value));
}
