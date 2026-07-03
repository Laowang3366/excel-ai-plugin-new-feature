import type { ChatMessage, StreamChatParams } from "../../providers/aiClient";
import type {
  RuntimeLongTermMemoryRecord,
  RuntimeMemoryCitation,
} from "../stateRuntimeTypes";
import type { Thread, Turn, TurnItem } from "../../shared/types";
import {
  TOOL_WRITABLE_MEMORY_KINDS,
  type MemoryWriteInput,
} from "./memoryTypes";
import {
  parseStageOneOutput,
  type ExtractedMemoryCandidate,
} from "./memoryExtraction";
import type { MemorySearchOptions } from "./memoryStore";

export interface MemoryAutoExtractionAIClient {
  chat(params: StreamChatParams): Promise<{ content?: string }>;
}

export interface MemoryAutoExtractionStore {
  write(input: MemoryWriteInput): Promise<RuntimeLongTermMemoryRecord>;
  search?(options?: MemorySearchOptions): Promise<RuntimeLongTermMemoryRecord[]>;
}

export interface ExtractAndWriteTurnMemoriesOptions {
  aiClient: MemoryAutoExtractionAIClient;
  memoryStore: MemoryAutoExtractionStore;
  thread: Thread;
  turn: Turn;
}

export interface ExtractAndWriteTurnMemoriesResult {
  candidates: ExtractedMemoryCandidate[];
  written: RuntimeLongTermMemoryRecord[];
  skippedDuplicates: number;
}

const MAX_TRANSCRIPT_CHARS = 12_000;
const MAX_ITEM_CHARS = 2_000;

export async function extractAndWriteTurnMemories(
  options: ExtractAndWriteTurnMemoriesOptions,
): Promise<ExtractAndWriteTurnMemoriesResult> {
  if (options.turn.status !== "completed") {
    return { candidates: [], written: [], skippedDuplicates: 0 };
  }

  const transcript = buildMemoryExtractionTranscript(options.turn);
  if (!transcript) {
    return { candidates: [], written: [], skippedDuplicates: 0 };
  }

  const result = await options.aiClient.chat({
    messages: buildMemoryExtractionMessages(options.thread, options.turn, transcript),
    maxTokens: 1000,
    temperature: 0,
    reasoningMode: "off",
  });

  const candidates = parseStageOneOutput(stripJsonFence(result.content ?? ""));
  const written: RuntimeLongTermMemoryRecord[] = [];
  let skippedDuplicates = 0;

  for (const candidate of candidates) {
    if (await hasExactDuplicate(options.memoryStore, candidate)) {
      skippedDuplicates += 1;
      continue;
    }

    written.push(
      await options.memoryStore.write({
        kind: candidate.kind,
        namespace: candidate.namespace,
        content: candidate.content,
        confidence: candidate.confidence,
        source: "extraction",
        sourceThreadId: options.thread.metadata.threadId,
        metadata: {
          ...(candidate.metadata ?? {}),
          extraction: {
            threadId: options.thread.metadata.threadId,
            turnId: options.turn.turnId,
          },
        },
        citations: normalizeCandidateCitations(
          candidate.citations,
          options.thread.metadata.threadId,
          options.turn.turnId,
        ),
      }),
    );
  }

  return { candidates, written, skippedDuplicates };
}

function buildMemoryExtractionMessages(
  thread: Thread,
  turn: Turn,
  transcript: string,
): ChatMessage[] {
  return [
    {
      role: "system",
      content: [
        "Extract durable long-term user memories from the completed turn.",
        "Return strict JSON only: {\"memories\":[]}.",
        `Allowed kinds: ${TOOL_WRITABLE_MEMORY_KINDS.join(", ")}.`,
        "Write only explicit user preferences, long-term constraints, user corrections, document style preferences, operation/tooling preferences, or low-sensitivity file impressions.",
        "Ignore one-off task instructions, progress, tool statistics, generated document bodies, table dumps, temporary paths, secrets, credentials, and sensitive personal data.",
        "Do not infer or invent memories. Preserve the user's language. Prefer namespace \"global\" unless the user clearly scopes the memory.",
      ].join("\n"),
    },
    {
      role: "user",
      content: [
        `threadId: ${thread.metadata.threadId}`,
        `turnId: ${turn.turnId}`,
        "Completed turn transcript:",
        transcript,
      ].join("\n\n"),
    },
  ];
}

export function buildMemoryExtractionTranscript(turn: Turn): string {
  const lines: string[] = [];
  let totalLength = 0;

  for (const item of turn.items) {
    const line = serializeTurnItemForMemoryExtraction(item);
    if (!line) continue;
    const nextLength = totalLength + line.length + 1;
    if (nextLength > MAX_TRANSCRIPT_CHARS) {
      lines.push("[transcript truncated]");
      break;
    }
    lines.push(line);
    totalLength = nextLength;
  }

  return lines.join("\n");
}

function serializeTurnItemForMemoryExtraction(item: TurnItem): string | null {
  switch (item.type) {
    case "user_message": {
      const attachments = item.attachments?.length
        ? `\nattachments: ${item.attachments.map((att) => `${att.fileName} (${att.fileType})`).join(", ")}`
        : "";
      return `USER:\n${truncateText(item.content)}${attachments}`;
    }
    case "assistant_message":
      return `ASSISTANT${item.phase ? ` (${item.phase})` : ""}:\n${truncateText(item.content)}`;
    case "tool_call":
      return `TOOL CALL: ${item.toolName}`;
    case "tool_result":
      if (item.isError) {
        return `TOOL ERROR (${item.toolName}): ${truncateText(stringifyToolResult(item.result), 800)}`;
      }
      return `TOOL RESULT (${item.toolName}): completed`;
    case "error":
      return `ERROR: ${truncateText(item.message, 800)}`;
    default:
      return null;
  }
}

function stringifyToolResult(result: unknown): string {
  if (typeof result === "string") return result;
  try {
    return JSON.stringify(result);
  } catch {
    return String(result);
  }
}

function truncateText(text: string, maxChars = MAX_ITEM_CHARS): string {
  const normalized = text.trim();
  if (normalized.length <= maxChars) return normalized;
  return `${normalized.slice(0, maxChars)}\n[truncated]`;
}

function stripJsonFence(content: string): string {
  const trimmed = content.trim();
  const fenced = trimmed.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  if (fenced) return fenced[1].trim();

  const firstBrace = trimmed.indexOf("{");
  const lastBrace = trimmed.lastIndexOf("}");
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    return trimmed.slice(firstBrace, lastBrace + 1);
  }
  return trimmed;
}

async function hasExactDuplicate(
  store: MemoryAutoExtractionStore,
  candidate: ExtractedMemoryCandidate,
): Promise<boolean> {
  if (!store.search) return false;

  const existing = await store.search({
    namespace: candidate.namespace,
    kind: candidate.kind,
    query: candidate.content,
    limit: 20,
  });
  const normalizedCandidate = normalizeForDuplicateCheck(candidate.content);
  return existing.some(
    (memory) => normalizeForDuplicateCheck(memory.content) === normalizedCandidate,
  );
}

function normalizeCandidateCitations(
  citations: RuntimeMemoryCitation[] | undefined,
  threadId: string,
  turnId: string,
): RuntimeMemoryCitation[] {
  if (citations && citations.length > 0) return citations;
  return [{ threadId, turnId }];
}

function normalizeForDuplicateCheck(value: string): string {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}
