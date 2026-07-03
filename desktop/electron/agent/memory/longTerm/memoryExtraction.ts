import type {
  RuntimeMemoryCitation,
  RuntimeMemoryKind,
  RuntimeMemoryVisibility,
} from "../stateRuntimeTypes";
import { getMemoryVisibility, isToolWritableMemoryKind } from "./memoryTypes";

export interface ExtractedMemoryCandidate {
  kind: RuntimeMemoryKind;
  visibility: RuntimeMemoryVisibility;
  namespace: string;
  content: string;
  confidence?: number;
  metadata?: Record<string, unknown>;
  citations?: RuntimeMemoryCitation[];
}

interface StageOneMemoryShape {
  kind?: unknown;
  namespace?: unknown;
  content?: unknown;
  confidence?: unknown;
  metadata?: unknown;
  citations?: unknown;
}

const TEMP_PATH_PATTERNS = [
  /\\AppData\\Local\\Temp\\/i,
  /\/AppData\/Local\/Temp\//i,
  /(?:^|\s)(?:[A-Z]:)?\\Temp\\/i,
  /(?:^|\s)\/tmp\//i,
];

export function shouldIgnoreCandidateContent(content: string): boolean {
  const normalized = content.trim();
  if (!normalized) return true;
  if (TEMP_PATH_PATTERNS.some((pattern) => pattern.test(normalized))) return true;
  if (/完整正文|全文内容|表格明细|逐行明细/.test(normalized)) return true;
  if (looksLikeTableDump(normalized)) return true;
  return false;
}

export function parseStageOneOutput(output: string): ExtractedMemoryCandidate[] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(output);
  } catch {
    return [];
  }

  const memories = getRawMemories(parsed);
  const candidates: ExtractedMemoryCandidate[] = [];
  for (const raw of memories) {
    const memory = raw as StageOneMemoryShape;
    if (
      typeof memory.kind !== "string" ||
      !isToolWritableMemoryKind(memory.kind)
    ) continue;
    if (typeof memory.content !== "string") continue;

    const content = memory.content.trim();
    if (shouldIgnoreCandidateContent(content)) continue;

    candidates.push({
      kind: memory.kind,
      visibility: getMemoryVisibility(memory.kind),
      namespace:
        typeof memory.namespace === "string" && memory.namespace.trim()
          ? memory.namespace.trim()
          : "global",
      content,
      confidence: normalizeConfidence(memory.confidence),
      metadata: normalizeMetadata(memory.metadata),
      citations: normalizeCitations(memory.citations),
    });
  }
  return candidates;
}

function getRawMemories(parsed: unknown): unknown[] {
  if (
    parsed &&
    typeof parsed === "object" &&
    Array.isArray((parsed as { memories?: unknown }).memories)
  ) {
    return (parsed as { memories: unknown[] }).memories;
  }
  return [];
}

function normalizeConfidence(confidence: unknown): number | undefined {
  if (typeof confidence !== "number" || !Number.isFinite(confidence)) {
    return undefined;
  }
  return Math.min(1, Math.max(0, confidence));
}

function normalizeMetadata(metadata: unknown): Record<string, unknown> | undefined {
  if (!metadata || typeof metadata !== "object" || Array.isArray(metadata)) {
    return undefined;
  }
  return metadata as Record<string, unknown>;
}

function normalizeCitations(citations: unknown): RuntimeMemoryCitation[] | undefined {
  if (!Array.isArray(citations)) return undefined;

  const normalized = citations.flatMap((citation): RuntimeMemoryCitation[] => {
    if (!citation || typeof citation !== "object") return [];
    const threadId = (citation as { threadId?: unknown }).threadId;
    if (typeof threadId !== "string" || !threadId.trim()) return [];

    const normalizedCitation: RuntimeMemoryCitation = {
      threadId: threadId.trim(),
    };
    const eventId = (citation as { eventId?: unknown }).eventId;
    if (typeof eventId === "number" && Number.isFinite(eventId)) {
      normalizedCitation.eventId = eventId;
    }
    const turnId = (citation as { turnId?: unknown }).turnId;
    if (typeof turnId === "string") {
      normalizedCitation.turnId = turnId;
    }
    return [normalizedCitation];
  });

  return normalized.length > 0 ? normalized : undefined;
}

function looksLikeTableDump(content: string): boolean {
  const lines = content
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const tableLines = lines.filter((line) => line.split("|").length >= 4);
  if (tableLines.length >= 3) return true;
  return hasDelimitedRows(lines, "\t") || hasDelimitedRows(lines, ",");
}

function hasDelimitedRows(lines: string[], delimiter: string): boolean {
  if (lines.length < 3) return false;
  const columnCounts = lines.map((line) => line.split(delimiter).length);
  return columnCounts.filter((count) => count >= 3).length >= 3;
}
