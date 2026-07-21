/**
 * Request history budget for provider.streamChat copies.
 *
 * Token estimation mirrors desktop/electron/agent/memory/compaction.ts
 * (estimateTokens / estimateRequestTokens). Output reserve mirrors
 * desktop/electron/agent/core/agentLoop/maxTokens.ts off-mode sizing.
 *
 * Trimming never mutates the AgentLoop working history — only request copies.
 */

import type { ToolDefinition } from "../tools/types";
import type { AgentMessage } from "./types";

/** Default context window when provider leaves size unset (desktop DEFAULT_CONTEXT_WINDOW). */
export const DEFAULT_CONTEXT_WINDOW = 128_000;

export interface RequestTokenEstimateInput {
  messages?: AgentMessage[];
  systemPrompt?: string;
  tools?: ToolDefinition[];
}

export interface TrimMessagesForRequestInput {
  messages: AgentMessage[];
  systemPrompt: string;
  tools?: ToolDefinition[];
  contextWindowSize: number;
  /**
   * Index of the current-turn user message in `messages`.
   * Messages at and after this index are never dropped for old-history budget.
   */
  protectFromIndex: number;
}

/**
 * Rough token estimate (desktop compaction.ts):
 * Chinese ≈ 1.5 chars/token, other ≈ 4 chars/token.
 */
export function estimateTokens(text: string): number {
  const charCount = text.length;
  const chineseChars = (text.match(/[\u4e00-\u9fff]/g) || []).length;
  const otherChars = charCount - chineseChars;
  return Math.ceil(chineseChars / 1.5 + otherChars / 4);
}

function estimateMessageContentTokens(content: unknown): number {
  if (!content) return 0;
  if (typeof content === "string") return estimateTokens(content);
  if (Array.isArray(content)) {
    return content.reduce((sum, part) => sum + estimateContentPartTokens(part), 0);
  }
  return estimateTokens(JSON.stringify(content));
}

function estimateContentPartTokens(part: unknown): number {
  if (!part) return 0;
  if (typeof part === "string") return estimateTokens(part);
  if (typeof part !== "object") return estimateTokens(String(part));

  const value = part as Record<string, unknown>;
  if (typeof value.text === "string") return estimateTokens(value.text);
  if (typeof value.content === "string") return estimateTokens(value.content);
  if (typeof value.output_text === "string") return estimateTokens(value.output_text);
  if (value.image_url) return estimateTokens(JSON.stringify(value.image_url)) + 85;
  if (value.file) return estimateTokens(JSON.stringify(value.file));
  return estimateTokens(JSON.stringify(value));
}

/** Estimate request payload tokens: system prompt, messages, tool schemas. */
export function estimateRequestTokens(input: RequestTokenEstimateInput): number {
  let total = 0;

  if (input.systemPrompt) {
    total += estimateTokens(input.systemPrompt) + 4;
  }

  for (const message of input.messages || []) {
    total += 4;
    total += estimateTokens(message.role || "");
    total += estimateMessageContentTokens(message.content);
    if (message.contentParts?.length) {
      for (const part of message.contentParts) {
        if (part.type === "text") total += estimateTokens(part.text);
        if (part.type === "image") total += 85 + estimateTokens(part.mimeType);
      }
    }
    if (message.name) total += estimateTokens(message.name);
    if (message.toolCallId) total += estimateTokens(message.toolCallId);
    if (message.toolCalls?.length) {
      total += estimateTokens(JSON.stringify(message.toolCalls));
      total += message.toolCalls.length * 20;
    }
  }

  if (input.tools?.length) {
    total += estimateTokens(JSON.stringify(input.tools));
    total += input.tools.length * 10;
  }

  return total;
}

/**
 * Output token reserve aligned with desktop resolveMaxTokens for reasoningMode "off":
 * baseFromCtx = clamp(floor(ctx * 0.06), 4096, 16384), then min(max(base, 4096), ctx * 0.08).
 */
export function resolveOutputReserve(contextWindowSize: number): number {
  const ctx =
    Number.isFinite(contextWindowSize) && contextWindowSize > 0
      ? Math.floor(contextWindowSize)
      : DEFAULT_CONTEXT_WINDOW;
  const baseFromCtx = Math.min(Math.max(Math.floor(ctx * 0.06), 4_096), 16_384);
  return Math.min(Math.max(baseFromCtx, 4_096), ctx * 0.08);
}

export function resolveMessageTokenBudget(input: {
  contextWindowSize: number;
  systemPrompt: string;
  tools?: ToolDefinition[];
}): number {
  const ctx =
    Number.isFinite(input.contextWindowSize) && input.contextWindowSize > 0
      ? Math.floor(input.contextWindowSize)
      : DEFAULT_CONTEXT_WINDOW;
  const overhead = estimateRequestTokens({
    systemPrompt: input.systemPrompt,
    tools: input.tools,
    messages: [],
  });
  const reserve = resolveOutputReserve(ctx);
  return Math.max(0, ctx - overhead - reserve);
}

/**
 * Group messages into droppable atoms for history budget.
 *
 * Primary unit is a user-initiated history turn:
 *   user + following assistant/tool messages until the next user.
 * That keeps Q/A pairs intact and, by inclusion, keeps assistant toolCalls
 * with their matching tool results (AgentLoop appends tools after the
 * assistant that issued them, before the next user).
 *
 * Leading non-user messages (malformed / partial history) are bundled until
 * the next user so we never drop half of a tool chain.
 */
export function groupMessageAtoms(messages: AgentMessage[]): AgentMessage[][] {
  const atoms: AgentMessage[][] = [];
  let i = 0;
  while (i < messages.length) {
    const group: AgentMessage[] = [messages[i]];
    let j = i + 1;
    // Consume until the next user message starts a new turn.
    while (j < messages.length && messages[j].role !== "user") {
      group.push(messages[j]);
      j += 1;
    }
    atoms.push(group);
    i = j;
  }
  return atoms;
}

/**
 * Build a request-only message copy that fits the context budget.
 * Drops oldest history atoms first; never mutates inputs; never drops protected suffix.
 */
export function trimMessagesForRequest(
  input: TrimMessagesForRequestInput,
): AgentMessage[] {
  const source = input.messages;
  const protectFrom = Math.max(
    0,
    Math.min(
      Number.isInteger(input.protectFromIndex) ? input.protectFromIndex : 0,
      source.length,
    ),
  );

  const fullEstimate = estimateRequestTokens({
    systemPrompt: input.systemPrompt,
    tools: input.tools,
    messages: source,
  });
  const ctx =
    Number.isFinite(input.contextWindowSize) && input.contextWindowSize > 0
      ? Math.floor(input.contextWindowSize)
      : DEFAULT_CONTEXT_WINDOW;
  const reserve = resolveOutputReserve(ctx);
  const hardLimit = Math.max(0, ctx - reserve);

  if (fullEstimate <= hardLimit) {
    return source.slice();
  }

  const history = source.slice(0, protectFrom);
  const protectedSuffix = source.slice(protectFrom);
  const atoms = groupMessageAtoms(history);

  // Drop oldest atoms until estimate fits; protected suffix always kept.
  let dropCount = 0;
  while (dropCount <= atoms.length) {
    const keptHistory: AgentMessage[] = [];
    for (let a = dropCount; a < atoms.length; a += 1) {
      keptHistory.push(...atoms[a]);
    }
    const candidate = keptHistory.concat(protectedSuffix);
    const estimate = estimateRequestTokens({
      systemPrompt: input.systemPrompt,
      tools: input.tools,
      messages: candidate,
    });
    if (estimate <= hardLimit || dropCount === atoms.length) {
      return candidate;
    }
    dropCount += 1;
  }

  return protectedSuffix.slice();
}
