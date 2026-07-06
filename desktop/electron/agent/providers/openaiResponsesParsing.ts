import { desanitizeToolName, sanitizeToolName } from "./openaiCompatibleClient";
import type {
  AIStreamEvent,
  ChatMessage,
  ContentPart,
  ReasoningMode,
  ToolCallInfo,
} from "./aiClientTypes";
import type { TokenUsage } from "../shared/types";

export type ResponseToolState = {
  id: string;
  name: string;
  arguments: string;
  began: boolean;
};

export type ResponseParserState = {
  toolByItemId: Map<string, ResponseToolState>;
  itemIdByCallId: Map<string, string>;
  emittedText: string;
  textByPartKey: Map<string, string>;
};

export function createResponseParserState(): ResponseParserState {
  return {
    toolByItemId: new Map(),
    itemIdByCallId: new Map(),
    emittedText: "",
    textByPartKey: new Map(),
  };
}

export function responseContentPartFromContentPart(part: ContentPart): Record<string, unknown> {
  if (part.type === "text") {
    return { type: "input_text", text: part.text };
  }
  if (part.type === "image_url") {
    return {
      type: "input_image",
      image_url: part.image_url.url,
      detail: part.image_url.detail || "auto",
    };
  }
  return {
    type: "input_file",
    filename: part.file.filename,
    file_data: part.file.file_data,
  };
}

export function responseFunctionCallFromToolCall(tc: ToolCallInfo): Record<string, unknown> {
  return {
    type: "function_call",
    call_id: tc.id,
    name: sanitizeToolName(tc.function.name),
    arguments: tc.function.arguments || "{}",
    status: "completed",
  };
}

export function stringifyContent(content: ChatMessage["content"]): string {
  if (typeof content === "string") return content;
  return content
    .map((part) => {
      if (part.type === "text") return part.text;
      if (part.type === "image_url") return `[image: ${part.image_url.url}]`;
      return `[file: ${part.file.filename || "unnamed"}]`;
    })
    .join("\n");
}

export function upsertToolState(item: any, state: ResponseParserState): ResponseToolState {
  const itemId = item.id || item.call_id || "";
  const callId = item.call_id || item.id || itemId;
  if (itemId) state.itemIdByCallId.set(callId, itemId);

  const existing = state.toolByItemId.get(itemId);
  if (existing) {
    if (item.name) existing.name = desanitizeToolName(item.name);
    return existing;
  }

  const next: ResponseToolState = {
    id: callId,
    name: desanitizeToolName(item.name || ""),
    arguments: typeof item.arguments === "string" ? item.arguments : "",
    began: false,
  };
  state.toolByItemId.set(itemId, next);
  return next;
}

export function getToolState(
  data: any,
  state: ResponseParserState
): ResponseToolState | undefined {
  const itemId = data.item_id || state.itemIdByCallId.get(data.call_id);
  if (itemId && state.toolByItemId.has(itemId)) {
    return state.toolByItemId.get(itemId);
  }
  if (data.call_id) {
    return upsertToolState({ id: data.item_id || data.call_id, call_id: data.call_id }, state);
  }
  return undefined;
}

function responseTextPartKey(data: any, fallbackIndex = 0): string {
  const itemId = data?.item_id || data?.item?.id || `output:${data?.output_index ?? "unknown"}`;
  const contentIndex = data?.content_index ?? data?.part_index ?? fallbackIndex;
  return `${itemId}:${contentIndex}`;
}

export function appendResponseTextPart(
  data: any,
  delta: string,
  state: ResponseParserState
): void {
  const key = responseTextPartKey(data);
  state.textByPartKey.set(key, `${state.textByPartKey.get(key) || ""}${delta}`);
}

export function *emitMissingResponseTextPart(
  data: any,
  fullText: unknown,
  state: ResponseParserState,
  fallbackIndex = 0
): Generator<AIStreamEvent> {
  if (typeof fullText !== "string" || !fullText) return;

  const key = responseTextPartKey(data, fallbackIndex);
  const currentText = state.textByPartKey.get(key) || "";
  const delta = missingSuffix(fullText, currentText);
  state.textByPartKey.set(key, fullText);
  if (!delta) return;

  state.emittedText += delta;
  yield { type: "text_delta", delta };
}

export function *emitMissingResponseOutputItemText(
  item: any,
  state: ResponseParserState
): Generator<AIStreamEvent> {
  if (!Array.isArray(item?.content)) return;

  for (let index = 0; index < item.content.length; index++) {
    const text = extractResponseContentText(item.content[index]);
    yield* emitMissingResponseTextPart(
      { item_id: item.id, content_index: index },
      text,
      state,
      index
    );
  }
}

export function *emitMissingCompletedResponseText(
  response: any,
  state: ResponseParserState
): Generator<AIStreamEvent> {
  const fullText = extractResponsesText(response);
  const delta = missingSuffix(fullText, state.emittedText);
  if (!delta) return;

  state.emittedText += delta;
  yield { type: "text_delta", delta };
}

export function extractResponseContentText(content: any): string {
  if (typeof content?.text === "string") return content.text;
  if (typeof content?.content === "string") return content.content;
  if (typeof content?.output_text === "string") return content.output_text;
  return "";
}

function missingSuffix(fullText: string, currentText: string): string {
  if (!fullText) return "";
  if (!currentText) return fullText;
  if (fullText.startsWith(currentText)) return fullText.slice(currentText.length);
  return "";
}

export function toOpenAIResponsesReasoningEffort(mode: ReasoningMode): string {
  if (mode === "off") return "none";
  if (mode === "max") return "xhigh";
  return mode;
}

export function extractResponsesText(data: any): string {
  if (typeof data?.output_text === "string") {
    return data.output_text;
  }

  const parts: string[] = [];
  for (const item of data?.output || []) {
    if (item?.type === "message") {
      for (const content of item.content || []) {
        if (typeof content?.text === "string") parts.push(content.text);
      }
    }
  }
  return parts.join("");
}

export function normalizeResponsesUsage(usage: any): TokenUsage | undefined {
  if (!usage) return undefined;
  return {
    inputTokens: usage.input_tokens || 0,
    outputTokens: usage.output_tokens || 0,
    cachedInputTokens: usage.input_tokens_details?.cached_tokens || 0,
    reasoningOutputTokens: usage.output_tokens_details?.reasoning_tokens || 0,
  };
}

export function responseContainsFunctionCall(response: any): boolean {
  return Array.isArray(response?.output)
    && response.output.some((item: any) => item?.type === "function_call");
}
