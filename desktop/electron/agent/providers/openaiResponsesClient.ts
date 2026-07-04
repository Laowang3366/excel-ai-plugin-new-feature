import { OpenAICompatibleClient, desanitizeToolName, sanitizeToolName } from "./openaiCompatibleClient";
import type {
  AIClientConfig,
  AIStreamEvent,
  ChatMessage,
  ContentPart,
  ReasoningMode,
  StreamChatParams,
  ToolCallInfo,
} from "./aiClientTypes";
import type { TokenUsage } from "../shared/types";
import { formatProviderHttpError } from "./providerErrors";

type ResponseToolState = {
  id: string;
  name: string;
  arguments: string;
  began: boolean;
};

type ResponseParserState = {
  toolByItemId: Map<string, ResponseToolState>;
  itemIdByCallId: Map<string, string>;
  emittedText: string;
  textByPartKey: Map<string, string>;
};

export class OpenAIResponsesClient extends OpenAICompatibleClient {
  constructor(config: AIClientConfig) {
    super(config);
  }

  async *streamChat(params: StreamChatParams): AsyncGenerator<AIStreamEvent> {
    const requestBody = this.buildRequestBody(params, true);
    const response = await this.postResponses(requestBody, params.signal);

    if (!response.ok) {
      const errorText = await response.text();
      yield { type: "error", error: formatProviderHttpError("Responses API 请求失败", response.status, errorText) };
      return;
    }

    yield* this.parseResponsesSSEStream(response);
  }

  async chat(params: StreamChatParams): Promise<{ content: string; usage?: TokenUsage }> {
    const requestBody = this.buildRequestBody(params, false);
    const response = await this.postResponses(requestBody, params.signal);

    if (!response.ok) {
      const errorText = await response.text();
      return { content: formatProviderHttpError("Responses API 请求失败", response.status, errorText) };
    }

    const data: any = await response.json();
    return {
      content: extractResponsesText(data),
      usage: normalizeResponsesUsage(data.usage),
    };
  }

  protected buildRequestBody(params: StreamChatParams, stream: boolean): Record<string, unknown> {
    const requestBody: Record<string, unknown> = {
      model: this.config.model,
      input: this.buildResponseInput(params.messages),
      stream,
      max_output_tokens: params.maxTokens || this.config.maxTokens || 4096,
      temperature: params.temperature ?? 0.7,
    };

    if (params.systemPrompt) {
      requestBody.instructions = params.systemPrompt;
    }

    if (params.tools && params.tools.length > 0) {
      requestBody.tools = params.tools.map((t) => ({
        type: "function",
        name: sanitizeToolName(t.name),
        description: t.description,
        parameters: t.parameters,
      }));
    }

    const effectiveMode: ReasoningMode | undefined =
      params.reasoningMode ||
      this.config.reasoningMode ||
      (params.enableReasoning || this.config.enableReasoning
        ? (params.reasoningEffort || "high")
        : undefined);
    if (effectiveMode) {
      const effort = toOpenAIResponsesReasoningEffort(effectiveMode);
      requestBody.reasoning = {
        effort,
        ...(effort !== "none" ? { summary: "auto" } : {}),
      };
    }

    return requestBody;
  }

  protected buildResponseInput(messages: ChatMessage[]): Record<string, unknown>[] {
    const input: Record<string, unknown>[] = [];

    for (const msg of messages) {
      if (msg.role === "tool") {
        input.push({
          type: "function_call_output",
          call_id: msg.toolCallId,
          output: stringifyContent(msg.content),
        });
        continue;
      }

      if (msg.role === "assistant" && msg.toolCalls && msg.toolCalls.length > 0) {
        const textContent = stringifyContent(msg.content).trim();
        if (textContent) {
          input.push({
            role: "assistant",
            content: [{ type: "output_text", text: textContent }],
          });
        }
        for (const tc of msg.toolCalls) {
          input.push(responseFunctionCallFromToolCall(tc));
        }
        continue;
      }

      input.push({
        role: msg.role,
        content: Array.isArray(msg.content)
          ? msg.content.map(responseContentPartFromContentPart)
          : msg.content,
      });
    }

    return input;
  }

  private async postResponses(body: Record<string, unknown>, signal?: AbortSignal): Promise<Response> {
    const timeoutSignal = AbortSignal.timeout(5 * 60 * 1000);
    const combinedSignal = signal
      ? AbortSignal.any([signal, timeoutSignal])
      : timeoutSignal;

    return fetch(`${this.config.baseUrl.replace(/\/+$/, "")}/responses`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${this.config.apiKey}`,
        ...this.config.customHeaders,
      },
      body: JSON.stringify(body),
      signal: combinedSignal,
    });
  }

  private async *parseResponsesSSEStream(response: Response): AsyncGenerator<AIStreamEvent> {
    const reader = response.body?.getReader();
    if (!reader) {
      yield { type: "error", error: "无法读取 Responses 响应流" };
      return;
    }

    const decoder = new TextDecoder();
    const state: ResponseParserState = {
      toolByItemId: new Map(),
      itemIdByCallId: new Map(),
      emittedText: "",
      textByPartKey: new Map(),
    };
    let buffer = "";

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const chunks = buffer.split("\n\n");
        buffer = chunks.pop() || "";

        for (const chunk of chunks) {
          const dataLine = chunk
            .split("\n")
            .map((line) => line.trim())
            .find((line) => line.startsWith("data: "));
          if (!dataLine || dataLine === "data: [DONE]") continue;

          try {
            const data = JSON.parse(dataLine.slice(6));
            yield* this.processResponsesEvent(data, state);
          } catch {
            // Ignore malformed stream chunks and continue reading.
          }
        }
      }
    } finally {
      reader.releaseLock();
    }
  }

  private *processResponsesEvent(
    data: any,
    state: ResponseParserState
  ): Generator<AIStreamEvent> {
    switch (data.type) {
      case "response.output_text.delta":
        if (data.delta) {
          const delta = String(data.delta);
          appendResponseTextPart(data, delta, state);
          state.emittedText += delta;
          yield { type: "text_delta", delta };
        }
        return;

      case "response.output_text.done":
        yield* emitMissingResponseTextPart(data, data.text, state);
        return;

      case "response.content_part.done":
        yield* emitMissingResponseTextPart(data, extractResponseContentText(data.part), state);
        return;

      case "response.reasoning_summary_text.delta":
      case "response.reasoning_text.delta":
        if (data.delta) yield { type: "reasoning_delta", delta: data.delta };
        return;

      case "response.output_item.added": {
        const item = data.item;
        if (item?.type === "function_call") {
          const tool = upsertToolState(item, state);
          if (!tool.began && tool.name) {
            tool.began = true;
            yield { type: "tool_call_begin", toolCallId: tool.id, toolName: tool.name };
          }
        }
        return;
      }

      case "response.function_call_arguments.delta": {
        const tool = getToolState(data, state);
        if (tool) {
          tool.arguments += data.delta || "";
          if (tool.began) {
            yield { type: "tool_call_delta", toolCallId: tool.id, delta: data.delta || "" };
          }
        }
        return;
      }

      case "response.output_item.done": {
        const item = data.item;
        if (item?.type === "function_call") {
          const tool = upsertToolState(item, state);
          tool.arguments = typeof item.arguments === "string" ? item.arguments : tool.arguments;
          if (!tool.began) {
            tool.began = true;
            yield { type: "tool_call_begin", toolCallId: tool.id, toolName: tool.name };
          }
          yield {
            type: "tool_call_end",
            toolCallId: tool.id,
            toolName: tool.name,
            arguments: tool.arguments,
          };
        } else if (item?.type === "message") {
          yield* emitMissingResponseOutputItemText(item, state);
        }
        return;
      }

      case "response.completed": {
        yield* emitMissingCompletedResponseText(data.response, state);
        const usage = normalizeResponsesUsage(data.response?.usage);
        if (usage) yield { type: "usage", usage };
        const finishReason = responseContainsFunctionCall(data.response) ? "tool_calls" : "stop";
        yield { type: "done", finishReason };
        return;
      }

      case "response.incomplete":
        yield { type: "done", finishReason: data.response?.incomplete_details?.reason || "incomplete" };
        return;

      case "response.failed":
        yield {
          type: "error",
          error: data.response?.error?.message || data.error?.message || "Responses API 请求失败",
        };
        return;

      case "error":
        yield { type: "error", error: data.error?.message || data.message || "Responses API 请求失败" };
        return;

      default:
        return;
    }
  }
}

function responseContentPartFromContentPart(part: ContentPart): Record<string, unknown> {
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

function responseFunctionCallFromToolCall(tc: ToolCallInfo): Record<string, unknown> {
  return {
    type: "function_call",
    call_id: tc.id,
    name: sanitizeToolName(tc.function.name),
    arguments: tc.function.arguments || "{}",
    status: "completed",
  };
}

function stringifyContent(content: ChatMessage["content"]): string {
  if (typeof content === "string") return content;
  return content
    .map((part) => {
      if (part.type === "text") return part.text;
      if (part.type === "image_url") return `[image: ${part.image_url.url}]`;
      return `[file: ${part.file.filename || "unnamed"}]`;
    })
    .join("\n");
}

function upsertToolState(item: any, state: ResponseParserState): ResponseToolState {
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

function getToolState(data: any, state: ResponseParserState): ResponseToolState | undefined {
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

function appendResponseTextPart(data: any, delta: string, state: ResponseParserState): void {
  const key = responseTextPartKey(data);
  state.textByPartKey.set(key, `${state.textByPartKey.get(key) || ""}${delta}`);
}

function *emitMissingResponseTextPart(
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

function *emitMissingResponseOutputItemText(
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

function *emitMissingCompletedResponseText(
  response: any,
  state: ResponseParserState
): Generator<AIStreamEvent> {
  const fullText = extractResponsesText(response);
  const delta = missingSuffix(fullText, state.emittedText);
  if (!delta) return;

  state.emittedText += delta;
  yield { type: "text_delta", delta };
}

function extractResponseContentText(content: any): string {
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

function toOpenAIResponsesReasoningEffort(mode: ReasoningMode): string {
  if (mode === "off") return "none";
  if (mode === "max") return "xhigh";
  return mode;
}

function extractResponsesText(data: any): string {
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

function normalizeResponsesUsage(usage: any): TokenUsage | undefined {
  if (!usage) return undefined;
  return {
    inputTokens: usage.input_tokens || 0,
    outputTokens: usage.output_tokens || 0,
    cachedInputTokens: usage.input_tokens_details?.cached_tokens || 0,
    reasoningOutputTokens: usage.output_tokens_details?.reasoning_tokens || 0,
  };
}

function responseContainsFunctionCall(response: any): boolean {
  return Array.isArray(response?.output)
    && response.output.some((item: any) => item?.type === "function_call");
}
