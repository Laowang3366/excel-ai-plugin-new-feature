import { OpenAICompatibleClient, sanitizeToolName } from "./openaiCompatibleClient";
import type {
  AIClientConfig,
  AIStreamEvent,
  ChatMessage,
  ReasoningMode,
  StreamChatParams,
} from "./aiClientTypes";
import type { TokenUsage } from "../shared/types";
import { formatProviderHttpError } from "./providerErrors";
import {
  appendResponseTextPart,
  createResponseParserState,
  emitMissingCompletedResponseText,
  emitMissingResponseOutputItemText,
  emitMissingResponseTextPart,
  extractResponseContentText,
  extractResponsesText,
  getToolState,
  normalizeResponsesUsage,
  responseContainsFunctionCall,
  responseContentPartFromContentPart,
  responseFunctionCallFromToolCall,
  stringifyContent,
  toOpenAIResponsesReasoningEffort,
  upsertToolState,
  type ResponseParserState,
} from "./openaiResponsesParsing";

export class OpenAIResponsesClient extends OpenAICompatibleClient {
  constructor(config: AIClientConfig) {
    super(config);
  }

  async *streamChat(params: StreamChatParams): AsyncGenerator<AIStreamEvent> {
    const requestBody = this.buildRequestBody(params, true);
    const response = await this.postResponses(requestBody, params.signal);

    if (!response.ok) {
      const errorText = await response.text();
      yield {
        type: "error",
        error: formatProviderHttpError("Responses API 请求失败", response.status, errorText),
      };
      return;
    }

    yield* this.parseResponsesSSEStream(response);
  }

  async chat(params: StreamChatParams): Promise<{ content: string; usage?: TokenUsage }> {
    const requestBody = this.buildRequestBody(params, false);
    const response = await this.postResponses(requestBody, params.signal);

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        formatProviderHttpError("Responses API 请求失败", response.status, errorText),
      );
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
      params.reasoningMode || this.config.reasoningMode;
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

  private async postResponses(
    body: Record<string, unknown>,
    signal?: AbortSignal,
  ): Promise<Response> {
    const timeoutSignal = AbortSignal.timeout(5 * 60 * 1000);
    const combinedSignal = signal ? AbortSignal.any([signal, timeoutSignal]) : timeoutSignal;

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
    const state = createResponseParserState();
    let buffer = "";

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const chunks = buffer.split("\n\n");
        buffer = chunks.pop() || "";

        for (const chunk of chunks) {
          yield* this.processResponsesSSEChunk(chunk, state);
        }
      }
      if (buffer.trim()) {
        yield* this.processResponsesSSEChunk(buffer, state);
      }
    } finally {
      reader.releaseLock();
    }
  }

  private *processResponsesSSEChunk(
    chunk: string,
    state: ResponseParserState,
  ): Generator<AIStreamEvent> {
    const dataLine = chunk
      .split("\n")
      .map((line) => line.trim())
      .find((line) => line.startsWith("data: "));
    if (!dataLine || dataLine === "data: [DONE]") return;

    try {
      const data = JSON.parse(dataLine.slice(6));
      yield* this.processResponsesEvent(data, state);
    } catch {
      // Ignore malformed stream chunks and continue reading.
    }
  }

  private *processResponsesEvent(data: any, state: ResponseParserState): Generator<AIStreamEvent> {
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
        yield {
          type: "done",
          finishReason: data.response?.incomplete_details?.reason || "incomplete",
        };
        return;

      case "response.failed":
        yield {
          type: "error",
          error: data.response?.error?.message || data.error?.message || "Responses API 请求失败",
        };
        return;

      case "error":
        yield {
          type: "error",
          error: data.error?.message || data.message || "Responses API 请求失败",
        };
        return;

      default:
        return;
    }
  }
}
