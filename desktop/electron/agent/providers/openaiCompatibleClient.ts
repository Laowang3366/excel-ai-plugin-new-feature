/**
 * OpenAI 兼容客户端基类
 *
 * 覆盖大部分国内厂商（DeepSeek、Kimi、智谱、小米等）。
 * 提供：
 * - SSE 流式响应解析
 * - 工具调用（function calling）支持
 * - 推理/思考模式（通过特殊字段）
 * - 消息格式化与安全校验（孤立 tool_call 剔除）
 * - 非流式聊天（用于压缩时的摘要生成）
 *
 * 关联模块：
 * - aiClientTypes.ts — 提供 ChatMessage、StreamChatParams、AIStreamEvent 等类型
 * - anthropicClient.ts — 继承本类，重写消息格式和 SSE 解析
 * - providerClients.ts — 各厂商子类继承本类，覆盖 applyReasoningConfig
 * - agentLoop/streamCollector.ts — 调用 streamChat() 获取流式事件
 * - agentLoop/summaryGenerator.ts — 调用 chat() 获取非流式摘要
 *
 * 工具名称清洗：
 * - sanitizeToolName() 将点号转下划线（DeepSeek 等要求 ^[a-zA-Z0-9_-]+$）
 * - desanitizeToolName() 将下划线还原为点号
 * - 两个函数均被 AnthropicClient 使用，因此导出
 */

import type {
  ChatMessage,
  ContentPart,
  StreamChatParams,
  AIClientConfig,
  AIStreamEvent,
  ReasoningMode,
  ToolCallInfo,
} from "./aiClientTypes";
import type { TokenUsage } from "../shared/types";
import { formatProviderHttpError } from "./providerErrors";

type ChatCompletionToolState = {
  id: string;
  name: string;
  arguments: string;
  began: boolean;
};

type ChatCompletionsParserState = {
  currentToolCalls: Map<number, ChatCompletionToolState>;
  emittedText: string;
};

// ============================================================
// 工具名称清洗 — 兼容 DeepSeek 等要求 ^[a-zA-Z0-9_-]+$ 的 API
// ============================================================

/**
 * 将内部工具名（含点号，如 workbook.inspect）转换为 API 安全格式（下划线）
 * DeepSeek V4 等要求 function name 匹配 ^[a-zA-Z0-9_-]+$
 */
export function sanitizeToolName(name: string): string {
  return name.replace(/\./g, "_");
}

/**
 * 将 API 返回的工具名还原为内部格式（下划线→点号）
 */
export function desanitizeToolName(name: string): string {
  const compoundPrefixes: Record<string, string> = {
    office_action_: "office.action.",
    office_script_: "office.script.",
  };
  for (const [safePrefix, internalPrefix] of Object.entries(compoundPrefixes)) {
    if (name.startsWith(safePrefix)) {
      return internalPrefix + name.slice(safePrefix.length);
    }
  }

  // 还原已知前缀: workbook_ → workbook., word_ → word., 等
  // 通用策略：第一个下划线如果是已知命名空间则还原。
  const prefixes = [
    "workbook", "range", "selection", "formula", "vba", "sheet",
    "script", "ui", "file", "shell", "word", "presentation", "office", "knowledge",
    "web", "ocr", "memory", "python",
  ];
  for (const prefix of prefixes) {
    if (name.startsWith(prefix + "_")) {
      return prefix + "." + name.slice(prefix.length + 1);
    }
  }
  return name;
}

// ============================================================
// OpenAI 兼容客户端（覆盖大部分国内厂商）
// ============================================================

export class OpenAICompatibleClient {
  protected config: AIClientConfig;

  constructor(config: AIClientConfig) {
    this.config = config;
  }

  /**
   * 流式聊天 — 所有 OpenAI 兼容的 API 都走这个方法
   *
   * 支持的特性：
   * - SSE 流式响应
   * - 工具调用（function calling）
   * - 推理/思考模式（通过特殊字段）
   */
  async *streamChat(params: StreamChatParams): AsyncGenerator<AIStreamEvent> {
    const { messages, tools, systemPrompt, maxTokens, temperature, signal } = params;

    // 构建请求体
    const requestMessages = this.buildRequestMessages(messages, systemPrompt);
    const requestBody: Record<string, unknown> = {
      model: this.config.model,
      messages: requestMessages,
      stream: true,
      max_tokens: maxTokens || 4096,
      temperature: temperature ?? 0.7,
    };

    // 添加工具定义（名称清洗为 API 安全格式）
    if (tools && tools.length > 0) {
      requestBody.tools = tools.map((t) => ({
        type: "function",
        function: {
          name: sanitizeToolName(t.name),
          description: t.description,
          parameters: t.parameters,
        },
      }));
    }

    // 特定厂商的推理模式配置
    // 优先使用 reasoningMode，向后兼容 enableReasoning + reasoningEffort
    const effectiveMode: ReasoningMode | undefined =
      params.reasoningMode ||
      this.config.reasoningMode ||
      (params.enableReasoning || this.config.enableReasoning
        ? (params.reasoningEffort || "high")
        : undefined);
    if (effectiveMode && effectiveMode !== "off") {
      this.applyReasoningConfig(requestBody, effectiveMode);
    }

    // 发起请求
    const url = `${this.config.baseUrl}/chat/completions`;
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${this.config.apiKey}`,
      ...this.config.customHeaders,
    };

    // 组合用户取消信号与 HTTP 超时信号（5 分钟）
    // 流式请求可能持续较长（推理模式 + 大上下文），5 分钟是合理上限
    const timeoutSignal = AbortSignal.timeout(5 * 60 * 1000);
    const combinedSignal = signal
      ? AbortSignal.any([signal, timeoutSignal])
      : timeoutSignal;

    const response = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(requestBody),
      signal: combinedSignal,
    });

    if (!response.ok) {
      const errorText = await response.text();
      yield { type: "error", error: formatProviderHttpError("API 请求失败", response.status, errorText) };
      return;
    }

    // 解析 SSE 流
    yield* this.parseSSEStream(response);
  }

  /** 构建请求消息数组 */
  protected buildRequestMessages(
    messages: ChatMessage[],
    systemPrompt?: string
  ): Record<string, unknown>[] {
    const result: Record<string, unknown>[] = [];

    if (systemPrompt) {
      result.push({ role: "system", content: systemPrompt });
    }

    // ── 安全校验：确保每个 assistant tool_call 都有对应的 tool 消息 ──
    // 先收集所有 tool 消息的 tool_call_id
    const toolResultIds = new Set<string>();
    for (const msg of messages) {
      if (msg.role === "tool" && msg.toolCallId) {
        toolResultIds.add(msg.toolCallId);
      }
    }
    // 检查每个 assistant 消息的 toolCalls，剔除没有对应 tool 消息的孤立项
    const sanitizedMessages: ChatMessage[] = [];
    for (const msg of messages) {
      if (msg.role === "assistant" && msg.toolCalls && msg.toolCalls.length > 0) {
        // 过滤掉孤立的 tool_call（没有对应的 tool result 消息）
        const validToolCalls = msg.toolCalls.filter(tc => toolResultIds.has(tc.id));
        if (validToolCalls.length === 0) {
          // 所有 toolCalls 都是孤立的，只保留文本内容（跳过纯工具调用消息）
          if (msg.content && (typeof msg.content === "string" ? msg.content.trim() : true)) {
            sanitizedMessages.push({ ...msg, toolCalls: undefined });
          }
          // 否则完全跳过这条空壳消息
        } else if (validToolCalls.length < msg.toolCalls.length) {
          // 部分孤立，保留有配对的 toolCalls
          sanitizedMessages.push({ ...msg, toolCalls: validToolCalls });
        } else {
          // 全部有配对，原样保留
          sanitizedMessages.push(msg);
        }
      } else {
        sanitizedMessages.push(msg);
      }
    }

    // ── 二次校验：确保连续的 assistant(toolCalls) 后紧跟对应的 tool 消息 ──
    // 处理 tool 消息被意外删除或顺序错误导致中间断开的情况
    const finalMessages: ChatMessage[] = [];
    let i = 0;
    while (i < sanitizedMessages.length) {
      const msg = sanitizedMessages[i];
      if (msg.role === "assistant" && msg.toolCalls && msg.toolCalls.length > 0) {
        // 收集这条 assistant 消息后面的所有连续 tool 消息
        finalMessages.push(msg);
        i++;
        const expectedIds = new Set(msg.toolCalls.map(tc => tc.id));
        const foundIds = new Set<string>();
        // 消耗紧跟的 tool 消息
        while (i < sanitizedMessages.length && sanitizedMessages[i].role === "tool") {
          const toolMsg = sanitizedMessages[i];
          if (expectedIds.has(toolMsg.toolCallId || "")) {
            foundIds.add(toolMsg.toolCallId || "");
            finalMessages.push(toolMsg);
          }
          // 不属于当前 assistant 的 tool 消息也保留
          if (!expectedIds.has(toolMsg.toolCallId || "")) {
            finalMessages.push(toolMsg);
          }
          i++;
        }
        // 如果有 toolCall 没找到对应的 tool 消息，从 assistant 中移除
        if (foundIds.size < expectedIds.size) {
          const assistantMsg = finalMessages[finalMessages.length - (foundIds.size + 1)];
          if (assistantMsg && assistantMsg.role === "assistant" && assistantMsg.toolCalls) {
            const missingIds = new Set([...expectedIds].filter(id => !foundIds.has(id)));
            assistantMsg.toolCalls = assistantMsg.toolCalls.filter(tc => !missingIds.has(tc.id));
            // 如果全部被移除，且内容为空，则删除这条 assistant 消息
            if (assistantMsg.toolCalls.length === 0) {
              const idx = finalMessages.indexOf(assistantMsg);
              if (idx !== -1) {
                finalMessages.splice(idx, 1);
              }
            }
          }
        }
      } else {
        finalMessages.push(msg);
        i++;
      }
    }

    // ── 构建最终 API 格式 ──
    for (const msg of finalMessages) {
      if (msg.role === "assistant" && msg.toolCalls && msg.toolCalls.length > 0) {
        // 助手消息 + 工具调用（名称清洗为 API 安全格式）
        result.push({
          role: "assistant",
          content: typeof msg.content === "string" ? (msg.content || null) : msg.content,
          tool_calls: msg.toolCalls.map((tc) => ({
            id: tc.id,
            type: "function",
            function: {
              name: sanitizeToolName(tc.function.name),
              arguments: tc.function.arguments,
            },
          })),
        });
      } else if (msg.role === "tool") {
        result.push({
          role: "tool",
          content: msg.content,
          tool_call_id: msg.toolCallId,
        });
      } else {
        result.push({
          role: msg.role,
          content: msg.content,
        });
      }
    }

    return result;
  }

  /** 应用推理模式配置（子类可覆盖） */
  protected applyReasoningConfig(
    _body: Record<string, unknown>,
    _mode?: ReasoningMode
  ): void {
    // 默认不添加推理配置，子类按需覆盖
  }

  /** 解析 SSE 流（通用实现） */
  protected async *parseSSEStream(response: Response): AsyncGenerator<AIStreamEvent> {
    const reader = response.body?.getReader();
    if (!reader) {
      yield { type: "error", error: "无法读取响应流" };
      return;
    }

    const decoder = new TextDecoder();
    let buffer = "";
    const state: ChatCompletionsParserState = {
      currentToolCalls: new Map(),
      emittedText: "",
    };

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop() || "";

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed || trimmed === "data: [DONE]") continue;
          if (!trimmed.startsWith("data: ")) continue;

          try {
            const data = JSON.parse(trimmed.slice(6));
            yield* this.processStreamChunk(data, state);
          } catch {
            // 跳过解析失败的行
          }
        }
      }
    } finally {
      reader.releaseLock();
    }
  }

  /** 处理单个 SSE chunk */
  protected *processStreamChunk(
    data: any,
    state: ChatCompletionsParserState
  ): Generator<AIStreamEvent> {
    const choice = data.choices?.[0];
    if (!choice) return;

    const delta = choice.delta;

    // 处理推理/思考内容
    // 注意：某些 API 可能同时返回 reasoning_content 和 reasoning.content，
    // 优先使用 reasoning_content（DeepSeek/Kimi 格式），避免重复 yield
    if (delta?.reasoning_content) {
      yield { type: "reasoning_delta", delta: delta.reasoning_content };
    } else if (delta?.reasoning?.content) {
      yield { type: "reasoning_delta", delta: delta.reasoning.content };
    }

    // 处理普通文本
    if (delta?.content) {
      yield* emitChatCompletionsDelta(extractChatCompletionsText(delta.content), state);
    }

    const choiceMessageText = extractChatCompletionsText(choice.message?.content);
    if (choiceMessageText) {
      yield* emitMissingChatCompletionsText(choiceMessageText, state);
    }

    // 处理工具调用（名称还原为内部格式）
    if (delta?.tool_calls) {
      for (const tc of delta.tool_calls) {
        const idx = tc.index ?? 0;
        if (!state.currentToolCalls.has(idx)) {
          state.currentToolCalls.set(idx, {
            id: "",
            name: "",
            arguments: "",
            began: false,
          });
        }
        const current = state.currentToolCalls.get(idx)!;
        if (tc.id) current.id = tc.id;
        if (tc.function?.name) current.name = desanitizeToolName(tc.function.name);
        if (current.id && current.name && !current.began) {
          current.began = true;
          yield {
            type: "tool_call_begin",
            toolCallId: current.id,
            toolName: current.name,
          };
        }
        if (tc.function?.arguments) {
          current.arguments += tc.function.arguments;
          if (current.began) {
            yield {
              type: "tool_call_delta",
              toolCallId: current.id,
              delta: tc.function.arguments,
            };
          }
        }
      }
    }

    // 完成
    if (choice.finish_reason) {
      // 先发送所有工具调用的 end 事件
      for (const [, tc] of state.currentToolCalls) {
        if (!tc.id && !tc.name && !tc.arguments) continue;
        yield {
          type: "tool_call_end",
          toolCallId: tc.id,
          toolName: tc.name,
          arguments: tc.arguments,
        };
      }
      state.currentToolCalls.clear();

      // 使用量
      if (data.usage) {
        yield {
          type: "usage",
          usage: {
            inputTokens: data.usage.prompt_tokens || 0,
            outputTokens: data.usage.completion_tokens || 0,
            cachedInputTokens: data.usage.prompt_tokens_details?.cached_tokens || 0,
            reasoningOutputTokens: data.usage.completion_tokens_details?.reasoning_tokens || 0,
          },
        };
      }

      yield { type: "done", finishReason: choice.finish_reason };
    }
  }

  /** 非流式聊天（用于压缩时的摘要生成） */
  async chat(params: StreamChatParams): Promise<{ content: string; usage?: TokenUsage }> {
    const events: AIStreamEvent[] = [];
    for await (const event of this.streamChat(params)) {
      events.push(event);
    }

    let content = "";
    let usage: TokenUsage | undefined;

    for (const event of events) {
      if (event.type === "text_delta") content += event.delta;
      if (event.type === "usage") usage = event.usage;
    }

    return { content, usage };
  }
}

function extractChatCompletionsText(content: unknown): string {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";

  return content
    .map((part) => {
      if (typeof part === "string") return part;
      if (!part || typeof part !== "object") return "";

      const value = part as Record<string, unknown>;
      if (typeof value.text === "string") return value.text;
      if (typeof value.output_text === "string") return value.output_text;
      if (typeof value.content === "string") return value.content;
      if (Array.isArray(value.content)) return extractChatCompletionsText(value.content);
      return "";
    })
    .join("");
}

function *emitChatCompletionsDelta(
  text: string,
  state: ChatCompletionsParserState
): Generator<AIStreamEvent> {
  if (!text) return;

  const delta = state.emittedText && text.startsWith(state.emittedText)
    ? text.slice(state.emittedText.length)
    : text;
  if (!delta) return;

  state.emittedText += delta;
  yield { type: "text_delta", delta };
}

function *emitMissingChatCompletionsText(
  fullText: string,
  state: ChatCompletionsParserState
): Generator<AIStreamEvent> {
  if (!fullText) return;

  if (!state.emittedText) {
    state.emittedText = fullText;
    yield { type: "text_delta", delta: fullText };
    return;
  }

  if (!fullText.startsWith(state.emittedText)) return;

  const delta = fullText.slice(state.emittedText.length);
  if (!delta) return;

  state.emittedText = fullText;
  yield { type: "text_delta", delta };
}
