/**
 * Anthropic (Claude) 客户端
 *
 * 继承 OpenAICompatibleClient，重写：
 * - 消息格式（tool_use/tool_result 块，system prompt 单独传）
 * - SSE 解析（content_block 事件）
 * - 认证头（x-api-key 而非 Bearer）
 * - 多模态图片格式（base64 source 块）
 *
 * 关联模块：
 * - openaiCompatibleClient.ts — 基类，提供 sanitizeToolName/desanitizeToolName
 * - aiClientTypes.ts — ChatMessage、StreamChatParams、AIStreamEvent 等类型
 * - aiClient.ts — 工厂函数 createAIClient 根据 apiFormat 路由到本类
 */

import { OpenAICompatibleClient, sanitizeToolName, desanitizeToolName } from "./openaiCompatibleClient";
import {
  type AIClientConfig,
  type AIStreamEvent,
  type ChatMessage,
  type ReasoningMode,
  type StreamChatParams,
} from "./aiClientTypes";
import { formatProviderHttpError } from "./providerErrors";

// ============================================================
// Anthropic 客户端
// ============================================================

export class AnthropicClient extends OpenAICompatibleClient {
  constructor(config: AIClientConfig) {
    super(config);
    // Anthropic 使用不同的认证头
    this.config.customHeaders = {
      ...config.customHeaders,
      "anthropic-version": "2023-06-01",
      "x-api-key": config.apiKey,
    };
  }

  protected buildRequestMessages(
    messages: ChatMessage[],
    systemPrompt?: string
  ): Record<string, unknown>[] {
    // Anthropic 格式不同：system prompt 单独传，messages 中不能有 system role
    const result: Record<string, unknown>[] = [];

    for (const msg of messages) {
      if (msg.role === "system") continue; // 跳过，放到 system 参数中
      if (msg.role === "assistant" && msg.toolCalls && msg.toolCalls.length > 0) {
        result.push({
          role: "assistant",
          content: [
            ...(typeof msg.content === "string" && msg.content ? [{ type: "text", text: msg.content }] : []),
            ...msg.toolCalls.map((tc) => ({
              type: "tool_use",
              id: tc.id,
              name: sanitizeToolName(tc.function.name),
              input: JSON.parse(tc.function.arguments || "{}"),
            })),
          ],
        });
      } else if (msg.role === "tool") {
        result.push({
          role: "user",
          content: [
            {
              type: "tool_result",
              tool_use_id: msg.toolCallId,
              content: msg.content,
            },
          ],
        });
      } else if (msg.role === "user" && Array.isArray(msg.content)) {
        // 多模态用户消息：转换为 Anthropic content block 格式
        const contentBlocks: Record<string, unknown>[] = [];
        for (const part of msg.content) {
          if (part.type === "text") {
            contentBlocks.push({ type: "text", text: part.text });
          } else if (part.type === "image_url") {
            const url = part.image_url.url;
            if (url.startsWith("data:")) {
              // data:image/png;base64,xxxx → Anthropic source 格式
              const match = url.match(/^data:([^;]+);base64,(.+)$/);
              if (match) {
                const mediaType = match[1];
                const base64Data = match[2];
                contentBlocks.push({
                  type: "image",
                  source: {
                    type: "base64",
                    media_type: mediaType,
                    data: base64Data,
                  },
                });
              }
            } else {
              // 非 base64 图片 URL，作为文本描述
              contentBlocks.push({ type: "text", text: `[图片: ${url}]` });
            }
          } else if (part.type === "file") {
            const fileData = part.file.file_data;
            const match = fileData.match(/^data:([^;]+);base64,(.+)$/);
            const mediaType = part.file.mime_type || match?.[1] || "application/pdf";
            const base64Data = match?.[2] || fileData;
            if (mediaType === "application/pdf") {
              contentBlocks.push({
                type: "document",
                source: {
                  type: "base64",
                  media_type: mediaType,
                  data: base64Data,
                },
              });
            } else {
              contentBlocks.push({
                type: "text",
                text: `[文件: ${part.file.filename || "未命名文件"} (${mediaType})]`,
              });
            }
          }
        }
        result.push({ role: "user", content: contentBlocks });
      } else {
        result.push({ role: msg.role, content: msg.content });
      }
    }

    return result;
  }

  async *streamChat(params: StreamChatParams): AsyncGenerator<AIStreamEvent> {
    const { messages, tools, systemPrompt, maxTokens, temperature, signal } = params;
    const requestMessages = this.buildRequestMessages(messages, systemPrompt);

    const requestBody: Record<string, unknown> = {
      model: this.config.model,
      messages: requestMessages,
      max_tokens: maxTokens || 4096,
      temperature: temperature ?? 0.7,
      stream: true,
    };

    if (systemPrompt) {
      requestBody.system = systemPrompt;
    }

    if (tools && tools.length > 0) {
      requestBody.tools = tools.map((t) => ({
        name: sanitizeToolName(t.name),
        description: t.description,
        input_schema: t.parameters,
      }));
    }

    // Anthropic adaptive thinking + effort level
    const effectiveMode: ReasoningMode | undefined =
      params.reasoningMode ||
      this.config.reasoningMode ||
      (params.enableReasoning || this.config.enableReasoning ? "high" : undefined);
    if (effectiveMode && effectiveMode !== "off") {
      requestBody.thinking = { type: "adaptive" };
      requestBody.output_config = {
        effort: effectiveMode === "max" ? "xhigh" : effectiveMode,
      };
    }

    const url = `${this.config.baseUrl}/messages`;
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      "anthropic-version": "2023-06-01",
      "x-api-key": this.config.apiKey,
      ...this.config.customHeaders,
    };
    // 移除 Authorization header（Anthropic 不用 Bearer）
    delete headers.Authorization;

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
      yield { type: "error", error: formatProviderHttpError("Anthropic API 错误", response.status, errorText) };
      return;
    }

    yield* this.parseAnthropicSSEStream(response);
  }

  private async *parseAnthropicSSEStream(response: Response): AsyncGenerator<AIStreamEvent> {
    const reader = response.body?.getReader();
    if (!reader) {
      yield { type: "error", error: "无法读取响应流" };
      return;
    }

    const decoder = new TextDecoder();
    let buffer = "";
    let currentToolId = "";
    let currentToolName = "";
    let currentToolInput = "";

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop() || "";

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed.startsWith("data: ")) continue;
          const jsonStr = trimmed.slice(6);

          try {
            const data = JSON.parse(jsonStr);

            switch (data.type) {
              case "content_block_start": {
                if (data.content_block?.type === "thinking") {
                  // 推理/思考内容开始
                } else if (data.content_block?.type === "tool_use") {
                  currentToolId = data.content_block.id;
                  currentToolName = desanitizeToolName(data.content_block.name);
                  currentToolInput = "";
                  yield {
                    type: "tool_call_begin",
                    toolCallId: currentToolId,
                    toolName: currentToolName,
                  };
                }
                break;
              }
              case "content_block_delta": {
                if (data.delta?.type === "thinking_delta") {
                  yield { type: "reasoning_delta", delta: data.delta.thinking };
                } else if (data.delta?.type === "text_delta") {
                  yield { type: "text_delta", delta: data.delta.text };
                } else if (data.delta?.type === "input_json_delta") {
                  currentToolInput += data.delta.partial_json;
                  yield {
                    type: "tool_call_delta",
                    toolCallId: currentToolId,
                    delta: data.delta.partial_json,
                  };
                }
                break;
              }
              case "content_block_stop": {
                if (currentToolId) {
                  yield {
                    type: "tool_call_end",
                    toolCallId: currentToolId,
                    toolName: currentToolName,
                    arguments: currentToolInput,
                  };
                  currentToolId = "";
                  currentToolName = "";
                  currentToolInput = "";
                }
                break;
              }
              case "message_delta": {
                if (data.usage) {
                  yield {
                    type: "usage",
                    usage: {
                      inputTokens: data.usage.input_tokens || 0,
                      outputTokens: data.usage.output_tokens || 0,
                    },
                  };
                }
                if (data.delta?.stop_reason) {
                  yield { type: "done", finishReason: data.delta.stop_reason };
                }
                break;
              }
            }
          } catch {
            // 跳过
          }
        }
      }
    } finally {
      reader.releaseLock();
    }
  }
}
