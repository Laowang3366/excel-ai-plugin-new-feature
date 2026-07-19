import {
  classifyNetworkError,
  joinUrl,
  readErrorMessage,
  type ProviderFetch,
} from "./client";
import { isAbortError } from "../agent/streamProvider";
import type {
  AgentStreamEvent,
  AgentStreamProvider,
  StreamChatRequest,
} from "../agent/types";
import { encodeAnthropicMessagesBody } from "./anthropicMessagesEncode";
import { AnthropicMessagesStreamAssembler } from "./anthropicMessagesStreamParse";
import { SseByteParser, type SseParseResult } from "./openaiSse";
import { buildToolNameMaps, isToolNameMaps } from "./openaiToolNameMap";

function redactSecrets(message: string, apiKey: string): string {
  if (!apiKey) return message;
  return message.split(apiKey).join("[REDACTED]");
}

export interface AnthropicMessagesStreamProviderOptions {
  baseUrl: string;
  apiKey: string;
  model: string;
  /** Defaults to 4096. Must be an integer >= 1. */
  maxTokens?: number;
  fetchImpl?: ProviderFetch;
}

export class AnthropicMessagesStreamProvider implements AgentStreamProvider {
  private readonly baseUrl: string;
  private readonly apiKey: string;
  private readonly model: string;
  private readonly maxTokens: number;
  private readonly fetchImpl: ProviderFetch;

  constructor(options: AnthropicMessagesStreamProviderOptions) {
    this.baseUrl = options.baseUrl;
    this.apiKey = options.apiKey;
    this.model = options.model;
    this.maxTokens = options.maxTokens ?? 4096;
    this.fetchImpl = options.fetchImpl ?? fetch.bind(globalThis);
  }

  async *streamChat(request: StreamChatRequest): AsyncIterable<AgentStreamEvent> {
    if (request.signal?.aborted) {
      yield { type: "error", message: "aborted", kind: "aborted" };
      return;
    }
    if (!this.apiKey.trim()) {
      yield { type: "error", message: "API key 未设置，无法发起请求", kind: "missing_key" };
      return;
    }
    if (!this.baseUrl.trim()) {
      yield { type: "error", message: "Base URL 未设置", kind: "parse" };
      return;
    }
    if (!this.model.trim()) {
      yield { type: "error", message: "model 未设置", kind: "parse" };
      return;
    }
    if (
      typeof this.maxTokens !== "number" ||
      !Number.isInteger(this.maxTokens) ||
      this.maxTokens < 1
    ) {
      yield {
        type: "error",
        message: "maxTokens must be an integer >= 1",
        kind: "parse",
      };
      return;
    }

    const mapsResult = buildToolNameMaps(request.tools);
    if (!isToolNameMaps(mapsResult)) {
      yield { type: "error", message: mapsResult.error, kind: "parse" };
      return;
    }
    const encoded = encodeAnthropicMessagesBody(
      request.systemPrompt,
      request.messages,
      request.tools,
      mapsResult,
    );
    if ("error" in encoded) {
      yield { type: "error", message: encoded.error, kind: "parse" };
      return;
    }

    const url = joinUrl(this.baseUrl, "/messages");
    const apiKey = this.apiKey;
    const body: Record<string, unknown> = {
      model: this.model,
      max_tokens: this.maxTokens,
      stream: true,
      messages: encoded.messages,
    };
    if (encoded.system != null) body.system = encoded.system;
    if (encoded.tools.length > 0) body.tools = encoded.tools;

    let response: Response;
    try {
      response = await this.fetchImpl(url, {
        method: "POST",
        headers: {
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "Content-Type": "application/json",
          Accept: "text/event-stream",
        },
        body: JSON.stringify(body),
        signal: request.signal,
      });
    } catch (error) {
      if (isAbortError(error) || request.signal?.aborted) {
        yield { type: "error", message: "aborted", kind: "aborted" };
        return;
      }
      const classified = classifyNetworkError(error, url);
      yield {
        type: "error",
        message: redactSecrets(classified.error, apiKey),
        kind: classified.kind === "cors" ? "cors" : "network",
        url,
      };
      return;
    }

    if (!response.ok) {
      const message = await readErrorMessage(response);
      yield {
        type: "error",
        message: redactSecrets(message, apiKey),
        kind: "http",
        status: response.status,
        url,
      };
      return;
    }
    if (!response.body) {
      yield { type: "error", message: "response body is empty", kind: "parse", url };
      return;
    }

    const assembler = new AnthropicMessagesStreamAssembler(mapsResult);
    const sse = new SseByteParser();
    const reader = response.body.getReader();

    const emitFinalize = function* (): Generator<AgentStreamEvent> {
      const final = assembler.finalize();
      if ("error" in final) {
        yield { type: "error", message: final.error, kind: "parse", url };
        return;
      }
      for (const e of final) yield e;
    };

    const handleParts = function* (
      parts: SseParseResult[],
    ): Generator<AgentStreamEvent, boolean> {
      for (const part of parts) {
        if (part.kind === "done") {
          yield* emitFinalize();
          return true;
        }
        let json: unknown;
        try {
          json = JSON.parse(part.data);
        } catch {
          yield { type: "error", message: "malformed SSE JSON", kind: "parse", url };
          return true;
        }
        const ingested = assembler.ingest(json);
        if ("error" in ingested) {
          const provider = (ingested as { provider?: boolean }).provider === true;
          yield {
            type: "error",
            message: redactSecrets(ingested.error, apiKey),
            kind: provider ? "provider" : "parse",
            url,
          };
          return true;
        }
        for (const e of ingested) yield e;
      }
      return false;
    };

    try {
      while (true) {
        if (request.signal?.aborted) {
          yield { type: "error", message: "aborted", kind: "aborted" };
          return;
        }
        const { done, value } = await reader.read();
        if (done) {
          const flushed = sse.flush();
          const stop = yield* handleParts(flushed);
          if (!stop) yield* emitFinalize();
          return;
        }
        if (value) {
          const stop = yield* handleParts(sse.push(value));
          if (stop) return;
        }
      }
    } catch (error) {
      if (isAbortError(error) || request.signal?.aborted) {
        yield { type: "error", message: "aborted", kind: "aborted" };
        return;
      }
      const classified = classifyNetworkError(error, url);
      yield {
        type: "error",
        message: redactSecrets(classified.error, apiKey),
        kind: classified.kind === "cors" ? "cors" : "network",
        url,
      };
    }
  }
}
