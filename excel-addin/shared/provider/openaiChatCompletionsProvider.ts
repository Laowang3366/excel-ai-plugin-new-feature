import {
  classifyNetworkError,
  joinUrl,
  readErrorMessage,
  type ProviderFetch,
} from "./client";
import { isAbortError, throwIfAborted } from "../agent/streamProvider";
import type {
  AgentStreamEvent,
  AgentStreamProvider,
  StreamChatRequest,
} from "../agent/types";
import { encodeChatCompletionsBody } from "./openaiChatEncode";
import { OpenAiChatStreamAssembler } from "./openaiChatStreamParse";
import { SseByteParser } from "./openaiSse";
import { buildToolNameMaps, isToolNameMaps } from "./openaiToolNameMap";

export interface OpenAIChatCompletionsStreamProviderOptions {
  baseUrl: string;
  apiKey: string;
  model: string;
  fetchImpl?: ProviderFetch;
}

export class OpenAIChatCompletionsStreamProvider implements AgentStreamProvider {
  private readonly baseUrl: string;
  private readonly apiKey: string;
  private readonly model: string;
  private readonly fetchImpl: ProviderFetch;

  constructor(options: OpenAIChatCompletionsStreamProviderOptions) {
    this.baseUrl = options.baseUrl;
    this.apiKey = options.apiKey;
    this.model = options.model;
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

    const mapsResult = buildToolNameMaps(request.tools);
    if (!isToolNameMaps(mapsResult)) {
      yield { type: "error", message: mapsResult.error, kind: "parse" };
      return;
    }
    const encoded = encodeChatCompletionsBody(
      request.systemPrompt,
      request.messages,
      request.tools,
      mapsResult,
    );
    if ("error" in encoded) {
      yield { type: "error", message: encoded.error, kind: "parse" };
      return;
    }

    const url = joinUrl(this.baseUrl, "/chat/completions");
    const body = {
      model: this.model,
      stream: true,
      messages: encoded.messages,
      ...(encoded.tools.length > 0 ? { tools: encoded.tools } : {}),
      stream_options: { include_usage: true },
    };

    let response: Response;
    try {
      response = await this.fetchImpl(url, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${this.apiKey}`,
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
        message: classified.error,
        kind: classified.kind === "cors" ? "cors" : "network",
        url,
      };
      return;
    }

    if (!response.ok) {
      const message = await readErrorMessage(response);
      yield { type: "error", message, kind: "http", status: response.status, url };
      return;
    }
    if (!response.body) {
      yield { type: "error", message: "response body is empty", kind: "parse", url };
      return;
    }

    const assembler = new OpenAiChatStreamAssembler(mapsResult);
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

    try {
      while (true) {
        throwIfAborted(request.signal);
        const { done, value } = await reader.read();
        if (done) break;
        if (!value) continue;
        for (const part of sse.push(value)) {
          if (part.kind === "done") {
            yield* emitFinalize();
            return;
          }
          let json: unknown;
          try {
            json = JSON.parse(part.data);
          } catch {
            yield { type: "error", message: "malformed SSE JSON", kind: "parse", url };
            return;
          }
          const ingested = assembler.ingest(json);
          if ("error" in ingested) {
            yield { type: "error", message: ingested.error, kind: "parse", url };
            return;
          }
          for (const e of ingested) yield e;
        }
      }
      yield* emitFinalize();
    } catch (error) {
      if (isAbortError(error) || request.signal?.aborted) {
        yield { type: "error", message: "aborted", kind: "aborted" };
        return;
      }
      const classified = classifyNetworkError(error, url);
      yield {
        type: "error",
        message: classified.error,
        kind: classified.kind === "cors" ? "cors" : "network",
        url,
      };
    } finally {
      try {
        reader.releaseLock();
      } catch {
        /* ignore */
      }
    }
  }
}
