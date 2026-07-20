import type { AgentStreamProvider } from "../agent/types";
import type { ProviderFetch } from "./client";
import type { ApiFormat, ConnectionMode, ReasoningMode } from "./types";
import type { ProviderStore } from "./store";
import {
  normalizeConnectionMode,
  resolveProviderEndpoint,
  streamKindForApiFormat,
} from "./endpointResolve";
import { OpenAIChatCompletionsStreamProvider } from "./openaiChatCompletionsProvider";
import { OpenAIResponsesStreamProvider } from "./openaiResponsesProvider";
import { AnthropicMessagesStreamProvider } from "./anthropicMessagesProvider";

export type CreateStreamProviderInput = {
  apiFormat: ApiFormat | string;
  baseUrl: string;
  apiKey: string;
  model: string;
  /** Vendor id (openai/deepseek/...); used for chat-completions reasoning mapping. */
  provider?: string;
  connectionMode?: ConnectionMode | string;
  gatewayBaseUrl?: string;
  gatewayUpstreamId?: string;
  reasoningMode?: ReasoningMode;
  /** Anthropic only; defaults to provider 4096 when omitted. */
  maxTokens?: number;
  fetchImpl?: ProviderFetch;
};

export type CreateStreamProviderOk = {
  ok: true;
  provider: AgentStreamProvider;
};

export type CreateStreamProviderErr = {
  ok: false;
  kind: "parse" | "missing_key";
  error: string;
};

export type CreateStreamProviderResult =
  CreateStreamProviderOk | CreateStreamProviderErr;

const SUPPORTED_FORMATS = new Set<ApiFormat>([
  "openai",
  "responses",
  "anthropic",
]);

function isApiFormat(value: string): value is ApiFormat {
  return SUPPORTED_FORMATS.has(value as ApiFormat);
}

/**
 * Route transport by apiFormat; chat-completions reasoning fields also
 * use the vendor `provider` id (desktop providerClients semantics).
 * Does not throw; returns a discriminant result for fetch-prep validation.
 */
export function createStreamProvider(
  input: CreateStreamProviderInput,
): CreateStreamProviderResult {
  let connectionMode: ConnectionMode;
  try {
    connectionMode = normalizeConnectionMode(input.connectionMode);
  } catch (error) {
    return {
      ok: false,
      kind: "parse",
      error: error instanceof Error ? error.message : String(error),
    };
  }
  const apiKey = typeof input.apiKey === "string" ? input.apiKey.trim() : "";
  const baseUrl = typeof input.baseUrl === "string" ? input.baseUrl.trim() : "";
  const gatewayBaseUrl =
    typeof input.gatewayBaseUrl === "string" ? input.gatewayBaseUrl.trim() : "";
  const model = typeof input.model === "string" ? input.model.trim() : "";
  const gatewayUpstreamId =
    typeof input.gatewayUpstreamId === "string"
      ? input.gatewayUpstreamId.trim()
      : "";
  const apiFormatRaw =
    typeof input.apiFormat === "string" ? input.apiFormat.trim() : "";

  if (connectionMode === "direct" && !apiKey) {
    return {
      ok: false,
      kind: "missing_key",
      error: "API key 未设置，无法发起请求",
    };
  }
  if (connectionMode === "direct" && !baseUrl) {
    return { ok: false, kind: "parse", error: "Base URL 未设置" };
  }
  if (!model) {
    return { ok: false, kind: "parse", error: "model 未设置" };
  }
  if (!apiFormatRaw || !isApiFormat(apiFormatRaw)) {
    return {
      ok: false,
      kind: "parse",
      error: `unknown apiFormat: ${apiFormatRaw || "(empty)"}`,
    };
  }

  const endpoint = resolveProviderEndpoint({
    connectionMode,
    baseUrl,
    gatewayBaseUrl,
    apiKey,
    apiFormat: apiFormatRaw,
    gatewayUpstreamId,
    kind: streamKindForApiFormat(apiFormatRaw),
    acceptEventStream: true,
  });
  if (!endpoint.ok) {
    return { ok: false, kind: endpoint.kind, error: endpoint.error };
  }

  const provider =
    typeof input.provider === "string" ? input.provider.trim() : "";

  const common = {
    baseUrl,
    apiKey: connectionMode === "gateway" ? "" : apiKey,
    model,
    provider,
    connectionMode,
    gatewayBaseUrl,
    gatewayUpstreamId,
    reasoningMode: input.reasoningMode,
    fetchImpl: input.fetchImpl,
  };

  switch (apiFormatRaw) {
    case "openai":
      return {
        ok: true,
        provider: new OpenAIChatCompletionsStreamProvider(common),
      };
    case "responses":
      return {
        ok: true,
        provider: new OpenAIResponsesStreamProvider(common),
      };
    case "anthropic":
      return {
        ok: true,
        provider: new AnthropicMessagesStreamProvider({
          ...common,
          ...(input.maxTokens != null ? { maxTokens: input.maxTokens } : {}),
        }),
      };
    default:
      return {
        ok: false,
        kind: "parse",
        error: `unknown apiFormat: ${apiFormatRaw}`,
      };
  }
}

/**
 * Thin wrapper: re-read active config + secret on every call and construct a
 * fresh provider. Never caches a provider singleton.
 */
export function createStreamProviderFromStore(
  store: ProviderStore,
  options?: { fetchImpl?: ProviderFetch; maxTokens?: number },
): CreateStreamProviderResult {
  const active = store.getActive();
  if (!active) {
    return {
      ok: false,
      kind: "parse",
      error: "no active provider configured",
    };
  }
  return createStreamProvider({
    apiFormat: active.apiFormat,
    baseUrl: active.baseUrl,
    apiKey: active.apiKey,
    model: active.model,
    provider: active.provider,
    connectionMode: active.connectionMode,
    gatewayBaseUrl: active.gatewayBaseUrl,
    gatewayUpstreamId: active.gatewayUpstreamId,
    reasoningMode: active.reasoningMode,
    fetchImpl: options?.fetchImpl,
    maxTokens: options?.maxTokens,
  });
}
