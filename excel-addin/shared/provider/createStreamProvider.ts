import type { AgentStreamProvider } from "../agent/types";
import type { ProviderFetch } from "./client";
import type { ApiFormat } from "./types";
import type { ProviderStore } from "./store";
import { OpenAIChatCompletionsStreamProvider } from "./openaiChatCompletionsProvider";
import { OpenAIResponsesStreamProvider } from "./openaiResponsesProvider";
import { AnthropicMessagesStreamProvider } from "./anthropicMessagesProvider";

export type CreateStreamProviderInput = {
  apiFormat: ApiFormat | string;
  baseUrl: string;
  apiKey: string;
  model: string;
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
  | CreateStreamProviderOk
  | CreateStreamProviderErr;

const SUPPORTED_FORMATS = new Set<ApiFormat>(["openai", "responses", "anthropic"]);

function isApiFormat(value: string): value is ApiFormat {
  return SUPPORTED_FORMATS.has(value as ApiFormat);
}

/**
 * Route by apiFormat only — never by vendor/provider display name.
 * Does not throw; returns a discriminant result for fetch-prep validation.
 */
export function createStreamProvider(
  input: CreateStreamProviderInput,
): CreateStreamProviderResult {
  const apiKey = typeof input.apiKey === "string" ? input.apiKey.trim() : "";
  const baseUrl = typeof input.baseUrl === "string" ? input.baseUrl.trim() : "";
  const model = typeof input.model === "string" ? input.model.trim() : "";
  const apiFormatRaw =
    typeof input.apiFormat === "string" ? input.apiFormat.trim() : "";

  if (!apiKey) {
    return {
      ok: false,
      kind: "missing_key",
      error: "API key 未设置，无法发起请求",
    };
  }
  if (!baseUrl) {
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

  const common = {
    baseUrl,
    apiKey,
    model,
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
      // Exhaustiveness: isApiFormat already narrowed.
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
    fetchImpl: options?.fetchImpl,
    maxTokens: options?.maxTokens,
  });
}
