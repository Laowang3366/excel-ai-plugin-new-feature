import {
  normalizeConnectionMode,
  resolveProviderEndpoint,
  type EndpointResolveErr,
  type EndpointKind,
} from "./endpointResolve";
import type { ApiFormat, ConnectionMode } from "./types";

export type ProviderClientErrorKind =
  | "missing_key"
  | "http"
  | "network"
  | "cors"
  | "parse"
  | "unsupported";

export interface ProviderEndpointConfig {
  baseUrl: string;
  apiKey: string;
  apiFormat: ApiFormat;
  model?: string;
  connectionMode?: ConnectionMode;
  gatewayBaseUrl?: string;
  gatewayUpstreamId?: string;
}

export interface ConnectionTestResult {
  latencyMs: number;
  url: string;
  status: number;
}

export interface ListModelsResult {
  models: string[];
  url: string;
}

export type ProviderClientOk<T> = { ok: true; data: T };
export type ProviderClientErr = {
  ok: false;
  error: string;
  kind: ProviderClientErrorKind;
  url?: string;
  status?: number;
};
export type ProviderClientResult<T> = ProviderClientOk<T> | ProviderClientErr;

export interface ProviderFetch {
  (input: string, init?: RequestInit): Promise<Response>;
}

export { joinUrl } from "./endpointResolve";

type TestRequest = {
  url: string;
  headers: Record<string, string>;
  body: Record<string, unknown>;
};

type ModelsRequest = { url: string; headers: Record<string, string> };
type ResolvedRequest<T> = T & { redactionSecret: string };
type BuildResult<T> = { ok: true; data: T } | EndpointResolveErr;

function missingKey(): ProviderClientErr {
  return { ok: false, kind: "missing_key", error: "API key 未设置，无法发起请求" };
}

function redactSecret(message: string, secret: string): string {
  if (!secret) return message;
  return message.split(secret).join("[REDACTED]");
}

export function classifyNetworkError(error: unknown, url: string): ProviderClientErr {
  const message = error instanceof Error ? error.message : String(error);
  const lower = message.toLowerCase();
  const looksCors =
    lower.includes("failed to fetch") ||
    lower.includes("networkerror") ||
    lower.includes("cors") ||
    lower.includes("access-control-allow-origin");
  return {
    ok: false,
    kind: looksCors ? "cors" : "network",
    error: looksCors
      ? `浏览器 CORS/网络拦截：${message}（Office 任务窗格直连第三方 API 常被 CORS 拒绝）`
      : `网络错误：${message}`,
    url,
  };
}

export async function readErrorMessage(response: Response): Promise<string> {
  const text = await response.text().catch(() => "");
  if (!text) return `HTTP ${response.status}`;
  try {
    const json = JSON.parse(text) as { error?: { message?: string }; message?: string };
    return json.error?.message || json.message || `HTTP ${response.status}`;
  } catch {
    return text.slice(0, 200) || `HTTP ${response.status}`;
  }
}

function testEndpointKind(apiFormat: ApiFormat): EndpointKind {
  if (apiFormat === "anthropic") return "messages";
  if (apiFormat === "responses") return "responses";
  return "chat";
}

function resolveTestConnectionRequest(
  config: ProviderEndpointConfig,
): BuildResult<ResolvedRequest<TestRequest>> {
  const model =
    config.model || (config.apiFormat === "anthropic" ? "claude-sonnet-4-6" : "gpt-4o");
  const resolved = resolveProviderEndpoint({
    connectionMode: config.connectionMode,
    baseUrl: config.baseUrl,
    gatewayBaseUrl: config.gatewayBaseUrl,
    apiKey: config.apiKey,
    apiFormat: config.apiFormat,
    gatewayUpstreamId: config.gatewayUpstreamId,
    kind: testEndpointKind(config.apiFormat),
    acceptEventStream: false,
  });
  if (!resolved.ok) return resolved;

  let body: Record<string, unknown>;
  if (config.apiFormat === "anthropic") {
    body = {
      model,
      max_tokens: 1,
      messages: [{ role: "user", content: "Hi" }],
    };
  } else if (config.apiFormat === "responses") {
    body = { model, input: "Hi", max_output_tokens: 1 };
  } else {
    body = {
      model,
      max_tokens: 1,
      messages: [{ role: "user", content: "Hi" }],
    };
  }
  return {
    ok: true,
    data: {
      url: resolved.data.url,
      headers: resolved.data.headers,
      body,
      redactionSecret: resolved.data.redactionSecret,
    },
  };
}

/** Build request contract for connection test (aligned with desktop ai:testConnection). */
export function buildTestConnectionRequest(config: ProviderEndpointConfig): TestRequest {
  const result = resolveTestConnectionRequest(config);
  if (!result.ok) throw new Error(result.error);
  return {
    url: result.data.url,
    headers: result.data.headers,
    body: result.data.body,
  };
}

function resolveListModelsRequest(
  config: ProviderEndpointConfig,
): BuildResult<ResolvedRequest<ModelsRequest> | null> {
  let mode: ConnectionMode;
  try {
    mode = normalizeConnectionMode(config.connectionMode);
  } catch (error) {
    return {
      ok: false,
      kind: "parse",
      error: error instanceof Error ? error.message : String(error),
    };
  }
  if (mode === "direct" && config.apiFormat === "anthropic") {
    return { ok: true, data: null };
  }
  const resolved = resolveProviderEndpoint({
    connectionMode: mode,
    baseUrl: config.baseUrl,
    gatewayBaseUrl: config.gatewayBaseUrl,
    apiKey: config.apiKey,
    apiFormat: config.apiFormat,
    gatewayUpstreamId: config.gatewayUpstreamId,
    kind: "models",
    acceptEventStream: false,
  });
  if (!resolved.ok) return resolved;
  return {
    ok: true,
    data: {
      url: resolved.data.url,
      headers: resolved.data.headers,
      redactionSecret: resolved.data.redactionSecret,
    },
  };
}

/** Build request contract for model list; direct Anthropic remains unsupported. */
export function buildListModelsRequest(
  config: ProviderEndpointConfig,
): ModelsRequest | null {
  const result = resolveListModelsRequest(config);
  if (!result.ok) throw new Error(result.error);
  if (!result.data) return null;
  return { url: result.data.url, headers: result.data.headers };
}

function modeOrError(
  value: ProviderEndpointConfig["connectionMode"],
): { ok: true; mode: ConnectionMode } | ProviderClientErr {
  try {
    return { ok: true, mode: normalizeConnectionMode(value) };
  } catch (error) {
    return {
      ok: false,
      kind: "parse",
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

export class ProviderClient {
  constructor(private readonly fetchImpl: ProviderFetch = fetch.bind(globalThis)) {}

  async testConnection(
    config: ProviderEndpointConfig,
    options?: { signal?: AbortSignal },
  ): Promise<ProviderClientResult<ConnectionTestResult>> {
    const parsedMode = modeOrError(config.connectionMode);
    if (!parsedMode.ok) return parsedMode;
    if (parsedMode.mode === "direct" && !config.apiKey.trim()) return missingKey();
    if (parsedMode.mode === "direct" && !config.baseUrl.trim()) {
      return { ok: false, kind: "http", error: "Base URL 未设置" };
    }

    const request = resolveTestConnectionRequest(config);
    if (!request.ok) return request;
    const started = Date.now();
    try {
      const response = await this.fetchImpl(request.data.url, {
        method: "POST",
        headers: request.data.headers,
        body: JSON.stringify(request.data.body),
        signal: options?.signal,
      });
      const latencyMs = Date.now() - started;
      if (!response.ok) {
        return {
          ok: false,
          kind: "http",
          error: redactSecret(
            await readErrorMessage(response),
            request.data.redactionSecret,
          ),
          url: request.data.url,
          status: response.status,
        };
      }
      return {
        ok: true,
        data: { latencyMs, url: request.data.url, status: response.status },
      };
    } catch (error) {
      if (options?.signal?.aborted) {
        return { ok: false, kind: "network", error: "aborted", url: request.data.url };
      }
      const classified = classifyNetworkError(error, request.data.url);
      return {
        ...classified,
        error: redactSecret(classified.error, request.data.redactionSecret),
      };
    }
  }

  async listModels(
    config: ProviderEndpointConfig,
    options?: { signal?: AbortSignal },
  ): Promise<ProviderClientResult<ListModelsResult>> {
    const parsedMode = modeOrError(config.connectionMode);
    if (!parsedMode.ok) return parsedMode;
    if (parsedMode.mode === "direct" && !config.apiKey.trim()) return missingKey();
    if (parsedMode.mode === "direct" && !config.baseUrl.trim()) {
      return { ok: false, kind: "http", error: "Base URL 未设置" };
    }
    if (parsedMode.mode === "direct" && config.apiFormat === "anthropic") {
      return {
        ok: false,
        kind: "unsupported",
        error: "Anthropic Messages 格式不提供 /models 列表接口（与桌面端一致）",
      };
    }

    const request = resolveListModelsRequest(config);
    if (!request.ok) return request;
    if (!request.data) {
      return {
        ok: false,
        kind: "unsupported",
        error: "Anthropic Messages 格式不提供 /models 列表接口（与桌面端一致）",
      };
    }

    try {
      const response = await this.fetchImpl(request.data.url, {
        method: "GET",
        headers: request.data.headers,
        signal: options?.signal,
      });
      if (!response.ok) {
        return {
          ok: false,
          kind: "http",
          error: redactSecret(
            await readErrorMessage(response),
            request.data.redactionSecret,
          ),
          url: request.data.url,
          status: response.status,
        };
      }
      const data = (await response.json()) as unknown;
      return {
        ok: true,
        data: { models: parseModelIds(data), url: request.data.url },
      };
    } catch (error) {
      if (options?.signal?.aborted) {
        return { ok: false, kind: "network", error: "aborted", url: request.data.url };
      }
      if (error instanceof SyntaxError) {
        return {
          ok: false,
          kind: "parse",
          error: `模型列表 JSON 解析失败：${error.message}`,
          url: request.data.url,
        };
      }
      const classified = classifyNetworkError(error, request.data.url);
      return {
        ...classified,
        error: redactSecret(classified.error, request.data.redactionSecret),
      };
    }
  }
}

function parseModelIds(data: unknown): string[] {
  if (Array.isArray(data)) {
    return data
      .map((item) =>
        typeof item === "string"
          ? item
          : (item as { id?: string; name?: string })?.id ||
            (item as { name?: string })?.name ||
            "",
      )
      .filter((id) => id.length > 0)
      .sort();
  }
  if (data && typeof data === "object" && Array.isArray((data as { data?: unknown }).data)) {
    return ((data as { data: unknown[] }).data
      .map((item) => {
        if (!item || typeof item !== "object") return "";
        const row = item as { id?: string; name?: string };
        return row.id || row.name || "";
      })
      .filter((id) => id.length > 0) as string[]).sort();
  }
  if (data && typeof data === "object" && Array.isArray((data as { models?: unknown }).models)) {
    return ((data as { models: unknown[] }).models
      .map((item) => {
        if (typeof item === "string") return item;
        if (!item || typeof item !== "object") return "";
        return (item as { id?: string; name?: string }).id ||
          (item as { name?: string }).name ||
          "";
      })
      .filter((id) => id.length > 0) as string[]).sort();
  }
  return [];
}
