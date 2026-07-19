import type { ApiFormat } from "./types";

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

function trimBase(baseUrl: string): string {
  return baseUrl.replace(/\/+$/, "");
}

export function joinUrl(baseUrl: string, path: string): string {
  const base = trimBase(baseUrl);
  if (base.endsWith(path)) return base;
  return `${base}${path.startsWith("/") ? path : `/${path}`}`;
}

function missingKey(): ProviderClientErr {
  return { ok: false, kind: "missing_key", error: "API key 未设置，无法发起请求" };
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

/** Build request contract for connection test (aligned with desktop ai:testConnection). */
export function buildTestConnectionRequest(config: ProviderEndpointConfig): {
  url: string;
  headers: Record<string, string>;
  body: Record<string, unknown>;
} {
  const model = config.model || (config.apiFormat === "anthropic" ? "claude-sonnet-4-6" : "gpt-4o");
  if (config.apiFormat === "anthropic") {
    return {
      url: joinUrl(config.baseUrl, "/messages"),
      headers: {
        "x-api-key": config.apiKey,
        "anthropic-version": "2023-06-01",
        "Content-Type": "application/json",
      },
      body: {
        model,
        max_tokens: 1,
        messages: [{ role: "user", content: "Hi" }],
      },
    };
  }
  if (config.apiFormat === "responses") {
    return {
      url: joinUrl(config.baseUrl, "/responses"),
      headers: {
        Authorization: `Bearer ${config.apiKey}`,
        "Content-Type": "application/json",
      },
      body: { model, input: "Hi", max_output_tokens: 1 },
    };
  }
  return {
    url: joinUrl(config.baseUrl, "/chat/completions"),
    headers: {
      Authorization: `Bearer ${config.apiKey}`,
      "Content-Type": "application/json",
    },
    body: {
      model,
      max_tokens: 1,
      messages: [{ role: "user", content: "Hi" }],
    },
  };
}

/** Build request contract for model list (OpenAI-compatible /models). */
export function buildListModelsRequest(config: ProviderEndpointConfig): {
  url: string;
  headers: Record<string, string>;
} | null {
  if (config.apiFormat === "anthropic") return null;
  const base = trimBase(config.baseUrl);
  const url = base.endsWith("/models") ? base : `${base}/models`;
  return {
    url,
    headers: {
      Authorization: `Bearer ${config.apiKey}`,
      "Content-Type": "application/json",
    },
  };
}

export class ProviderClient {
  constructor(private readonly fetchImpl: ProviderFetch = fetch.bind(globalThis)) {}

  async testConnection(
    config: ProviderEndpointConfig,
  ): Promise<ProviderClientResult<ConnectionTestResult>> {
    if (!config.apiKey.trim()) return missingKey();
    if (!config.baseUrl.trim()) {
      return { ok: false, kind: "http", error: "Base URL 未设置" };
    }
    const request = buildTestConnectionRequest(config);
    const started = Date.now();
    try {
      const response = await this.fetchImpl(request.url, {
        method: "POST",
        headers: request.headers,
        body: JSON.stringify(request.body),
      });
      const latencyMs = Date.now() - started;
      if (!response.ok) {
        return {
          ok: false,
          kind: "http",
          error: await readErrorMessage(response),
          url: request.url,
          status: response.status,
        };
      }
      return {
        ok: true,
        data: { latencyMs, url: request.url, status: response.status },
      };
    } catch (error) {
      return classifyNetworkError(error, request.url);
    }
  }

  async listModels(
    config: ProviderEndpointConfig,
  ): Promise<ProviderClientResult<ListModelsResult>> {
    if (!config.apiKey.trim()) return missingKey();
    if (!config.baseUrl.trim()) {
      return { ok: false, kind: "http", error: "Base URL 未设置" };
    }
    const request = buildListModelsRequest(config);
    if (!request) {
      return {
        ok: false,
        kind: "unsupported",
        error: "Anthropic Messages 格式不提供 /models 列表接口（与桌面端一致）",
      };
    }
    try {
      const response = await this.fetchImpl(request.url, {
        method: "GET",
        headers: request.headers,
      });
      if (!response.ok) {
        return {
          ok: false,
          kind: "http",
          error: await readErrorMessage(response),
          url: request.url,
          status: response.status,
        };
      }
      const data = (await response.json()) as unknown;
      const models = parseModelIds(data);
      return { ok: true, data: { models, url: request.url } };
    } catch (error) {
      if (error instanceof SyntaxError) {
        return { ok: false, kind: "parse", error: `模型列表 JSON 解析失败：${error.message}` };
      }
      return classifyNetworkError(error, request.url);
    }
  }
}

function parseModelIds(data: unknown): string[] {
  if (Array.isArray(data)) {
    return data
      .map((item) => (typeof item === "string" ? item : (item as { id?: string; name?: string })?.id || (item as { name?: string })?.name || ""))
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
  return [];
}
