/** Central URL/header resolution for direct APIs and the same-origin AI gateway. */

import type { ApiFormat, ConnectionMode } from "./types";

export type EndpointKind = "models" | "chat" | "responses" | "messages";

export type ResolvedEndpoint = {
  url: string;
  headers: Record<string, string>;
  /** Used only to redact direct-provider errors; never transmitted in gateway mode. */
  redactionSecret: string;
};

export type EndpointResolveOk = { ok: true; data: ResolvedEndpoint };
export type EndpointResolveErr = {
  ok: false;
  kind: "missing_key" | "parse";
  error: string;
};
export type EndpointResolveResult = EndpointResolveOk | EndpointResolveErr;

export type ResolveEndpointInput = {
  connectionMode?: ConnectionMode | string | null;
  /** Direct-provider API root. Ignored in gateway mode. */
  baseUrl: string;
  /** Gateway origin. Empty means the current page origin. */
  gatewayBaseUrl?: string | null;
  apiKey?: string | null;
  apiFormat: ApiFormat | string;
  kind: EndpointKind;
  gatewayUpstreamId?: string | null;
  extraHeaders?: Record<string, string | undefined>;
  acceptEventStream?: boolean;
};

const UPSTREAM_ID_RE = /^[a-z][a-z0-9_-]{0,63}$/;

function trimString(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

/** Undefined remains backward-compatible direct mode; unknown values fail closed. */
export function normalizeConnectionMode(value: unknown): ConnectionMode {
  if (value == null || value === "" || value === "direct") return "direct";
  if (value === "gateway") return "gateway";
  throw new Error(`unknown connectionMode: ${String(value)}`);
}

export function assertSafeGatewayUpstreamId(
  raw: unknown,
): { ok: true; id: string } | { ok: false; error: string } {
  const id = trimString(raw);
  if (!id) return { ok: false, error: "gatewayUpstreamId 未设置" };
  if (!UPSTREAM_ID_RE.test(id)) {
    return {
      ok: false,
      error: "gatewayUpstreamId 非法：必须以小写字母开头，且仅包含小写字母、数字、下划线或短横线（最多 64 字符）",
    };
  }
  return { ok: true, id };
}

function validateAbsoluteUrl(
  input: string,
): { ok: true; url: URL } | { ok: false; error: string } {
  let url: URL;
  try {
    url = new URL(input);
  } catch {
    return { ok: false, error: "Base URL 非法" };
  }
  if (url.protocol !== "http:" && url.protocol !== "https:") {
    return { ok: false, error: "Base URL 仅支持 http/https" };
  }
  if (url.username || url.password) {
    return { ok: false, error: "Base URL 不得包含用户名或密码" };
  }
  if (url.search) return { ok: false, error: "Base URL 不得包含 query" };
  if (url.hash) return { ok: false, error: "Base URL 不得包含 hash" };
  return { ok: true, url };
}

/**
 * Direct mode accepts an absolute API root with a path. Gateway mode accepts
 * only an origin (or empty/current-origin), because its /api/ai/v1 path is fixed.
 */
export function assertSafeBaseUrl(
  raw: unknown,
  mode: ConnectionMode,
): { ok: true; baseUrl: string } | { ok: false; error: string } {
  const input = trimString(raw);
  if (mode === "gateway") {
    if (!input || input === "/") return { ok: true, baseUrl: "" };
    if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(input)) {
      return {
        ok: false,
        error: "gatewayBaseUrl 必须留空（同源）或填写仅含 origin 的绝对 http(s) 地址",
      };
    }
    const parsed = validateAbsoluteUrl(input);
    if (!parsed.ok) return parsed;
    if (parsed.url.pathname !== "/") {
      return { ok: false, error: "gatewayBaseUrl 不得包含路径" };
    }
    return { ok: true, baseUrl: parsed.url.origin };
  }

  if (!input) return { ok: false, error: "Base URL 未设置" };
  if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(input)) {
    return { ok: false, error: "Base URL 必须是绝对 http(s) 地址" };
  }
  const parsed = validateAbsoluteUrl(input);
  if (!parsed.ok) return parsed;
  return { ok: true, baseUrl: parsed.url.toString().replace(/\/+$/, "") };
}

export function trimBase(baseUrl: string): string {
  return baseUrl.replace(/\/+$/, "");
}

export function joinUrl(baseUrl: string, path: string): string {
  const base = trimBase(baseUrl);
  const suffix = path.startsWith("/") ? path : `/${path}`;
  if (!base) return suffix;
  if (base.endsWith(suffix)) return base;
  return `${base}${suffix}`;
}

function endpointPath(kind: EndpointKind, format: ApiFormat): string | null {
  if (kind === "models") return "/models";
  if (kind === "chat" && format === "openai") return "/chat/completions";
  if (kind === "responses" && format === "responses") return "/responses";
  if (kind === "messages" && format === "anthropic") return "/messages";
  return null;
}

export function streamKindForApiFormat(format: ApiFormat): EndpointKind {
  if (format === "anthropic") return "messages";
  if (format === "responses") return "responses";
  return "chat";
}

function applyOptionalHeaders(
  headers: Record<string, string>,
  format: ApiFormat,
  extra?: Record<string, string | undefined>,
): void {
  if (!extra) return;
  const allowed =
    format === "anthropic"
      ? (["anthropic-beta"] as const)
      : (["openai-organization", "openai-project"] as const);
  for (const key of allowed) {
    const value = extra[key];
    if (value && !/[\r\n]/.test(value)) headers[key] = value;
  }
}

function buildHeaders(
  mode: ConnectionMode,
  format: ApiFormat,
  apiKey: string,
  acceptEventStream: boolean,
  extra?: Record<string, string | undefined>,
): Record<string, string> {
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    Accept: acceptEventStream ? "text/event-stream" : "application/json",
  };
  if (format === "anthropic") headers["anthropic-version"] = "2023-06-01";
  applyOptionalHeaders(headers, format, extra);
  if (mode === "direct") {
    if (format === "anthropic") headers["x-api-key"] = apiKey;
    else headers.Authorization = `Bearer ${apiKey}`;
  }
  return headers;
}

function isApiFormat(value: string): value is ApiFormat {
  return value === "openai" || value === "anthropic" || value === "responses";
}

export function resolveProviderEndpoint(
  input: ResolveEndpointInput,
): EndpointResolveResult {
  let mode: ConnectionMode;
  try {
    mode = normalizeConnectionMode(input.connectionMode);
  } catch (error) {
    return {
      ok: false,
      kind: "parse",
      error: error instanceof Error ? error.message : String(error),
    };
  }

  const formatRaw = trimString(input.apiFormat);
  if (!isApiFormat(formatRaw)) {
    return {
      ok: false,
      kind: "parse",
      error: `unknown apiFormat: ${formatRaw || "(empty)"}`,
    };
  }
  const path = endpointPath(input.kind, formatRaw);
  if (!path) {
    return {
      ok: false,
      kind: "parse",
      error: `apiFormat ${formatRaw} does not support endpoint kind ${input.kind}`,
    };
  }

  const base = assertSafeBaseUrl(
    mode === "gateway" ? input.gatewayBaseUrl : input.baseUrl,
    mode,
  );
  if (!base.ok) return { ok: false, kind: "parse", error: base.error };

  const apiKey = trimString(input.apiKey);
  if (mode === "direct" && !apiKey) {
    return {
      ok: false,
      kind: "missing_key",
      error: "API key 未设置，无法发起请求",
    };
  }

  let url: string;
  if (mode === "gateway") {
    const upstream = assertSafeGatewayUpstreamId(input.gatewayUpstreamId);
    if (!upstream.ok) return { ok: false, kind: "parse", error: upstream.error };
    url = joinUrl(
      base.baseUrl,
      `/api/ai/v1/${upstream.id}/${path.replace(/^\//, "")}`,
    );
  } else if (input.kind === "models") {
    const directBase = trimBase(base.baseUrl);
    url = directBase.endsWith("/models") ? directBase : `${directBase}/models`;
  } else {
    url = joinUrl(base.baseUrl, path);
  }

  return {
    ok: true,
    data: {
      url,
      headers: buildHeaders(
        mode,
        formatRaw,
        apiKey,
        input.acceptEventStream === true,
        input.extraHeaders,
      ),
      redactionSecret: mode === "direct" ? apiKey : "",
    },
  };
}
