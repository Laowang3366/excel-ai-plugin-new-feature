/**
 * Fixed-path upstream proxy: byte-transparent stream with header allowlists.
 * Timers: connect covers pre-headers only; total spans the whole transfer and
 * is owned by the caller (not cleared when headers arrive).
 */

const FORWARDED_REQUEST_HEADERS = new Set([
  "content-type",
  "accept",
  "anthropic-version",
  "anthropic-beta",
  "openai-organization",
  "openai-project",
]);

const FORWARDED_RESPONSE_HEADERS = new Set([
  "content-type",
  "cache-control",
  "x-request-id",
  "openai-organization",
  "openai-processing-ms",
  "openai-version",
  "anthropic-ratelimit-requests-limit",
  "anthropic-ratelimit-requests-remaining",
  "anthropic-ratelimit-requests-reset",
  "anthropic-ratelimit-tokens-limit",
  "anthropic-ratelimit-tokens-remaining",
  "anthropic-ratelimit-tokens-reset",
  "retry-after",
]);

const ENDPOINT_SUFFIXES = new Set([
  "models",
  "chat/completions",
  "responses",
  "messages",
]);

/**
 * @param {string} baseUrl
 * @param {string} suffix
 */
export function joinUpstreamUrl(baseUrl, suffix) {
  if (!ENDPOINT_SUFFIXES.has(suffix)) {
    throw new Error("unsupported endpoint");
  }
  return `${baseUrl}/${suffix}`;
}

/**
 * @param {import('node:http').IncomingHttpHeaders} headers
 * @returns {Record<string, string>}
 */
export function filterRequestHeaders(headers) {
  /** @type {Record<string, string>} */
  const out = {};
  for (const [rawName, rawValue] of Object.entries(headers)) {
    if (rawValue === undefined) continue;
    const name = rawName.toLowerCase();
    if (!FORWARDED_REQUEST_HEADERS.has(name)) continue;
    const value = Array.isArray(rawValue) ? rawValue.join(", ") : String(rawValue);
    if (/[\r\n]/.test(value)) continue;
    out[name] = value;
  }
  return out;
}

/**
 * @param {Headers} headers
 * @returns {Record<string, string>}
 */
export function filterResponseHeaders(headers) {
  /** @type {Record<string, string>} */
  const out = {};
  for (const [name, value] of headers.entries()) {
    const lower = name.toLowerCase();
    if (!FORWARDED_RESPONSE_HEADERS.has(lower)) continue;
    if (/[\r\n]/.test(value)) continue;
    out[lower] = value;
  }
  return out;
}

/**
 * @param {{
 *   upstream: { baseUrl: string, authHeaderName: string, authHeaderValue: string },
 *   suffix: string,
 *   method: 'GET' | 'POST',
 *   headers: import('node:http').IncomingHttpHeaders,
 *   body: Buffer | null,
 *   connectTimeoutMs: number,
 *   signal: AbortSignal,
 *   fetchImpl?: typeof fetch,
 * }} opts
 */
export async function proxyToUpstream(opts) {
  const fetchImpl = opts.fetchImpl || globalThis.fetch;
  const url = joinUpstreamUrl(opts.upstream.baseUrl, opts.suffix);
  const headers = filterRequestHeaders(opts.headers);
  // Fail-closed: never forward client credentials (allowlist should already drop them).
  delete headers.authorization;
  delete headers["x-api-key"];
  headers[opts.upstream.authHeaderName] = opts.upstream.authHeaderValue;
  headers.accept = headers.accept || "application/json";

  if (opts.method === "POST") {
    headers["content-type"] = headers["content-type"] || "application/json";
  }

  // Connect timeout only until response headers arrive. Total timeout / client
  // abort remain on opts.signal for the entire stream lifetime.
  const connectController = new AbortController();
  const onParentAbort = () => connectController.abort();
  if (opts.signal.aborted) {
    connectController.abort();
  } else {
    opts.signal.addEventListener("abort", onParentAbort, { once: true });
  }
  const connectTimer = setTimeout(() => connectController.abort(), opts.connectTimeoutMs);

  try {
    const init = {
      method: opts.method,
      headers,
      signal: connectController.signal,
      redirect: "manual",
    };
    if (opts.method === "POST" && opts.body) {
      init.body = opts.body;
    }

    const response = await fetchImpl(url, init);
    // Headers received: drop connect timer only. Keep parent signal for body.
    clearTimeout(connectTimer);
    opts.signal.removeEventListener("abort", onParentAbort);

    // If parent already aborted while headers landed, cancel body immediately.
    if (opts.signal.aborted) {
      try {
        await response.body?.cancel();
      } catch {
        // ignore
      }
      const error = new Error("upstream timeout or aborted");
      error.code = "UPSTREAM_ABORTED";
      throw error;
    }

    // Bridge parent abort -> cancel response body for mid-stream disconnect.
    const onBodyAbort = () => {
      response.body?.cancel().catch(() => {});
    };
    opts.signal.addEventListener("abort", onBodyAbort, { once: true });

    return {
      status: response.status,
      headers: filterResponseHeaders(response.headers),
      body: response.body,
      response,
      release() {
        opts.signal.removeEventListener("abort", onBodyAbort);
      },
    };
  } catch (err) {
    clearTimeout(connectTimer);
    opts.signal.removeEventListener("abort", onParentAbort);
    if (opts.signal.aborted || connectController.signal.aborted) {
      const error = new Error("upstream timeout or aborted");
      error.code = "UPSTREAM_ABORTED";
      throw error;
    }
    const error = new Error("upstream request failed");
    error.code = "UPSTREAM_FAILED";
    throw error;
  }
}

export { FORWARDED_REQUEST_HEADERS, FORWARDED_RESPONSE_HEADERS, ENDPOINT_SUFFIXES };
