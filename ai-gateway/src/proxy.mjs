/**
 * Fixed-path upstream proxy: byte-transparent stream with header allowlists.
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
    // Reject CR/LF injection
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
 *   totalTimeoutMs: number,
 *   signal: AbortSignal,
 *   fetchImpl?: typeof fetch,
 * }} opts
 */
export async function proxyToUpstream(opts) {
  const fetchImpl = opts.fetchImpl || globalThis.fetch;
  const url = joinUpstreamUrl(opts.upstream.baseUrl, opts.suffix);
  const headers = filterRequestHeaders(opts.headers);
  headers[opts.upstream.authHeaderName] = opts.upstream.authHeaderValue;
  headers.accept = headers.accept || "application/json";

  if (opts.method === "POST") {
    headers["content-type"] = headers["content-type"] || "application/json";
  }

  const controller = new AbortController();
  const onAbort = () => controller.abort();
  if (opts.signal.aborted) {
    controller.abort();
  } else {
    opts.signal.addEventListener("abort", onAbort, { once: true });
  }

  const totalTimer = setTimeout(() => controller.abort(), opts.totalTimeoutMs);
  let connectTimer = setTimeout(() => controller.abort(), opts.connectTimeoutMs);

  try {
    const init = {
      method: opts.method,
      headers,
      signal: controller.signal,
      // Node undici: redirect must not open SSRF via 3xx to private hosts
      redirect: "manual",
    };
    if (opts.method === "POST" && opts.body) {
      init.body = opts.body;
    }

    const response = await fetchImpl(url, init);
    clearTimeout(connectTimer);
    connectTimer = null;

    return {
      status: response.status,
      headers: filterResponseHeaders(response.headers),
      body: response.body,
      // Keep response for tests that need arrayBuffer; streaming path uses body
      response,
    };
  } catch (err) {
    if (controller.signal.aborted) {
      const error = new Error("upstream timeout or aborted");
      error.code = "UPSTREAM_ABORTED";
      throw error;
    }
    const error = new Error("upstream request failed");
    error.code = "UPSTREAM_FAILED";
    throw error;
  } finally {
    clearTimeout(totalTimer);
    if (connectTimer) clearTimeout(connectTimer);
    opts.signal.removeEventListener("abort", onAbort);
  }
}

export { FORWARDED_REQUEST_HEADERS, FORWARDED_RESPONSE_HEADERS, ENDPOINT_SUFFIXES };
