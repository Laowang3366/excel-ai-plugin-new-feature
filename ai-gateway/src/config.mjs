/**
 * Fail-closed gateway configuration.
 * Upstream targets come only from AI_GATEWAY_UPSTREAMS_JSON (no client baseUrl).
 */

const MAX_BODY_BYTES = 4 * 1024 * 1024;

/** @typedef {{ type: 'bearer' | 'x-api-key', env: string }} AuthConfig */
/** @typedef {{ id: string, baseUrl: string, authHeaderName: string, authHeaderValue: string }} Upstream */

/**
 * @param {NodeJS.ProcessEnv} [env]
 * @returns {{
 *   host: string,
 *   port: number,
 *   connectTimeoutMs: number,
 *   totalTimeoutMs: number,
 *   maxConcurrent: number,
 *   rateLimitMax: number,
 *   rateLimitWindowMs: number,
 *   allowLocalUpstreams: boolean,
 *   maxBodyBytes: number,
 *   upstreams: Map<string, Upstream>,
 * }}
 */
export function loadConfig(env = process.env) {
  const allowLocalUpstreams = parseBool(env.AI_GATEWAY_ALLOW_LOCAL_UPSTREAMS, false);
  const host = (env.AI_GATEWAY_HOST || "127.0.0.1").trim();
  const port = parsePositiveInt(env.AI_GATEWAY_PORT, 8787, "AI_GATEWAY_PORT");
  const connectTimeoutMs = parsePositiveInt(
    env.AI_GATEWAY_CONNECT_TIMEOUT_MS,
    10_000,
    "AI_GATEWAY_CONNECT_TIMEOUT_MS",
  );
  const totalTimeoutMs = parsePositiveInt(
    env.AI_GATEWAY_TOTAL_TIMEOUT_MS,
    120_000,
    "AI_GATEWAY_TOTAL_TIMEOUT_MS",
  );
  const maxConcurrent = parsePositiveInt(
    env.AI_GATEWAY_MAX_CONCURRENT,
    32,
    "AI_GATEWAY_MAX_CONCURRENT",
  );
  const rateLimitMax = parsePositiveInt(
    env.AI_GATEWAY_RATE_LIMIT_MAX,
    60,
    "AI_GATEWAY_RATE_LIMIT_MAX",
  );
  const rateLimitWindowMs = parsePositiveInt(
    env.AI_GATEWAY_RATE_LIMIT_WINDOW_MS,
    60_000,
    "AI_GATEWAY_RATE_LIMIT_WINDOW_MS",
  );

  if (totalTimeoutMs < connectTimeoutMs) {
    throw new Error("AI_GATEWAY_TOTAL_TIMEOUT_MS must be >= AI_GATEWAY_CONNECT_TIMEOUT_MS");
  }

  const raw = env.AI_GATEWAY_UPSTREAMS_JSON;
  if (typeof raw !== "string" || raw.trim() === "") {
    throw new Error("AI_GATEWAY_UPSTREAMS_JSON is required");
  }

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error("AI_GATEWAY_UPSTREAMS_JSON must be valid JSON");
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error("AI_GATEWAY_UPSTREAMS_JSON must be a JSON object");
  }

  /** @type {Map<string, Upstream>} */
  const upstreams = new Map();
  for (const [id, value] of Object.entries(parsed)) {
    validateUpstreamId(id);
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      throw new Error(`upstream ${id}: config must be an object`);
    }
    const baseUrl = assertSafeBaseUrl(value.baseUrl, { allowLocalUpstreams, id });
    const auth = value.auth;
    if (!auth || typeof auth !== "object" || Array.isArray(auth)) {
      throw new Error(`upstream ${id}: auth is required`);
    }
    if (auth.type !== "bearer" && auth.type !== "x-api-key") {
      throw new Error(`upstream ${id}: auth.type must be bearer or x-api-key`);
    }
    if (typeof auth.env !== "string" || !/^[A-Z][A-Z0-9_]*$/.test(auth.env)) {
      throw new Error(`upstream ${id}: auth.env must be an uppercase env var name`);
    }
    const secret = env[auth.env];
    if (typeof secret !== "string" || secret.trim() === "") {
      throw new Error(`upstream ${id}: env ${auth.env} is missing or empty`);
    }
    const authHeaderName = auth.type === "bearer" ? "authorization" : "x-api-key";
    const authHeaderValue = auth.type === "bearer" ? `Bearer ${secret.trim()}` : secret.trim();
    upstreams.set(id, {
      id,
      baseUrl,
      authHeaderName,
      authHeaderValue,
    });
  }

  if (upstreams.size === 0) {
    throw new Error("AI_GATEWAY_UPSTREAMS_JSON must define at least one upstream");
  }

  return {
    host,
    port,
    connectTimeoutMs,
    totalTimeoutMs,
    maxConcurrent,
    rateLimitMax,
    rateLimitWindowMs,
    allowLocalUpstreams,
    maxBodyBytes: MAX_BODY_BYTES,
    upstreams,
  };
}

/**
 * @param {string} id
 */
export function validateUpstreamId(id) {
  if (typeof id !== "string" || !/^[a-z][a-z0-9_-]{0,63}$/.test(id)) {
    throw new Error(`invalid upstream id: ${JSON.stringify(id)}`);
  }
}

/**
 * @param {unknown} raw
 * @param {{ allowLocalUpstreams: boolean, id: string }} opts
 * @returns {string}
 */
export function assertSafeBaseUrl(raw, opts) {
  if (typeof raw !== "string" || raw.trim() === "") {
    throw new Error(`upstream ${opts.id}: baseUrl is required`);
  }
  let url;
  try {
    url = new URL(raw);
  } catch {
    throw new Error(`upstream ${opts.id}: baseUrl is not a valid URL`);
  }
  if (url.username || url.password) {
    throw new Error(`upstream ${opts.id}: baseUrl must not include credentials`);
  }
  if (url.hash) {
    throw new Error(`upstream ${opts.id}: baseUrl must not include a fragment`);
  }
  if (url.search) {
    throw new Error(`upstream ${opts.id}: baseUrl must not include a query string`);
  }
  if (url.protocol !== "https:" && url.protocol !== "http:") {
    throw new Error(`upstream ${opts.id}: baseUrl must use http or https`);
  }

  const loopback = isLoopbackHost(url.hostname);
  const privateOrLocal = isPrivateOrLocalHost(url.hostname);

  if (opts.allowLocalUpstreams) {
    // Test/dev only: loopback HTTP or loopback HTTPS. Never open private LAN.
    if (!loopback) {
      throw new Error(
        `upstream ${opts.id}: AI_GATEWAY_ALLOW_LOCAL_UPSTREAMS only permits loopback hosts`,
      );
    }
  } else {
    if (url.protocol === "http:") {
      throw new Error(
        `upstream ${opts.id}: http baseUrl requires AI_GATEWAY_ALLOW_LOCAL_UPSTREAMS=1`,
      );
    }
    if (privateOrLocal) {
      throw new Error(`upstream ${opts.id}: private/local baseUrl host is not allowed`);
    }
  }

  // Normalize: strip trailing slash for fixed endpoint join
  return url.toString().replace(/\/+$/, "");
}

/**
 * @param {string} hostname
 */
export function isLoopbackHost(hostname) {
  const h = hostname.toLowerCase().replace(/^\[|\]$/g, "");
  return h === "localhost" || h === "127.0.0.1" || h === "::1" || h === "0:0:0:0:0:0:0:1";
}

/**
 * Loopback, RFC1918, link-local, CGNAT, unspecified.
 * @param {string} hostname
 */
export function isPrivateOrLocalHost(hostname) {
  const h = hostname.toLowerCase().replace(/^\[|\]$/g, "");
  if (isLoopbackHost(h) || h === "0.0.0.0" || h === "::" || h === "0:0:0:0:0:0:0:0") {
    return true;
  }
  // IPv4 dotted
  const m = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(h);
  if (m) {
    const a = Number(m[1]);
    const b = Number(m[2]);
    if ([a, b, Number(m[3]), Number(m[4])].some((n) => n > 255)) return true;
    if (a === 10) return true;
    if (a === 127) return true;
    if (a === 0) return true;
    if (a === 169 && b === 254) return true;
    if (a === 192 && b === 168) return true;
    if (a === 172 && b >= 16 && b <= 31) return true;
    // CGNAT 100.64.0.0/10
    if (a === 100 && b >= 64 && b <= 127) return true;
    return false;
  }
  // IPv6 unique local / link-local
  if (h.startsWith("fc") || h.startsWith("fd") || h.startsWith("fe80:")) return true;
  if (h.startsWith("::ffff:")) {
    const mapped = h.slice("::ffff:".length);
    if (isPrivateOrLocalHost(mapped)) return true;
  }
  return false;
}

/**
 * @param {string | undefined} raw
 * @param {number} fallback
 * @param {string} name
 */
function parsePositiveInt(raw, fallback, name) {
  if (raw === undefined || raw === "") return fallback;
  const n = Number(raw);
  if (!Number.isInteger(n) || n <= 0) {
    throw new Error(`${name} must be a positive integer`);
  }
  return n;
}

/**
 * @param {string | undefined} raw
 * @param {boolean} fallback
 */
function parseBool(raw, fallback) {
  if (raw === undefined || raw === "") return fallback;
  const v = raw.trim().toLowerCase();
  if (v === "1" || v === "true" || v === "yes") return true;
  if (v === "0" || v === "false" || v === "no") return false;
  throw new Error("AI_GATEWAY_ALLOW_LOCAL_UPSTREAMS must be 0 or 1");
}

export { MAX_BODY_BYTES };
