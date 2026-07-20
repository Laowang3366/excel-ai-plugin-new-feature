/**
 * Same-origin AI gateway HTTP server (Node 22, no framework).
 */

import http from "node:http";
import { ConcurrencyLimiter, RateLimiter } from "./limits.mjs";
import { proxyToUpstream } from "./proxy.mjs";

const API_PREFIX = "/api/ai/v1/";
const POST_ENDPOINTS = new Set(["chat/completions", "responses", "messages"]);

/**
 * @param {ReturnType<import('./config.mjs').loadConfig>} config
 * @param {{ fetchImpl?: typeof fetch }} [deps]
 */
export function createServer(config, deps = {}) {
  const concurrency = new ConcurrencyLimiter(config.maxConcurrent);
  const rateLimiter = new RateLimiter(config.rateLimitMax, config.rateLimitWindowMs);
  const fetchImpl = deps.fetchImpl || globalThis.fetch;

  const server = http.createServer(async (req, res) => {
    try {
      await handleRequest(req, res, config, { concurrency, rateLimiter, fetchImpl });
    } catch (err) {
      if (!res.headersSent) {
        sendJson(res, 500, { error: "internal_error" });
      } else {
        res.destroy();
      }
      logEvent("error", { event: "request_failed", code: err?.code || "INTERNAL" });
    }
  });

  server.keepAliveTimeout = 65_000;
  server.headersTimeout = 70_000;
  server.requestTimeout = config.totalTimeoutMs + 5_000;
  return server;
}

/**
 * @param {http.IncomingMessage} req
 * @param {http.ServerResponse} res
 * @param {ReturnType<import('./config.mjs').loadConfig>} config
 * @param {{ concurrency: ConcurrencyLimiter, rateLimiter: RateLimiter, fetchImpl: typeof fetch }} runtime
 */
async function handleRequest(req, res, config, runtime) {
  const method = (req.method || "GET").toUpperCase();
  const url = new URL(req.url || "/", "http://gateway.local");
  const path = url.pathname;

  if (method === "GET" && path === "/healthz") {
    sendJson(res, 200, { status: "ok" });
    return;
  }

  if (!path.startsWith(API_PREFIX)) {
    sendJson(res, 404, { error: "not_found" });
    return;
  }

  const rest = path.slice(API_PREFIX.length);
  const slash = rest.indexOf("/");
  if (slash <= 0) {
    sendJson(res, 404, { error: "not_found" });
    return;
  }
  const upstreamId = rest.slice(0, slash);
  const suffix = rest.slice(slash + 1);

  if (!config.upstreams.has(upstreamId)) {
    sendJson(res, 404, { error: "unknown_upstream" });
    return;
  }

  const isModelsGet = method === "GET" && suffix === "models";
  const isPost = method === "POST" && POST_ENDPOINTS.has(suffix);
  if (!isModelsGet && !isPost) {
    sendJson(res, 404, { error: "not_found" });
    return;
  }

  const clientKey = clientRateKey(req);
  if (!runtime.rateLimiter.allow(clientKey)) {
    sendJson(res, 429, { error: "rate_limited" });
    return;
  }
  if (!runtime.concurrency.tryAcquire()) {
    sendJson(res, 503, { error: "concurrency_limited" });
    return;
  }

  const ac = new AbortController();
  const abortUpstream = () => {
    if (!ac.signal.aborted) ac.abort();
  };
  // Socket close is the reliable client-disconnect signal. Ignore after a full
  // response completes (writableFinished).
  const onSocketClose = () => {
    if (!res.writableFinished) abortUpstream();
  };
  const onResClose = () => {
    if (!res.writableFinished) abortUpstream();
  };
  req.on("aborted", abortUpstream);
  req.socket?.on("close", onSocketClose);
  res.on("close", onResClose);

  try {
    let body = null;
    if (isPost) {
      body = await readBodyLimited(req, config.maxBodyBytes);
      if (body === null) {
        sendJson(res, 413, { error: "payload_too_large" });
        return;
      }
      if (!isJsonContentType(req.headers["content-type"])) {
        sendJson(res, 415, { error: "unsupported_media_type" });
        return;
      }
      if (!looksLikeJsonObjectOrArray(body)) {
        sendJson(res, 400, { error: "invalid_json_body" });
        return;
      }
    } else if (method === "GET") {
      // models GET: reject unexpected body
      const peek = await readBodyLimited(req, 1);
      if (peek && peek.length > 0) {
        sendJson(res, 400, { error: "body_not_allowed" });
        return;
      }
    }

    const upstream = config.upstreams.get(upstreamId);
    const proxied = await proxyToUpstream({
      upstream,
      suffix,
      method: isPost ? "POST" : "GET",
      headers: req.headers,
      body,
      connectTimeoutMs: config.connectTimeoutMs,
      totalTimeoutMs: config.totalTimeoutMs,
      signal: ac.signal,
      fetchImpl: runtime.fetchImpl,
    });

    res.writeHead(proxied.status, {
      ...proxied.headers,
      connection: "close",
    });

    if (!proxied.body) {
      res.end();
      return;
    }

    const reader = proxied.body.getReader();
    const onSignalAbort = () => {
      reader.cancel().catch(() => {});
    };
    ac.signal.addEventListener("abort", onSignalAbort, { once: true });
    try {
      while (!ac.signal.aborted) {
        const { done, value } = await reader.read();
        if (done) break;
        if (ac.signal.aborted || res.destroyed || !res.writable) {
          await reader.cancel().catch(() => {});
          abortUpstream();
          break;
        }
        const ok = res.write(Buffer.from(value));
        if (!ok) {
          if (ac.signal.aborted || res.destroyed) {
            await reader.cancel().catch(() => {});
            abortUpstream();
            break;
          }
          await onceDrain(res);
        }
      }
    } catch {
      abortUpstream();
    } finally {
      ac.signal.removeEventListener("abort", onSignalAbort);
      if (!res.writableEnded) {
        res.end();
      }
    }
  } catch (err) {
    if (res.headersSent) {
      res.destroy();
      return;
    }
    if (err?.code === "UPSTREAM_ABORTED") {
      sendJson(res, 504, { error: "upstream_timeout" });
      return;
    }
    sendJson(res, 502, { error: "upstream_error" });
  } finally {
    req.off("aborted", abortUpstream);
    req.socket?.off("close", onSocketClose);
    res.off("close", onResClose);
    runtime.concurrency.release();
  }
}

/**
 * @param {http.IncomingMessage} req
 * @param {number} maxBytes
 * @returns {Promise<Buffer | null>} null if over limit
 */
function readBodyLimited(req, maxBytes) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    let settled = false;
    let overLimit = false;

    const finish = (value) => {
      if (settled) return;
      settled = true;
      resolve(value);
    };

    req.on("data", (chunk) => {
      if (overLimit || settled) return;
      total += chunk.length;
      if (total > maxBytes) {
        overLimit = true;
        // Stop consuming; drain remaining so we can still answer 413 cleanly.
        req.resume();
        finish(null);
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => {
      if (overLimit) {
        finish(null);
        return;
      }
      finish(Buffer.concat(chunks, total));
    });
    req.on("error", (err) => {
      if (settled) return;
      settled = true;
      reject(err);
    });
  });
}

/**
 * @param {string | string[] | undefined} contentType
 */
function isJsonContentType(contentType) {
  if (!contentType) return false;
  const raw = Array.isArray(contentType) ? contentType[0] : contentType;
  const base = raw.split(";")[0].trim().toLowerCase();
  return base === "application/json" || base.endsWith("+json");
}

/**
 * @param {Buffer} body
 */
function looksLikeJsonObjectOrArray(body) {
  if (!body || body.length === 0) return false;
  // Reject non-JSON before proxy; still fail closed without logging body
  const text = body.toString("utf8").trimStart();
  if (!(text.startsWith("{") || text.startsWith("["))) return false;
  try {
    JSON.parse(body.toString("utf8"));
    return true;
  } catch {
    return false;
  }
}

/**
 * @param {http.IncomingMessage} req
 */
function clientRateKey(req) {
  const xf = req.headers["x-forwarded-for"];
  if (typeof xf === "string" && xf.trim()) {
    return xf.split(",")[0].trim();
  }
  return req.socket.remoteAddress || "unknown";
}

/**
 * @param {http.ServerResponse} res
 */
function onceDrain(res) {
  return new Promise((resolve) => res.once("drain", resolve));
}

/**
 * @param {http.ServerResponse} res
 * @param {number} status
 * @param {Record<string, unknown>} obj
 */
function sendJson(res, status, obj) {
  const payload = Buffer.from(JSON.stringify(obj), "utf8");
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "content-length": String(payload.length),
    "cache-control": "no-store",
    connection: "close",
  });
  res.end(payload);
}

/**
 * Structured logs without secrets, bodies, or full URL query.
 * @param {'info'|'error'} level
 * @param {Record<string, unknown>} fields
 */
function logEvent(level, fields) {
  const line = JSON.stringify({
    ts: new Date().toISOString(),
    level,
    ...fields,
  });
  if (level === "error") {
    console.error(line);
  } else {
    console.log(line);
  }
}
