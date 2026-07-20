import http from "node:http";
import { loadConfig } from "../src/config.mjs";
import { createServer } from "../src/server.mjs";

/**
 * Start a local fake upstream that records requests and can stream.
 * @param {{
 *   status?: number,
 *   headers?: Record<string, string>,
 *   body?: string | Buffer,
 *   streamChunks?: (string | Buffer)[],
 *   delayMs?: number,
 *   hang?: boolean,
 * }} [opts]
 */
export function startFakeUpstream(opts = {}) {
  /** @type {import('node:http').IncomingMessage[]} */
  const requests = [];
  const server = http.createServer(async (req, res) => {
    const chunks = [];
    for await (const c of req) chunks.push(c);
    const body = Buffer.concat(chunks);
    requests.push({
      method: req.method,
      url: req.url,
      headers: { ...req.headers },
      body,
    });

    if (opts.hang) {
      // never respond (no headers)
      return;
    }
    if (opts.delayMs) {
      await new Promise((r) => setTimeout(r, opts.delayMs));
    }

    const status = opts.status ?? 200;
    const headers = {
      "content-type": "application/json",
      "x-request-id": "fake-req-1",
      "set-cookie": "secret=1",
      "x-should-strip": "nope",
      ...(opts.headers || {}),
    };
    res.writeHead(status, headers);

    if (opts.hangAfterHeaders) {
      // headers sent, body never completes
      if (opts.hangAfterHeadersWrite) {
        res.write(opts.hangAfterHeadersWrite);
      }
      return;
    }

    if (opts.streamChunks) {
      for (const chunk of opts.streamChunks) {
        res.write(typeof chunk === "string" ? chunk : chunk);
        await new Promise((r) => setTimeout(r, 5));
      }
      res.end();
      return;
    }

    const bodyOut = opts.body ?? JSON.stringify({ ok: true, path: req.url });
    res.end(typeof bodyOut === "string" ? bodyOut : bodyOut);
  });

  return new Promise((resolve) => {
    server.listen(0, "127.0.0.1", () => {
      const addr = server.address();
      const port = typeof addr === "object" && addr ? addr.port : 0;
      resolve({
        server,
        port,
        baseUrl: `http://127.0.0.1:${port}/v1`,
        requests,
        close: () =>
          new Promise((r) => {
            server.close(() => r());
          }),
      });
    });
  });
}

/**
 * @param {{
 *   upstreamPort: number,
 *   basePath?: string,
 *   env?: Record<string, string>,
 *   fetchImpl?: typeof fetch,
 * }} opts
 */
export function startGateway(opts) {
  const basePath = opts.basePath || "/v1";
  const env = {
    AI_GATEWAY_HOST: "127.0.0.1",
        AI_GATEWAY_ALLOW_LOCAL_UPSTREAMS: "1",
    AI_GATEWAY_CONNECT_TIMEOUT_MS: "2000",
    AI_GATEWAY_TOTAL_TIMEOUT_MS: "5000",
    AI_GATEWAY_MAX_CONCURRENT: "4",
    AI_GATEWAY_RATE_LIMIT_MAX: "1000",
    AI_GATEWAY_RATE_LIMIT_WINDOW_MS: "60000",
    FAKE_API_KEY: "test-secret-key",
    AI_GATEWAY_UPSTREAMS_JSON: JSON.stringify({
      fake: {
        baseUrl: `http://127.0.0.1:${opts.upstreamPort}${basePath}`,
        auth: { type: "bearer", env: "FAKE_API_KEY" },
      },
    }),
    ...(opts.env || {}),
  };
  const config = loadConfig(env);
  // force ephemeral port
  config.port = 0;
  const server = createServer(config, { fetchImpl: opts.fetchImpl });
  return new Promise((resolve) => {
    server.listen(0, "127.0.0.1", () => {
      const addr = server.address();
      const port = typeof addr === "object" && addr ? addr.port : 0;
      resolve({
        server,
        port,
        base: `http://127.0.0.1:${port}`,
        config,
        close: () =>
          new Promise((r) => {
            server.close(() => r());
          }),
      });
    });
  });
}

/**
 * @param {string} url
 * @param {RequestInit} [init]
 */
export async function request(url, init = {}) {
  const res = await fetch(url, init);
  const buf = Buffer.from(await res.arrayBuffer());
  return {
    status: res.status,
    headers: Object.fromEntries(res.headers.entries()),
    body: buf,
    text: buf.toString("utf8"),
    json: () => {
      try {
        return JSON.parse(buf.toString("utf8"));
      } catch {
        return null;
      }
    },
  };
}
