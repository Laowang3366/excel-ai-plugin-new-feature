/**
 * Process entry: load config fail-closed, listen on configured host/port.
 */

import { loadConfig } from "./config.mjs";
import { createServer } from "./server.mjs";

function main() {
  let config;
  try {
    config = loadConfig(process.env);
  } catch (err) {
    console.error(
      JSON.stringify({
        ts: new Date().toISOString(),
        level: "error",
        event: "config_invalid",
        message: err instanceof Error ? err.message : "invalid configuration",
      }),
    );
    process.exit(1);
  }

  const server = createServer(config);
  server.listen(config.port, config.host, () => {
    console.log(
      JSON.stringify({
        ts: new Date().toISOString(),
        level: "info",
        event: "listen",
        host: config.host,
        port: config.port,
        upstreamCount: config.upstreams.size,
      }),
    );
  });

  const shutdown = (signal) => {
    console.log(
      JSON.stringify({
        ts: new Date().toISOString(),
        level: "info",
        event: "shutdown",
        signal,
      }),
    );
    server.close(() => process.exit(0));
    setTimeout(() => process.exit(1), 10_000).unref();
  };
  process.on("SIGINT", () => shutdown("SIGINT"));
  process.on("SIGTERM", () => shutdown("SIGTERM"));
}

main();
