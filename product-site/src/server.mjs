import path from "node:path";

import cookie from "@fastify/cookie";
import helmet from "@fastify/helmet";
import rateLimit from "@fastify/rate-limit";
import fastifyStatic from "@fastify/static";
import Fastify from "fastify";

import { clearAdminSession, hasValidAdminSession, setAdminSession, verifyPassword } from "./auth.mjs";
import { loadConfig } from "./config.mjs";
import { createAnalyticsDatabase } from "./database.mjs";
import { createReleaseStore } from "./releaseStore.mjs";

export async function buildServer(overrides = {}) {
  const config = loadConfig(overrides);
  const app = Fastify({ logger: overrides.logger ?? true, trustProxy: true, bodyLimit: 64 * 1024 });
  const analytics = createAnalyticsDatabase(config.databasePath, config.analyticsSalt);
  const releases = createReleaseStore(config.releasesDir);

  await app.register(cookie, { secret: config.cookieSecret });
  await app.register(helmet, {
    contentSecurityPolicy: {
      directives: {
        defaultSrc: ["'self'"],
        scriptSrc: ["'self'"],
        styleSrc: ["'self'"],
        imgSrc: ["'self'", "data:"],
        connectSrc: ["'self'"],
        objectSrc: ["'none'"],
        frameAncestors: ["'none'"],
      },
    },
  });
  await app.register(rateLimit, { max: 240, timeWindow: "1 minute" });
  await app.register(fastifyStatic, { root: config.publicDir, prefix: "/" });
  await app.register(fastifyStatic, {
    root: config.releasesDir,
    prefix: "/releases/windows/",
    decorateReply: false,
    index: false,
  });

  app.addHook("onClose", async () => analytics.close());
  app.get("/healthz", async () => ({ status: "ok" }));

  app.get("/api/v1/releases/current", async (_request, reply) => {
    try {
      reply.header("Cache-Control", "public, max-age=60");
      return await releases.readPublicRelease();
    } catch {
      return reply.code(404).send({ error: "暂无可用版本" });
    }
  });

  app.get("/api/v1/updates/check", async (_request, reply) => {
    try {
      reply.header("Cache-Control", "no-store");
      return await releases.readSignedManifest();
    } catch {
      return reply.code(404).send({ error: "暂无可用更新" });
    }
  });

  app.get("/download/windows", async (request, reply) => {
    try {
      const installer = await releases.getInstaller();
      analytics.recordDownload({
        version: installer.release.version,
        artifact: installer.fileName,
        ip: request.ip,
        userAgent: request.headers["user-agent"],
        referer: request.headers.referer,
      });
      reply.header("Content-Type", "application/vnd.microsoft.portable-executable");
      reply.header("Content-Disposition", `attachment; filename="${installer.fileName}"`);
      reply.header("Content-Length", installer.size);
      reply.header("Cache-Control", "private, no-store");
      if (config.useAccelRedirect) {
        reply.header("X-Accel-Redirect", `/_wenge_releases/${encodeURIComponent(installer.fileName)}`);
        return reply.send();
      }
      return reply.send(installer.stream());
    } catch {
      return reply.code(404).send({ error: "安装包尚未发布" });
    }
  });

  app.post("/api/admin/login", { config: { rateLimit: { max: 8, timeWindow: "15 minutes" } } }, async (request, reply) => {
    const password = typeof request.body?.password === "string" ? request.body.password : "";
    if (!(await verifyPassword(password, config.adminPasswordHash))) {
      return reply.code(401).send({ error: "账号或密码错误" });
    }
    setAdminSession(reply);
    return { success: true };
  });

  const requireAdmin = async (request, reply) => {
    if (!hasValidAdminSession(request)) return reply.code(401).send({ error: "未登录" });
  };

  app.post("/api/admin/logout", { preHandler: requireAdmin }, async (_request, reply) => {
    clearAdminSession(reply);
    return { success: true };
  });
  app.get("/api/admin/session", { preHandler: requireAdmin }, async () => ({ authenticated: true }));
  app.get("/api/admin/stats", { preHandler: requireAdmin }, async (request) => {
    return analytics.getStats(request.query?.days);
  });
  app.get("/api/admin/release", { preHandler: requireAdmin }, async (_request, reply) => {
    try {
      return await releases.readPublicRelease();
    } catch {
      return reply.code(404).send({ error: "暂无可用版本" });
    }
  });

  app.get("/admin", async (_request, reply) => reply.redirect("/admin/"));
  return { app, config };
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(import.meta.filename)) {
  const { app, config } = await buildServer();
  await app.listen({ host: config.host, port: config.port });
}
