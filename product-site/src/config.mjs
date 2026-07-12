import path from "node:path";

function requiredInProduction(name, fallback) {
  const value = process.env[name] || fallback;
  if (process.env.NODE_ENV === "production" && !process.env[name]) {
    throw new Error(`生产环境缺少 ${name}`);
  }
  return value;
}

export function loadConfig(overrides = {}) {
  const root = path.resolve(import.meta.dirname, "..");
  const dataDir = path.resolve(overrides.dataDir || process.env.DATA_DIR || path.join(root, ".local", "data"));
  return {
    host: overrides.host || process.env.HOST || "127.0.0.1",
    port: Number(overrides.port || process.env.PORT || 18120),
    publicDir: path.resolve(overrides.publicDir || process.env.PUBLIC_DIR || path.join(root, "public")),
    dataDir,
    releasesDir: path.resolve(overrides.releasesDir || process.env.RELEASES_DIR || path.join(root, ".local", "releases")),
    databasePath: path.resolve(overrides.databasePath || process.env.DATABASE_PATH || path.join(dataDir, "analytics.sqlite")),
    adminPasswordHash: overrides.adminPasswordHash || requiredInProduction("ADMIN_PASSWORD_HASH", "scrypt$YpRPhq5bCIWtwFLGCwBvZQ==$rPJ4dulEcgeS+t95RvP8Hv6ZDDtgZu47u19j7boQHVI="),
    cookieSecret: overrides.cookieSecret || requiredInProduction("COOKIE_SECRET", "development-cookie-secret-change-before-deploy-1234"),
    analyticsSalt: overrides.analyticsSalt || requiredInProduction("ANALYTICS_SALT", "development-analytics-salt"),
    useAccelRedirect: overrides.useAccelRedirect ?? process.env.USE_ACCEL_REDIRECT === "true",
  };
}
