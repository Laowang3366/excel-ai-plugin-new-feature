import path from "node:path";

function requiredInProduction(name, fallback) {
  const value = process.env[name] || fallback;
  if (process.env.NODE_ENV === "production" && !process.env[name]) {
    throw new Error(`生产环境缺少 ${name}`);
  }
  return value;
}

function validateProductionSecret(name, value, minimumLength) {
  if (process.env.NODE_ENV !== "production") return value;
  if (typeof value !== "string" || value.length < minimumLength || new Set(value).size < 8) {
    throw new Error(`生产环境 ${name} 强度不足`);
  }
  return value;
}

function readBoundedInteger(name, overrideValue, fallback, minimum, maximum) {
  const rawValue = overrideValue ?? process.env[name];
  if (rawValue === undefined || rawValue === "") return fallback;
  const value = Number(rawValue);
  if (!Number.isInteger(value) || value < minimum || value > maximum) {
    throw new Error(`${name} 必须是 ${minimum}-${maximum} 之间的整数`);
  }
  return value;
}

export function loadConfig(overrides = {}) {
  const root = path.resolve(import.meta.dirname, "..");
  const dataDir = path.resolve(overrides.dataDir || process.env.DATA_DIR || path.join(root, ".local", "data"));
  const config = {
    host: overrides.host || process.env.HOST || "127.0.0.1",
    port: Number(overrides.port || process.env.PORT || 18120),
    publicDir: path.resolve(overrides.publicDir || process.env.PUBLIC_DIR || path.join(root, "public")),
    dataDir,
    releasesDir: path.resolve(overrides.releasesDir || process.env.RELEASES_DIR || path.join(root, ".local", "releases")),
    databasePath: path.resolve(overrides.databasePath || process.env.DATABASE_PATH || path.join(dataDir, "analytics.sqlite")),
    adminPasswordHash: overrides.adminPasswordHash || requiredInProduction("ADMIN_PASSWORD_HASH", "scrypt$YpRPhq5bCIWtwFLGCwBvZQ==$rPJ4dulEcgeS+t95RvP8Hv6ZDDtgZu47u19j7boQHVI="),
    cookieSecret: overrides.cookieSecret || requiredInProduction("COOKIE_SECRET", "development-cookie-secret-change-before-deploy-1234"),
    analyticsSalt: overrides.analyticsSalt || requiredInProduction("ANALYTICS_SALT", "development-analytics-salt"),
    analyticsRetentionDays: readBoundedInteger("ANALYTICS_RETENTION_DAYS", overrides.analyticsRetentionDays, 90, 1, 3650),
    analyticsIpRotationDays: readBoundedInteger("ANALYTICS_IP_ROTATION_DAYS", overrides.analyticsIpRotationDays, 30, 1, 365),
    useAccelRedirect: overrides.useAccelRedirect ?? process.env.USE_ACCEL_REDIRECT === "true",
  };
  validateProductionSecret("COOKIE_SECRET", config.cookieSecret, 32);
  validateProductionSecret("ANALYTICS_SALT", config.analyticsSalt, 16);
  if (process.env.NODE_ENV === "production" && !/^scrypt\$[^$]+\$[^$]+$/.test(config.adminPasswordHash)) {
    throw new Error("生产环境 ADMIN_PASSWORD_HASH 格式无效");
  }
  return config;
}
