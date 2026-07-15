import assert from "node:assert/strict";
import test from "node:test";

import { loadConfig } from "../src/config.mjs";

const validOverrides = {
  adminPasswordHash: "scrypt$salt$hash",
  cookieSecret: "test-cookie-secret-with-at-least-32-characters",
  analyticsSalt: "test-analytics-salt",
};

test("analytics retention and rotation use privacy-preserving defaults", () => {
  const config = loadConfig(validOverrides);
  assert.equal(config.analyticsRetentionDays, 90);
  assert.equal(config.analyticsIpRotationDays, 30);
});

test("analytics retention and rotation reject invalid bounds", () => {
  assert.throws(
    () => loadConfig({ ...validOverrides, analyticsRetentionDays: 0 }),
    /ANALYTICS_RETENTION_DAYS/,
  );
  assert.throws(
    () => loadConfig({ ...validOverrides, analyticsIpRotationDays: 1.5 }),
    /ANALYTICS_IP_ROTATION_DAYS/,
  );
});
