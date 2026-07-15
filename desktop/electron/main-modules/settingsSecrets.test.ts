import { describe, expect, it } from "vitest";

import {
  SETTINGS_SECRET_MASK,
  decryptProviderForRuntime,
  decryptSettingValueForRuntime,
  migrateSettingsSecrets,
  protectSettingValueForStorage,
  sanitizeSettingsForRenderer,
  type SettingsSecretCipher,
} from "./settingsSecrets";

const cipher: SettingsSecretCipher = {
  isAvailable: () => true,
  encrypt: (value) => Buffer.from(value, "utf8").toString("base64"),
  decrypt: (value) => Buffer.from(value, "base64").toString("utf8"),
};

describe("settings secret storage", () => {
  it("migrates plaintext credentials and exposes only masks to the renderer", () => {
    const migrated = migrateSettingsSecrets(
      {
        aiProviders: {
          p1: {
            id: "p1",
            apiKey: "canary-api-key",
            customHeaders: { "x-private-token": "canary-header" },
          },
        },
        mineruApiToken: "canary-mineru",
        compactionConfig: {
          enabled: true,
          remoteCompactApiKey: "canary-compact-key",
        },
      },
      cipher,
    );
    const serialized = JSON.stringify(migrated);
    expect(serialized).not.toContain("canary-api-key");
    expect(serialized).not.toContain("canary-header");
    expect(serialized).not.toContain("canary-mineru");
    expect(serialized).not.toContain("canary-compact-key");

    const renderer = sanitizeSettingsForRenderer(migrated);
    expect(renderer).toMatchObject({
      aiProviders: {
        p1: {
          apiKey: SETTINGS_SECRET_MASK,
          customHeaders: { "x-private-token": SETTINGS_SECRET_MASK },
        },
      },
      mineruApiToken: SETTINGS_SECRET_MASK,
      compactionConfig: {
        enabled: true,
        remoteCompactApiKey: SETTINGS_SECRET_MASK,
      },
    });
    expect(JSON.stringify(renderer)).not.toContain("canary");
    expect(
      decryptSettingValueForRuntime("compactionConfig", migrated.compactionConfig, cipher),
    ).toMatchObject({
      enabled: true,
      remoteCompactApiKey: "canary-compact-key",
    });
  });

  it("preserves the encrypted remote compaction key when the renderer sends a mask", () => {
    const current = migrateSettingsSecrets(
      {
        aiProviders: {},
        mineruApiToken: "",
        compactionConfig: { enabled: true, remoteCompactApiKey: "original-compact-secret" },
      },
      cipher,
    );
    const next = protectSettingValueForStorage(
      "compactionConfig",
      {
        enabled: false,
        remoteCompactApiKey: SETTINGS_SECRET_MASK,
      },
      current.compactionConfig,
      cipher,
    );

    expect(JSON.stringify(next)).not.toContain("original-compact-secret");
    expect(decryptSettingValueForRuntime("compactionConfig", next, cipher)).toMatchObject({
      enabled: false,
      remoteCompactApiKey: "original-compact-secret",
    });
  });

  it("preserves an existing encrypted secret when the renderer sends a mask", () => {
    const current = migrateSettingsSecrets(
      {
        aiProviders: { p1: { id: "p1", apiKey: "original-secret" } },
        mineruApiToken: "",
      },
      cipher,
    );
    const next = protectSettingValueForStorage(
      "aiProviders",
      {
        p1: { id: "p1", apiKey: SETTINGS_SECRET_MASK, model: "gpt-5" },
      },
      current.aiProviders,
      cipher,
    ) as Record<string, Record<string, unknown>>;

    expect(decryptProviderForRuntime(next.p1, cipher)).toMatchObject({
      apiKey: "original-secret",
      model: "gpt-5",
    });
  });

  it("fails closed when secure storage is unavailable", () => {
    const unavailableCipher = { ...cipher, isAvailable: () => false };
    expect(() =>
      migrateSettingsSecrets(
        {
          aiProviders: { p1: { apiKey: "secret" } },
          mineruApiToken: "",
        },
        unavailableCipher,
      ),
    ).toThrow("secure_storage_unavailable");
    expect(() =>
      migrateSettingsSecrets(
        {
          aiProviders: {},
          mineruApiToken: "",
          compactionConfig: { remoteCompactApiKey: "remote-secret" },
        },
        unavailableCipher,
      ),
    ).toThrow("secure_storage_unavailable");
  });
});
