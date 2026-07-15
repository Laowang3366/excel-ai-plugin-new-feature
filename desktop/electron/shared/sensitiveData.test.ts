import { describe, expect, it } from "vitest";

import {
  findHighConfidenceSensitiveData,
  isSensitiveFieldName,
  redactSensitiveText,
  redactSensitiveValue,
  summarizeValueForAudit,
} from "./sensitiveData";

const CANARY = "sk-1234567890abcdefghijklmnop";

describe("sensitiveData", () => {
  it("detects and redacts high-confidence credentials", () => {
    expect(findHighConfidenceSensitiveData([`Authorization: Bearer ${CANARY}`])).toEqual([
      "openai-style-key",
    ]);
    expect(redactSensitiveText(`value=${CANARY}`)).toBe("value=[REDACTED:openai-style-key]");
  });

  it("redacts an entire private key block instead of only its header", () => {
    const privateKey = [
      "-----BEGIN PRIVATE KEY-----",
      "sensitive-private-key-body",
      "-----END PRIVATE KEY-----",
    ].join("\n");

    const redacted = redactSensitiveText(privateKey);

    expect(redacted).toBe("[REDACTED:private-key]");
    expect(redacted).not.toContain("sensitive-private-key-body");
  });

  it("redacts sensitive fields recursively and handles cycles", () => {
    const input: Record<string, unknown> = {
      apiKey: CANARY,
      nested: { authorization: `Bearer ${CANARY}`, note: `token ${CANARY}` },
    };
    input.self = input;

    const serialized = JSON.stringify(redactSensitiveValue(input));

    expect(serialized).not.toContain(CANARY);
    expect(serialized).toContain("[REDACTED]");
    expect(serialized).toContain("[Circular]");
    expect(serialized).toContain("[REDACTED:openai-style-key]");
  });

  it("recognizes credential field names without treating ordinary fields as secrets", () => {
    expect(isSensitiveFieldName("remoteCompactApiKey")).toBe(true);
    expect(isSensitiveFieldName("mineruApiToken")).toBe(true);
    expect(isSensitiveFieldName("tokenUsage")).toBe(false);
    expect(isSensitiveFieldName("status")).toBe(false);
  });

  it("summarizes value shape without preserving values or secret-derived hashes", () => {
    const summary = summarizeValueForAudit({
      apiKey: CANARY,
      query: "confidential quarterly forecast",
      rows: [
        [1, 2],
        [3, 4],
      ],
      success: true,
    });

    expect(summary).not.toContain(CANARY);
    expect(summary).not.toContain("confidential quarterly forecast");
    expect(JSON.parse(summary)).toMatchObject({
      type: "object",
      fields: 4,
      stats: {
        redactedFields: 1,
        arrays: 3,
        numbers: 4,
      },
    });
    expect(JSON.parse(summary).fingerprint).toMatch(/^[a-f0-9]{64}$/);
  });
});
