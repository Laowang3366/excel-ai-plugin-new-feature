import { describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  chat: vi.fn(),
  createAIClient: vi.fn(),
  getActiveAIConfig: vi.fn(() => ({
    provider: "openai",
    baseUrl: "https://api.example.test/v1",
    apiKey: "test-key",
    model: "test-model",
  })),
}));

vi.mock("../agent/providers/aiClient", () => ({
  createAIClient: mocks.createAIClient,
}));

vi.mock("./settingsManager", () => ({
  getActiveAIConfig: mocks.getActiveAIConfig,
}));

import { buildOcrResultFromDocuments, normalizeOcrVisionResult } from "./ocrDocumentResultBuilder";

describe("normalizeOcrVisionResult", () => {
  it("normalizes fenced JSON and coerces structured values", () => {
    const result = normalizeOcrVisionResult(
      "invoice",
      '```json\n{"fields":{"发票号码":123,"备注":null},"rows":[["金额",106]],"invoices":[{"filename":"a.pdf","fields":{"税额":6}}],"errors":[" warning ",3]}\n```',
    );

    expect(result).toMatchObject({
      kind: "invoice",
      fields: { 发票号码: "123", 备注: "" },
      rows: [["金额", "106"]],
      invoices: [{ filename: "a.pdf", fields: { 税额: "6" } }],
      errors: ["warning"],
    });
  });

  it("falls back to plain invoice text when model JSON is invalid", () => {
    const result = normalizeOcrVisionResult("invoice", "发票号码：001\n价税合计：106.00");

    expect(result.kind).toBe("invoice");
    expect(result.fields).toMatchObject({ 发票号码: "001", 价税合计: "106.00" });
  });
});

describe("buildOcrResultFromDocuments", () => {
  it("builds an image result without invoking invoice extraction", async () => {
    const result = await buildOcrResultFromDocuments(
      [
        {
          filename: "notes.txt",
          text: "普通会议记录",
          rows: [["主题", "预算"]],
        },
      ],
      "image",
    );

    expect(result).toMatchObject({
      kind: "image",
      text: "## notes.txt\n普通会议记录",
      rows: [["主题", "预算"]],
      fields: {},
    });
  });

  it("merges invoice extraction with local fallback fields", async () => {
    mocks.chat.mockResolvedValue({
      content: JSON.stringify({
        fields: { 发票号码: "001", 税额: "6.00" },
        invoices: [
          {
            filename: "invoice.txt",
            fields: { 发票号码: "001", 税额: "6.00" },
            rows: [],
          },
        ],
        errors: [],
      }),
    });
    mocks.createAIClient.mockReturnValue({ chat: mocks.chat });

    const result = await buildOcrResultFromDocuments(
      [
        {
          filename: "invoice.txt",
          text: "发票号码 001\n价税合计 106.00",
          rows: [],
        },
      ],
      "invoice",
    );

    expect(mocks.chat).toHaveBeenCalledOnce();
    expect(result).toMatchObject({
      kind: "invoice",
      fields: { 发票号码: "001", 税额: "6.00" },
      invoices: [
        {
          filename: "invoice.txt",
          fields: { 发票号码: "001", 税额: "6.00", 价税合计: "106.00" },
        },
      ],
      remoteProcessing: [
        {
          operation: "invoice-extraction",
          destination: "api.example.test",
        },
      ],
    });
  });
});
