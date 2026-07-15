import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  createAIClient: vi.fn(),
  getActiveAIConfig: vi.fn(() => ({
    provider: "openai",
    baseUrl: "https://api.example.test/v1",
    apiKey: "test-key",
    model: "test-model",
  })),
  getRuntimeSettingValue: vi.fn(),
  parseFilesLocally: vi.fn(),
  parseFilesWithMineru: vi.fn(),
  parseFilesWithMineruAgent: vi.fn(),
}));

vi.mock("../agent/providers/aiClient", () => ({
  createAIClient: mocks.createAIClient,
}));

vi.mock("../agent/tools/executors/localDocumentParser", () => ({
  parseFilesLocally: mocks.parseFilesLocally,
}));

vi.mock("./settingsManager", () => ({
  getActiveAIConfig: mocks.getActiveAIConfig,
  getRuntimeSettingValue: mocks.getRuntimeSettingValue,
}));

vi.mock("./mineruOcr", () => ({
  parseFilesWithMineru: mocks.parseFilesWithMineru,
  parseFilesWithMineruAgent: mocks.parseFilesWithMineruAgent,
}));

import { recognizeWithOcrFallbacks } from "./ipcOcrHandlers";

describe("ipc OCR remote-data policy", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mocks.getRuntimeSettingValue.mockReturnValue(false);
    mocks.parseFilesLocally.mockResolvedValue([{
      filename: "invoice.txt",
      text: "发票号码 001\n价税合计 100.00",
      rows: [],
      provider: "local",
      sourceType: "txt",
      warnings: [],
    }]);
  });

  it("uses local parsing and makes no remote call when remote processing is disabled", async () => {
    const result = await recognizeWithOcrFallbacks("invoice", ["C:\\docs\\invoice.txt"]);

    expect(mocks.parseFilesWithMineru).not.toHaveBeenCalled();
    expect(mocks.parseFilesWithMineruAgent).not.toHaveBeenCalled();
    expect(mocks.createAIClient).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      kind: "invoice",
      fields: { 发票号码: "001" },
      remoteProcessing: [],
    });
  });

  it("blocks sensitive OCR text before invoice AI extraction", async () => {
    mocks.getRuntimeSettingValue.mockReturnValue(true);
    mocks.parseFilesLocally.mockResolvedValue([{
      filename: "invoice.txt",
      text: "发票号码 001\nsecret sk-1234567890abcdefghijklmnop",
      rows: [],
      provider: "local",
      sourceType: "txt",
      warnings: [],
    }]);

    const result = await recognizeWithOcrFallbacks("invoice", ["C:\\docs\\invoice.txt"]);

    expect(mocks.createAIClient).not.toHaveBeenCalled();
    expect(result.errors).toEqual(expect.arrayContaining([
      expect.stringContaining("高置信敏感凭据"),
    ]));
  });
});
