import { describe, expect, it } from "vitest";

import {
  AgentStartTurnInput,
  AgentInterruptInput,
  AppLogInput,
  ExcelWriteRangeInput,
  ExcelReadRangeInput,
  FileWriteTempFileInput,
  IPC_MAX_ATTACHMENTS,
  IPC_MAX_CHAT_CONTENT_CHARS,
  IPC_MAX_EXCEL_CELLS,
  IPC_MAX_OCR_FILES,
  LaunchOfficeApplicationInput,
  OcrRecognizeInput,
  SettingsGetInput,
  SettingsSetInput,
  OfficeAutomationTemplateRunInput,
  estimateBase64DecodedBytes,
  isBase64PayloadWithinLimit,
  validateInput,
} from "./ipcSchemas";

describe("ipcSchemas", () => {
  it("validates IPC inputs that previously relied on ad hoc checks", () => {
    expect(validateInput(AgentInterruptInput, { threadId: null })).toEqual({ threadId: null });
    expect(validateInput(ExcelReadRangeInput, {
      sheetName: "Sheet1",
      range: "A1:B2",
      expand: "spill",
    })).toMatchObject({ expand: "spill" });
    expect(validateInput(OcrRecognizeInput, {
      mode: "invoice",
      filePaths: ["C:\\tmp\\invoice.pdf"],
    })).toMatchObject({ mode: "invoice" });
    expect(validateInput(LaunchOfficeApplicationInput, "powerpoint")).toBe("powerpoint");
  });

  it("rejects malformed structured IPC inputs", () => {
    expect(() => validateInput(FileWriteTempFileInput, { data: "" })).toThrow("IPC 参数校验失败");
    expect(() => validateInput(OcrRecognizeInput, { mode: "bad", filePaths: [] })).toThrow("IPC 参数校验失败");
    expect(() => validateInput(LaunchOfficeApplicationInput, "cmd")).toThrow("IPC 参数校验失败");
    expect(() => validateInput(SettingsGetInput, "arbitrary-secret-key")).toThrow("IPC 参数校验失败");
    expect(() => validateInput(AppLogInput, {
      level: "info",
      tag: "renderer",
      message: "x".repeat(50_001),
    })).toThrow("IPC 参数校验失败");
  });

  it("validates setting values according to their setting key", () => {
    expect(validateInput(SettingsSetInput, ["theme", "dark"])).toEqual(["theme", "dark"]);
    expect(validateInput(SettingsSetInput, ["windowOpacity", 0.75])).toEqual([
      "windowOpacity",
      0.75,
    ]);
    expect(() => validateInput(SettingsSetInput, ["theme", true])).toThrow("IPC 参数校验失败");
    expect(() => validateInput(SettingsSetInput, ["windowOpacity", 2])).toThrow("IPC 参数校验失败");
    expect(() => validateInput(SettingsSetInput, ["aiProviders", {
      provider_1: {
        id: "provider_1",
        name: "Provider",
        provider: "custom",
        apiKey: "secret",
        baseUrl: "https://example.test/v1",
        model: "model",
        unexpected: true,
      },
    }])).toThrow("IPC 参数校验失败");
  });

  it("accepts the persisted provider, compaction and pinned-folder shapes", () => {
    expect(() => validateInput(SettingsSetInput, ["aiProviders", {
      provider_1: {
        id: "provider_1",
        name: "Provider",
        provider: "custom",
        apiKey: "••••••••",
        baseUrl: "https://example.test/v1",
        model: "model-a",
        models: ["model-a", "model-b"],
        modelConfigs: [{
          name: "model-a",
          contextWindowSize: 128_000,
          compHash: "family-a",
          reasoningMode: "high",
        }],
        defaultBaseUrl: "https://example.test/v1",
        defaultModel: "model-a",
        apiFormat: "openai",
        customHeaders: { "x-api-key": "••••••••" },
        contextWindowSize: 128_000,
        compHash: "family-a",
        reasoningMode: "high",
      },
    }])).not.toThrow();
    expect(() => validateInput(SettingsSetInput, ["compactionConfig", {
      enabled: true,
      autoCompactThresholdPercent: 80,
      retainedUserMessageMaxTokens: 20_000,
      summaryRetryCount: 1,
      midTurnThresholdRatio: 0.9,
      compactionProvider: "remote",
      remoteCompactUrl: "https://compact.example.test/v2",
      remoteCompactApiKey: "remote-key",
      remoteCompactModel: "compact-model",
    }])).not.toThrow();
    expect(() => validateInput(SettingsSetInput, ["pinnedFolders", [{
      path: "C:\\workspace",
      name: "workspace",
      addedAt: Date.now(),
      pinnedFiles: ["C:\\workspace\\book.xlsx"],
    }]])).not.toThrow();
  });

  it("rejects oversized open JSON variable bags", () => {
    expect(() => validateInput(OfficeAutomationTemplateRunInput, {
      templateId: "3fbb8f2f-20e8-49ff-8150-a45789f4f624",
      variables: { rows: Array.from({ length: 20_001 }, () => 1) },
    })).toThrow("JSON 数组不能超过 20000 项");
  });

  it("rejects oversized chat, attachment and OCR requests", () => {
    const attachment = {
      filePath: "C:\\tmp\\file.pdf",
      fileName: "file.pdf",
      fileType: "document" as const,
      size: 1,
    };
    expect(() => validateInput(AgentStartTurnInput, {
      content: "x".repeat(IPC_MAX_CHAT_CONTENT_CHARS + 1),
    })).toThrow("IPC 参数校验失败");
    expect(() => validateInput(AgentStartTurnInput, {
      content: "ok",
      attachments: Array.from({ length: IPC_MAX_ATTACHMENTS + 1 }, () => attachment),
    })).toThrow("IPC 参数校验失败");
    expect(() => validateInput(OcrRecognizeInput, {
      filePaths: Array.from({ length: IPC_MAX_OCR_FILES + 1 }, (_, index) => `C:\\tmp\\${index}.pdf`),
    })).toThrow("IPC 参数校验失败");
  });

  it("rejects Excel matrices above the total cell budget", () => {
    const rows = Array.from(
      { length: Math.ceil((IPC_MAX_EXCEL_CELLS + 1) / 1_000) },
      () => Array.from({ length: 1_000 }, () => 1)
    );
    expect(() => validateInput(ExcelWriteRangeInput, {
      sheetName: "Sheet1",
      range: "A1",
      values: rows,
    })).toThrow(`${IPC_MAX_EXCEL_CELLS} 个单元格`);
  });

  it("estimates Base64 decoded bytes without allocating a Buffer", () => {
    expect(estimateBase64DecodedBytes("AAAA")).toBe(3);
    expect(estimateBase64DecodedBytes("TQ==")).toBe(1);
    expect(isBase64PayloadWithinLimit("AAAA", 3)).toBe(true);
    expect(isBase64PayloadWithinLimit("AAAA", 2)).toBe(false);
  });
});
