import { describe, expect, it } from "vitest";

import {
  AgentStartTurnInput,
  AgentInterruptInput,
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
