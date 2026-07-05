import { describe, expect, it } from "vitest";

import {
  AgentInterruptInput,
  ExcelReadRangeInput,
  FileWriteTempFileInput,
  OcrRecognizeInput,
  SandboxUserRulesInput,
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
  });

  it("rejects malformed structured IPC inputs", () => {
    expect(() => validateInput(FileWriteTempFileInput, { data: "" })).toThrow("IPC 参数校验失败");
    expect(() => validateInput(SandboxUserRulesInput, [{ decision: "allow" }])).toThrow("IPC 参数校验失败");
    expect(() => validateInput(OcrRecognizeInput, { mode: "bad", filePaths: [] })).toThrow("IPC 参数校验失败");
  });
});
