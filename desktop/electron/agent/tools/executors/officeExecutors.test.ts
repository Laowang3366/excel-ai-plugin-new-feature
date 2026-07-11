import { describe, expect, it, vi } from "vitest";

import type { ToolExecutor } from "../../shared/types";
import type {
  OfficeActionBridge,
  OfficeScriptBridge,
  PresentationBridge,
  WordDocumentBridge,
} from "../contracts/office";
import { addOfficeExecutors } from "./officeExecutors";

function createTarget(deps: Parameters<typeof addOfficeExecutors>[1]): Map<string, ToolExecutor> {
  const target = new Map<string, ToolExecutor>();
  addOfficeExecutors(target, deps);
  return target;
}

describe("addOfficeExecutors", () => {
  it("accepts common Office tool-name aliases emitted by some models", async () => {
    const excelBridge = {
      detectStatus: vi.fn(async () => ({ connected: false, host: "excel", version: "12.0" })),
    };
    const target = createTarget({ excelBridge: excelBridge as any });

    const result = await target.get("office.connection_status")!.execute({ app: "excel" });

    expect(result).toEqual({
      success: true,
      data: { connected: false, host: "excel" },
    });
    expect(excelBridge.detectStatus).toHaveBeenCalledTimes(1);
  });

  it("validates required Word arguments before calling the bridge", async () => {
    const wordBridge = {
      openDocument: vi.fn(),
    } as unknown as WordDocumentBridge;
    const target = createTarget({ wordBridge });

    const result = await target.get("word.open")!.execute({});

    expect(result).toMatchObject({
      success: false,
      error: "缺少必填参数: filePath",
    });
    expect(wordBridge.openDocument).not.toHaveBeenCalled();
  });

  it("routes Word insert heading arguments to the Word bridge", async () => {
    const wordBridge = {
      insertHeading: vi.fn(async () => ({ ok: true })),
    } as unknown as WordDocumentBridge;
    const target = createTarget({ wordBridge });

    const result = await target.get("word.insertHeading")!.execute({
      text: "季度报告",
      level: 2,
      position: "end",
    });

    expect(result).toEqual({ success: true, data: { ok: true } });
    expect(wordBridge.insertHeading).toHaveBeenCalledWith("季度报告", 2, "end");
  });

  it("validates PowerPoint slide index before reading slides", async () => {
    const presentationBridge = {
      readSlide: vi.fn(),
    } as unknown as PresentationBridge;
    const target = createTarget({ presentationBridge });

    const result = await target.get("presentation.readSlide")!.execute({});

    expect(result).toMatchObject({
      success: false,
      error: "缺少必填参数: slideIndex",
    });
    expect(presentationBridge.readSlide).not.toHaveBeenCalled();
  });

  it("routes Office action inspect through the unified action bridge", async () => {
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "done" as const,
        engine: "openxml" as const,
        app: input.app,
        action: input.action,
        operation: input.operation,
        summary: "已检查",
        changes: [],
      })),
    };
    const target = createTarget({ officeActionBridge });

    const result = await target.get("office.action.inspect")!.execute({
      app: "word",
      operation: "inspectFile",
      filePath: "C:/tmp/a.docx",
    });

    expect(result.success).toBe(true);
    expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
      app: "word",
      action: "inspect",
      operation: "inspectFile",
      filePath: "C:/tmp/a.docx",
    });
  });

  it.each([
    ["office.action.inspect", "excel", "writeRange"],
    ["office.action.validate", "word", "setHeaderFooter"],
    ["office.action.inspect", "presentation", "addSlides"],
    ["office.action.inspect", "presentation", "snapshot"],
  ])("%s rejects mutation operation %s/%s", async (toolName, app, operation) => {
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "done" as const,
        engine: "openxml" as const,
        app: input.app,
        action: input.action,
        operation: input.operation,
        summary: "不应执行",
        changes: [],
      })),
    };
    const target = createTarget({ officeActionBridge });

    const result = await target.get(toolName)!.execute({
      app,
      operation,
      filePath: "C:/tmp/input.office",
    });

    expect(result.success).toBe(false);
    expect(result.error).toContain("office.action.apply");
    expect(officeActionBridge.executeAction).not.toHaveBeenCalled();
  });

  it("reports non-done Office action statuses as unsuccessful so callers can fall back", async () => {
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "unsupported" as const,
        engine: "openxml" as const,
        app: input.app,
        action: input.action,
        operation: input.operation,
        summary: "Open XML 暂不支持该操作",
        changes: [],
      })),
    };
    const target = createTarget({ officeActionBridge });

    const result = await target.get("office.action.apply")!.execute({
      app: "presentation",
      action: "edit",
      operation: "editAnimationTimeline",
      filePath: "C:/tmp/a.pptx",
    });

    expect(result).toMatchObject({
      success: false,
      error: "Open XML 暂不支持该操作",
      data: {
        status: "unsupported",
        app: "presentation",
        action: "edit",
        operation: "editAnimationTimeline",
      },
    });
  });

  it("rejects unsupported Office script apps", async () => {
    const officeScriptBridge = {
      executeScript: vi.fn(),
    } as unknown as OfficeScriptBridge;
    const target = createTarget({ officeScriptBridge });

    const result = await target.get("office.script.execute")!.execute({
      app: "excel",
      code: "$app.ActiveWorkbook",
    });

    expect(result).toEqual({
      success: false,
      error: "参数 app 必须是 word 或 presentation",
    });
    expect(officeScriptBridge.executeScript).not.toHaveBeenCalled();
  });
});
