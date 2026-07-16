import { access, mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { describe, expect, it, vi } from "vitest";

import type { ToolExecutor } from "../../shared/types";
import type {
  OfficeActionBridge,
  OfficeDocumentManagerBridge,
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

  it("reports the WPS regex dialect without exposing a host version", async () => {
    const excelBridge = {
      detectStatus: vi.fn(async () => ({ connected: true, host: "wps", version: "12.0" })),
    };
    const target = createTarget({ excelBridge: excelBridge as any });

    const result = await target.get("office.connection.status")!.execute({ app: "excel" });

    expect(result).toEqual({
      success: true,
      data: {
        connected: true,
        host: "wps",
        formulaDialect: {
          regexFunction: "REGEXP",
          guidance:
            "WPS 正则提取使用 REGEXP；不要使用 Excel 方言的 REGEXEXTRACT/REGEXREPLACE/REGEXTEST 函数名",
        },
      },
    });
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

  it("falls back to Open XML inspection when Word cannot open the file", async () => {
    const wordBridge = {
      openDocument: vi.fn(async () => ({ success: false, error: "Word 不可用" })),
    } as unknown as WordDocumentBridge;
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "done",
        engine: "openxml",
        ...input,
        summary: "已检查 DOCX 文件结构",
        changes: [],
      })),
    };
    const target = createTarget({ wordBridge, officeActionBridge });

    const result = await target.get("word.open")!.execute({ filePath: "D:/docs/report.docx" });

    expect(result).toMatchObject({
      success: false,
      error: "Word 不可用",
      data: { fileReadable: true, openedInApp: false, fallback: "openxml" },
    });
    expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
      app: "word",
      action: "inspect",
      operation: "inspectFile",
      filePath: "D:/docs/report.docx",
    });
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

  it("falls back to Open XML inspection when PowerPoint cannot open the file", async () => {
    const presentationBridge = {
      openPresentation: vi.fn(async () => ({ success: false, error: "PowerPoint 不可用" })),
    } as unknown as PresentationBridge;
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "done",
        engine: "openxml",
        ...input,
        summary: "已检查 PPTX 文件结构",
        changes: [],
      })),
    };
    const target = createTarget({ presentationBridge, officeActionBridge });

    const result = await target
      .get("presentation.open")!
      .execute({ filePath: "D:/docs/talk.pptx" });

    expect(result).toMatchObject({
      success: false,
      error: "PowerPoint 不可用",
      data: { fileReadable: true, openedInApp: false, fallback: "openxml" },
    });
    expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
      app: "presentation",
      action: "inspect",
      operation: "inspectFile",
      filePath: "D:/docs/talk.pptx",
    });
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

  it("returns the concrete Office action error instead of the generic summary", async () => {
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "failed" as const,
        engine: "com" as const,
        app: input.app,
        action: input.action,
        operation: input.operation,
        summary: "Office action 执行失败",
        changes: [],
        error: "当前 WPS 宿主不支持持久化 Power Query",
        data: { code: "power_query_unavailable" },
      })),
    };
    const target = createTarget({ officeActionBridge });

    const result = await target.get("office.action.apply")!.execute({
      app: "excel",
      action: "edit",
      operation: "createPowerQuery",
      filePath: "C:/tmp/a.xlsx",
      params: {
        advancedIntent: "refreshable-etl",
        sourceKind: "external",
        name: "Query1",
        mFormula: "let Source = 1 in Source",
        loadMode: "connectionOnly",
      },
    });

    expect(result).toMatchObject({
      success: false,
      error: "当前 WPS 宿主不支持持久化 Power Query",
      data: { data: { code: "power_query_unavailable" } },
    });
  });

  it("requires filePath for file-level office.action.apply calls", async () => {
    const officeActionBridge: OfficeActionBridge = { executeAction: vi.fn() };
    const target = createTarget({ officeActionBridge });

    const result = await target.get("office.action.apply")!.execute({
      app: "word",
      action: "edit",
      operation: "replaceText",
      params: { findText: "A", replaceText: "B" },
    });

    expect(result).toEqual({ success: false, error: "缺少必填参数: filePath" });
    expect(officeActionBridge.executeAction).not.toHaveBeenCalled();
  });

  it.each([
    {
      operation: "createPowerQuery",
      target: undefined,
      params: { name: "SimpleWrite", mFormula: "let Source = #table({}, {}) in Source" },
      expected: "refreshable-etl",
    },
    {
      operation: "createPivotTable",
      target: "range:Sheet1!A1:B10",
      params: { rowFields: ["Department"] },
      expected: "interactive-pivot",
    },
  ])(
    "rejects advanced Excel operation $operation without an explicit advanced intent",
    async ({ operation, target: actionTarget, params, expected }) => {
      const officeActionBridge: OfficeActionBridge = { executeAction: vi.fn() };
      const target = createTarget({ officeActionBridge });

      const result = await target.get("office.action.apply")!.execute({
        app: "excel",
        action: "insert",
        operation,
        filePath: "C:/tmp/book.xlsx",
        target: actionTarget,
        params,
      });

      expect(result).toMatchObject({ success: false });
      expect(result.error).toContain(expected);
      expect(officeActionBridge.executeAction).not.toHaveBeenCalled();
    },
  );

  it("allows a fully declared external refreshable Power Query request", async () => {
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "done",
        engine: "com",
        ...input,
        summary: "done",
        changes: [],
      })),
    };
    const target = createTarget({ officeActionBridge });
    const params = {
      advancedIntent: "refreshable-etl",
      sourceKind: "external",
      name: "SalesImport",
      mFormula: 'let Source = Csv.Document(File.Contents("C:/data/sales.csv")) in Source',
      loadMode: "worksheet",
      destination: "QueryOutput!A1",
    };

    const result = await target.get("office.action.apply")!.execute({
      app: "excel",
      action: "edit",
      operation: "createPowerQuery",
      filePath: "C:/tmp/book.xlsx",
      params,
    });

    expect(result.success).toBe(true);
    expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
      app: "excel",
      action: "edit",
      operation: "createPowerQuery",
      filePath: "C:/tmp/book.xlsx",
      params,
    });
  });

  it("does not expose an arbitrary PowerShell Office script executor", () => {
    expect(createTarget({}).has("office.script.execute")).toBe(false);
  });

  it("lists and activates explicit Office document windows", async () => {
    const officeDocumentBridge: OfficeDocumentManagerBridge = {
      listDocuments: vi.fn(async () => [
        {
          app: "word" as const,
          name: "report.docx",
          fullName: "C:/tmp/report.docx",
          index: 1,
          active: true,
          progId: "Word.Application",
          host: "microsoft-office" as const,
          instanceId: "word:100:200",
        },
      ]),
      activateDocument: vi.fn(async (input) => ({
        app: input.app,
        name: input.name || "report.docx",
        fullName: input.filePath,
        index: input.index || 1,
        active: true,
        progId: "Word.Application",
        host: "microsoft-office" as const,
        instanceId: input.instanceId || "word:100:200",
      })),
      listObjects: vi.fn(async () => [
        {
          app: "word" as const,
          documentPath: "C:/tmp/report.docx",
          kind: "bookmark",
          name: "Summary",
          locator: "bookmark:Summary",
        },
      ]),
      activateObject: vi.fn(async (input) => ({
        app: input.app,
        filePath: input.filePath,
        kind: "document",
        name: input.locator,
        locator: input.locator,
        active: true,
      })),
      prepareTransaction: vi.fn(async () => []),
      restoreTransactionFiles: vi.fn(async () => []),
    };
    const target = createTarget({ officeDocumentBridge });

    const listed = await target.get("office.documents.list")!.execute({ app: "word" });
    const activated = await target.get("office.documents.activate")!.execute({
      app: "word",
      filePath: "C:/tmp/report.docx",
    });
    const objects = await target.get("office.objects.list")!.execute({
      app: "word",
      filePath: "C:/tmp/report.docx",
      kind: "bookmark",
    });
    const selected = await target.get("office.objects.activate")!.execute({
      app: "word",
      filePath: "C:/tmp/report.docx",
      locator: "bookmark:Summary",
    });

    expect(listed).toMatchObject({ success: true, data: { count: 1 } });
    expect(activated).toMatchObject({ success: true, data: { active: true } });
    expect(objects).toMatchObject({ success: true, data: { count: 1 } });
    expect(selected).toMatchObject({ success: true, data: { locator: "bookmark:Summary" } });
    expect(officeDocumentBridge.activateDocument).toHaveBeenCalledWith({
      app: "word",
      filePath: "C:/tmp/report.docx",
      name: undefined,
      index: undefined,
      instanceId: undefined,
    });
    expect(officeDocumentBridge.listObjects).toHaveBeenCalledWith({
      app: "word",
      filePath: "C:/tmp/report.docx",
      instanceId: undefined,
      kind: "bookmark",
    });
  });

  it("runs validated multi-step Office workflows through the action bridge", async () => {
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "done",
        engine: "com",
        ...input,
        summary: "done",
        changes: [],
      })),
    };
    const target = createTarget({ officeActionBridge });

    const result = await target.get("office.workflow.run")!.execute({
      steps: [
        {
          app: "excel",
          action: "style",
          operation: "applyWorkbookTemplate",
          filePath: "C:/tmp/book.xlsx",
        },
      ],
    });

    expect(result).toMatchObject({ success: true, data: { status: "done", completedSteps: 1 } });
    expect(officeActionBridge.executeAction).toHaveBeenCalledTimes(1);
  });

  it("saves, lists, runs, and deletes reusable Office workflow templates", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-template-executors-"));
    try {
      const officeActionBridge: OfficeActionBridge = {
        executeAction: vi.fn(async (input) => ({
          status: "done",
          engine: "com",
          ...input,
          summary: "done",
          changes: [],
        })),
      };
      const target = createTarget({ officeActionBridge, workflowRoot: root });
      const saved = await target.get("office.workflow.template.save")!.execute({
        name: "月报",
        steps: [
          {
            app: "excel",
            action: "inspect",
            operation: "inspectCharts",
            filePath: "{{vars.source}}",
          },
        ],
      });
      const templateId = (saved.data as { id: string }).id;
      const listed = await target.get("office.workflow.template.list")!.execute({});
      const run = await target.get("office.workflow.run")!.execute({
        templateId,
        variables: { source: "C:/tmp/monthly.xlsx" },
      });
      const deleted = await target.get("office.workflow.template.delete")!.execute({ templateId });

      expect(saved).toMatchObject({ success: true, data: { name: "月报" } });
      expect(listed).toMatchObject({ success: true, data: { count: 1 } });
      expect(run).toMatchObject({ success: true, data: { status: "done" } });
      expect(officeActionBridge.executeAction).toHaveBeenCalledWith(
        expect.objectContaining({ filePath: "C:/tmp/monthly.xlsx" }),
      );
      expect(deleted).toEqual({ success: true, data: { deleted: true } });
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it("exposes persistent workflow status and group transaction undo/redo", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-executors-"));
    try {
      const workflowRoot = path.join(root, "workflows");
      const transactionRoot = path.join(root, "transactions");
      const sourcePath = path.join(root, "book.xlsx");
      const outputPath = path.join(root, "report.docx");
      await writeFile(sourcePath, "source", "utf8");
      const officeActionBridge: OfficeActionBridge = {
        executeAction: vi.fn(async (input) => {
          await writeFile(input.outputPath!, "report", "utf8");
          return {
            status: "done",
            engine: "com",
            ...input,
            summary: "done",
            changes: [{ kind: "create", target: input.outputPath, detail: "生成报告" }],
          };
        }),
      };
      const target = createTarget({ officeActionBridge, workflowRoot, transactionRoot });
      const run = await target.get("office.workflow.run")!.execute({
        steps: [
          {
            app: "word",
            action: "insert",
            operation: "buildReport",
            filePath: sourcePath,
            outputPath,
          },
        ],
      });
      const data = run.data as { workflowId: string; transactionId: string };

      const status = await target
        .get("office.workflow.status")!
        .execute({ workflowId: data.workflowId });
      const inspected = await target
        .get("office.transaction.inspect")!
        .execute({ transactionId: data.transactionId });
      const undone = await target
        .get("office.transaction.undo")!
        .execute({ transactionId: data.transactionId });
      await expect(access(outputPath)).rejects.toThrow();
      const redone = await target
        .get("office.transaction.redo")!
        .execute({ transactionId: data.transactionId });

      expect(run).toMatchObject({ success: true, data: { status: "done", completedSteps: 1 } });
      expect(status).toMatchObject({ success: true, data: { status: "done", nextStep: 2 } });
      expect(inspected).toMatchObject({
        success: true,
        data: { status: "applied", artifacts: [outputPath] },
      });
      expect(undone).toMatchObject({ success: true, data: { status: "undone" } });
      expect(redone).toMatchObject({ success: true, data: { status: "applied" } });
      await expect(access(outputPath)).resolves.toBeUndefined();
      expect(officeActionBridge.executeAction).toHaveBeenCalledTimes(1);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });
});
