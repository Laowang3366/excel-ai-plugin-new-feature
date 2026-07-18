import { copyFile, mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import { describe, expect, it, vi } from "vitest";
import type { OfficeDocumentManagerBridge, OfficeFileBridge } from "../contracts/office";
import { createOfficeActionBridge } from "./officeActionAdapter";
import type { OfficeActionInput } from "./types";

describe("createOfficeActionBridge", () => {
  it("preserves structured Office Worker errors for capability handling", async () => {
    const error = Object.assign(new Error("当前 WPS 宿主不支持持久化 Power Query"), {
      code: "power_query_unavailable",
      details: { host: "wps" },
    });
    const bridge = createOfficeActionBridge({
      officeComActionBridge: {
        executeAction: vi.fn(async () => {
          throw error;
        }),
      },
    });

    const result = await bridge.executeAction({
      app: "excel",
      action: "edit",
      operation: "createPowerQuery",
      filePath: "C:/tmp/report.xlsx",
    });

    expect(result).toMatchObject({
      status: "failed",
      error: "当前 WPS 宿主不支持持久化 Power Query",
      data: {
        code: "power_query_unavailable",
        details: { host: "wps" },
      },
    });
  });

  it("evaluates explicit validation conditions instead of returning route success", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-validation-"));
    const filePath = path.join(tempDir, "report.docx");
    await writeFile(filePath, "fixture", "utf8");
    try {
      const officeFileBridge: OfficeFileBridge = {
        inspectFile: vi.fn(async () => ({ textPartCount: 2, textPreview: "季度销售报告" })),
        replaceText: vi.fn(),
        inspectLayout: vi.fn(),
        inspectTable: vi.fn(),
        applyTableStyle: vi.fn(),
        snapshot: vi.fn(),
      };
      const bridge = createOfficeActionBridge({ officeFileBridge });

      const passed = await bridge.executeAction({
        app: "word",
        action: "validate",
        operation: "inspectFile",
        filePath,
        params: { containsText: "销售报告", countPath: "textPartCount", expectedCount: 2 },
      });
      expect(passed.validation).toEqual(expect.objectContaining({ ok: true }));
      expect(passed.validation?.checks.map((check) => check.name)).toEqual([
        "file-exists",
        "contains-text",
        "expected-count",
      ]);

      const failed = await bridge.executeAction({
        app: "word",
        action: "validate",
        operation: "inspectFile",
        filePath,
        params: { countPath: "textPartCount", expectedCount: 3 },
      });
      expect(failed.status).toBe("done");
      expect(failed.validation).toEqual(expect.objectContaining({ ok: false }));
      expect(failed.summary).toBe("Office 验证未通过");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("routes file inspect and text replacement through unified Office actions", async () => {
    const officeFileBridge: OfficeFileBridge = {
      inspectFile: vi.fn(async () => ({
        engine: "openxml",
        operation: "inspect",
        documentType: "word",
        filePath: "D:\\docs\\report.docx",
        textPartCount: 1,
      })),
      replaceText: vi.fn(async () => ({
        engine: "openxml",
        operation: "replaceText",
        documentType: "word",
        filePath: "D:\\docs\\report.docx",
        outputPath: "D:\\docs\\report-edited.docx",
        replacements: 2,
        changedParts: [{ partName: "word/document.xml", replacements: 2 }],
      })),
      inspectLayout: vi.fn(),
      inspectTable: vi.fn(),
      applyTableStyle: vi.fn(),
      snapshot: vi.fn(),
    };
    const bridge = createOfficeActionBridge({ officeFileBridge });

    const inspectResult = await bridge.executeAction({
      app: "word",
      action: "inspect",
      operation: "inspectFile",
      filePath: "D:\\docs\\report.docx",
    });
    const replaceResult = await bridge.executeAction({
      app: "word",
      action: "edit",
      operation: "replaceText",
      filePath: "D:\\docs\\report.docx",
      params: { findText: "旧标题", replaceText: "新标题" },
    });

    expect(inspectResult.status).toBe("done");
    expect(replaceResult.status).toBe("done");
    expect(officeFileBridge.inspectFile).toHaveBeenCalledWith("D:\\docs\\report.docx");
    expect(officeFileBridge.replaceText).toHaveBeenCalledWith({
      filePath: "D:\\docs\\report.docx",
      findText: "旧标题",
      replaceText: "新标题",
      outputPath: undefined,
      matchCase: undefined,
    });
  });

  it("routes table style actions to the Open XML file bridge", async () => {
    const officeFileBridge: OfficeFileBridge = {
      inspectFile: vi.fn(),
      replaceText: vi.fn(),
      inspectLayout: vi.fn(),
      inspectTable: vi.fn(),
      applyTableStyle: vi.fn(async () => ({
        engine: "openxml",
        operation: "applyTableStyle",
        documentType: "spreadsheet",
        filePath: "D:\\docs\\book.xlsx",
        outputPath: "D:\\docs\\book-styled.xlsx",
        changedParts: ["xl/worksheets/sheet1.xml"],
      })),
      snapshot: vi.fn(),
    };

    const bridge = createOfficeActionBridge({ officeFileBridge });
    const result = await bridge.executeAction({
      app: "excel",
      action: "style",
      operation: "styleTable",
      filePath: "D:\\docs\\book.xlsx",
      target: "table:1",
      params: { style: "professional" },
    });

    expect(result.status).toBe("done");
    expect(result.engine).toBe("openxml");
    expect(result.validation).toBeUndefined();
    expect(officeFileBridge.applyTableStyle).toHaveBeenCalledWith({
      filePath: "D:\\docs\\book.xlsx",
      target: "table:1",
      style: "professional",
      outputPath: undefined,
    });
  });

  it.each([
    ["inspect", "excel", "writeRange"],
    ["validate", "word", "setHeaderFooter"],
    ["inspect", "presentation", "addSlides"],
    ["inspect", "presentation", "snapshot"],
  ] as const)("rejects %s action routing mutation %s/%s", async (action, app, operation) => {
    const officeComActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "done" as const,
        engine: "com" as const,
        app: input.app,
        action: input.action,
        operation: input.operation,
        summary: "不应执行",
        changes: [],
      })),
    };
    const bridge = createOfficeActionBridge({ officeComActionBridge } as any);

    const result = await bridge.executeAction({
      app,
      action,
      operation,
      preferEngine: "com",
      filePath: "D:\\docs\\input.office",
    });

    expect(result.status).toBe("failed");
    expect(result.error).toContain("office.action.apply");
    expect(officeComActionBridge.executeAction).not.toHaveBeenCalled();
  });

  it("returns needsCom when the Open XML snapshot bridge reports an unsupported placeholder", async () => {
    const officeFileBridge: OfficeFileBridge = {
      inspectFile: vi.fn(),
      replaceText: vi.fn(),
      inspectLayout: vi.fn(),
      inspectTable: vi.fn(),
      applyTableStyle: vi.fn(),
      snapshot: vi.fn(async () => ({
        engine: "openxml",
        operation: "snapshot",
        documentType: "word",
        filePath: "D:\\docs\\report.docx",
        outputPath: "D:\\docs\\report-snapshot.png",
        supported: false,
        error: "Open XML/headless 实际导出尚未接入",
      })),
    };

    const bridge = createOfficeActionBridge({ officeFileBridge });
    const result = await bridge.executeAction({
      app: "word",
      action: "edit",
      operation: "snapshot",
      filePath: "D:\\docs\\report.docx",
    });

    expect(result.status).toBe("needsCom");
    expect(result.summary).toContain("Open XML/headless 实际导出尚未接入");
    expect(officeFileBridge.snapshot).toHaveBeenCalled();
  });

  it("falls back to COM when Open XML reports needsCom", async () => {
    const officeComActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "done" as const,
        engine: "com" as const,
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: "COM 已插入图表",
        changes: [{ kind: "com-object", target: input.target, detail: "已创建图表" }],
      })),
    };

    const bridge = createOfficeActionBridge({ officeComActionBridge } as any);
    const result = await bridge.executeAction({
      app: "excel",
      action: "insert",
      operation: "insertChart",
      filePath: "D:\\docs\\book.xlsx",
      target: "range:Sheet1!A1:B5",
      params: { chartType: "column" },
    });

    expect(result.status).toBe("done");
    expect(result.engine).toBe("com");
    expect(officeComActionBridge.executeAction).toHaveBeenCalledWith({
      app: "excel",
      action: "insert",
      operation: "insertChart",
      filePath: "D:\\docs\\book.xlsx",
      target: "range:Sheet1!A1:B5",
      params: { chartType: "column" },
    });
  });

  it("routes directly to COM when preferEngine is com", async () => {
    const officeFileBridge: OfficeFileBridge = {
      inspectFile: vi.fn(),
      replaceText: vi.fn(),
      inspectLayout: vi.fn(),
      inspectTable: vi.fn(),
      applyTableStyle: vi.fn(),
      snapshot: vi.fn(),
    };
    const officeComActionBridge = {
      executeAction: vi.fn(async (input) => ({
        status: "done" as const,
        engine: "com" as const,
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: "COM 已生成快照",
        changes: [],
      })),
    };

    const bridge = createOfficeActionBridge({ officeFileBridge, officeComActionBridge } as any);
    const result = await bridge.executeAction({
      app: "presentation",
      action: "snapshot",
      operation: "snapshot",
      preferEngine: "com",
      filePath: "D:\\docs\\slides.pptx",
      outputPath: "D:\\docs\\slides.png",
      target: "slide:1",
    });

    expect(result.status).toBe("done");
    expect(result.engine).toBe("com");
    expect(officeFileBridge.snapshot).not.toHaveBeenCalled();
    expect(officeComActionBridge.executeAction).toHaveBeenCalled();
  });

  it.each([
    ["excel", "edit", "setDataValidation"],
    ["word", "style", "applyHeadingStyles"],
    ["presentation", "style", "applyTheme"],
    ["presentation", "edit", "deleteSlides"],
    ["presentation", "insert", "addSlide"],
  ] as const)(
    "routes %s advanced action %s through the .NET Open XML bridge",
    async (app, action, operation) => {
      const executeAction = vi.fn(async (input: OfficeActionInput) => ({
        status: "done" as const,
        engine: "openxml" as const,
        app: input.app,
        action: input.action,
        operation: input.operation,
        summary: "SDK action complete",
        changes: [],
      }));
      const officeFileBridge = { executeAction } as unknown as OfficeFileBridge;
      const bridge = createOfficeActionBridge({ officeFileBridge });
      const input = {
        app,
        action,
        operation,
        filePath: "D:\\docs\\input.office",
      };

      await expect(bridge.executeAction(input)).resolves.toMatchObject({
        status: "done",
        engine: "openxml",
      });
      expect(executeAction).toHaveBeenCalledWith(input);
    },
  );

  it.each([
    ["excel", "createWorkbook", ".xlsx"],
    ["word", "createDocument", ".docx"],
    ["presentation", "createPresentation", ".pptx"],
  ] as const)(
    "creates a new %s file through Open XML without backing up a nonexistent source",
    async (app, operation, extension) => {
      const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-create-action-"));
      try {
        const filePath = path.join(tempDir, `new-file${extension}`);
        const executeAction = vi.fn(async (input: OfficeActionInput) => {
          await writeFile(input.filePath!, "created", "utf8");
          return {
            status: "done" as const,
            engine: "openxml" as const,
            app: input.app,
            action: input.action,
            operation: input.operation,
            filePath: input.filePath,
            outputPath: input.filePath,
            summary: "created",
            changes: [],
          };
        });
        const bridge = createOfficeActionBridge({
          officeFileBridge: { executeAction } as unknown as OfficeFileBridge,
          backupRoot: path.join(tempDir, "backups"),
        });
        const input: OfficeActionInput = {
          app,
          action: "insert",
          operation,
          filePath,
        };

        await expect(bridge.executeAction(input)).resolves.toMatchObject({
          status: "done",
          engine: "openxml",
        });
        expect(executeAction).toHaveBeenCalledWith(input);
        expect(await readFile(filePath, "utf8")).toBe("created");
      } finally {
        await rm(tempDir, { recursive: true, force: true });
      }
    },
  );

  it("creates and restores a transaction backup for an in-place COM edit", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-action-transaction-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      const backupRoot = path.join(tempDir, "backups");
      await writeFile(filePath, "before", "utf8");
      const officeComActionBridge = {
        executeAction: vi.fn(async (input) => {
          await writeFile(input.filePath!, "after", "utf8");
          return {
            status: "done" as const,
            engine: "com" as const,
            app: input.app,
            action: input.action,
            operation: input.operation,
            filePath: input.filePath,
            outputPath: input.filePath,
            summary: "edited",
            changes: [],
          };
        }),
      };
      const bridge = createOfficeActionBridge({ officeComActionBridge, backupRoot });

      const edited = await bridge.executeAction({
        app: "excel",
        action: "style",
        operation: "formatChart",
        preferEngine: "com",
        filePath,
      });
      const transaction = (edited.data as { transaction: { backupPath: string } }).transaction;
      expect(await readFile(filePath, "utf8")).toBe("after");
      expect(transaction.backupPath).toContain(backupRoot);

      const restored = await bridge.executeAction({
        app: "excel",
        action: "edit",
        operation: "restoreBackup",
        filePath,
        params: { backupPath: transaction.backupPath },
      });
      expect(restored.status).toBe("done");
      expect(await readFile(filePath, "utf8")).toBe("before");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("does not create a transaction backup when writing an explicit output copy", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-action-copy-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      const outputPath = path.join(tempDir, "book-copy.xlsx");
      const backupRoot = path.join(tempDir, "backups");
      await writeFile(filePath, "source", "utf8");
      const officeComActionBridge = {
        executeAction: vi.fn(async (input) => ({
          status: "done" as const,
          engine: "com" as const,
          app: input.app,
          action: input.action,
          operation: input.operation,
          filePath: input.filePath,
          outputPath: input.outputPath,
          summary: "copied",
          changes: [],
        })),
      };
      const bridge = createOfficeActionBridge({ officeComActionBridge, backupRoot });

      const result = await bridge.executeAction({
        app: "excel",
        action: "style",
        operation: "formatChart",
        preferEngine: "com",
        filePath,
        outputPath,
      });
      const backups = await bridge.executeAction({
        app: "excel",
        action: "inspect",
        operation: "listBackups",
        filePath,
      });

      expect(result.data).toBeUndefined();
      expect(backups.data).toEqual({ records: [] });
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("restores every report target when a standalone incremental cross-office update fails", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-cross-transaction-"));
    try {
      const sourcePath = path.join(tempDir, "source.xlsx");
      const wordPath = path.join(tempDir, "report.docx");
      const presentationPath = path.join(tempDir, "report.pptx");
      const transactionRoot = path.join(tempDir, "transactions");
      await writeFile(sourcePath, "source", "utf8");
      await writeFile(wordPath, "word-before", "utf8");
      await writeFile(presentationPath, "ppt-before", "utf8");

      const officeDocumentBridge = createFileTransactionBridge();
      const officeComActionBridge = {
        executeAction: vi.fn(async () => {
          await writeFile(wordPath, "word-after", "utf8");
          await writeFile(presentationPath, "ppt-after", "utf8");
          return {
            status: "failed" as const,
            engine: "com" as const,
            app: "excel" as const,
            action: "insert" as const,
            operation: "buildReportPackage",
            filePath: sourcePath,
            summary: "模拟第二个输出失败",
            changes: [],
            error: "模拟第二个输出失败",
          };
        }),
      };
      const bridge = createOfficeActionBridge({
        officeComActionBridge,
        officeDocumentBridge,
        transactionRoot,
      });

      const result = await bridge.executeAction({
        app: "excel",
        action: "insert",
        operation: "buildReportPackage",
        filePath: sourcePath,
        params: {
          updateExisting: true,
          wordOutputPath: wordPath,
          presentationOutputPath: presentationPath,
          sections: [{ linkId: "sales", range: "A1:B3" }],
        },
      });

      expect(result.status).toBe("failed");
      expect(await readFile(wordPath, "utf8")).toBe("word-before");
      expect(await readFile(presentationPath, "utf8")).toBe("ppt-before");
      expect(officeDocumentBridge.prepareTransaction).toHaveBeenCalled();
      expect(officeDocumentBridge.restoreTransactionFiles).toHaveBeenCalled();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});

function createFileTransactionBridge(): OfficeDocumentManagerBridge {
  return {
    listDocuments: vi.fn(async () => []),
    activateDocument: vi.fn(async () => {
      throw new Error("not used");
    }),
    listObjects: vi.fn(async () => []),
    activateObject: vi.fn(async () => {
      throw new Error("not used");
    }),
    prepareTransaction: vi.fn(async () => []),
    restoreTransactionFiles: vi.fn(async (files) => {
      for (const file of files) {
        if (file.existed && file.snapshotPath) await copyFile(file.snapshotPath, file.filePath);
        else await rm(file.filePath, { force: true });
      }
      return [];
    }),
  };
}
