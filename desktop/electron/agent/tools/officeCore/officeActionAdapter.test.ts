import { copyFile, mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";
import type { OfficeDocumentManagerBridge, OfficeFileBridge } from "../contracts/office";
import { createOfficeActionBridge } from "./officeActionAdapter";

async function writeZip(filePath: string, files: Record<string, string>): Promise<void> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) zip.file(name, content);
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}

async function readZipText(filePath: string, partName: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(filePath));
  const part = zip.file(partName);
  if (!part) throw new Error(`missing ${partName}`);
  return part.async("text");
}

describe("createOfficeActionBridge", () => {
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
    expect(result.validation?.ok).toBe(true);
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

  it("routes Excel advanced actions to the Open XML implementation", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-action-excel-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      const outputPath = path.join(tempDir, "book-action.xlsx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "xl/worksheets/sheet1.xml": "<worksheet><sheetData /></worksheet>",
      });

      const bridge = createOfficeActionBridge({});
      const result = await bridge.executeAction({
        app: "excel",
        action: "edit",
        operation: "setDataValidation",
        filePath,
        outputPath,
        target: "range:Sheet1!A2:A10",
        params: { values: ["通过", "失败"] },
      });

      const sheetXml = await readZipText(outputPath, "xl/worksheets/sheet1.xml");
      expect(result.status).toBe("done");
      expect(sheetXml).toContain("<dataValidations");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("routes Word advanced actions to the Open XML implementation", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-action-word-"));
    try {
      const filePath = path.join(tempDir, "report.docx");
      const outputPath = path.join(tempDir, "report-action.docx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "word/document.xml": "<w:document><w:body><w:p><w:r><w:t>一、概览</w:t></w:r></w:p></w:body></w:document>",
      });

      const bridge = createOfficeActionBridge({});
      const result = await bridge.executeAction({
        app: "word",
        action: "style",
        operation: "applyHeadingStyles",
        filePath,
        outputPath,
        params: { startsWith: "一、", level: 1 },
      });

      const xml = await readZipText(outputPath, "word/document.xml");
      expect(result.status).toBe("done");
      expect(xml).toContain('w:val="Heading1"');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("routes presentation advanced actions to the Open XML implementation", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-action-ppt-"));
    try {
      const filePath = path.join(tempDir, "slides.pptx");
      const outputPath = path.join(tempDir, "slides-action.pptx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "ppt/slides/slide1.xml": '<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>标题</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
      });

      const bridge = createOfficeActionBridge({});
      const result = await bridge.executeAction({
        app: "presentation",
        action: "style",
        operation: "applyTheme",
        filePath,
        outputPath,
        params: { accentColor: "1F4E79" },
      });

      const xml = await readZipText(outputPath, "ppt/slides/slide1.xml");
      expect(result.status).toBe("done");
      expect(xml).toContain('val="1F4E79"');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("routes presentation slide deletion through the Open XML implementation", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-action-ppt-delete-"));
    try {
      const filePath = path.join(tempDir, "slides.pptx");
      const outputPath = path.join(tempDir, "slides-action.pptx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "ppt/presentation.xml": '<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst></p:presentation>',
        "ppt/_rels/presentation.xml.rels": '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Target="slides/slide2.xml"/></Relationships>',
        "ppt/slides/slide1.xml": "<p:sld>one</p:sld>",
        "ppt/slides/slide2.xml": "<p:sld>two</p:sld>",
      });

      const bridge = createOfficeActionBridge({});
      const result = await bridge.executeAction({
        app: "presentation",
        action: "edit",
        operation: "deleteSlides",
        filePath,
        outputPath,
        params: { from: 2, to: 2 },
      });

      const zip = await JSZip.loadAsync(await readFile(outputPath));
      expect(result.status).toBe("done");
      expect(result.engine).toBe("openxml");
      expect(zip.file("ppt/slides/slide1.xml")).toBeTruthy();
      expect(zip.file("ppt/slides/slide2.xml")).toBeNull();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("routes presentation addSlide through the Open XML implementation", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "office-action-ppt-add-"));
    try {
      const filePath = path.join(tempDir, "slides.pptx");
      await writeZip(filePath, {
        "[Content_Types].xml": [
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
          '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>',
          "</Types>",
        ].join(""),
        "ppt/presentation.xml": '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="wide"/></p:presentation>',
        "ppt/_rels/presentation.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
        "ppt/slides/slide1.xml": "<p:sld>one</p:sld>",
      });

      const bridge = createOfficeActionBridge({});
      const result = await bridge.executeAction({
        app: "presentation",
        action: "insert",
        operation: "addSlide",
        filePath,
        params: { title: "新增内容", body: "第一点\n第二点" },
      });

      const zip = await JSZip.loadAsync(await readFile(filePath));
      const presentationXml = await readZipText(filePath, "ppt/presentation.xml");
      const slideXml = await readZipText(filePath, "ppt/slides/slide2.xml");
      expect(result.status).toBe("done");
      expect(result.engine).toBe("openxml");
      expect(zip.file("ppt/slides/slide2.xml")).toBeTruthy();
      expect(presentationXml).toContain('r:id="rId2"');
      expect(slideXml).toContain("新增内容");
      expect(slideXml).toContain("第一点");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

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
      const bridge = createOfficeActionBridge({ officeComActionBridge, officeDocumentBridge, transactionRoot });

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
    activateDocument: vi.fn(async () => { throw new Error("not used"); }),
    listObjects: vi.fn(async () => []),
    activateObject: vi.fn(async () => { throw new Error("not used"); }),
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
