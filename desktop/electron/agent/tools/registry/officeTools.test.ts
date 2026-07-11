import { describe, expect, it, vi } from "vitest";
import { createToolExecutors } from "../executors/createToolExecutors";
import { ALL_TOOL_DEFINITIONS, TOOL_DEFINITIONS_MAP } from "./toolDefinitions";
import type { Retriever } from "../../knowledge/retriever";
import type {
  ExcelScriptBridge,
  ExcelUiBridge,
  ExcelVbaBridge,
  ExcelWorkbookBridge,
} from "../contracts/excel";
import type {
  OfficeActionBridge,
  OfficeScriptBridge,
  PresentationBridge,
  WordDocumentBridge,
} from "../contracts/office";
import type { OfficeActionResult } from "../officeCore/types";

type ObjectToolParameters = {
  required?: string[];
  properties: Record<string, { description?: string; enum?: string[] }>;
};

function fakeExcelBridge(): ExcelWorkbookBridge {
  return {
    detectStatus: vi.fn(async () => ({ connected: true, host: "excel", workbookName: "book.xlsx" })),
    connect: vi.fn(),
    selectHost: vi.fn(),
    isConnected: vi.fn(),
    getHostInfo: vi.fn(),
    inspectWorkbook: vi.fn(),
    readRange: vi.fn(),
    writeRange: vi.fn(),
    clearRange: vi.fn(),
    getSelection: vi.fn(),
    getSelectionAddress: vi.fn(),
    getFormulaContext: vi.fn(),
    sheetOperation: vi.fn(),
    openWorkbook: vi.fn(),
    createWorkbook: vi.fn(),
    saveWorkbook: vi.fn(),
    switchWorkbook: vi.fn(),
  } as unknown as ExcelWorkbookBridge;
}

function fakeVbaBridge(): ExcelVbaBridge {
  return {
    detectCapabilities: vi.fn(),
    runMacro: vi.fn(),
    writeModule: vi.fn(),
    executeCode: vi.fn(),
  };
}

function fakeScriptBridge(): ExcelScriptBridge {
  return {
    detectEnvironment: vi.fn(),
    executeScript: vi.fn(),
  };
}

function fakeUiBridge(): ExcelUiBridge {
  return {
    addControl: vi.fn(),
    removeControl: vi.fn(),
    listControls: vi.fn(),
    createForm: vi.fn(),
    addMenu: vi.fn(),
  };
}

describe("Office Word/PPT tool definitions", () => {
  it("preserves the full tool definition catalog", () => {
    const names = ALL_TOOL_DEFINITIONS.map((tool) => tool.name);

    expect(names).toEqual([
      "workbook.inspect",
      "range.read",
      "range.write",
      "range.clear",
      "selection.get",
      "formula.context",
      "vba.runMacro",
      "vba.writeModule",
      "formula.search",
      "sheet.operation",
      "script.detect",
      "script.execute",
      "ui.addControl",
      "ui.removeControl",
      "ui.listControls",
      "ui.createForm",
      "ui.addMenu",
      "file.getPaths",
      "workbook.open",
      "workbook.create",
      "workbook.save",
      "shell.execute",
      "python.execute",
      "workbook.switch",
      "knowledge.search",
      "knowledge.listSources",
      "knowledge.write",
      "knowledge.updateSource",
      "knowledge.deleteSource",
      "web.search",
      "ocr.parseDocument",
      "memory.write",
      "memory.search",
      "memory.list",
      "memory.delete",
      "office.connection.status",
      "word.open",
      "word.create",
      "word.inspect",
      "word.readText",
      "word.insertText",
      "word.insertHeading",
      "word.replaceText",
      "word.save",
      "presentation.open",
      "presentation.create",
      "presentation.inspect",
      "presentation.readSlide",
      "presentation.addSlide",
      "presentation.setShapeText",
      "presentation.replaceText",
      "presentation.save",
      "office.action.inspect",
      "office.action.apply",
      "office.action.validate",
      "office.script.execute",
    ]);
    expect(names).not.toContain("range_read");
    expect(names).not.toContain("office_action_apply");
  });

  it("requires explicit action for office.action.apply", () => {
    const applyTool = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === "office.action.apply");
    const parameters = applyTool?.parameters as ObjectToolParameters | undefined;

    expect(parameters?.required).toEqual(["app", "action", "operation"]);
    expect(parameters?.properties.action.description).toContain("必填");
  });

  it("documents snapshot as an approved apply operation", () => {
    const inspectTool = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === "office.action.inspect");
    const applyTool = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === "office.action.apply");
    const inspectParameters = inspectTool?.parameters as ObjectToolParameters | undefined;
    const applyParameters = applyTool?.parameters as ObjectToolParameters | undefined;

    expect(inspectParameters?.properties.operation.description).not.toContain("snapshot");
    expect(applyParameters?.properties.operation.description).toContain("snapshot");
  });

  it("exposes range.read spill expansion for dynamic array validation", () => {
    const readTool = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === "range.read");
    const parameters = readTool?.parameters as ObjectToolParameters | undefined;

    expect(parameters?.properties.expand.enum).toEqual(["none", "spill", "currentArray", "currentRegion"]);
    expect(readTool?.description).toContain('expand:"spill"');
  });

  it("describes knowledge.search as a scene-and-difficulty gated tool", () => {
    const knowledgeTool = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === "knowledge.search");
    const parameters = knowledgeTool?.parameters as ObjectToolParameters | undefined;

    expect(knowledgeTool?.description).toContain("判断场景难度");
    expect(knowledgeTool?.description).toContain("简单问答");
    expect(knowledgeTool?.description).toContain("中高复杂度");
    expect(parameters?.properties.query.description).toContain("场景摘要");
    expect(parameters?.properties.query.description).toContain("目标输出");
  });

  it("does not expose internal memory kinds in the memory.write schema", () => {
    const writeTool = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === "memory.write");
    const parameters = writeTool?.parameters as ObjectToolParameters | undefined;

    expect(parameters?.properties.kind.enum).toEqual([
      "preference",
      "constraint",
      "correction",
      "style_preference",
      "operation_preference",
      "file_impression",
    ]);
    expect(parameters?.properties.kind.enum).not.toContain("tool_success_profile");
  });

  it("requires memoryId for the memory.delete schema", () => {
    const deleteTool = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === "memory.delete");
    const parameters = deleteTool?.parameters as ObjectToolParameters | undefined;

    expect(parameters?.required).toEqual(["memoryId"]);
    expect(parameters?.properties.memoryId).toEqual({ type: "string" });
    expect(deleteTool?.description).toContain("memory.list");
    expect(deleteTool?.description).toContain("知识库");
  });

  it("registers first-class Word and PowerPoint editing tools", () => {
    const names = ALL_TOOL_DEFINITIONS.map((tool) => tool.name);

    expect(names).toEqual(expect.arrayContaining([
      "word.open",
      "word.create",
      "word.inspect",
      "word.readText",
      "word.insertText",
      "word.insertHeading",
      "word.replaceText",
      "word.save",
      "presentation.open",
      "presentation.create",
      "presentation.inspect",
      "presentation.readSlide",
      "presentation.addSlide",
      "presentation.setShapeText",
      "presentation.replaceText",
      "presentation.save",
      "office.action.inspect",
      "office.action.apply",
      "office.action.validate",
      "office.script.execute",
      "python.execute",
    ]));
  });
});

describe("Office Word/PPT tool executors", () => {
  it("registers core Excel, file, knowledge, Word, PowerPoint, and Office script executors", () => {
    const knowledgeRetriever = {
      search: vi.fn(),
      formatForToolResult: vi.fn(),
    } as unknown as Retriever;
    const wordBridge: WordDocumentBridge = {
      openDocument: vi.fn(),
      createDocument: vi.fn(),
      inspectDocument: vi.fn(),
      readText: vi.fn(),
      insertText: vi.fn(),
      insertHeading: vi.fn(),
      replaceText: vi.fn(),
      saveDocument: vi.fn(),
    };
    const presentationBridge: PresentationBridge = {
      openPresentation: vi.fn(),
      createPresentation: vi.fn(),
      inspectPresentation: vi.fn(),
      readSlide: vi.fn(),
      addSlide: vi.fn(),
      setShapeText: vi.fn(),
      replaceText: vi.fn(),
      savePresentation: vi.fn(),
    };
    const officeScriptBridge: OfficeScriptBridge = {
      executeScript: vi.fn(),
    };
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(),
    };

    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      "D:\\temp",
      knowledgeRetriever,
      wordBridge,
      presentationBridge,
      officeScriptBridge,
      officeActionBridge
    );

    const names = [...executors.keys()];
    expect(names).toEqual(expect.arrayContaining([
      "workbook.inspect",
      "workbook_inspect",
      "range.write",
      "range_write",
      "file.getPaths",
      "knowledge.search",
      "knowledge.write",
      "memory.delete",
      "memory_delete",
      "web.search",
      "ocr.parseDocument",
      "ocr_parseDocument",
      "python.execute",
      "python_execute",
      "office.connection.status",
      "office_connection_status",
      "word.replaceText",
      "word.insertHeading",
      "presentation.addSlide",
      "presentation.replaceText",
      "office.action.inspect",
      "office.action_inspect",
      "office_action_inspect",
      "office.action.apply",
      "office_action_apply",
      "office.action.validate",
      "office.script.execute",
    ]));
    expect(names).not.toEqual(expect.arrayContaining([
      "office.file.inspect",
      "office.file.replaceText",
      "office.layout.inspect",
      "office.table.inspect",
      "office.table.applyStyle",
      "office.visual.snapshot",
    ]));
  });

  it("maps underscore aliases back to canonical tool definitions", () => {
    expect(TOOL_DEFINITIONS_MAP.get("range_read")?.name).toBe("range.read");
    expect(TOOL_DEFINITIONS_MAP.get("range_read")?.riskLevel).toBe("safe");
    expect(TOOL_DEFINITIONS_MAP.get("office.action_apply")?.name).toBe("office.action.apply");
    expect(TOOL_DEFINITIONS_MAP.get("office_action_apply")?.name).toBe("office.action.apply");
  });

  it("detects Office connection status before choosing editing tools", async () => {
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge()
    );

    const result = await executors.get("office.connection.status")!.execute({
      app: "excel",
    });

    expect(result).toEqual({
      success: true,
      data: { connected: true, host: "excel", workbookName: "book.xlsx" },
    });
  });

  it("forwards Word tool calls to the Word bridge", async () => {
    const wordBridge: WordDocumentBridge = {
      openDocument: vi.fn(async () => ({ success: true, documentName: "demo.docx" })),
      createDocument: vi.fn(),
      inspectDocument: vi.fn(),
      readText: vi.fn(),
      insertText: vi.fn(),
      insertHeading: vi.fn(async () => ({ inserted: true, level: 2 })),
      replaceText: vi.fn(async () => ({ replacements: 2 })),
      saveDocument: vi.fn(),
    };
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      wordBridge
    );

    const result = await executors.get("word.replaceText")!.execute({
      findText: "旧标题",
      replaceText: "新标题",
    });

    expect(result).toEqual({ success: true, data: { replacements: 2 } });
    expect(wordBridge.replaceText).toHaveBeenCalledWith("旧标题", "新标题", undefined);

    const headingResult = await executors.get("word.insertHeading")!.execute({
      text: "阶段总结",
      level: 2,
      position: "start",
    });

    expect(headingResult).toEqual({ success: true, data: { inserted: true, level: 2 } });
    expect(wordBridge.insertHeading).toHaveBeenCalledWith("阶段总结", 2, "start");
  });

  it("forwards PowerPoint tool calls to the presentation bridge", async () => {
    const presentationBridge: PresentationBridge = {
      openPresentation: vi.fn(),
      createPresentation: vi.fn(),
      inspectPresentation: vi.fn(),
      readSlide: vi.fn(),
      addSlide: vi.fn(async () => ({ slideIndex: 3 })),
      setShapeText: vi.fn(),
      replaceText: vi.fn(async () => ({ replacements: 4 })),
      savePresentation: vi.fn(),
    };
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      undefined,
      presentationBridge
    );

    const result = await executors.get("presentation.addSlide")!.execute({
      title: "季度总结",
      body: "收入增长 18%",
    });

    expect(result).toEqual({ success: true, data: { slideIndex: 3 } });
    expect(presentationBridge.addSlide).toHaveBeenCalledWith("季度总结", "收入增长 18%", undefined);

    const replaceResult = await executors.get("presentation.replaceText")!.execute({
      findText: "旧产品名",
      replaceText: "新产品名",
      matchCase: true,
    });

    expect(replaceResult).toEqual({ success: true, data: { replacements: 4 } });
    expect(presentationBridge.replaceText).toHaveBeenCalledWith("旧产品名", "新产品名", true);
  });

  it("creates PowerPoint files through unified Office action before COM fallback", async () => {
    const presentationBridge: PresentationBridge = {
      openPresentation: vi.fn(),
      createPresentation: vi.fn(),
      inspectPresentation: vi.fn(),
      readSlide: vi.fn(),
      addSlide: vi.fn(),
      setShapeText: vi.fn(),
      replaceText: vi.fn(),
      savePresentation: vi.fn(),
    };
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async () => ({
        status: "done" as const,
        engine: "openxml" as const,
        app: "presentation" as const,
        action: "insert" as const,
        operation: "createPresentation",
        filePath: "D:\\docs\\talk.pptx",
        summary: "已创建 PPTX",
        changes: [],
      })),
    };
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      undefined,
      presentationBridge,
      undefined,
      officeActionBridge
    );

    const result = await executors.get("presentation.create")!.execute({
      filePath: "D:\\docs\\talk.pptx",
    });

    expect(result.success).toBe(true);
    expect(result.data).toMatchObject({ engine: "openxml", operation: "createPresentation" });
    expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
      app: "presentation",
      action: "insert",
      operation: "createPresentation",
      filePath: "D:\\docs\\talk.pptx",
    });
    expect(presentationBridge.createPresentation).not.toHaveBeenCalled();
  });

  it("creates Word files through unified Office action before COM fallback", async () => {
    const wordBridge: WordDocumentBridge = {
      openDocument: vi.fn(),
      createDocument: vi.fn(),
      inspectDocument: vi.fn(),
      readText: vi.fn(),
      insertText: vi.fn(),
      insertHeading: vi.fn(),
      replaceText: vi.fn(),
      saveDocument: vi.fn(),
    };
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async () => ({
        status: "done" as const,
        engine: "openxml" as const,
        app: "word" as const,
        action: "insert" as const,
        operation: "createDocument",
        filePath: "D:\\docs\\report.docx",
        summary: "已创建 DOCX",
        changes: [],
      })),
    };
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      wordBridge,
      undefined,
      undefined,
      officeActionBridge
    );

    const result = await executors.get("word.create")!.execute({
      filePath: "D:\\docs\\report.docx",
    });

    expect(result.success).toBe(true);
    expect(result.data).toMatchObject({ engine: "openxml", operation: "createDocument" });
    expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
      app: "word",
      action: "insert",
      operation: "createDocument",
      filePath: "D:\\docs\\report.docx",
      params: undefined,
    });
    expect(wordBridge.createDocument).not.toHaveBeenCalled();
  });

  it("falls back to Open XML inspection when PowerPoint app open fails", async () => {
    const presentationBridge: PresentationBridge = {
      openPresentation: vi.fn(async () => ({
        success: false,
        error: "PowerPoint 不可用",
      })),
      createPresentation: vi.fn(),
      inspectPresentation: vi.fn(),
      readSlide: vi.fn(),
      addSlide: vi.fn(),
      setShapeText: vi.fn(),
      replaceText: vi.fn(),
      savePresentation: vi.fn(),
    };
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async () => ({
        status: "done" as const,
        engine: "openxml" as const,
        app: "presentation" as const,
        action: "inspect" as const,
        operation: "inspectFile",
        filePath: "D:\\docs\\talk.pptx",
        summary: "已检查 PPTX 文件结构",
        changes: [],
      })),
    };
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      undefined,
      presentationBridge,
      undefined,
      officeActionBridge
    );

    const result = await executors.get("presentation.open")!.execute({
      filePath: "D:\\docs\\talk.pptx",
    });

    expect(result.success).toBe(false);
    expect(result.error).toBe("PowerPoint 不可用");
    expect(result.data).toMatchObject({
      success: false,
      fileReadable: true,
      openedInApp: false,
      fallback: "openxml",
      openError: "PowerPoint 不可用",
      inspection: { engine: "openxml", operation: "inspectFile" },
    });
    expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
      app: "presentation",
      action: "inspect",
      operation: "inspectFile",
      filePath: "D:\\docs\\talk.pptx",
    });
  });

  it("falls back to Open XML inspection when Word app open fails", async () => {
    const wordBridge: WordDocumentBridge = {
      openDocument: vi.fn(async () => ({
        success: false,
        error: "Word 不可用",
      })),
      createDocument: vi.fn(),
      inspectDocument: vi.fn(),
      readText: vi.fn(),
      insertText: vi.fn(),
      insertHeading: vi.fn(),
      replaceText: vi.fn(),
      saveDocument: vi.fn(),
    };
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async () => ({
        status: "done" as const,
        engine: "openxml" as const,
        app: "word" as const,
        action: "inspect" as const,
        operation: "inspectFile",
        filePath: "D:\\docs\\report.docx",
        summary: "已检查 DOCX 文件结构",
        changes: [],
      })),
    };
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      wordBridge,
      undefined,
      undefined,
      officeActionBridge
    );

    const result = await executors.get("word.open")!.execute({
      filePath: "D:\\docs\\report.docx",
    });

    expect(result.success).toBe(false);
    expect(result.error).toBe("Word 不可用");
    expect(result.data).toMatchObject({
      success: false,
      fileReadable: true,
      openedInApp: false,
      fallback: "openxml",
      openError: "Word 不可用",
      inspection: { engine: "openxml", operation: "inspectFile" },
    });
    expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
      app: "word",
      action: "inspect",
      operation: "inspectFile",
      filePath: "D:\\docs\\report.docx",
    });
  });

  it("forwards generic Office script calls to the Office script bridge", async () => {
    const officeScriptBridge: OfficeScriptBridge = {
      executeScript: vi.fn(async () => ({
        success: true,
        output: "OK",
        app: "word",
        engine: "powershell",
      })),
    };
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      undefined,
      undefined,
      officeScriptBridge
    );

    const result = await executors.get("office.script.execute")!.execute({
      app: "word",
      code: "$doc = $app.ActiveDocument",
    });

    expect(result).toEqual({
      success: true,
      data: {
        success: true,
        output: "OK",
        app: "word",
        engine: "powershell",
      },
    });
    expect(officeScriptBridge.executeScript).toHaveBeenCalledWith("word", "$doc = $app.ActiveDocument");
  });

  it("does not register legacy Open XML file executors", async () => {
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined
    );

    expect(executors.has("office.file.inspect")).toBe(false);
    expect(executors.has("office.file.replaceText")).toBe(false);
    expect(executors.has("office.layout.inspect")).toBe(false);
    expect(executors.has("office.table.inspect")).toBe(false);
    expect(executors.has("office.table.applyStyle")).toBe(false);
    expect(executors.has("office.visual.snapshot")).toBe(false);
  });

  it("forwards unified Office action calls to the Office action bridge", async () => {
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => {
        const result: OfficeActionResult = {
          status: "done",
          engine: "openxml",
          app: input.app,
          action: input.action,
          operation: input.operation,
          summary: "ok",
          changes: [],
        };
        return result;
      }),
    };
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      officeActionBridge
    );

    const result = await executors.get("office.action.apply")!.execute({
      app: "excel",
      action: "insert",
      operation: "insertChart",
      filePath: "D:\\docs\\book.xlsx",
      target: "range:Sheet1!A1:B5",
      params: { chartType: "column" },
    });

    expect(result.success).toBe(true);
    expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
      app: "excel",
      action: "insert",
      operation: "insertChart",
      filePath: "D:\\docs\\book.xlsx",
      target: "range:Sheet1!A1:B5",
      params: { chartType: "column" },
    });
  });

  it("rejects office.action.apply calls without an explicit action", async () => {
    const officeActionBridge: OfficeActionBridge = {
      executeAction: vi.fn(),
    };
    const executors = createToolExecutors(
      fakeExcelBridge(),
      fakeVbaBridge(),
      fakeScriptBridge(),
      fakeUiBridge(),
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      officeActionBridge
    );

    const result = await executors.get("office.action.apply")!.execute({
      app: "word",
      operation: "applyHeadingStyles",
      filePath: "D:\\docs\\report.docx",
    });

    expect(result).toEqual({
      success: false,
      error: "参数 action 必须是 inspect、edit、style、insert、snapshot 或 validate",
    });
    expect(officeActionBridge.executeAction).not.toHaveBeenCalled();
  });
});
