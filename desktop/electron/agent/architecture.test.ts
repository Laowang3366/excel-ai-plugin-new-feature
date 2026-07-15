import { existsSync, readFileSync } from "fs";
import path from "path";
import { describe, expect, it } from "vitest";

const agentRoot = path.resolve(__dirname);

function expectAgentPath(relativePath: string): void {
  expect(existsSync(path.join(agentRoot, relativePath)), relativePath).toBe(true);
}

describe("agent folder architecture", () => {
  it("classifies agent files into explicit architecture layers", () => {
    [
      "core/agentLoop/agentLoop.ts",
      "core/agentLoop/maxTokens.ts",
      "core/agentLoop/summaryGenerator.ts",
      "core/agentLoop/sessionCompactionConfig.ts",
      "core/agentLoop/turnState.ts",
      "core/agentLoop/threadLifecycle.ts",
      "core/agentLoop/turnRunner.ts",
      "interaction/README.md",
      "interaction/eventForwarder.ts",
      "interaction/ipcAgentHandlers.ts",
      "runtime/README.md",
      "runtime/agentRuntime.ts",
      "runtime/bridgeRegistry.ts",
      "runtime/compactionRuntime.ts",
      "runtime/knowledgeRuntime.ts",
      "memory/sessionStore.ts",
      "memory/compaction.ts",
      "knowledge/retriever.ts",
      "knowledge/knowledgeWriter.ts",
      "tools/contracts/README.md",
      "tools/contracts/excel.ts",
      "tools/contracts/office.ts",
      "tools/executors/README.md",
      "tools/executors/createToolExecutors.ts",
      "tools/executors/validation.ts",
      "tools/executors/excelExecutors.ts",
      "tools/executors/excelMacroExecutors.ts",
      "tools/executors/excelUiExecutors.ts",
      "tools/executors/fileExecutors.ts",
      "tools/executors/knowledgeExecutors.ts",
      "tools/executors/localDocumentParser.ts",
      "tools/executors/webSearchExecutors.ts",
      "tools/executors/ocrExecutors.ts",
      "tools/executors/officeExecutors.ts",
      "tools/registry/toolDefinitions.ts",
      "tools/registry/workbook.ts",
      "tools/registry/range.ts",
      "tools/registry/formula.ts",
      "tools/registry/sheet.ts",
      "tools/registry/macro.ts",
      "tools/registry/ui.ts",
      "tools/registry/file.ts",
      "tools/registry/knowledge.ts",
      "tools/registry/web.ts",
      "tools/registry/ocr.ts",
      "tools/registry/office.ts",
      "officeWorker/officeWorkerClient.ts",
      "officeWorker/dotNetOpenXmlBridge.ts",
      "officeWorker/dotNetOfficeActionBridge.ts",
      "providers/openaiCompatibleClient.ts",
      "prompts/promptComposer.ts",
      "prompts/systemPrompt.ts",
      "prompts/sections/folderContextPrompt.ts",
      "prompts/templates/system/base.zh-CN.md",
      "prompts/templates/system/security.zh-CN.md",
      "prompts/templates/scenarios/formula.zh-CN.md",
      "prompts/templates/scenarios/office-tools.zh-CN.md",
      "prompts/templates/scenarios/ocr-invoice.zh-CN.md",
      "prompts/templates/scenarios/general-office.zh-CN.md",
      "prompts/templates/runtime/environment.zh-CN.md",
      "prompts/templates/runtime/folder.zh-CN.md",
      "attachments/imageAttachmentResolver.ts",
      "shared/types.ts",
      "shared/messageBuilder.ts",
    ].forEach(expectAgentPath);
  });

  it("packages the WPS JSA add-in used for internal macro writes", () => {
    const addonRoot = path.resolve(agentRoot, "../../public/wps-jsa-bridge");
    ["index.html", "main.js", "manifest.xml", "ribbon.xml"].forEach((fileName) => {
      expect(existsSync(path.join(addonRoot, fileName)), fileName).toBe(true);
    });
  });

  it("documents architecture layer responsibilities with README files", () => {
    [
      "prompts/README.md",
      "memory/README.md",
      "knowledge/README.md",
      "providers/README.md",
      "tools/registry/README.md",
    ].forEach(expectAgentPath);
  });

  it("keeps layer README files aligned with current module paths", () => {
    const readmes = [
      "tools/contracts/README.md",
      "tools/executors/README.md",
      "tools/registry/README.md",
    ].map((relativePath) => readFileSync(path.join(agentRoot, relativePath), "utf8"));
    const text = readmes.join("\n");

    [
      "tools/registry/interfaces.ts",
      "tools/registry/executors.ts",
      "tools/sandbox",
      "后续拆分",
      "后续迁入",
    ].forEach((staleText) => {
      expect(text).not.toContain(staleText);
    });
  });

  it("keeps core agent loop independent from the tools layer", () => {
    const toolExecutor = readFileSync(
      path.join(agentRoot, "core/agentLoop/toolExecutor.ts"),
      "utf8",
    );

    expect(toolExecutor).not.toContain("tools/sandbox");
  });

  it("keeps AgentLoop runtime dependencies explicit", () => {
    const agentLoop = readFileSync(path.join(agentRoot, "core/agentLoop/agentLoop.ts"), "utf8");

    expect(agentLoop).not.toContain("await import(");
  });

  it("keeps system prompt prose in external templates", () => {
    const systemPrompt = readFileSync(path.join(agentRoot, "prompts/systemPrompt.ts"), "utf8");

    expect(systemPrompt).not.toContain("Office 连接预检铁律");
    expect(systemPrompt).not.toContain("场景化操作指南：公式助手");
    expect(systemPrompt).not.toContain("已开启，这是公式函数能力的权威先验");
    expect(systemPrompt).toContain("templates/system/base.zh-CN.md?raw");
    expect(systemPrompt).toContain("templates/scenarios/formula.zh-CN.md?raw");
    expect(systemPrompt).toContain("templates/runtime/dynamic-array-enabled.zh-CN.md?raw");
    expect(systemPrompt).toContain("resolvePromptScenarios");
  });

  it("keeps IPC outside concrete Excel bridge implementations", () => {
    const ipcHandlers = readFileSync(
      path.join(agentRoot, "../main-modules/ipcHandlers.ts"),
      "utf8",
    );

    expect(ipcHandlers).not.toContain("tools/implementations/excel");
    expect(ipcHandlers).not.toContain("new ExcelComBridge");
  });
});
