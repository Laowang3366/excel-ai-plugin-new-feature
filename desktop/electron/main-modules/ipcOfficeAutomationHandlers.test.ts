import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import type { OfficeActionBridge, OfficeDocumentManagerBridge } from "../agent/tools/contracts/office";
import { runOfficeWorkflow } from "../agent/tools/officeCore/workflow";
import { createOfficeAutomationService } from "./ipcOfficeAutomationHandlers";

const roots: string[] = [];

describe("Office automation IPC service", () => {
  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
  });

  it("lists and activates documents and objects through the shared document bridge", async () => {
    const documentBridge = createDocumentBridge();
    const service = createOfficeAutomationService({ getDataPath: () => "C:\\data", documentBridge });

    await service.listDocuments("excel");
    await service.activateDocument({ app: "excel", filePath: "C:\\book.xlsx", instanceId: "excel:1" });
    await service.listObjects({ app: "excel", filePath: "C:\\book.xlsx", instanceId: "excel:1" });
    await service.activateObject({ app: "excel", filePath: "C:\\book.xlsx", instanceId: "excel:1", locator: "sheet:Sheet1" });

    expect(documentBridge.listDocuments).toHaveBeenCalledWith("excel");
    expect(documentBridge.activateDocument).toHaveBeenCalledWith(expect.objectContaining({ instanceId: "excel:1" }));
    expect(documentBridge.activateObject).toHaveBeenCalledWith(expect.objectContaining({ locator: "sheet:Sheet1" }));
  });

  it("saves a template only from a persisted workflow and runs it with variables", async () => {
    const dataPath = await mkdtemp(path.join(os.tmpdir(), "office-ipc-service-"));
    roots.push(dataPath);
    const workflowRoot = path.join(dataPath, "office-automation", "workflows");
    const sourcePath = path.join(dataPath, "source.xlsx");
    await writeFile(sourcePath, "source", "utf8");
    const executeAction = vi.fn(async (input) => ({
      status: "done" as const,
      engine: "com" as const,
      ...input,
      summary: "done",
      changes: [],
    }));
    const actionBridge: OfficeActionBridge = { executeAction };
    const first = await runOfficeWorkflow(actionBridge, [
      { app: "excel", action: "edit", operation: "format", filePath: "{{vars.source}}" },
    ], { workflowRoot, variables: { source: sourcePath } });
    const service = createOfficeAutomationService({
      getDataPath: () => dataPath,
      documentBridge: createDocumentBridge(),
      createActionBridge: () => actionBridge,
    });

    const template = await service.saveTemplateFromWorkflow({ workflowId: first.workflowId!, name: "月报" });
    const run = await service.runTemplate(template.id, { source: sourcePath });

    expect((await service.listTemplates()).map((item) => item.name)).toEqual(["月报"]);
    expect(run.status).toBe("done");
    expect(executeAction).toHaveBeenLastCalledWith(expect.objectContaining({ filePath: sourcePath, transactionContext: "workflow" }));
  });
});

function createDocumentBridge(): OfficeDocumentManagerBridge {
  return {
    listDocuments: vi.fn(async () => []),
    activateDocument: vi.fn(async (input) => ({ ...input, name: "book.xlsx", fullName: input.filePath, index: 1, active: true, progId: "Excel.Application", host: "microsoft-office", instanceId: input.instanceId || "excel:1" })),
    listObjects: vi.fn(async () => []),
    activateObject: vi.fn(async (input) => ({ ...input, documentPath: input.filePath, kind: "sheet", name: "Sheet1", selected: true })),
    prepareTransaction: vi.fn(async () => []),
    restoreTransactionFiles: vi.fn(async () => []),
  };
}
