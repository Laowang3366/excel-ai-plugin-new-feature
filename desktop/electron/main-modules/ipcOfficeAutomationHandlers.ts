import path from "node:path";

import { ipcMain } from "electron";

import { getOrCreateOfficeBridges } from "../agent/runtime/bridgeRegistry";
import type { OfficeActionBridge, OfficeDocumentManagerBridge } from "../agent/tools/contracts/office";
import { OfficeComActionBridge } from "../agent/tools/implementations/office/officeComActionBridge";
import { OfficeDocumentComBridge } from "../agent/tools/implementations/office/officeDocumentComBridge";
import { createOfficeActionBridge } from "../agent/tools/officeCore/officeActionAdapter";
import { getOfficeTransaction, listOfficeTransactions, redoOfficeTransaction, undoOfficeTransaction } from "../agent/tools/officeCore/transactionJournal";
import { getOfficeWorkflow, listOfficeWorkflows, requestOfficeWorkflowCancellation, runOfficeWorkflow } from "../agent/tools/officeCore/workflow";
import {
  deleteOfficeWorkflowTemplate,
  getOfficeWorkflowTemplate,
  listOfficeWorkflowTemplates,
  saveOfficeWorkflowTemplate,
} from "../agent/tools/officeCore/workflowTemplates";
import type { OfficeActionApp } from "../agent/tools/officeCore/types";
import {
  OfficeAutomationDocumentInput,
  OfficeAutomationDocumentsListInput,
  OfficeAutomationForceInput,
  OfficeAutomationIdInput,
  OfficeAutomationObjectActivateInput,
  OfficeAutomationObjectsListInput,
  OfficeAutomationTemplateRunInput,
  OfficeAutomationTemplateSaveInput,
  validateInput,
} from "../shared/ipcSchemas";

interface OfficeAutomationServiceDeps {
  getDataPath: () => string;
  documentBridge?: OfficeDocumentManagerBridge;
  createActionBridge?: (transactionRoot: string, documentBridge: OfficeDocumentManagerBridge) => OfficeActionBridge;
}

export function createOfficeAutomationService(deps: OfficeAutomationServiceDeps) {
  const documentBridge = deps.documentBridge || new OfficeDocumentComBridge();
  const roots = () => {
    const automationRoot = path.join(deps.getDataPath(), "office-automation");
    return {
      workflowRoot: path.join(automationRoot, "workflows"),
      transactionRoot: path.join(automationRoot, "transactions"),
    };
  };
  const actionBridge = (transactionRoot: string) => deps.createActionBridge?.(transactionRoot, documentBridge)
    || createOfficeActionBridge({
      officeFileBridge: getOrCreateOfficeBridges().officeFileBridge,
      officeComActionBridge: new OfficeComActionBridge(),
      officeDocumentBridge: documentBridge,
      backupRoot: path.join(deps.getDataPath(), "office-backups"),
      transactionRoot,
    });
  const restoreOptions = (force = false) => ({
    force,
    prepareFiles: (filePaths: string[]) => documentBridge.prepareTransaction(filePaths),
    restoreFiles: (files: Parameters<OfficeDocumentManagerBridge["restoreTransactionFiles"]>[0]) => documentBridge.restoreTransactionFiles(files),
  });

  return {
    listDocuments: (app?: OfficeActionApp) => documentBridge.listDocuments(app),
    activateDocument: (input: { app: OfficeActionApp; filePath: string; instanceId?: string }) => documentBridge.activateDocument(input),
    listObjects: (input: { app: OfficeActionApp; filePath: string; instanceId?: string; kind?: string }) => documentBridge.listObjects(input),
    activateObject: (input: { app: OfficeActionApp; filePath: string; instanceId?: string; locator: string }) => documentBridge.activateObject(input),
    listWorkflows: () => listOfficeWorkflows(roots().workflowRoot),
    getWorkflow: (id: string) => getOfficeWorkflow(roots().workflowRoot, id),
    cancelWorkflow: (id: string) => requestOfficeWorkflowCancellation(roots().workflowRoot, id),
    resumeWorkflow: async (id: string) => {
      const { workflowRoot, transactionRoot } = roots();
      return runOfficeWorkflow(actionBridge(transactionRoot), [], {
        workflowRoot,
        transactionRoot,
        workflowId: id,
        resume: true,
        prepareTransaction: (filePaths) => documentBridge.prepareTransaction(filePaths),
        restoreTransaction: (files) => documentBridge.restoreTransactionFiles(files),
      });
    },
    listTemplates: () => listOfficeWorkflowTemplates(roots().workflowRoot),
    saveTemplateFromWorkflow: async (input: { workflowId: string; templateId?: string; name: string; description?: string }) => {
      const { workflowRoot } = roots();
      const workflow = await getOfficeWorkflow(workflowRoot, input.workflowId);
      return saveOfficeWorkflowTemplate({ root: workflowRoot, id: input.templateId, name: input.name, description: input.description, steps: workflow.sourceSteps || workflow.steps });
    },
    deleteTemplate: (id: string) => deleteOfficeWorkflowTemplate(roots().workflowRoot, id),
    runTemplate: async (id: string, variables?: Record<string, unknown>) => {
      const { workflowRoot, transactionRoot } = roots();
      const template = await getOfficeWorkflowTemplate(workflowRoot, id);
      return runOfficeWorkflow(actionBridge(transactionRoot), template.steps, {
        workflowRoot,
        transactionRoot,
        variables,
        prepareTransaction: (filePaths) => documentBridge.prepareTransaction(filePaths),
        restoreTransaction: (files) => documentBridge.restoreTransactionFiles(files),
      });
    },
    listTransactions: () => listOfficeTransactions(roots().transactionRoot),
    getTransaction: (id: string) => getOfficeTransaction(roots().transactionRoot, id),
    undoTransaction: (id: string, force = false) => undoOfficeTransaction(roots().transactionRoot, id, restoreOptions(force)),
    redoTransaction: (id: string, force = false) => redoOfficeTransaction(roots().transactionRoot, id, actionBridge(roots().transactionRoot), restoreOptions(force)),
  };
}

export function registerOfficeAutomationIpcHandlers(deps: OfficeAutomationServiceDeps): void {
  const service = createOfficeAutomationService(deps);
  const handle = <T>(channel: string, operation: (input: T) => Promise<unknown> | unknown) => {
    ipcMain.handle(channel, async (_event, input: T) => {
      try { return { success: true, data: await operation(input) }; }
      catch (error) { return { success: false, error: error instanceof Error ? error.message : String(error) }; }
    });
  };

  handle("office:automation:documents:list", (input: unknown) => service.listDocuments(validateInput(OfficeAutomationDocumentsListInput, input)?.app));
  handle("office:automation:documents:activate", (input: unknown) => service.activateDocument(validateInput(OfficeAutomationDocumentInput, input)));
  handle("office:automation:objects:list", (input: unknown) => service.listObjects(validateInput(OfficeAutomationObjectsListInput, input)));
  handle("office:automation:objects:activate", (input: unknown) => service.activateObject(validateInput(OfficeAutomationObjectActivateInput, input)));
  handle("office:automation:workflows:list", () => service.listWorkflows());
  handle("office:automation:workflows:get", (input: unknown) => service.getWorkflow(validateInput(OfficeAutomationIdInput, input).id));
  handle("office:automation:workflows:cancel", (input: unknown) => service.cancelWorkflow(validateInput(OfficeAutomationIdInput, input).id));
  handle("office:automation:workflows:resume", (input: unknown) => service.resumeWorkflow(validateInput(OfficeAutomationIdInput, input).id));
  handle("office:automation:templates:list", () => service.listTemplates());
  handle("office:automation:templates:saveFromWorkflow", (input: unknown) => service.saveTemplateFromWorkflow(validateInput(OfficeAutomationTemplateSaveInput, input)));
  handle("office:automation:templates:delete", (input: unknown) => service.deleteTemplate(validateInput(OfficeAutomationIdInput, input).id));
  handle("office:automation:templates:run", (input: unknown) => { const value = validateInput(OfficeAutomationTemplateRunInput, input); return service.runTemplate(value.templateId, value.variables); });
  handle("office:automation:transactions:list", () => service.listTransactions());
  handle("office:automation:transactions:get", (input: unknown) => service.getTransaction(validateInput(OfficeAutomationIdInput, input).id));
  handle("office:automation:transactions:undo", (input: unknown) => { const value = validateInput(OfficeAutomationForceInput, input); return service.undoTransaction(value.id, value.force); });
  handle("office:automation:transactions:redo", (input: unknown) => { const value = validateInput(OfficeAutomationForceInput, input); return service.redoTransaction(value.id, value.force); });
}
