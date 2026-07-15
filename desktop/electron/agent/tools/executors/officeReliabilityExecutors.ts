import type { ToolExecutor } from "../../shared/types";
import type { OfficeActionBridge, OfficeDocumentManagerBridge } from "../contracts/office";
import {
  getOfficeTransaction,
  listOfficeTransactions,
  redoOfficeTransaction,
  undoOfficeTransaction,
} from "../officeCore/transactionJournal";
import { getOfficeWorkflow, listOfficeWorkflows, requestOfficeWorkflowCancellation, runOfficeWorkflow } from "../officeCore/workflow";
import type { OfficeWorkflowCondition, OfficeWorkflowStepInput } from "../officeCore/workflowStepExecution";
import {
  deleteOfficeWorkflowTemplate,
  getOfficeWorkflowTemplate,
  listOfficeWorkflowTemplates,
  saveOfficeWorkflowTemplate,
} from "../officeCore/workflowTemplates";
import type { OfficeActionApp, OfficeActionEngine, OfficeActionKind } from "../officeCore/types";
import { officeAdvancedOperationError } from "../officeCore/operationPolicy";

export interface OfficeReliabilityExecutorDeps {
  officeActionBridge?: OfficeActionBridge;
  officeDocumentBridge?: OfficeDocumentManagerBridge;
  workflowRoot?: string;
  transactionRoot?: string;
}

export function addOfficeReliabilityExecutors(
  target: Map<string, ToolExecutor>,
  deps: OfficeReliabilityExecutorDeps,
): void {
  const { officeActionBridge, officeDocumentBridge, workflowRoot, transactionRoot } = deps;
  if (officeActionBridge) {
    target.set("office.workflow.run", {
      name: "office.workflow.run",
      execute: async (args) => {
        const resume = args.resume === true;
        let rawSteps = args.steps;
        if (!resume && typeof args.templateId === "string") {
          if (!workflowRoot) return { success: false, error: "Office 工作流目录未配置" };
          try { rawSteps = (await getOfficeWorkflowTemplate(workflowRoot, args.templateId)).steps; }
          catch (error) { return { success: false, error: errorMessage(error) }; }
        }
        const steps = rawSteps === undefined && resume ? [] : parseWorkflowSteps(rawSteps);
        if (typeof steps === "string") return { success: false, error: steps };
        const workflowId = typeof args.workflowId === "string" ? args.workflowId : undefined;
        if (resume && !workflowId) return { success: false, error: "继续工作流需要 workflowId" };
        const result = await runOfficeWorkflow(officeActionBridge, steps, {
          workflowRoot,
          transactionRoot,
          workflowId,
          resume,
          recoverRunning: args.recoverRunning === true,
          leaseMs: typeof args.leaseMs === "number" ? args.leaseMs : undefined,
          failureMode: args.failureMode === "rollback" ? "rollback" : workflowRoot ? "pause" : "rollback",
          cancellationMode: args.cancellationMode === "rollback" ? "rollback" : "pause",
          variables: args.variables && typeof args.variables === "object" && !Array.isArray(args.variables)
            ? args.variables as Record<string, unknown>
            : undefined,
          prepareTransaction: officeDocumentBridge
            ? (filePaths) => officeDocumentBridge.prepareTransaction(filePaths)
            : undefined,
          restoreTransaction: officeDocumentBridge
            ? (files) => officeDocumentBridge.restoreTransactionFiles(files)
            : undefined,
        });
        return {
          success: result.status === "done",
          data: result,
          ...(result.status === "done" ? {} : { error: result.error || (result.status === "paused" ? "Office 工作流已暂停，可从失败步骤继续" : "Office 工作流执行失败") }),
        };
      },
    });

    target.set("office.workflow.status", {
      name: "office.workflow.status",
      execute: async (args) => {
        if (!workflowRoot) return { success: false, error: "Office 工作流目录未配置" };
        if (typeof args.workflowId === "string") {
          try { return { success: true, data: await getOfficeWorkflow(workflowRoot, args.workflowId) }; }
          catch (error) { return { success: false, error: errorMessage(error) }; }
        }
        const records = await listOfficeWorkflows(workflowRoot);
        return { success: true, data: { workflows: records, count: records.length } };
      },
    });

    target.set("office.workflow.cancel", {
      name: "office.workflow.cancel",
      execute: async (args) => {
        if (!workflowRoot) return { success: false, error: "Office 工作流目录未配置" };
        if (typeof args.workflowId !== "string") return { success: false, error: "缺少必填参数: workflowId" };
        try { return { success: true, data: await requestOfficeWorkflowCancellation(workflowRoot, args.workflowId) }; }
        catch (error) { return { success: false, error: errorMessage(error) }; }
      },
    });

    target.set("office.workflow.template.list", {
      name: "office.workflow.template.list",
      execute: async () => {
        if (!workflowRoot) return { success: false, error: "Office 工作流目录未配置" };
        const templates = await listOfficeWorkflowTemplates(workflowRoot);
        return { success: true, data: { templates, count: templates.length } };
      },
    });

    target.set("office.workflow.template.save", {
      name: "office.workflow.template.save",
      execute: async (args) => {
        if (!workflowRoot) return { success: false, error: "Office 工作流目录未配置" };
        if (typeof args.name !== "string") return { success: false, error: "缺少必填参数: name" };
        const steps = parseWorkflowSteps(args.steps);
        if (typeof steps === "string") return { success: false, error: steps };
        try {
          const template = await saveOfficeWorkflowTemplate({
            root: workflowRoot,
            id: typeof args.templateId === "string" ? args.templateId : undefined,
            name: args.name,
            description: typeof args.description === "string" ? args.description : undefined,
            steps,
          });
          return { success: true, data: template };
        } catch (error) { return { success: false, error: errorMessage(error) }; }
      },
    });

    target.set("office.workflow.template.delete", {
      name: "office.workflow.template.delete",
      execute: async (args) => {
        if (!workflowRoot) return { success: false, error: "Office 工作流目录未配置" };
        if (typeof args.templateId !== "string") return { success: false, error: "缺少必填参数: templateId" };
        try { return { success: true, data: { deleted: await deleteOfficeWorkflowTemplate(workflowRoot, args.templateId) } }; }
        catch (error) { return { success: false, error: errorMessage(error) }; }
      },
    });

    target.set("office.transaction.list", {
      name: "office.transaction.list",
      execute: async () => {
        if (!transactionRoot) return { success: false, error: "Office 事务目录未配置" };
        const records = await listOfficeTransactions(transactionRoot);
        return { success: true, data: { transactions: records, count: records.length } };
      },
    });

    target.set("office.transaction.inspect", {
      name: "office.transaction.inspect",
      execute: async (args) => transactionRecordResult(transactionRoot, args.transactionId, "inspect"),
    });

    target.set("office.transaction.undo", {
      name: "office.transaction.undo",
      execute: async (args) => transactionRecordResult(transactionRoot, args.transactionId, "undo", args.force === true),
    });

    target.set("office.transaction.redo", {
      name: "office.transaction.redo",
      execute: async (args) => {
        if (!transactionRoot) return { success: false, error: "Office 事务目录未配置" };
        if (typeof args.transactionId !== "string") return { success: false, error: "缺少必填参数: transactionId" };
        try {
          const record = await redoOfficeTransaction(
            transactionRoot,
            args.transactionId,
            officeActionBridge,
            transactionRestoreOptions(args.force === true),
          );
          return { success: record.status === "applied", data: record, ...(record.status === "conflicted" ? { error: record.error } : {}) };
        }
        catch (error) { return { success: false, error: errorMessage(error) }; }
      },
    });
  }

  if (officeDocumentBridge) {
    target.set("office.documents.list", {
      name: "office.documents.list",
      execute: async (args) => {
        if (args.app !== undefined && !isOfficeActionApp(args.app)) return { success: false, error: "参数 app 必须是 excel、word 或 presentation" };
        const documents = await officeDocumentBridge.listDocuments(args.app as OfficeActionApp | undefined);
        return { success: true, data: { documents, count: documents.length } };
      },
    });

    target.set("office.documents.activate", {
      name: "office.documents.activate",
      execute: async (args) => {
        if (!isOfficeActionApp(args.app)) return { success: false, error: "参数 app 必须是 excel、word 或 presentation" };
        if (typeof args.filePath !== "string" && typeof args.name !== "string" && typeof args.index !== "number") return { success: false, error: "需要 filePath、name 或 index 之一" };
        const document = await officeDocumentBridge.activateDocument({
          app: args.app,
          filePath: typeof args.filePath === "string" ? args.filePath : undefined,
          name: typeof args.name === "string" ? args.name : undefined,
          index: typeof args.index === "number" ? args.index : undefined,
          instanceId: typeof args.instanceId === "string" ? args.instanceId : undefined,
        });
        return { success: true, data: document };
      },
    });

    target.set("office.objects.list", {
      name: "office.objects.list",
      execute: async (args) => {
        if (!isOfficeActionApp(args.app)) return { success: false, error: "参数 app 必须是 excel、word 或 presentation" };
        if (typeof args.filePath !== "string") return { success: false, error: "缺少必填参数: filePath" };
        const objects = await officeDocumentBridge.listObjects({
          app: args.app,
          filePath: args.filePath,
          instanceId: typeof args.instanceId === "string" ? args.instanceId : undefined,
          kind: typeof args.kind === "string" ? args.kind : undefined,
        });
        return { success: true, data: { objects, count: objects.length } };
      },
    });

    target.set("office.objects.activate", {
      name: "office.objects.activate",
      execute: async (args) => {
        if (!isOfficeActionApp(args.app)) return { success: false, error: "参数 app 必须是 excel、word 或 presentation" };
        if (typeof args.filePath !== "string" || typeof args.locator !== "string") return { success: false, error: "需要 filePath 和 locator" };
        return {
          success: true,
          data: await officeDocumentBridge.activateObject({
            app: args.app,
            filePath: args.filePath,
            instanceId: typeof args.instanceId === "string" ? args.instanceId : undefined,
            locator: args.locator,
          }),
        };
      },
    });
  }

  async function transactionRecordResult(root: string | undefined, id: unknown, action: "inspect" | "undo", force = false) {
    if (!root) return { success: false, error: "Office 事务目录未配置" };
    if (typeof id !== "string") return { success: false, error: "缺少必填参数: transactionId" };
    try {
      const record = action === "undo"
        ? await undoOfficeTransaction(root, id, transactionRestoreOptions(force))
        : await getOfficeTransaction(root, id);
      return { success: record.status !== "conflicted", data: record, ...(record.status === "conflicted" ? { error: record.error } : {}) };
    } catch (error) {
      return { success: false, error: errorMessage(error) };
    }
  }

  function transactionRestoreOptions(force: boolean) {
    return {
      force,
      prepareFiles: officeDocumentBridge
        ? (filePaths: string[]) => officeDocumentBridge.prepareTransaction(filePaths)
        : undefined,
      restoreFiles: officeDocumentBridge
        ? (files: Parameters<OfficeDocumentManagerBridge["restoreTransactionFiles"]>[0]) => officeDocumentBridge.restoreTransactionFiles(files)
        : undefined,
    };
  }
}

function parseWorkflowSteps(value: unknown): OfficeWorkflowStepInput[] | string {
  if (!Array.isArray(value)) return "参数 steps 必须是数组";
  const steps: OfficeWorkflowStepInput[] = [];
  for (let index = 0; index < value.length; index++) {
    const raw = value[index];
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return `工作流第 ${index + 1} 步必须是对象`;
    const item = raw as Record<string, unknown>;
    if (!isOfficeActionApp(item.app)) return `工作流第 ${index + 1} 步 app 无效`;
    if (!isOfficeActionKind(item.action)) return `工作流第 ${index + 1} 步 action 无效`;
    if (typeof item.operation !== "string" || typeof item.filePath !== "string") return `工作流第 ${index + 1} 步需要 operation 和 filePath`;
    const step: OfficeWorkflowStepInput = { app: item.app, action: item.action, operation: item.operation, filePath: item.filePath };
    if (typeof item.outputPath === "string") step.outputPath = item.outputPath;
    if (typeof item.target === "string") step.target = item.target;
    if (isOfficeActionEngine(item.preferEngine)) step.preferEngine = item.preferEngine;
    if (item.params && typeof item.params === "object" && !Array.isArray(item.params)) step.params = item.params as Record<string, unknown>;
    const advancedOperationError = officeAdvancedOperationError(step);
    if (advancedOperationError) return `工作流第 ${index + 1} 步: ${advancedOperationError}`;
    if (typeof item.id === "string") step.id = item.id;
    if (typeof item.parallelGroup === "string") step.parallelGroup = item.parallelGroup;
    if (typeof item.timeoutMs === "number") step.timeoutMs = item.timeoutMs;
    if (item.retry !== undefined) {
      if (!item.retry || typeof item.retry !== "object" || Array.isArray(item.retry)) return `工作流第 ${index + 1} 步 retry 必须是对象`;
      const retry = item.retry as Record<string, unknown>;
      step.retry = {
        maxAttempts: typeof retry.maxAttempts === "number" ? retry.maxAttempts : undefined,
        delayMs: typeof retry.delayMs === "number" ? retry.delayMs : undefined,
      };
    }
    if (item.when !== undefined) {
      const condition = parseWorkflowCondition(item.when);
      if (typeof condition === "string") return `工作流第 ${index + 1} 步 ${condition}`;
      step.when = condition;
    }
    steps.push(step);
  }
  return steps;
}

function parseWorkflowCondition(value: unknown): OfficeWorkflowCondition | string {
  if (!value || typeof value !== "object" || Array.isArray(value)) return "when 必须是对象";
  const item = value as Record<string, unknown>;
  if (typeof item.step !== "number" && typeof item.step !== "string") return "when.step 必须是步骤序号或 ID";
  if (item.status !== undefined && !["done", "failed", "skipped"].includes(String(item.status))) return "when.status 无效";
  return {
    step: item.step,
    status: item.status as OfficeWorkflowCondition["status"],
    dataPath: typeof item.dataPath === "string" ? item.dataPath : undefined,
    ...(Object.prototype.hasOwnProperty.call(item, "equals") ? { equals: item.equals } : {}),
    exists: typeof item.exists === "boolean" ? item.exists : undefined,
  };
}

function isOfficeActionApp(value: unknown): value is OfficeActionApp {
  return value === "excel" || value === "word" || value === "presentation";
}

function isOfficeActionKind(value: unknown): value is OfficeActionKind {
  return value === "inspect" || value === "edit" || value === "style" || value === "insert" || value === "snapshot" || value === "validate";
}

function isOfficeActionEngine(value: unknown): value is OfficeActionEngine {
  return value === "openxml" || value === "com";
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
