/**
 * Office 工具执行器
 *
 * 只注册 Word、PowerPoint 和通用 Office 脚本工具。
 */

import type { ToolExecutor } from "../../shared/types";
import type { ExcelConnectionBridge } from "../contracts/excel";
import type {
  WordDocumentBridge,
  PresentationBridge,
  OfficeActionBridge,
  OfficeDocumentManagerBridge,
} from "../contracts/office";
import {
  officeActionOperationError,
  officeAdvancedOperationError,
} from "../officeCore/operationPolicy";
import type {
  OfficeActionApp,
  OfficeActionEngine,
  OfficeActionInput,
  OfficeActionKind,
} from "../officeCore/types";
import { addOfficeReliabilityExecutors } from "./officeReliabilityExecutors";
import { addOfficeWordExecutors } from "./officeWordExecutors";
import { validateArgs } from "./validation";
import { omitVersionMetadata, toModelFacingSpreadsheetMetadata } from "./modelFacingMetadata";

export interface OfficeExecutorDeps {
  excelBridge?: ExcelConnectionBridge;
  wordBridge?: WordDocumentBridge;
  presentationBridge?: PresentationBridge;
  officeActionBridge?: OfficeActionBridge;
  officeDocumentBridge?: OfficeDocumentManagerBridge;
  workflowRoot?: string;
  transactionRoot?: string;
}

function addToolAlias(
  target: Map<string, ToolExecutor>,
  alias: string,
  canonicalName: string,
): void {
  const executor = target.get(canonicalName);
  if (!executor || target.has(alias)) return;
  target.set(alias, { ...executor, name: alias });
}

export function addOfficeExecutors(
  target: Map<string, ToolExecutor>,
  deps: OfficeExecutorDeps,
): void {
  const { excelBridge, wordBridge, presentationBridge, officeActionBridge, officeDocumentBridge } =
    deps;

  target.set("office.connection.status", {
    name: "office.connection.status",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { app: "string" });
      if (err) return { success: false, error: err };
      const app = args.app;
      if (app === "excel") {
        if (typeof excelBridge?.detectStatus === "function") {
          const status = await excelBridge.detectStatus();
          return { success: true, data: toModelFacingSpreadsheetMetadata(status) };
        }
        return {
          success: true,
          data: { connected: false, host: "unknown", error: "Excel bridge not available" },
        };
      }
      if (app === "word") {
        const detectStatus = (wordBridge as { detectStatus?: () => Promise<unknown> } | undefined)
          ?.detectStatus;
        if (typeof detectStatus === "function") {
          return { success: true, data: omitVersionMetadata(await detectStatus.call(wordBridge)) };
        }
        return {
          success: true,
          data: { connected: false, host: "unknown", error: "Word bridge not available" },
        };
      }
      if (app === "presentation") {
        const detectStatus = (
          presentationBridge as { detectStatus?: () => Promise<unknown> } | undefined
        )?.detectStatus;
        if (typeof detectStatus === "function") {
          return {
            success: true,
            data: omitVersionMetadata(await detectStatus.call(presentationBridge)),
          };
        }
        return {
          success: true,
          data: { connected: false, host: "unknown", error: "Presentation bridge not available" },
        };
      }
      return { success: false, error: "参数 app 必须是 excel、word 或 presentation" };
    },
  });

  if (wordBridge) {
    addOfficeWordExecutors(target, {
      wordBridge,
      officeActionBridge,
      inspectFileWithOpenXml,
    });
  }

  if (presentationBridge) {
    target.set("presentation.open", {
      name: "presentation.open",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { filePath: "string" });
        if (err) return { success: false, error: err };
        const filePath = args.filePath as string;
        const result = await presentationBridge.openPresentation(filePath);
        if (!result.success && officeActionBridge) {
          const fallback = await inspectFileWithOpenXml(
            officeActionBridge,
            "presentation",
            filePath,
            result.error,
          );
          if (fallback) return fallback;
        }
        return { success: result.success, data: result, error: result.error };
      },
    });

    target.set("presentation.inspect", {
      name: "presentation.inspect",
      execute: async (_args: Record<string, unknown>) => {
        const result = await presentationBridge.inspectPresentation();
        return { success: true, data: result };
      },
    });

    target.set("presentation.readSlide", {
      name: "presentation.readSlide",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { slideIndex: "number" });
        if (err) return { success: false, error: err };
        const result = await presentationBridge.readSlide(args.slideIndex as number);
        return { success: true, data: result };
      },
    });

    target.set("presentation.addSlide", {
      name: "presentation.addSlide",
      execute: async (args: Record<string, unknown>) => {
        const result = await presentationBridge.addSlide(
          args.title as string | undefined,
          args.body as string | undefined,
          args.layout as string | undefined,
        );
        return { success: true, data: result };
      },
    });

    target.set("presentation.setShapeText", {
      name: "presentation.setShapeText",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { slideIndex: "number", text: "string" });
        if (err) return { success: false, error: err };
        const result = await presentationBridge.setShapeText(
          args.slideIndex as number,
          args.text as string,
          args.shapeName as string | undefined,
          typeof args.shapeIndex === "number" ? args.shapeIndex : undefined,
        );
        return { success: true, data: result };
      },
    });

    target.set("presentation.replaceText", {
      name: "presentation.replaceText",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { findText: "string", replaceText: "string" });
        if (err) return { success: false, error: err };
        const result = await presentationBridge.replaceText(
          args.findText as string,
          args.replaceText as string,
          typeof args.matchCase === "boolean" ? args.matchCase : undefined,
        );
        return { success: true, data: result };
      },
    });

    target.set("presentation.save", {
      name: "presentation.save",
      execute: async (args: Record<string, unknown>) => {
        const result = await presentationBridge.savePresentation(
          args.saveAsPath as string | undefined,
        );
        return { success: result.success, data: result, error: result.error };
      },
    });
  }

  if (officeActionBridge) {
    target.set("office.action.inspect", {
      name: "office.action.inspect",
      execute: async (args: Record<string, unknown>) =>
        executeOfficeAction(args, officeActionBridge, "inspect"),
    });

    target.set("office.action.apply", {
      name: "office.action.apply",
      execute: async (args: Record<string, unknown>) =>
        executeOfficeAction(args, officeActionBridge),
    });

    target.set("office.action.validate", {
      name: "office.action.validate",
      execute: async (args: Record<string, unknown>) =>
        executeOfficeAction(args, officeActionBridge, "validate"),
    });
  }

  addOfficeReliabilityExecutors(target, {
    officeActionBridge,
    officeDocumentBridge,
    workflowRoot: deps.workflowRoot,
    transactionRoot: deps.transactionRoot,
  });

  addToolAlias(target, "office.connection_status", "office.connection.status");
  addToolAlias(target, "office_connection_status", "office.connection.status");
  addToolAlias(target, "office.action_inspect", "office.action.inspect");
  addToolAlias(target, "office.action_apply", "office.action.apply");
  addToolAlias(target, "office.action_validate", "office.action.validate");
}

async function inspectFileWithOpenXml(
  officeActionBridge: OfficeActionBridge,
  app: "word" | "presentation",
  filePath: string,
  openError?: string,
) {
  const inspection = await officeActionBridge.executeAction({
    app,
    action: "inspect",
    operation: "inspectFile",
    filePath,
  });
  if (inspection.status !== "done") return undefined;

  return {
    success: false,
    error: openError || `${app === "word" ? "Word" : "PowerPoint"} 未能打开文件`,
    data: {
      success: false,
      fileReadable: true,
      openedInApp: false,
      fallback: "openxml",
      openError,
      inspection,
    },
  };
}

function isOfficeActionApp(value: unknown): value is OfficeActionApp {
  return value === "excel" || value === "word" || value === "presentation";
}

function isOfficeActionKind(value: unknown): value is OfficeActionKind {
  return (
    value === "inspect" ||
    value === "edit" ||
    value === "style" ||
    value === "insert" ||
    value === "snapshot" ||
    value === "validate"
  );
}

function isOfficeActionEngine(value: unknown): value is OfficeActionEngine {
  return value === "openxml" || value === "com";
}

async function executeOfficeAction(
  args: Record<string, unknown>,
  officeActionBridge: OfficeActionBridge,
  defaultAction?: OfficeActionKind,
) {
  const err = validateArgs(args, { app: "string", operation: "string" });
  if (err) return { success: false, error: err };
  if (!defaultAction) {
    const filePathError = validateArgs(args, { filePath: "string" });
    if (filePathError) return { success: false, error: filePathError };
  }
  if (!isOfficeActionApp(args.app)) {
    return { success: false, error: "参数 app 必须是 excel、word 或 presentation" };
  }
  const action = defaultAction || args.action;
  if (!isOfficeActionKind(action)) {
    return {
      success: false,
      error: "参数 action 必须是 inspect、edit、style、insert、snapshot 或 validate",
    };
  }
  const operation = args.operation as string;
  const operationError = officeActionOperationError(action, operation);
  if (operationError) {
    return { success: false, error: operationError };
  }

  const input: OfficeActionInput = {
    app: args.app,
    action,
    operation,
  };
  if (typeof args.filePath === "string") input.filePath = args.filePath;
  if (typeof args.outputPath === "string") input.outputPath = args.outputPath;
  if (typeof args.target === "string") input.target = args.target;
  if (isOfficeActionEngine(args.preferEngine)) input.preferEngine = args.preferEngine;
  if (args.params && typeof args.params === "object" && !Array.isArray(args.params)) {
    input.params = args.params as Record<string, unknown>;
  }
  const advancedOperationError = officeAdvancedOperationError(input);
  if (advancedOperationError) return { success: false, error: advancedOperationError };

  const result = await officeActionBridge.executeAction(input);
  const success = result.status === "done";
  return {
    success,
    data: result,
    ...(success
      ? {}
      : { error: result.summary || `Office action returned status: ${result.status}` }),
  };
}
