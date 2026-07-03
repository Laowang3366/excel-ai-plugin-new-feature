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
  OfficeScriptBridge,
  OfficeActionBridge,
} from "../contracts/office";
import type { OfficeActionApp, OfficeActionEngine, OfficeActionInput, OfficeActionKind } from "../officeCore/types";
import { validateArgs } from "./validation";

export interface OfficeExecutorDeps {
  excelBridge?: ExcelConnectionBridge;
  wordBridge?: WordDocumentBridge;
  presentationBridge?: PresentationBridge;
  officeScriptBridge?: OfficeScriptBridge;
  officeActionBridge?: OfficeActionBridge;
}

function addToolAlias(target: Map<string, ToolExecutor>, alias: string, canonicalName: string): void {
  const executor = target.get(canonicalName);
  if (!executor || target.has(alias)) return;
  target.set(alias, { ...executor, name: alias });
}

export function addOfficeExecutors(target: Map<string, ToolExecutor>, deps: OfficeExecutorDeps): void {
  const { excelBridge, wordBridge, presentationBridge, officeScriptBridge, officeActionBridge } = deps;

  target.set("office.connection.status", {
    name: "office.connection.status",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { app: "string" });
      if (err) return { success: false, error: err };
      const app = args.app;
      if (app === "excel") {
        if (typeof excelBridge?.detectStatus === "function") {
          const status = await excelBridge.detectStatus();
          return { success: true, data: status };
        }
        return { success: true, data: { connected: false, host: "unknown", error: "Excel bridge not available" } };
      }
      if (app === "word") {
        const detectStatus = (wordBridge as { detectStatus?: () => Promise<unknown> } | undefined)?.detectStatus;
        if (typeof detectStatus === "function") {
          return { success: true, data: await detectStatus.call(wordBridge) };
        }
        return { success: true, data: { connected: false, host: "unknown", error: "Word bridge not available" } };
      }
      if (app === "presentation") {
        const detectStatus = (presentationBridge as { detectStatus?: () => Promise<unknown> } | undefined)?.detectStatus;
        if (typeof detectStatus === "function") {
          return { success: true, data: await detectStatus.call(presentationBridge) };
        }
        return { success: true, data: { connected: false, host: "unknown", error: "Presentation bridge not available" } };
      }
      return { success: false, error: "参数 app 必须是 excel、word 或 presentation" };
    },
  });

  if (wordBridge) {
    target.set("word.open", {
      name: "word.open",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { filePath: "string" });
        if (err) return { success: false, error: err };
        const filePath = args.filePath as string;
        const result = await wordBridge.openDocument(filePath);
        if (!result.success && officeActionBridge) {
          const fallback = await inspectFileWithOpenXml(officeActionBridge, "word", filePath, result.error);
          if (fallback) return fallback;
        }
        return { success: result.success, data: result, error: result.error };
      },
    });

    target.set("word.create", {
      name: "word.create",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { filePath: "string" });
        if (err) return { success: false, error: err };
        if (officeActionBridge) {
          const openXmlResult = await officeActionBridge.executeAction({
            app: "word",
            action: "insert",
            operation: "createDocument",
            filePath: args.filePath as string,
            params: args.params && typeof args.params === "object" && !Array.isArray(args.params)
              ? args.params as Record<string, unknown>
              : undefined,
          });
          if (openXmlResult.status === "done") {
            return { success: true, data: openXmlResult };
          }
        }
        const result = await wordBridge.createDocument(args.filePath as string);
        return { success: result.success, data: result, error: result.error };
      },
    });

    target.set("word.inspect", {
      name: "word.inspect",
      execute: async (_args: Record<string, unknown>) => {
        const result = await wordBridge.inspectDocument();
        return { success: true, data: result };
      },
    });

    target.set("word.readText", {
      name: "word.readText",
      execute: async (args: Record<string, unknown>) => {
        const maxChars = typeof args.maxChars === "number" ? args.maxChars : undefined;
        const result = await wordBridge.readText(maxChars);
        return { success: true, data: result };
      },
    });

    target.set("word.insertText", {
      name: "word.insertText",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { text: "string" });
        if (err) return { success: false, error: err };
        const result = await wordBridge.insertText(args.text as string, args.position as string | undefined);
        return { success: true, data: result };
      },
    });

    target.set("word.insertHeading", {
      name: "word.insertHeading",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { text: "string" });
        if (err) return { success: false, error: err };
        const result = await wordBridge.insertHeading(
          args.text as string,
          typeof args.level === "number" ? args.level : undefined,
          args.position as string | undefined
        );
        return { success: true, data: result };
      },
    });

    target.set("word.replaceText", {
      name: "word.replaceText",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { findText: "string", replaceText: "string" });
        if (err) return { success: false, error: err };
        const result = await wordBridge.replaceText(
          args.findText as string,
          args.replaceText as string,
          typeof args.matchCase === "boolean" ? args.matchCase : undefined
        );
        return { success: true, data: result };
      },
    });

    target.set("word.save", {
      name: "word.save",
      execute: async (args: Record<string, unknown>) => {
        const result = await wordBridge.saveDocument(args.saveAsPath as string | undefined);
        return { success: result.success, data: result, error: result.error };
      },
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
            result.error
          );
          if (fallback) return fallback;
        }
        return { success: result.success, data: result, error: result.error };
      },
    });

    target.set("presentation.create", {
      name: "presentation.create",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { filePath: "string" });
        if (err) return { success: false, error: err };
        if (officeActionBridge) {
          const openXmlResult = await officeActionBridge.executeAction({
            app: "presentation",
            action: "insert",
            operation: "createPresentation",
            filePath: args.filePath as string,
          });
          if (openXmlResult.status === "done") {
            return { success: true, data: openXmlResult };
          }
        }
        const result = await presentationBridge.createPresentation(args.filePath as string);
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
          args.layout as string | undefined
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
          typeof args.shapeIndex === "number" ? args.shapeIndex : undefined
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
          typeof args.matchCase === "boolean" ? args.matchCase : undefined
        );
        return { success: true, data: result };
      },
    });

    target.set("presentation.save", {
      name: "presentation.save",
      execute: async (args: Record<string, unknown>) => {
        const result = await presentationBridge.savePresentation(args.saveAsPath as string | undefined);
        return { success: result.success, data: result, error: result.error };
      },
    });
  }

  if (officeActionBridge) {
    target.set("office.action.inspect", {
      name: "office.action.inspect",
      execute: async (args: Record<string, unknown>) => executeOfficeAction(args, officeActionBridge, "inspect"),
    });

    target.set("office.action.apply", {
      name: "office.action.apply",
      execute: async (args: Record<string, unknown>) => executeOfficeAction(args, officeActionBridge),
    });

    target.set("office.action.validate", {
      name: "office.action.validate",
      execute: async (args: Record<string, unknown>) => executeOfficeAction(args, officeActionBridge, "validate"),
    });
  }

  if (officeScriptBridge) {
    target.set("office.script.execute", {
      name: "office.script.execute",
      execute: async (args: Record<string, unknown>) => {
        const err = validateArgs(args, { app: "string", code: "string" });
        if (err) return { success: false, error: err };
        const app = args.app as string;
        if (app !== "word" && app !== "presentation") {
          return { success: false, error: "参数 app 必须是 word 或 presentation" };
        }
        const result = await officeScriptBridge.executeScript(app, args.code as string);
        return { success: true, data: result };
      },
    });
  }

  addToolAlias(target, "office.connection_status", "office.connection.status");
  addToolAlias(target, "office_connection_status", "office.connection.status");
  addToolAlias(target, "office.action_inspect", "office.action.inspect");
  addToolAlias(target, "office.action_apply", "office.action.apply");
  addToolAlias(target, "office.action_validate", "office.action.validate");
  addToolAlias(target, "office.script_execute", "office.script.execute");
}

async function inspectFileWithOpenXml(
  officeActionBridge: OfficeActionBridge,
  app: "word" | "presentation",
  filePath: string,
  openError?: string
) {
  const inspection = await officeActionBridge.executeAction({
    app,
    action: "inspect",
    operation: "inspectFile",
    filePath,
  });
  if (inspection.status !== "done") return undefined;

  return {
    success: true,
    data: {
      success: true,
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
  return value === "inspect" ||
    value === "edit" ||
    value === "style" ||
    value === "insert" ||
    value === "snapshot" ||
    value === "validate";
}

function isOfficeActionEngine(value: unknown): value is OfficeActionEngine {
  return value === "openxml" || value === "com";
}

async function executeOfficeAction(
  args: Record<string, unknown>,
  officeActionBridge: OfficeActionBridge,
  defaultAction?: OfficeActionKind
) {
  const err = validateArgs(args, { app: "string", operation: "string" });
  if (err) return { success: false, error: err };
  if (!isOfficeActionApp(args.app)) {
    return { success: false, error: "参数 app 必须是 excel、word 或 presentation" };
  }
  const action = defaultAction || args.action;
  if (!isOfficeActionKind(action)) {
    return { success: false, error: "参数 action 必须是 inspect、edit、style、insert、snapshot 或 validate" };
  }

  const input: OfficeActionInput = {
    app: args.app,
    action,
    operation: args.operation as string,
  };
  if (typeof args.filePath === "string") input.filePath = args.filePath;
  if (typeof args.outputPath === "string") input.outputPath = args.outputPath;
  if (typeof args.target === "string") input.target = args.target;
  if (isOfficeActionEngine(args.preferEngine)) input.preferEngine = args.preferEngine;
  if (args.params && typeof args.params === "object" && !Array.isArray(args.params)) {
    input.params = args.params as Record<string, unknown>;
  }

  const result = await officeActionBridge.executeAction(input);
  const success = result.status === "done";
  return {
    success,
    data: result,
    ...(success ? {} : { error: result.summary || `Office action returned status: ${result.status}` }),
  };
}
