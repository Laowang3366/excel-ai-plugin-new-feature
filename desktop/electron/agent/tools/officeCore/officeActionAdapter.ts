/**
 * Office action 统一适配器
 *
 * 关联模块：
 * - contracts/office.ts: OfficeActionBridge 与 OfficeFileBridge 契约。
 * - officeWorker/dotNetOpenXmlBridge: .NET Open XML SDK 文件级执行入口。
 * - executors/officeExecutors.ts: 校验工具参数后把统一 action 转发到这里。
 */

import type { OfficeActionBridge, OfficeFileBridge } from "../contracts/office";
import { findOfficeCapability } from "./capabilities";
import {
  executeOfficeActionWithTransaction,
  type OfficeActionTransactionDeps,
} from "./officeActionTransactionAdapter";
import { officeActionOperationError } from "./operationPolicy";
import { doneResult, failedResult, needsComResult, unsupportedResult } from "./results";
import type { OfficeActionInput, OfficeActionResult } from "./types";
import { withValidationChecks } from "./officeActionValidation";

type TableStylePreset = "professional" | "compact" | "financial";

export interface OfficeActionAdapterDeps extends OfficeActionTransactionDeps {
  officeFileBridge?: OfficeFileBridge;
  /** COM 兜底执行器：Open XML 不适合处理动态对象、导出快照和应用内刷新时使用。 */
  officeComActionBridge?: OfficeActionBridge;
}

const TABLE_STYLES = new Set<TableStylePreset>(["professional", "compact", "financial"]);

export function createOfficeActionBridge(deps: OfficeActionAdapterDeps): OfficeActionBridge {
  return {
    executeAction: (input) =>
      executeOfficeActionWithTransaction(input, deps, async (actionInput) => {
        const result = await executeOfficeAction(actionInput, deps);
        return actionInput.action === "validate"
          ? withValidationChecks(actionInput, result)
          : result;
      }),
  };
}

async function executeOfficeAction(
  input: OfficeActionInput,
  deps: OfficeActionAdapterDeps,
): Promise<OfficeActionResult> {
  try {
    const operationError = officeActionOperationError(input.action, input.operation);
    if (operationError) {
      return failedResult(input, operationError);
    }

    if (input.preferEngine === "com") {
      return await routeComAction(input, deps);
    }

    if (input.app === "excel" && isExcelAdvancedOperation(input.operation)) {
      if (!input.filePath) {
        return failedResult(input, "缺少 filePath，无法执行文件级 Office action");
      }
      const result = await routeAdvancedOpenXml(input, deps);
      return await withComFallback(input, deps, result);
    }

    if (input.app === "word" && isWordAdvancedOperation(input.operation)) {
      if (!input.filePath) {
        return failedResult(input, "缺少 filePath，无法执行文件级 Office action");
      }
      const result = await routeAdvancedOpenXml(input, deps);
      return await withComFallback(input, deps, result);
    }

    if (input.app === "presentation" && isPresentationAdvancedOperation(input.operation)) {
      if (!input.filePath) {
        return failedResult(input, "缺少 filePath，无法执行文件级 Office action");
      }
      const result = await routeAdvancedOpenXml(input, deps);
      return await withComFallback(input, deps, result);
    }

    const capability = findOfficeCapability(input.app, input.operation);
    if (capability?.preferredEngine === "com") {
      return await routeComAction(input, deps);
    }

    if (!deps.officeFileBridge) {
      return needsComResult({
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: "未配置 Open XML 文件桥，需要 COM 执行",
      });
    }
    if (!input.filePath) {
      return failedResult(input, "缺少 filePath，无法执行文件级 Office action");
    }

    return await withComFallback(
      input,
      deps,
      await routeOpenXmlAction(input, deps.officeFileBridge),
    );
  } catch (error) {
    return failedResult(input, error);
  }
}

async function withComFallback(
  input: OfficeActionInput,
  deps: OfficeActionAdapterDeps,
  result: OfficeActionResult,
): Promise<OfficeActionResult> {
  if (result.status !== "needsCom") return result;
  if (!deps.officeComActionBridge) return result;
  return routeComAction(input, deps);
}

async function routeAdvancedOpenXml(
  input: OfficeActionInput,
  deps: OfficeActionAdapterDeps,
): Promise<OfficeActionResult> {
  if (deps.officeFileBridge?.executeAction) return deps.officeFileBridge.executeAction(input);
  return needsComResult({
    app: input.app,
    action: input.action,
    operation: input.operation,
    filePath: input.filePath,
    outputPath: input.outputPath,
    target: input.target,
    summary: ".NET Open XML Worker 未配置，需要 COM 执行",
  });
}

async function routeComAction(
  input: OfficeActionInput,
  deps: OfficeActionAdapterDeps,
): Promise<OfficeActionResult> {
  if (!deps.officeComActionBridge) {
    return needsComResult({
      app: input.app,
      action: input.action,
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
      summary: "需要 COM 执行，但当前未配置 COM action 桥",
    });
  }
  return deps.officeComActionBridge.executeAction(input);
}

function isExcelAdvancedOperation(operation: string): boolean {
  return (
    operation === "createWorkbook" ||
    operation === "writeRange" ||
    operation === "setDataValidation" ||
    operation === "applyConditionalFormatting" ||
    operation === "insertChart"
  );
}

function isWordAdvancedOperation(operation: string): boolean {
  return (
    operation === "createDocument" ||
    operation === "applyHeadingStyles" ||
    operation === "styleTables" ||
    operation === "setHeaderFooter" ||
    operation === "insertOrUpdateToc" ||
    operation === "insertOrReplaceImage"
  );
}

function isPresentationAdvancedOperation(operation: string): boolean {
  return (
    operation === "createPresentation" ||
    operation === "addSlide" ||
    operation === "addSlides" ||
    operation === "appendSlide" ||
    operation === "appendSlides" ||
    operation === "addSlideContent" ||
    operation === "applyTheme" ||
    operation === "deleteSlides" ||
    operation === "normalizeLayouts" ||
    operation === "insertChart" ||
    operation === "replacePictureSlot" ||
    operation === "alignShapes"
  );
}

async function routeOpenXmlAction(
  input: OfficeActionInput,
  officeFileBridge: OfficeFileBridge,
): Promise<OfficeActionResult> {
  if (
    (input.action === "inspect" || input.action === "validate") &&
    input.operation === "inspectFile"
  ) {
    const data = await officeFileBridge.inspectFile(input.filePath!);
    return doneFromBridge(input, "已检查 Office 文件结构", data);
  }

  if (input.action === "edit" && input.operation === "replaceText") {
    const findText = stringParam(input.params, "findText");
    const replaceText = stringParam(input.params, "replaceText");
    if (!findText || replaceText === undefined) {
      return failedResult(input, "replaceText 操作需要 params.findText 和 params.replaceText");
    }
    const data = await officeFileBridge.replaceText({
      filePath: input.filePath!,
      findText,
      replaceText,
      outputPath: input.outputPath,
      matchCase: booleanParam(input.params, "matchCase"),
    });
    return doneFromBridge(input, "已替换 Office 文件文本", data);
  }

  if ((input.action === "inspect" || input.action === "validate") && input.operation === "layout") {
    const data = await officeFileBridge.inspectLayout({
      filePath: input.filePath!,
      target: input.target,
    });
    return doneFromBridge(input, "已检查 Office 布局对象", data);
  }

  if ((input.action === "inspect" || input.action === "validate") && input.operation === "tables") {
    const data = await officeFileBridge.inspectTable({
      filePath: input.filePath!,
      target: input.target,
    });
    return doneFromBridge(input, "已检查 Office 表格结构", data);
  }

  if (input.action === "style" && input.operation === "styleTable") {
    const style = normalizeTableStyle(input.params?.style);
    const data = await officeFileBridge.applyTableStyle({
      filePath: input.filePath!,
      target: input.target,
      style,
      outputPath: input.outputPath,
    });
    return doneFromBridge(input, "已应用 Open XML 表格样式", data);
  }

  // snapshot 是跨 action 的只读特例，允许 action 不精确时仍按 operation 路由。
  if (input.action === "snapshot" || input.operation === "snapshot") {
    const data = await officeFileBridge.snapshot({
      filePath: input.filePath!,
      target: input.target,
      outputPath: input.outputPath,
      preferEngine: input.preferEngine,
    });
    if (isUnsupportedSnapshot(data)) {
      return needsComResult({
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: data.error || "Open XML 快照暂不可用，需要显式 COM 兜底",
        data,
      });
    }
    return doneFromBridge(input, "已生成 Office 视觉快照", data);
  }

  const capability = findOfficeCapability(input.app, input.operation);
  if (capability?.fallback === "needsCom") {
    return needsComResult({
      app: input.app,
      action: input.action,
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
      summary: `Open XML 暂未覆盖 ${input.operation}，需要显式 COM 兜底`,
    });
  }

  return unsupportedResult({
    app: input.app,
    action: input.action,
    operation: input.operation,
    filePath: input.filePath,
    outputPath: input.outputPath,
    target: input.target,
    summary: `暂不支持 Office action: ${input.app}/${input.operation}`,
  });
}

function normalizeTableStyle(value: unknown): TableStylePreset {
  return typeof value === "string" && TABLE_STYLES.has(value as TableStylePreset)
    ? (value as TableStylePreset)
    : "professional";
}

function isUnsupportedSnapshot(data: unknown): data is { supported: false; error?: string } {
  return (
    data !== null &&
    typeof data === "object" &&
    "supported" in data &&
    (data as { supported?: unknown }).supported === false
  );
}

function doneFromBridge(
  input: OfficeActionInput,
  summary: string,
  data: unknown,
): OfficeActionResult {
  const outputPath = extractString(data, "outputPath") || input.outputPath;
  const changedParts = extractStringArray(data, "changedParts");
  return doneResult({
    engine: "openxml",
    app: input.app,
    action: input.action,
    operation: input.operation,
    filePath: input.filePath,
    outputPath,
    target: input.target,
    summary,
    data,
    changes: changedParts.map((partName) => ({
      kind: "openxml-part",
      target: partName,
      detail: `已更新 ${partName}`,
    })),
  });
}

function extractString(value: unknown, key: string): string | undefined {
  return value &&
    typeof value === "object" &&
    key in value &&
    typeof (value as Record<string, unknown>)[key] === "string"
    ? (value as Record<string, string>)[key]
    : undefined;
}

function extractStringArray(value: unknown, key: string): string[] {
  if (!value || typeof value !== "object" || !(key in value)) return [];
  const raw = (value as Record<string, unknown>)[key];
  return Array.isArray(raw) ? raw.filter((item): item is string => typeof item === "string") : [];
}

function stringParam(params: Record<string, unknown> | undefined, key: string): string | undefined {
  return typeof params?.[key] === "string" ? params[key] : undefined;
}

function booleanParam(
  params: Record<string, unknown> | undefined,
  key: string,
): boolean | undefined {
  return typeof params?.[key] === "boolean" ? params[key] : undefined;
}
