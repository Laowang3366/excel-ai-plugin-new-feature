/**
 * Office action 统一适配器
 *
 * 关联模块：
 * - contracts/office.ts: OfficeActionBridge 与 OfficeFileBridge 契约。
 * - implementations/officeOpenXml/*: 实际 Open XML 文件级检查、样式和快照能力。
 * - executors/officeExecutors.ts: 校验工具参数后把统一 action 转发到这里。
 */

import path from "node:path";
import type { OfficeActionBridge, OfficeDocumentManagerBridge, OfficeFileBridge } from "../contracts/office";
import { applyExcelAdvancedAction } from "../implementations/officeOpenXml/advancedExcel";
import { applyPresentationAdvancedAction } from "../implementations/officeOpenXml/advancedPresentation";
import { applyWordAdvancedAction } from "../implementations/officeOpenXml/advancedWord";
import { findOfficeCapability } from "./capabilities";
import { officeActionOperationError } from "./operationPolicy";
import { doneResult, failedResult, needsComResult, unsupportedResult } from "./results";
import {
  createOfficeBackup,
  listOfficeBackups,
  restoreOfficeBackup,
  type OfficeBackupRecord,
} from "./transactions";
import {
  beginOfficeTransaction,
  finalizeOfficeTransaction,
  listOfficeTransactionPaths,
  recordOfficeTransactionResult,
  undoOfficeTransaction,
} from "./transactionJournal";
import type { OfficeActionInput, OfficeActionResult } from "./types";

type TableStylePreset = "professional" | "compact" | "financial";

export interface OfficeActionAdapterDeps {
  officeFileBridge?: OfficeFileBridge;
  /** COM 兜底执行器：Open XML 不适合处理动态对象、导出快照和应用内刷新时使用。 */
  officeComActionBridge?: OfficeActionBridge;
  officeDocumentBridge?: OfficeDocumentManagerBridge;
  backupRoot?: string;
  transactionRoot?: string;
}

const TABLE_STYLES = new Set<TableStylePreset>(["professional", "compact", "financial"]);

export function createOfficeActionBridge(deps: OfficeActionAdapterDeps): OfficeActionBridge {
  return {
    executeAction: (input) => executeOfficeActionWithTransaction(input, deps),
  };
}

async function executeOfficeActionWithTransaction(
  input: OfficeActionInput,
  deps: OfficeActionAdapterDeps,
): Promise<OfficeActionResult> {
  if (input.operation === "listBackups") {
    if (!deps.backupRoot) return failedResult(input, "Office 事务备份目录未配置");
    const records = await listOfficeBackups(deps.backupRoot, input.filePath);
    return doneResult({
      engine: "openxml",
      app: input.app,
      action: input.action,
      operation: input.operation,
      filePath: input.filePath,
      summary: `已列出 ${records.length} 个 Office 事务备份`,
      data: { records },
    });
  }

  if (input.operation === "restoreBackup") {
    if (!deps.backupRoot) return failedResult(input, "Office 事务备份目录未配置");
    const backupPath = stringParam(input.params, "backupPath");
    if (!input.filePath || !backupPath) return failedResult(input, "restoreBackup 需要 filePath 和 params.backupPath");
    try {
      await restoreOfficeBackup({ backupRoot: deps.backupRoot, backupPath, destinationPath: input.filePath });
      return doneResult({
        engine: "openxml",
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        summary: "已恢复 Office 事务备份",
        data: { backupPath },
        changes: [{ kind: "transaction-restore", target: input.filePath, detail: `已从 ${backupPath} 恢复` }],
      });
    } catch (error) {
      return failedResult(input, error);
    }
  }

  if (requiresStandaloneCrossOfficeTransaction(input)) {
    if (!deps.transactionRoot || !deps.officeDocumentBridge) {
      return failedResult(input, "增量跨软件更新需要 Office 事务和文档协调器");
    }
    return executeStandaloneCrossOfficeTransaction(input, {
      ...deps,
      transactionRoot: deps.transactionRoot,
      officeDocumentBridge: deps.officeDocumentBridge,
    });
  }

  let backup: OfficeBackupRecord | undefined;
  if (deps.backupRoot && shouldCreateBackup(input)) {
    try {
      backup = await createOfficeBackup({
        backupRoot: deps.backupRoot,
        app: input.app,
        operation: input.operation,
        sourcePath: input.filePath!,
      });
    } catch (error) {
      return failedResult(input, error);
    }
  }

  const result = await executeOfficeAction(input, deps);
  if (!backup) return result;
  const data = result.data && typeof result.data === "object" && !Array.isArray(result.data)
    ? result.data as Record<string, unknown>
    : {};
  return {
    ...result,
    data: { ...data, transaction: backup },
    changes: [
      { kind: "transaction-backup", target: backup.backupPath, detail: `已备份原文件 ${backup.sourcePath}` },
      ...result.changes,
    ],
  };
}

async function executeStandaloneCrossOfficeTransaction(
  input: OfficeActionInput,
  deps: OfficeActionAdapterDeps & { transactionRoot: string; officeDocumentBridge: OfficeDocumentManagerBridge },
): Promise<OfficeActionResult> {
  const paths = listOfficeTransactionPaths([input]);
  try {
    await deps.officeDocumentBridge.prepareTransaction(paths);
  } catch (error) {
    return failedResult(input, `准备已打开的 Office 文档失败: ${errorMessage(error)}`);
  }

  let transaction: Awaited<ReturnType<typeof beginOfficeTransaction>> | undefined;
  try {
    transaction = await beginOfficeTransaction({ root: deps.transactionRoot, steps: [input] });
    const result = await executeOfficeAction(input, deps);
    await recordOfficeTransactionResult(deps.transactionRoot, transaction, result);
    if (result.status !== "done") {
      await undoOfficeTransaction(deps.transactionRoot, transaction.id, transactionRestoreOptions(deps.officeDocumentBridge, true));
      return withGroupTransaction(result, transaction.id, "已自动恢复跨软件更新前的文件");
    }
    const completed = await finalizeOfficeTransaction(deps.transactionRoot, transaction);
    return withGroupTransaction(result, completed.id, "已创建可整体撤销的跨软件事务");
  } catch (error) {
    let rollbackError = "";
    if (transaction) {
      try {
        await undoOfficeTransaction(deps.transactionRoot, transaction.id, transactionRestoreOptions(deps.officeDocumentBridge, true));
      } catch (rollback) {
        rollbackError = `；自动恢复失败: ${errorMessage(rollback)}`;
      }
    }
    return failedResult(input, `${errorMessage(error)}${rollbackError}`);
  }
}

function requiresStandaloneCrossOfficeTransaction(input: OfficeActionInput): boolean {
  return input.transactionContext !== "workflow"
    && input.params?.updateExisting === true
    && ["exportRangeToWord", "exportRangeToPresentation", "buildReportPackage"].includes(input.operation);
}

function transactionRestoreOptions(bridge: OfficeDocumentManagerBridge, force: boolean) {
  return {
    force,
    prepareFiles: (filePaths: string[]) => bridge.prepareTransaction(filePaths),
    restoreFiles: (files: Parameters<OfficeDocumentManagerBridge["restoreTransactionFiles"]>[0]) => bridge.restoreTransactionFiles(files),
  };
}

function withGroupTransaction(result: OfficeActionResult, transactionId: string, detail: string): OfficeActionResult {
  const data = result.data && typeof result.data === "object" && !Array.isArray(result.data)
    ? result.data as Record<string, unknown>
    : {};
  return {
    ...result,
    data: { ...data, transaction: { id: transactionId, kind: "office-group" } },
    changes: [{ kind: "office-group-transaction", target: transactionId, detail }, ...result.changes],
  };
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function shouldCreateBackup(input: OfficeActionInput): boolean {
  if (!input.filePath || input.action === "inspect" || input.action === "validate" || input.action === "snapshot") return false;
  if ([
    "createWorkbook", "createDocument", "createPresentation", "restoreBackup",
    "exportPdf", "exportSheetsToPdf", "exportHandouts", "exportRangeToWord", "exportRangeToPresentation",
    "buildReportPackage", "traceFormulaDependencies", "inspectFormulaDependencies",
    "inspectFormulaBackups", "inspectFormulaProtection", "inspectPrintSettings",
    "inspectDocumentFormatting", "inspectReferences", "inspectRevisions", "inspectContentControls",
    "mailMerge", "batchMailMerge", "compareDocuments",
  ].includes(input.operation)) return false;
  if (input.outputPath && path.resolve(input.outputPath) !== path.resolve(input.filePath)) return false;
  return true;
}

async function executeOfficeAction(
  input: OfficeActionInput,
  deps: OfficeActionAdapterDeps
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
      return await withComFallback(input, deps, await applyExcelAdvancedAction({ ...input, filePath: input.filePath }));
    }

    if (input.app === "word" && isWordAdvancedOperation(input.operation)) {
      if (!input.filePath) {
        return failedResult(input, "缺少 filePath，无法执行文件级 Office action");
      }
      return await withComFallback(input, deps, await applyWordAdvancedAction({ ...input, filePath: input.filePath }));
    }

    if (input.app === "presentation" && isPresentationAdvancedOperation(input.operation)) {
      if (!input.filePath) {
        return failedResult(input, "缺少 filePath，无法执行文件级 Office action");
      }
      return await withComFallback(input, deps, await applyPresentationAdvancedAction({ ...input, filePath: input.filePath }));
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

    return await withComFallback(input, deps, await routeOpenXmlAction(input, deps.officeFileBridge));
  } catch (error) {
    return failedResult(input, error);
  }
}

async function withComFallback(
  input: OfficeActionInput,
  deps: OfficeActionAdapterDeps,
  result: OfficeActionResult
): Promise<OfficeActionResult> {
  if (result.status !== "needsCom") return result;
  if (!deps.officeComActionBridge) return result;
  return routeComAction(input, deps);
}

async function routeComAction(input: OfficeActionInput, deps: OfficeActionAdapterDeps): Promise<OfficeActionResult> {
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
  return operation === "createWorkbook" ||
    operation === "writeRange" ||
    operation === "setDataValidation" ||
    operation === "applyConditionalFormatting" ||
    operation === "insertChart";
}

function isWordAdvancedOperation(operation: string): boolean {
  return operation === "createDocument" ||
    operation === "applyHeadingStyles" ||
    operation === "styleTables" ||
    operation === "setHeaderFooter" ||
    operation === "insertOrUpdateToc" ||
    operation === "insertOrReplaceImage";
}

function isPresentationAdvancedOperation(operation: string): boolean {
  return operation === "createPresentation" ||
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
    operation === "alignShapes";
}

async function routeOpenXmlAction(input: OfficeActionInput, officeFileBridge: OfficeFileBridge): Promise<OfficeActionResult> {
  if (input.action === "inspect" && input.operation === "inspectFile") {
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

  if (input.action === "inspect" && input.operation === "layout") {
    const data = await officeFileBridge.inspectLayout({ filePath: input.filePath!, target: input.target });
    return doneFromBridge(input, "已检查 Office 布局对象", data);
  }

  if (input.action === "inspect" && input.operation === "tables") {
    const data = await officeFileBridge.inspectTable({ filePath: input.filePath!, target: input.target });
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
    ? value as TableStylePreset
    : "professional";
}

function isUnsupportedSnapshot(data: unknown): data is { supported: false; error?: string } {
  return data !== null &&
    typeof data === "object" &&
    "supported" in data &&
    (data as { supported?: unknown }).supported === false;
}

function doneFromBridge(input: OfficeActionInput, summary: string, data: unknown): OfficeActionResult {
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
    validation: {
      ok: true,
      checks: [{ name: "openxml-route", ok: true, message: "Open XML 路由执行完成" }],
    },
    changes: changedParts.map((partName) => ({
      kind: "openxml-part",
      target: partName,
      detail: `已更新 ${partName}`,
    })),
  });
}

function extractString(value: unknown, key: string): string | undefined {
  return value && typeof value === "object" && key in value && typeof (value as Record<string, unknown>)[key] === "string"
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

function booleanParam(params: Record<string, unknown> | undefined, key: string): boolean | undefined {
  return typeof params?.[key] === "boolean" ? params[key] : undefined;
}
