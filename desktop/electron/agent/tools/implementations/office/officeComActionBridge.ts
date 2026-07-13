/**
 * Office COM 高级 action 兜底执行器。
 *
 * 关联模块：
 * - officeCore/officeActionAdapter.ts: Open XML 不支持或显式 preferEngine=com 时转到这里。
 * - automation/powershell.ts: 负责 PowerShell 执行和变量安全注入。
 */

import { executePowerShell } from "../../../automation/powershell";
import { safeJsonParse } from "../../../automation/json";
import { doneResult, failedResult } from "../../officeCore/results";
import type { OfficeActionBridge } from "../../contracts/office";
import type { OfficeActionInput, OfficeActionResult } from "../../officeCore/types";
import { buildComScript } from "./officeComActionScripts";

type ComChange = { kind: string; target?: string; detail: string };

const EXCEL_COM_OPERATIONS = new Set([
  "insertChart", "applyConditionalFormatting", "setDataValidation", "styleTable",
  "snapshot",
  "createPivotTable", "refreshPivotTables", "addSlicer", "createPowerQuery",
  "inspectPowerQueries", "managePowerQuery", "inspectCharts", "formatChart",
  "inspectWorkbookObjects", "manageWorkbookObject", "manageWorksheetObjects",
  "captureWorkbookTemplate", "inspectWorkbookFormatting", "applyWorkbookTemplate",
  "inspectPrintSettings", "configurePrint", "exportSheetsToPdf",
  "exportPdf", "traceFormulaDependencies", "inspectFormulaDependencies",
  "repairFormulaReferences", "convertFormulasToValues", "inspectFormulaBackups",
  "restoreFormulas", "inspectFormulaProtection", "manageFormulaProtection",
  "exportRangeToWord", "exportRangeToPresentation", "buildReportPackage",
]);
const WORD_COM_OPERATIONS = new Set([
  "applyHeadingStyles", "insertOrUpdateToc", "styleTables", "setHeaderFooter",
  "insertOrReplaceImage", "snapshot", "formatLongDocument", "manageReferences",
  "inspectDocumentFormatting", "inspectReferences", "inspectRevisions", "manageRevisions",
  "compareDocuments", "applyTrackedChanges", "prepareMailMergeTemplate", "mailMerge", "batchMailMerge",
  "inspectContentControls", "populateContentControls", "manageContentControls", "exportPdf",
  "inspectLinkedOfficeContent", "refreshLinkedOfficeContent", "relinkLinkedOfficeContent",
]);
const PPT_COM_OPERATIONS = new Set([
  "applyTheme", "deleteSlides", "normalizeLayouts", "insertChart", "insertTable", "replacePictureSlot",
  "alignShapes", "snapshot", "applyMasterBranding", "layoutElements",
  "inspectPresentationTheme", "inspectSlideElements", "inspectAnimations", "inspectSpeakerNotes",
  "configureAnimations", "configureSlideShow", "setSpeakerNotes", "exportHandouts",
  "inspectLinkedOfficeContent", "refreshLinkedOfficeContent", "relinkLinkedOfficeContent",
]);

export class OfficeComActionBridge implements OfficeActionBridge {
  async executeAction(input: OfficeActionInput): Promise<OfficeActionResult> {
    if (!input.filePath) {
      return failedResult({ ...input, preferEngine: "com" }, "缺少 filePath，无法执行 COM Office action");
    }

    if (!supportsComAction(input)) {
      return {
        status: "unsupported",
        engine: "com",
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: `暂不支持 COM Office action: ${input.app}/${input.operation}`,
        changes: [],
      };
    }

    try {
      const script = buildComScript(input);
      const requestedTimeout = Number(input.params?.actionTimeoutMs);
      const timeout = Number.isFinite(requestedTimeout)
        ? Math.min(600_000, Math.max(5_000, Math.trunc(requestedTimeout)))
        : 120_000;
      const output = await executePowerShell(script, timeout);
      const data = safeJsonParse<{ outputPath?: string; changes?: ComChange[] }>(output, "powershell", "执行 COM Office action");
      const outputPath = data.outputPath || input.outputPath || input.filePath;
      return doneResult({
        engine: "com",
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        outputPath,
        target: input.target,
        summary: `已通过 COM 执行 ${input.app === "presentation" ? "PowerPoint" : input.app[0].toUpperCase() + input.app.slice(1)} ${input.operation}`,
        data,
        validation: {
          ok: true,
          checks: [{ name: "com-route", ok: true, message: "COM 兜底执行完成" }],
        },
        changes: Array.isArray(data.changes) ? data.changes : [],
      });
    } catch (error) {
      return failedResult({ ...input, preferEngine: "com" }, error);
    }
  }
}

function supportsComAction(input: OfficeActionInput): boolean {
  if (input.app === "excel") return EXCEL_COM_OPERATIONS.has(input.operation);
  if (input.app === "word") return WORD_COM_OPERATIONS.has(input.operation);
  return PPT_COM_OPERATIONS.has(input.operation);
}
