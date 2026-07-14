import { access, mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  DotNetOfficeActionBridge as OfficeComActionBridge,
  applyExcelAdvancedAction,
  applyPresentationAdvancedAction,
  applyWordAdvancedAction,
  disposeOfficeWorker,
} from "./officeWorkerSmokeHelpers";
import type { OfficeActionInput } from "../electron/agent/tools/officeCore/types";

async function main(): Promise<void> {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-office-smoke-"));
  const excelPath = path.join(tempDir, "pivot.xlsx");
  const wordPath = path.join(tempDir, "report.docx");
  const presentationPath = path.join(tempDir, "slides.pptx");
  const pdfPath = path.join(tempDir, "selected-sheets.pdf");
  const keepArtifacts = process.env.KEEP_OFFICE_SMOKE === "1";
  const operationFilter = new Set(
    (process.env.OFFICE_SMOKE_OPERATIONS || "")
      .split(",")
      .map((operation) => operation.trim())
      .filter(Boolean),
  );
  try {
    await createFixtures(excelPath, wordPath, presentationPath);
    const bridge = new OfficeComActionBridge();
    const actions: OfficeActionInput[] = [
      {
        app: "excel",
        action: "insert",
        operation: "manageWorkbookObject",
        filePath: excelPath,
        target: "range:Sheet1!A1:B4",
        params: { objectType: "table", command: "create", name: "DataTable", style: "TableStyleMedium2" },
      },
      {
        app: "excel",
        action: "insert",
        operation: "manageWorkbookObject",
        filePath: excelPath,
        params: { objectType: "worksheet", command: "add", name: "QueryOutput" },
      },
      {
        app: "excel",
        action: "edit",
        operation: "createPowerQuery",
        filePath: excelPath,
        params: {
          name: "DataPipeline",
          description: "Office advanced smoke query",
          mFormula: "let Source = Excel.CurrentWorkbook(){[Name=\"DataTable\"]}[Content] in Source",
          loadMode: "worksheet",
          destination: "QueryOutput!A1",
          tableName: "PipelineOutput",
        },
      },
      {
        app: "excel",
        action: "edit",
        operation: "managePowerQuery",
        filePath: excelPath,
        params: { command: "rename", name: "DataPipeline", newName: "DataPipelineRenamed" },
      },
      {
        app: "excel",
        action: "edit",
        operation: "managePowerQuery",
        filePath: excelPath,
        params: { command: "rename", name: "DataPipelineRenamed", newName: "DataPipeline" },
      },
      {
        app: "excel",
        action: "edit",
        operation: "managePowerQuery",
        filePath: excelPath,
        params: { command: "duplicate", name: "DataPipeline", newName: "DataPipelineCopy" },
      },
      {
        app: "excel",
        action: "edit",
        operation: "managePowerQuery",
        filePath: excelPath,
        params: { command: "rename", name: "DataPipelineCopy", newName: "DataPipelineArchive" },
      },
      {
        app: "excel",
        action: "edit",
        operation: "managePowerQuery",
        filePath: excelPath,
        params: { command: "load", name: "DataPipelineArchive", loadMode: "connectionOnly" },
      },
      {
        app: "excel",
        action: "edit",
        operation: "managePowerQuery",
        filePath: excelPath,
        params: { command: "refresh", name: "DataPipeline" },
      },
      {
        app: "excel",
        action: "edit",
        operation: "managePowerQuery",
        filePath: excelPath,
        params: { command: "unload", name: "DataPipelineArchive" },
      },
      {
        app: "excel",
        action: "edit",
        operation: "managePowerQuery",
        filePath: excelPath,
        params: { command: "delete", name: "DataPipelineArchive" },
      },
      {
        app: "excel",
        action: "insert",
        operation: "insertChart",
        filePath: excelPath,
        target: "range:Sheet1!A1:B4",
        params: { chartType: "column" },
      },
      {
        app: "excel",
        action: "style",
        operation: "formatChart",
        filePath: excelPath,
        params: {
          chartIndex: 1,
          name: "RevenueChart",
          title: "Amount by department",
          showLegend: false,
          width: 480,
          height: 280,
          chartArea: { fillColor: "FFFFFF", borderColor: "D1D5DB" },
          series: [{ index: 1, fillColor: "2563EB", dataLabels: { enabled: true, showValue: true } }],
          axes: [{ kind: "value", title: "Amount", minimum: 0, numberFormat: "0" }],
        },
      },
      {
        app: "excel",
        action: "edit",
        operation: "manageWorkbookObject",
        filePath: excelPath,
        params: { objectType: "name", command: "upsert", name: "AmountValues", refersTo: "=Sheet1!$B$2:$B$4" },
      },
      {
        app: "excel",
        action: "style",
        operation: "applyWorkbookTemplate",
        filePath: excelPath,
        params: {
          preset: "professional",
          allSheets: true,
          template: {
            version: 1,
            theme: { fontName: "Microsoft YaHei", fontSize: 10, fontColor: "202124" },
            defaultSheet: {
              headerRows: 1,
              freezeRows: 1,
              showGridlines: false,
              autoFit: true,
              bandedRows: true,
              bandedRowColor: "EAF2F8",
              headerStyle: { fillColor: "1F4E79", fontColor: "FFFFFF", bold: true, rowHeight: 24 },
              columns: [{ index: 2, numberFormat: "#,##0", width: 14 }],
            },
          },
        },
      },
      {
        app: "excel",
        action: "insert",
        operation: "createPivotTable",
        filePath: excelPath,
        target: "range:Sheet1!A1:B4",
        params: {
          rowFields: ["Dept"],
          dataFields: [{ name: "Amount", function: "sum" }],
          destination: "Sheet1!D3",
        },
      },
      { app: "excel", action: "inspect", operation: "inspectPowerQueries", filePath: excelPath },
      { app: "excel", action: "inspect", operation: "inspectCharts", filePath: excelPath },
      { app: "excel", action: "inspect", operation: "inspectWorkbookObjects", filePath: excelPath },
      { app: "excel", action: "inspect", operation: "captureWorkbookTemplate", filePath: excelPath },
      {
        app: "excel",
        action: "edit",
        operation: "configurePrint",
        filePath: excelPath,
        params: {
          sheetNames: ["Sheet1", "Inputs"],
          paperSize: "A4",
          orientation: "landscape",
          margins: { top: 1.2, bottom: 1.2, left: 1, right: 1, header: 0.5, footer: 0.5 },
          marginUnit: "centimeters",
          repeatRows: "$1:$1",
          repeatColumns: "$A:$A",
          fitToOnePageWide: true,
          clearPageBreaks: true,
          horizontalPageBreaks: ["A4"],
          verticalPageBreaks: ["C1"],
        },
      },
      { app: "excel", action: "inspect", operation: "inspectPrintSettings", filePath: excelPath, params: { sheetNames: ["Sheet1", "Inputs"] } },
      { app: "excel", action: "inspect", operation: "inspectFormulaDependencies", filePath: excelPath, params: { expectBroken: true } },
      {
        app: "excel",
        action: "edit",
        operation: "repairFormulaReferences",
        filePath: excelPath,
        params: { replacements: [{ sheetName: "Sheet1", find: "#REF!", replace: "Inputs!A1" }] },
      },
      { app: "excel", action: "inspect", operation: "inspectFormulaDependencies", filePath: excelPath, params: { expectBroken: false } },
      {
        app: "excel",
        action: "edit",
        operation: "convertFormulasToValues",
        filePath: excelPath,
        target: "range:Sheet1!C2:D2",
        params: { backupId: "office-smoke-formulas" },
      },
      { app: "excel", action: "inspect", operation: "inspectFormulaBackups", filePath: excelPath },
      {
        app: "excel",
        action: "edit",
        operation: "restoreFormulas",
        filePath: excelPath,
        params: { backupId: "office-smoke-formulas" },
      },
      {
        app: "excel",
        action: "edit",
        operation: "manageFormulaProtection",
        filePath: excelPath,
        target: "range:Sheet1!A1:F4",
        params: { command: "lock", protectSheet: true, unlockInputs: true },
      },
      { app: "excel", action: "inspect", operation: "inspectFormulaProtection", filePath: excelPath, target: "range:Sheet1!A1:F4" },
      { app: "word", action: "style", operation: "formatLongDocument", filePath: wordPath },
      { app: "presentation", action: "style", operation: "layoutElements", filePath: presentationPath },
      {
        app: "excel",
        action: "edit",
        operation: "exportSheetsToPdf",
        filePath: excelPath,
        outputPath: pdfPath,
        params: { sheetNames: ["Sheet1", "Inputs"], mode: "combined", overwrite: true },
      },
    ];
    const selectedActions = operationFilter.size > 0
      ? actions.filter((action) => operationFilter.has(action.operation))
      : actions;
    const results = [];
    for (const action of selectedActions) {
      if (keepArtifacts) await writeFile(path.join(tempDir, `${action.app}-${action.operation}.json`), JSON.stringify(action, null, 2), "utf8");
      const result = await bridge.executeAction(action);
      results.push({ app: action.app, operation: action.operation, status: result.status, error: result.error });
      if (result.status !== "done") throw new Error(`${action.app}/${action.operation}: ${result.error || result.summary}`);
      verifyAdvancedExcelResult(action, result.data);
    }
    if (selectedActions.some((action) => action.operation === "exportSheetsToPdf")) await access(pdfPath);
    process.stdout.write(`${JSON.stringify({ ok: true, results }, null, 2)}\n`);
  } finally {
    await disposeOfficeWorker();
    if (keepArtifacts) process.stdout.write(`Office smoke artifacts: ${tempDir}\n`);
    else await rm(tempDir, { recursive: true, force: true });
  }
}

function verifyAdvancedExcelResult(action: OfficeActionInput, resultData: unknown): void {
  if (action.app !== "excel") return;
  const operationData = nestedData(resultData);
  if (action.operation === "inspectPowerQueries") {
    const queries = arrayField(operationData, "queries");
    const pipeline = queries.find((item) => item.name === "DataPipeline");
    if (!pipeline || !Array.isArray(pipeline.loads) || pipeline.loads.length === 0) {
      throw new Error("Power Query 冒烟验收失败：查询或加载目标不存在");
    }
    if (queries.some((item) => item.name === "DataPipelineArchive")) {
      throw new Error("Power Query 冒烟验收失败：已删除查询仍然存在");
    }
  }
  if (action.operation === "inspectCharts") {
    const charts = arrayField(operationData, "charts");
    const chart = charts.find((item) => item.name === "RevenueChart");
    if (!chart || chart.title !== "Amount by department" || !Array.isArray(chart.series) || chart.series.length === 0) {
      throw new Error("图表深度编辑冒烟验收失败");
    }
  }
  if (action.operation === "inspectWorkbookObjects") {
    const objects = objectField(operationData, "objects");
    const names = arrayField(objects, "names");
    const tables = arrayField(objects, "tables");
    if (!names.some((item) => String(item.name).endsWith("AmountValues")) || !tables.some((item) => item.name === "DataTable")) {
      throw new Error("工作簿对象冒烟验收失败");
    }
  }
  if (action.operation === "captureWorkbookTemplate") {
    const template = objectField(operationData, "template");
    if (!Array.isArray(template.sheets) || template.sheets.length < 2) {
      throw new Error("模板捕获冒烟验收失败");
    }
  }
  if (action.operation === "inspectPrintSettings") {
    const settings = arrayField(operationData, "settings");
    if (settings.length !== 2 || settings.some((item) => item.orientation !== "landscape" || Number(item.fitToPagesWide) !== 1)) {
      throw new Error("打印与页面设置冒烟验收失败");
    }
  }
  if (action.operation === "inspectFormulaDependencies") {
    const broken = arrayField(operationData, "brokenReferences");
    const cycles = operationData.cycles;
    if (action.params?.expectBroken === true && broken.length === 0) {
      throw new Error("公式错误引用检测冒烟验收失败");
    }
    if (action.params?.expectBroken === false && broken.length > 0) {
      throw new Error("公式错误引用修复后仍存在 #REF!");
    }
    if (!Array.isArray(cycles) || cycles.length === 0) {
      throw new Error("公式循环引用检测冒烟验收失败");
    }
    const edges = arrayField(operationData, "edges");
    if (!edges.some((edge) => edge.kind === "cross-sheet")) {
      throw new Error("公式跨表依赖检测冒烟验收失败");
    }
  }
  if (action.operation === "repairFormulaReferences" && Number(operationData.repairedCount) < 1) {
    throw new Error("公式错误引用修复冒烟验收失败");
  }
  if (action.operation === "convertFormulasToValues") {
    if (operationData.backupId !== "office-smoke-formulas" || Number(operationData.convertedFormulaCells) !== 2) {
      throw new Error("公式转值与备份冒烟验收失败");
    }
  }
  if (action.operation === "inspectFormulaBackups") {
    const backups = arrayField(operationData, "backups");
    if (!backups.some((item) => item.backupId === "office-smoke-formulas" && Number(item.formulaCount) === 2)) {
      throw new Error("公式备份清单冒烟验收失败");
    }
  }
  if (action.operation === "restoreFormulas" && Number(operationData.restoredCount) !== 2) {
    throw new Error("公式恢复冒烟验收失败");
  }
  if (action.operation === "inspectFormulaProtection") {
    const protection = arrayField(operationData, "protection");
    if (protection.length !== 1 || protection[0].protected !== true || Number(protection[0].lockedFormulaCount) < 1) {
      throw new Error(`公式区域保护冒烟验收失败: ${JSON.stringify(protection)}`);
    }
  }
  if (action.operation === "exportSheetsToPdf") {
    const outputs = operationData.outputPaths;
    if (!Array.isArray(outputs) || outputs.length !== 1) throw new Error("工作表批量 PDF 导出冒烟验收失败");
  }
}

function nestedData(value: unknown): Record<string, unknown> {
  return asRecord(value);
}

function objectField(value: Record<string, unknown>, key: string): Record<string, unknown> {
  return asRecord(value[key]);
}

function arrayField(value: Record<string, unknown>, key: string): Array<Record<string, unknown>> {
  const field = value[key];
  return Array.isArray(field) ? field.map(asRecord) : [];
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
}

async function createFixtures(excelPath: string, wordPath: string, presentationPath: string): Promise<void> {
  const excelResult = await applyExcelAdvancedAction({
      operation: "createWorkbook",
      filePath: excelPath,
      params: {
        sheetNames: ["Sheet1", "Inputs"],
        values: [
          ["Dept", "Amount", "CrossSheet", "Broken", "CycleA", "CycleB"],
          ["A", 10, "=Inputs!A1+B2", "=#REF!+B2", "=F2+1", "=E2+1"],
          ["B", 20, "=Inputs!A2+B3", "=B3*2", "", ""],
          ["A", 30, "=Inputs!A3+B4", "=B4*2", "", ""],
        ],
      },
    });
  if (excelResult.status !== "done") throw new Error(`创建 Excel 冒烟文件失败: ${excelResult.error || excelResult.summary}`);
  const inputResult = await applyExcelAdvancedAction({
    operation: "writeRange",
    filePath: excelPath,
    target: "range:Inputs!A1:A3",
    params: { values: [[1], [2], [3]] },
  });
  if (inputResult.status !== "done") throw new Error(`写入 Excel 冒烟输入失败: ${inputResult.error || inputResult.summary}`);
  const results = await Promise.all([
    applyWordAdvancedAction({
      operation: "createDocument",
      filePath: wordPath,
      params: { title: "一、概览", paragraphs: ["这是用于高级自动化冒烟测试的临时文档。"] },
    }),
    applyPresentationAdvancedAction({
      operation: "createPresentation",
      filePath: presentationPath,
      params: { title: "Office advanced automation smoke test" },
    }),
  ]);
  const failed = results.find((result) => result.status !== "done");
  if (failed) throw new Error(`创建 Office 冒烟文件失败: ${failed.error || failed.summary}`);
}

void main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exitCode = 1;
});
