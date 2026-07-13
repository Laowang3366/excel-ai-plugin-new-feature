import { access, mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { OfficeComActionBridge } from "../electron/agent/tools/implementations/office/officeComActionBridge";
import { buildComScript } from "../electron/agent/tools/implementations/office/officeComActionScripts";
import { applyExcelAdvancedAction } from "../electron/agent/tools/implementations/officeOpenXml/advancedExcel";
import { applyWordAdvancedAction } from "../electron/agent/tools/implementations/officeOpenXml/advancedWord";
import type { OfficeActionInput } from "../electron/agent/tools/officeCore/types";

async function main(): Promise<void> {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-word-smoke-"));
  const sourcePath = path.join(tempDir, "source.docx");
  const revisedPath = path.join(tempDir, "revised.docx");
  const controlPath = path.join(tempDir, "controls.docx");
  const mergeTemplatePath = path.join(tempDir, "merge-template.docx");
  const dataPath = path.join(tempDir, "merge-data.xlsx");
  const comparisonPath = path.join(tempDir, "comparison.docx");
  const singleMergePath = path.join(tempDir, "merged.docx");
  const outputDirectory = path.join(tempDir, "merged");
  const imagePath = path.join(tempDir, "photo.png");
  const keepArtifacts = process.env.KEEP_WORD_SMOKE === "1";
  const operationFilter = new Set(
    (process.env.WORD_SMOKE_OPERATIONS || "")
      .split(",")
      .map((operation) => operation.trim())
      .filter(Boolean),
  );
  try {
    await createFixtures({ sourcePath, revisedPath, controlPath, mergeTemplatePath, dataPath, imagePath });
    const bridge = new OfficeComActionBridge();
    const actions: OfficeActionInput[] = [
      { app: "word", action: "inspect", operation: "inspectDocumentFormatting", filePath: sourcePath },
      {
        app: "word",
        action: "style",
        operation: "formatLongDocument",
        filePath: sourcePath,
        params: {
          autoDetectHeadings: true,
          toc: "create",
          margins: { top: 2.5, bottom: 2.5, left: 2.8, right: 2.8 },
          headerFooter: { header: "高级 Word 自动化", footer: "内部测试" },
          pageNumbers: true,
        },
      },
      { app: "word", action: "insert", operation: "manageReferences", filePath: sourcePath, params: { command: "addBookmark", name: "Overview", start: 0, end: 3 } },
      { app: "word", action: "insert", operation: "manageReferences", filePath: sourcePath, params: { command: "addFootnote", text: "脚注来源" } },
      { app: "word", action: "insert", operation: "manageReferences", filePath: sourcePath, params: { command: "addEndnote", text: "尾注来源" } },
      { app: "word", action: "insert", operation: "manageReferences", filePath: sourcePath, params: { command: "addCaption", targetType: "range", label: "图", title: "测试题注" } },
      { app: "word", action: "insert", operation: "manageReferences", filePath: sourcePath, params: { command: "addCrossReference", referenceType: "bookmark", item: "Overview" } },
      { app: "word", action: "insert", operation: "manageReferences", filePath: sourcePath, params: { command: "addTableOfFigures", label: "图" } },
      { app: "word", action: "inspect", operation: "inspectReferences", filePath: sourcePath },
      {
        app: "word",
        action: "edit",
        operation: "applyTrackedChanges",
        filePath: sourcePath,
        params: { edits: [{ command: "replace", find: "旧内容", replace: "新内容", all: true }], keepTracking: true },
      },
      { app: "word", action: "inspect", operation: "inspectRevisions", filePath: sourcePath },
      { app: "word", action: "edit", operation: "compareDocuments", filePath: sourcePath, outputPath: comparisonPath, params: { revisedFilePath: revisedPath, author: "Smoke Compare" } },
      {
        app: "word",
        action: "edit",
        operation: "manageContentControls",
        filePath: controlPath,
        params: {
          command: "add",
          controls: [
            { type: "text", tag: "name", title: "姓名", placeholder: "请输入姓名" },
            { type: "checkBox", tag: "approved", title: "同意" },
            { type: "date", tag: "date", title: "日期" },
            { type: "dropDownList", tag: "department", title: "部门", entries: [{ text: "销售", value: "sales" }, { text: "研发", value: "rd" }] },
            { type: "picture", tag: "photo", title: "照片" },
          ],
        },
      },
      {
        app: "word",
        action: "edit",
        operation: "populateContentControls",
        filePath: controlPath,
        params: {
          values: { name: "张三", approved: true, date: { value: "2026-07-12", dateFormat: "yyyy-MM-dd" }, department: "rd", photo: { value: imagePath, kind: "image", width: 36 } },
          strictListValues: true,
        },
      },
      { app: "word", action: "inspect", operation: "inspectContentControls", filePath: controlPath },
      {
        app: "word",
        action: "edit",
        operation: "prepareMailMergeTemplate",
        filePath: mergeTemplatePath,
        params: { fields: [{ placeholder: "{{Name}}", field: "Name" }, { placeholder: "{{Id}}", field: "Id" }] },
      },
      {
        app: "word",
        action: "edit",
        operation: "mailMerge",
        filePath: mergeTemplatePath,
        outputPath: singleMergePath,
        params: { dataSourcePath: dataPath, outputFormat: "docx" },
      },
      {
        app: "word",
        action: "edit",
        operation: "batchMailMerge",
        filePath: mergeTemplatePath,
        outputPath: path.join(outputDirectory, "batch.docx"),
        params: {
          dataSourcePath: dataPath,
          outputDirectory,
          outputFormat: "both",
          fileNamePattern: "{Id}-{Name}",
          overwrite: true,
          conditions: [{ placeholder: "{{StatusText}}", field: "Status", operator: "eq", value: "active", trueText: "已生效", falseText: "待确认" }],
          imageFields: [{ placeholder: "{{Photo}}", field: "Photo", width: 36 }],
        },
      },
    ];

    const selectedActions = operationFilter.size > 0
      ? actions.filter((action) => operationFilter.has(action.operation))
      : actions;
    const results = [];
    for (const action of selectedActions) {
      const command = typeof action.params?.command === "string" ? `-${action.params.command}` : "";
      process.stdout.write(`Testing ${action.operation}${command}\n`);
      if (keepArtifacts) await writeFile(path.join(tempDir, `${action.operation}${command}.ps1`), buildComScript(action), "utf8");
      const result = await bridge.executeAction(action);
      results.push({ operation: `${action.operation}${command}`, status: result.status, error: result.error });
      if (result.status !== "done") throw new Error(`${action.operation}${command}: ${result.error || result.summary}`);
      verifyResult(action, result.data);
    }
    if (selectedActions.some((action) => action.operation === "compareDocuments")) await access(comparisonPath);
    if (selectedActions.some((action) => action.operation === "mailMerge")) await access(singleMergePath);
    const mergedFiles = selectedActions.some((action) => action.operation === "batchMailMerge")
      ? await readdir(outputDirectory)
      : [];
    if (selectedActions.some((action) => action.operation === "batchMailMerge")) {
      const docxCount = mergedFiles.filter((name) => name.endsWith(".docx")).length;
      const pdfCount = mergedFiles.filter((name) => name.endsWith(".pdf")).length;
      if (docxCount !== 2 || pdfCount !== 2) throw new Error(`批量邮件合并输出数量错误: docx=${docxCount}, pdf=${pdfCount}`);
    }
    process.stdout.write(`${JSON.stringify({ ok: true, results, mergedFiles }, null, 2)}\n`);
  } finally {
    if (keepArtifacts) process.stdout.write(`Word smoke artifacts: ${tempDir}\n`);
    else await rm(tempDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
}

function verifyResult(action: OfficeActionInput, data: unknown): void {
  const operationData = asRecord(asRecord(data).data);
  if (action.operation === "inspectDocumentFormatting") {
    const formatting = asRecord(operationData.formatting);
    if (Number(formatting.paragraphCount) < 2 || !Array.isArray(formatting.styles)) throw new Error("长文档格式检查失败");
  }
  if (action.operation === "inspectReferences") {
    const references = asRecord(operationData.references);
    if (!Array.isArray(references.bookmarks) || references.bookmarks.length < 1 || !Array.isArray(references.footnotes) || references.footnotes.length < 1 || !Array.isArray(references.endnotes) || references.endnotes.length < 1) {
      throw new Error(`引用检查失败: ${JSON.stringify(references)}`);
    }
  }
  if (action.operation === "inspectRevisions") {
    const review = asRecord(operationData.review);
    if (Number(review.revisionCount) < 1 || review.trackRevisions !== true) throw new Error(`修订检查失败: ${JSON.stringify(review)}`);
  }
  if (action.operation === "compareDocuments") {
    const summary = asRecord(operationData.summary);
    if (Number(summary.changeCount) < 1) throw new Error(`文档对比失败: ${JSON.stringify(summary)}`);
  }
  if (action.operation === "inspectContentControls") {
    const controls = Array.isArray(operationData.controls) ? operationData.controls.map(asRecord) : [];
    if (controls.length !== 5) throw new Error(`内容控件数量错误: ${controls.length}`);
    if (!controls.some((control) => control.tag === "approved" && control.checked === true)) throw new Error("复选框内容控件填充失败");
    if (!controls.some((control) => control.tag === "department" && String(control.text).includes("研发"))) throw new Error("下拉内容控件填充失败");
  }
  if (action.operation === "batchMailMerge" && Number(operationData.recordCount) !== 2) {
    throw new Error(`批量邮件合并记录数错误: ${JSON.stringify(operationData)}`);
  }
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
}

async function createFixtures(input: {
  sourcePath: string;
  revisedPath: string;
  controlPath: string;
  mergeTemplatePath: string;
  dataPath: string;
  imagePath: string;
}): Promise<void> {
  const image = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  await writeFile(input.imagePath, image);
  const results = await Promise.all([
    applyWordAdvancedAction({ operation: "createDocument", filePath: input.sourcePath, params: { title: "一、项目概览", paragraphs: ["这是旧内容，用于修订测试。", "（一）实施范围", "正文段落。"] } }),
    applyWordAdvancedAction({ operation: "createDocument", filePath: input.revisedPath, params: { title: "一、项目概览", paragraphs: ["这是修订后的内容。", "（一）实施范围", "新增段落。"] } }),
    applyWordAdvancedAction({ operation: "createDocument", filePath: input.controlPath, params: { title: "智能模板", paragraphs: ["以下字段由内容控件填充。"] } }),
    applyWordAdvancedAction({ operation: "createDocument", filePath: input.mergeTemplatePath, params: { title: "合同通知", paragraphs: ["客户：{{Name}}", "编号：{{Id}}", "状态：{{StatusText}}", "照片：{{Photo}}"] } }),
    applyExcelAdvancedAction({ operation: "createWorkbook", filePath: input.dataPath, params: { values: [["Name", "Id", "Status", "Photo"], ["张三", "A001", "active", input.imagePath], ["李四", "A002", "pending", input.imagePath]] } }),
  ]);
  const failed = results.find((result) => result.status !== "done");
  if (failed) throw new Error(`创建 Word 冒烟文件失败: ${failed.error || failed.summary}`);
}

void main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exitCode = 1;
});
