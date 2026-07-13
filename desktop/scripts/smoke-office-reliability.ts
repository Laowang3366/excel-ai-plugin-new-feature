import { access, mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { executePowerShell, psVar } from "../electron/agent/automation/powershell";
import type { OfficeActionBridge } from "../electron/agent/tools/contracts/office";
import { OfficeComActionBridge } from "../electron/agent/tools/implementations/office/officeComActionBridge";
import { OfficeDocumentComBridge, officeInstanceDiscoveryScript } from "../electron/agent/tools/implementations/office/officeDocumentComBridge";
import { applyExcelAdvancedAction } from "../electron/agent/tools/implementations/officeOpenXml/advancedExcel";
import { applyPresentationAdvancedAction } from "../electron/agent/tools/implementations/officeOpenXml/advancedPresentation";
import { applyWordAdvancedAction } from "../electron/agent/tools/implementations/officeOpenXml/advancedWord";
import { getOfficeTransaction, redoOfficeTransaction, undoOfficeTransaction } from "../electron/agent/tools/officeCore/transactionJournal";
import { getOfficeWorkflow, runOfficeWorkflow } from "../electron/agent/tools/officeCore/workflow";
import type { OfficeActionInput, OfficeActionResult } from "../electron/agent/tools/officeCore/types";

type OwnedProcesses = { excel: number[]; word: number[]; presentation: number[] };

async function main(): Promise<void> {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-office-reliability-"));
  const keepArtifacts = process.env.KEEP_OFFICE_RELIABILITY_SMOKE === "1";
  const sourcePath = path.join(tempDir, "source.xlsx");
  const secondWorkbookPath = path.join(tempDir, "second.xlsx");
  const linkedWordPath = path.join(tempDir, "linked-report.docx");
  const linkedPresentationPath = path.join(tempDir, "linked-slides.pptx");
  const pipelineWordPath = path.join(tempDir, "pipeline-report.docx");
  const pipelinePresentationPath = path.join(tempDir, "pipeline-slides.pptx");
  const selectionWordPath = path.join(tempDir, "selection.docx");
  const selectionPresentationPath = path.join(tempDir, "selection.pptx");
  const workflowRoot = path.join(tempDir, "workflows");
  const transactionRoot = path.join(tempDir, "transactions");
  const wordLinkId = "reliability-word";
  const presentationLinkId = "reliability-presentation";
  let ownedProcesses: OwnedProcesses | undefined;
  let primaryError: unknown;
  let summary: Record<string, unknown> | undefined;
  const officeProcessesBefore = await listMicrosoftOfficeProcessIds();
  const protectedWpsProcesses = await listVisibleWpsProcessIds();

  try {
    await createWorkbook(sourcePath, [["部门", "金额"], ["华东", 10], ["华南", 20], ["华北", 30]]);
    await createWorkbook(secondWorkbookPath, [["名称", "值"], ["备用", 1]]);
    await createSelectionFixtures(selectionWordPath, selectionPresentationPath);
    const nativeBridge = new OfficeComActionBridge();
    const bridge: OfficeActionBridge = {
      executeAction: (input) => nativeBridge.executeAction(routeToMicrosoftOffice(input)),
    };
    await runAction(bridge, {
      app: "excel", action: "insert", operation: "insertChart", filePath: sourcePath,
      target: "range:Sheet1!A1:B4", params: { chartType: "column" },
    });
    await runAction(bridge, {
      app: "excel", action: "style", operation: "formatChart", filePath: sourcePath,
      params: { chartIndex: 1, name: "RevenueChart", title: "部门金额" },
    });

    await runAction(bridge, {
      app: "excel", action: "insert", operation: "exportRangeToWord", filePath: sourcePath,
      outputPath: linkedWordPath, target: "range:Sheet1!A1:B4",
      params: { linked: true, linkId: wordLinkId, title: "部门金额明细", overwrite: true },
    });
    await runAction(bridge, {
      app: "excel", action: "insert", operation: "exportRangeToPresentation", filePath: sourcePath,
      outputPath: linkedPresentationPath,
      params: { linked: true, linkId: presentationLinkId, sourceType: "chart", chartName: "RevenueChart", title: "部门金额", overwrite: true },
    });

    const wordLinksBefore = linksFrom(await runAction(bridge, {
      app: "word", action: "inspect", operation: "inspectLinkedOfficeContent", filePath: linkedWordPath,
    }));
    const presentationLinksBefore = linksFrom(await runAction(bridge, {
      app: "presentation", action: "inspect", operation: "inspectLinkedOfficeContent", filePath: linkedPresentationPath,
    }));
    assertLinks(wordLinksBefore, sourcePath, "Word");
    assertLinks(presentationLinksBefore, sourcePath, "PowerPoint");

    await runAction(bridge, {
      app: "word", action: "insert", operation: "manageReferences", filePath: linkedWordPath,
      params: { command: "addBookmark", name: "ManualKeep" },
    });
    await runAction(bridge, {
      app: "word", action: "edit", operation: "applyTrackedChanges", filePath: linkedWordPath,
      params: { edits: [{ command: "replaceBookmark", name: "ManualKeep", text: "人工保留段落" }], keepTracking: false },
    });
    await runAction(bridge, {
      app: "presentation", action: "insert", operation: "insertTable", filePath: linkedPresentationPath,
      target: "slide:1", params: { name: "ManualKeep", values: [["人工保留表格"]], left: 520, top: 400, width: 180, height: 60 },
    });
    const managedShapeName = String(presentationLinksBefore[0]?.shapeName || presentationLinksBefore[0]?.name || "");
    if (!managedShapeName) throw new Error(`PowerPoint 链接缺少 shapeName: ${JSON.stringify(presentationLinksBefore)}`);
    await runAction(bridge, {
      app: "presentation", action: "style", operation: "layoutElements", filePath: linkedPresentationPath,
      target: "slide:1", params: { mode: "precise", shapeNames: [managedShapeName], edits: [{ shapeName: managedShapeName, left: 96, top: 112, width: 410, height: 236, preserveAspectRatio: true }] },
    });
    const managedLayoutBefore = findPresentationShape(await runAction(bridge, {
      app: "presentation", action: "inspect", operation: "inspectSlideElements", filePath: linkedPresentationPath,
      target: "slide:1", params: { allSlides: true },
    }), managedShapeName);

    const writeResult = await applyExcelAdvancedAction({
      operation: "writeRange",
      filePath: sourcePath,
      target: "range:Sheet1!B2:B4",
      params: { values: [[100], [200], [300]] },
    });
    if (writeResult.status !== "done") throw new Error(writeResult.error || writeResult.summary);
    const wordRefresh = await runAction(bridge, {
      app: "word", action: "edit", operation: "refreshLinkedOfficeContent", filePath: linkedWordPath,
    });
    const presentationRefresh = await runAction(bridge, {
      app: "presentation", action: "edit", operation: "refreshLinkedOfficeContent", filePath: linkedPresentationPath,
    });
    assertRefresh(wordRefresh, "Word");
    assertRefresh(presentationRefresh, "PowerPoint");
    assertSameLocators(wordLinksBefore, linksFrom(await runAction(bridge, {
      app: "word", action: "inspect", operation: "inspectLinkedOfficeContent", filePath: linkedWordPath,
    })), "Word");
    assertSameLocators(presentationLinksBefore, linksFrom(await runAction(bridge, {
      app: "presentation", action: "inspect", operation: "inspectLinkedOfficeContent", filePath: linkedPresentationPath,
    })), "PowerPoint");

    const wordIncremental = await runAction(bridge, {
      app: "excel", action: "insert", operation: "exportRangeToWord", filePath: sourcePath,
      outputPath: linkedWordPath, target: "range:Sheet1!A1:B4",
      params: { linked: true, updateExisting: true, linkId: wordLinkId, title: "部门金额明细（更新）" },
    });
    await runAction(bridge, {
      app: "excel", action: "insert", operation: "exportRangeToPresentation", filePath: sourcePath,
      outputPath: linkedPresentationPath,
      params: { linked: true, updateExisting: true, linkId: presentationLinkId, sourceType: "chart", chartName: "RevenueChart", title: "部门金额（更新）" },
    });
    assertWordBookmark(await runAction(bridge, {
      app: "word", action: "inspect", operation: "inspectReferences", filePath: linkedWordPath,
    }), "ManualKeep", "人工保留段落");
    const wordLinksAfterIncremental = linksFrom(await runAction(bridge, {
      app: "word", action: "inspect", operation: "inspectLinkedOfficeContent", filePath: linkedWordPath,
    }));
    if (wordLinksAfterIncremental.length !== 1 || String(wordLinksAfterIncremental[0].linkId) !== wordLinkId) {
      throw new Error(`Word 增量更新后的链接清单异常: ${JSON.stringify({ links: wordLinksAfterIncremental, update: operationData(wordIncremental) })}`);
    }
    const presentationAfterIncremental = await runAction(bridge, {
      app: "presentation", action: "inspect", operation: "inspectSlideElements", filePath: linkedPresentationPath,
      target: "slide:1", params: { allSlides: true },
    });
    findPresentationShape(presentationAfterIncremental, "ManualKeep");
    assertPresentationLayout(managedLayoutBefore, findPresentationShape(presentationAfterIncremental, managedShapeName));

    await runAction(bridge, {
      app: "excel", action: "insert", operation: "insertChart", filePath: secondWorkbookPath,
      target: "range:Sheet1!A1:B2", params: { chartType: "column" },
    });
    await runAction(bridge, {
      app: "excel", action: "style", operation: "formatChart", filePath: secondWorkbookPath,
      params: { chartIndex: 1, name: "RevenueChart", title: "备用数据" },
    });
    await assertTargetedRelink(bridge, "word", linkedWordPath, wordLinkId, secondWorkbookPath, sourcePath);
    await assertTargetedRelink(bridge, "presentation", linkedPresentationPath, presentationLinkId, secondWorkbookPath, sourcePath);

    const pipelineSteps: OfficeActionInput[] = [
      { app: "excel", action: "style", operation: "formatChart", filePath: sourcePath, params: { chartIndex: 1, name: "RevenueChart", title: "流水线部门金额" } },
      { app: "excel", action: "insert", operation: "exportRangeToWord", filePath: sourcePath, outputPath: pipelineWordPath, target: "range:Sheet1!A1:B4", params: { linked: true, overwrite: true } },
      { app: "excel", action: "insert", operation: "exportRangeToPresentation", filePath: sourcePath, outputPath: pipelinePresentationPath, params: { linked: true, sourceType: "chart", chartName: "RevenueChart", overwrite: true } },
    ];
    const pipeline = await runOfficeWorkflow(bridge, pipelineSteps, { workflowRoot, transactionRoot });
    if (pipeline.status !== "done" || !pipeline.transactionId) throw new Error(`办公流水线失败: ${pipeline.error || pipeline.status}`);
    await access(pipelineWordPath);
    await access(pipelinePresentationPath);
    const transaction = await getOfficeTransaction(transactionRoot, pipeline.transactionId);
    if (transaction.artifacts.length !== 2 || transaction.changes.length < 3) throw new Error("事务产物或修改清单不完整");
    await undoOfficeTransaction(transactionRoot, pipeline.transactionId);
    await expectMissing(pipelineWordPath);
    await expectMissing(pipelinePresentationPath);
    await redoOfficeTransaction(transactionRoot, pipeline.transactionId, bridge);
    await access(pipelineWordPath);
    await access(pipelinePresentationPath);

    let failOnce = true;
    const unstableBridge: OfficeActionBridge = {
      executeAction: async (input) => {
        if (input.operation === "inspectLinkedOfficeContent" && failOnce) {
          failOnce = false;
          throw new Error("模拟 Word 临时不可用");
        }
        return bridge.executeAction(input);
      },
    };
    const resumableSteps: OfficeActionInput[] = [
      { app: "excel", action: "inspect", operation: "inspectCharts", filePath: sourcePath },
      { app: "word", action: "inspect", operation: "inspectLinkedOfficeContent", filePath: pipelineWordPath },
      { app: "presentation", action: "inspect", operation: "inspectLinkedOfficeContent", filePath: pipelinePresentationPath },
    ];
    const paused = await runOfficeWorkflow(unstableBridge, resumableSteps, { workflowRoot, transactionRoot });
    if (paused.status !== "paused" || paused.completedSteps !== 1 || paused.nextStep !== 2 || !paused.workflowId) {
      throw new Error(`工作流未在失败步骤暂停: ${JSON.stringify(paused)}`);
    }
    const resumed = await runOfficeWorkflow(unstableBridge, [], {
      workflowRoot, transactionRoot, workflowId: paused.workflowId, resume: true,
    });
    if (resumed.status !== "done" || resumed.completedSteps !== 3) throw new Error(`工作流续跑失败: ${JSON.stringify(resumed)}`);
    const savedWorkflow = await getOfficeWorkflow(workflowRoot, paused.workflowId);
    if (savedWorkflow.stepRecords.some((step) => step.status !== "done")) throw new Error("工作流步骤记录不完整");

    ownedProcesses = await openFixtures([sourcePath, secondWorkbookPath], [selectionWordPath, linkedWordPath], [selectionPresentationPath, linkedPresentationPath]);
    const documentBridge = new OfficeDocumentComBridge();
    const documents = await documentBridge.listDocuments();
    const excelDocuments = documents.filter((document) => document.app === "excel" && [sourcePath, secondWorkbookPath].some((file) => samePath(file, document.fullName)));
    if (excelDocuments.length !== 2) throw new Error(`多工作簿识别失败: ${JSON.stringify(documents)}`);
    assertIndependentInstances(excelDocuments, "Excel");
    const wordDocuments = documents.filter((document) => document.app === "word" && [selectionWordPath, linkedWordPath].some((file) => samePath(file, document.fullName)));
    const presentationDocuments = documents.filter((document) => document.app === "presentation" && [selectionPresentationPath, linkedPresentationPath].some((file) => samePath(file, document.fullName)));
    if (wordDocuments.length !== 2) throw new Error(`多 Word 实例识别失败: ${JSON.stringify(documents)}`);
    if (presentationDocuments.length !== 2) throw new Error(`多 PowerPoint 实例识别失败: ${JSON.stringify(documents)}`);
    assertIndependentInstances(wordDocuments, "Word");
    assertIndependentInstances(presentationDocuments, "PowerPoint");
    const sourceDocument = excelDocuments.find((document) => samePath(sourcePath, document.fullName))!;
    const wordDocument = wordDocuments.find((document) => samePath(selectionWordPath, document.fullName))!;
    const presentationDocument = presentationDocuments.find((document) => samePath(selectionPresentationPath, document.fullName))!;
    const excelObjects = await documentBridge.listObjects({ app: "excel", filePath: sourcePath, instanceId: sourceDocument.instanceId });
    const chart = excelObjects.find((item) => item.kind === "chart" && item.name === "RevenueChart");
    if (!chart) throw new Error(`Excel 图表对象识别失败: ${JSON.stringify(excelObjects)}`);
    await documentBridge.activateObject({ app: "excel", filePath: sourcePath, instanceId: sourceDocument.instanceId, locator: chart.locator });
    const wordObjects = await documentBridge.listObjects({ app: "word", filePath: selectionWordPath, instanceId: wordDocument.instanceId });
    const wordPage = wordObjects.find((item) => item.kind === "page");
    if (!wordPage) throw new Error(`Word 页面识别失败: ${JSON.stringify(wordObjects)}`);
    await documentBridge.activateObject({ app: "word", filePath: selectionWordPath, instanceId: wordDocument.instanceId, locator: wordPage.locator });
    const presentationObjects = await documentBridge.listObjects({ app: "presentation", filePath: selectionPresentationPath, instanceId: presentationDocument.instanceId });
    const selectedSlide = presentationObjects.find((item) => item.kind === "slide");
    const presentationShape = presentationObjects.find((item) => item.kind === "shape" && item.parent === selectedSlide?.locator);
    if (!presentationShape) throw new Error(`PowerPoint 对象识别失败: ${JSON.stringify(presentationObjects)}`);
    await documentBridge.activateObject({ app: "presentation", filePath: selectionPresentationPath, instanceId: presentationDocument.instanceId, locator: presentationShape.locator });

    const dirtyTransaction = await verifyDirtyDocumentTransaction({
      bridge,
      documentBridge,
      workflowRoot,
      transactionRoot,
      filePath: selectionWordPath,
      instanceId: wordDocument.instanceId,
    });

    summary = {
      ok: true,
      linkedRefresh: { word: wordLinksBefore.length, presentation: presentationLinksBefore.length },
      pipeline: { workflowId: pipeline.workflowId, transactionId: pipeline.transactionId, artifacts: transaction.artifacts.length },
      resume: { workflowId: paused.workflowId, completedSteps: resumed.completedSteps },
      dirtyTransaction,
      selection: { documents: documents.length, instances: { excel: excelDocuments.length, word: wordDocuments.length, presentation: presentationDocuments.length }, excelObjects: excelObjects.length, wordObjects: wordObjects.length, presentationObjects: presentationObjects.length },
    };
  } catch (error) {
    primaryError = error;
  } finally {
    const cleanupErrors: unknown[] = [];
    const cleanup = async (operation: () => Promise<unknown>) => {
      try { await operation(); } catch (error) { cleanupErrors.push(error); }
    };
    if (ownedProcesses) await cleanup(() => closeFixtures(ownedProcesses!));
    await cleanup(() => assertProcessesStillRunning(protectedWpsProcesses, "测试前已存在的 WPS 窗口"));
    await cleanup(() => assertProcessesStillRunning(officeProcessesBefore, "测试前已存在的 Microsoft Office 进程"));
    if (ownedProcesses) await cleanup(() => assertOwnedMicrosoftOfficeProcessesStopped(ownedProcesses!));
    await cleanup(() => assertNoUnexpectedMicrosoftOfficeProcesses(officeProcessesBefore));
    if (keepArtifacts) process.stdout.write(`Office reliability smoke artifacts: ${tempDir}\n`);
    else await cleanup(() => rm(tempDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 300 }));
    if (cleanupErrors.length > 0) {
      const detail = cleanupErrors.map((error) => error instanceof Error ? error.message : String(error)).join("；");
      if (primaryError) process.stderr.write(`Office reliability cleanup: ${detail}\n`);
      else primaryError = new Error(detail);
    }
  }
  if (primaryError) throw primaryError;
  process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
}

async function createWorkbook(filePath: string, values: unknown[][]): Promise<void> {
  const result = await applyExcelAdvancedAction({ operation: "createWorkbook", filePath, params: { sheetNames: ["Sheet1"], values } });
  if (result.status !== "done") throw new Error(result.error || result.summary);
}

async function createSelectionFixtures(wordPath: string, presentationPath: string): Promise<void> {
  const word = await applyWordAdvancedAction({
    operation: "createDocument",
    filePath: wordPath,
    params: { title: "对象选择测试", paragraphs: ["用于验证 Word 页面和对象定位。"] },
  });
  if (word.status !== "done") throw new Error(word.error || word.summary);
  const presentation = await applyPresentationAdvancedAction({
    operation: "createPresentation",
    filePath: presentationPath,
    params: { title: "对象选择测试", subtitle: "用于验证幻灯片和形状定位" },
  });
  if (presentation.status !== "done") throw new Error(presentation.error || presentation.summary);
}

async function runAction(bridge: OfficeActionBridge, input: OfficeActionInput): Promise<OfficeActionResult> {
  const result = await bridge.executeAction(input);
  if (result.status !== "done") throw new Error(`${input.app}/${input.operation}: ${result.error || result.summary}`);
  return result;
}

function operationData(result: OfficeActionResult): Record<string, unknown> {
  const outer = asRecord(result.data);
  return asRecord(outer.data);
}

function linksFrom(result: OfficeActionResult): Array<Record<string, unknown>> {
  const links = operationData(result).links;
  return Array.isArray(links) ? links.map(asRecord) : [];
}

function assertLinks(links: Array<Record<string, unknown>>, sourcePath: string, label: string): void {
  if (links.length === 0 || !links.some((link) => String(link.source || "").toLowerCase().includes(path.basename(sourcePath).toLowerCase()))) {
    throw new Error(`${label} 未保留 Excel 链接来源: ${JSON.stringify(links)}`);
  }
}

function assertRefresh(result: OfficeActionResult, label: string): void {
  const data = operationData(result);
  if (Number(data.updated) < 1 || (Array.isArray(data.failures) && data.failures.length > 0)) {
    throw new Error(`${label} 链接刷新失败: ${JSON.stringify(data)}`);
  }
}

function assertSameLocators(before: Array<Record<string, unknown>>, after: Array<Record<string, unknown>>, label: string): void {
  const beforeLocators = before.map((item) => String(item.locator)).sort();
  const afterLocators = after.map((item) => String(item.locator)).sort();
  if (JSON.stringify(beforeLocators) !== JSON.stringify(afterLocators)) throw new Error(`${label} 刷新后对象定位发生变化`);
}

function assertWordBookmark(result: OfficeActionResult, name: string, text: string): void {
  const references = asRecord(operationData(result).references);
  const bookmarks = Array.isArray(references.bookmarks) ? references.bookmarks.map(asRecord) : [];
  const bookmark = bookmarks.find((item) => String(item.name) === name);
  if (!bookmark || !String(bookmark.text || "").includes(text)) {
    throw new Error(`Word 增量更新未保留人工书签内容: ${JSON.stringify(bookmarks)}`);
  }
}

function findPresentationShape(result: OfficeActionResult, name: string): Record<string, unknown> {
  const rawSlides = operationData(result).slides;
  const slides = (Array.isArray(rawSlides) ? rawSlides : rawSlides ? [rawSlides] : []).map(asRecord);
  const shapes = slides.flatMap((slide) => {
    const value = slide.shapes;
    return (Array.isArray(value) ? value : value ? [value] : []).map(asRecord);
  });
  const shape = shapes.find((item) => String(item.name) === name);
  if (!shape) throw new Error(`PowerPoint 增量更新未保留对象 ${name}: ${JSON.stringify(shapes)}`);
  return shape;
}

function assertPresentationLayout(before: Record<string, unknown>, after: Record<string, unknown>): void {
  for (const property of ["left", "top", "width", "height", "rotation", "zOrder"] as const) {
    const expected = Number(before[property]);
    const actual = Number(after[property]);
    if (!Number.isFinite(expected) || !Number.isFinite(actual) || Math.abs(expected - actual) > 1) {
      throw new Error(`PowerPoint 增量更新未保留 ${property}: ${expected} -> ${actual}`);
    }
  }
}

async function assertTargetedRelink(
  bridge: OfficeActionBridge,
  app: "word" | "presentation",
  filePath: string,
  linkId: string,
  temporarySource: string,
  originalSource: string,
): Promise<void> {
  await runAction(bridge, {
    app, action: "edit", operation: "relinkLinkedOfficeContent", filePath,
    params: { linkId, sourcePath: temporarySource },
  });
  const moved = linksFrom(await runAction(bridge, {
    app, action: "inspect", operation: "inspectLinkedOfficeContent", filePath,
  }));
  const movedLink = moved.find((item) => String(item.linkId) === linkId);
  if (!movedLink || !String(movedLink.source || "").toLowerCase().includes(path.basename(temporarySource).toLowerCase())) {
    throw new Error(`${app} 定向重绑未切换来源: ${JSON.stringify(moved)}`);
  }
  const refreshed = await runAction(bridge, {
    app, action: "edit", operation: "refreshLinkedOfficeContent", filePath,
    params: { linkId },
  });
  assertRefresh(refreshed, `${app} 定向刷新`);
  if (Number(operationData(refreshed).updated) !== 1) throw new Error(`${app} 定向刷新影响了多个链接对象`);
  await runAction(bridge, {
    app, action: "edit", operation: "relinkLinkedOfficeContent", filePath,
    params: { linkId, sourcePath: originalSource },
  });
}

async function verifyDirtyDocumentTransaction(input: {
  bridge: OfficeActionBridge;
  documentBridge: OfficeDocumentComBridge;
  workflowRoot: string;
  transactionRoot: string;
  filePath: string;
  instanceId: string;
}): Promise<{ transactionId: string; conflicts: number }> {
  await markOpenWordBookmarkDirty(input.filePath, input.instanceId, "DirtyBaseline", "脏文档基线");
  const workflow = await runOfficeWorkflow(input.bridge, [{
    app: "word",
    action: "edit",
    operation: "applyTrackedChanges",
    filePath: input.filePath,
    params: {
      instanceId: input.instanceId,
      edits: [{ command: "replaceBookmark", name: "DirtyBaseline", text: "事务修改结果" }],
      keepTracking: false,
    },
  }], {
    workflowRoot: input.workflowRoot,
    transactionRoot: input.transactionRoot,
    prepareTransaction: (filePaths) => input.documentBridge.prepareTransaction(filePaths),
    restoreTransaction: (files) => input.documentBridge.restoreTransactionFiles(files),
  });
  if (workflow.status !== "done" || !workflow.transactionId) {
    throw new Error(`脏文档事务执行失败: ${JSON.stringify(workflow)}`);
  }

  let instanceId = await currentDocumentInstanceId(input.documentBridge, "word", input.filePath);
  await assertWordBookmarkInFile(input.bridge, input.filePath, instanceId, "DirtyBaseline", "事务修改结果");
  const restoreOptions = {
    prepareFiles: (filePaths: string[]) => input.documentBridge.prepareTransaction(filePaths),
    restoreFiles: (files: Parameters<OfficeDocumentComBridge["restoreTransactionFiles"]>[0]) => input.documentBridge.restoreTransactionFiles(files),
  };
  await undoOfficeTransaction(input.transactionRoot, workflow.transactionId, restoreOptions);
  instanceId = await currentDocumentInstanceId(input.documentBridge, "word", input.filePath);
  await assertWordBookmarkInFile(input.bridge, input.filePath, instanceId, "DirtyBaseline", "脏文档基线");

  await redoOfficeTransaction(input.transactionRoot, workflow.transactionId, input.bridge, restoreOptions);
  instanceId = await currentDocumentInstanceId(input.documentBridge, "word", input.filePath);
  await assertWordBookmarkInFile(input.bridge, input.filePath, instanceId, "DirtyBaseline", "事务修改结果");

  await runAction(input.bridge, {
    app: "word", action: "edit", operation: "applyTrackedChanges", filePath: input.filePath,
    params: { instanceId, edits: [{ command: "replaceBookmark", name: "DirtyBaseline", text: "事务外修改" }], keepTracking: false },
  });
  const conflicted = await undoOfficeTransaction(input.transactionRoot, workflow.transactionId, restoreOptions);
  if (conflicted.status !== "conflicted" || !conflicted.conflicts?.length) {
    throw new Error(`事务外修改未被拦截: ${JSON.stringify(conflicted)}`);
  }
  instanceId = await currentDocumentInstanceId(input.documentBridge, "word", input.filePath);
  await assertWordBookmarkInFile(input.bridge, input.filePath, instanceId, "DirtyBaseline", "事务外修改");

  await undoOfficeTransaction(input.transactionRoot, workflow.transactionId, { ...restoreOptions, force: true });
  instanceId = await currentDocumentInstanceId(input.documentBridge, "word", input.filePath);
  await assertWordBookmarkInFile(input.bridge, input.filePath, instanceId, "DirtyBaseline", "脏文档基线");
  return { transactionId: workflow.transactionId, conflicts: conflicted.conflicts.length };
}

async function markOpenWordBookmarkDirty(filePath: string, instanceId: string, name: string, text: string): Promise<void> {
  const output = await executePowerShell(`
${officeInstanceDiscoveryScript()}
${psVar("_filePath", filePath)}
${psVar("_instanceId", instanceId)}
${psVar("_bookmarkName", name)}
${psVar("_bookmarkText", text)}
$handle = Resolve-OfficeDocumentHandle 'word' $_filePath $_instanceId '' 0
$document = $handle.document
$start = [Math]::Max(0, $document.Content.End - 1)
$target = $document.Range($start, $start)
$target.InsertAfter($_bookmarkText)
if ($document.Bookmarks.Exists($_bookmarkName)) { $document.Bookmarks.Item($_bookmarkName).Delete() }
[void]$document.Bookmarks.Add($_bookmarkName, $document.Range($start, $start + $_bookmarkText.Length))
[pscustomobject]@{ saved = [bool]$document.Saved; instanceId = [string]$handle.instanceId } | ConvertTo-Json -Compress
`);
  const result = JSON.parse(output) as { saved: boolean };
  if (result.saved) throw new Error("未能构造未保存的 Word 文档状态");
}

async function currentDocumentInstanceId(
  bridge: OfficeDocumentComBridge,
  app: "excel" | "word" | "presentation",
  filePath: string,
): Promise<string> {
  const matches = (await bridge.listDocuments(app)).filter((document) => samePath(filePath, document.fullName));
  if (matches.length !== 1) throw new Error(`恢复后无法唯一定位 ${app} 文档: ${JSON.stringify(matches)}`);
  return matches[0].instanceId;
}

async function assertWordBookmarkInFile(
  bridge: OfficeActionBridge,
  filePath: string,
  instanceId: string,
  name: string,
  text: string,
): Promise<void> {
  assertWordBookmark(await runAction(bridge, {
    app: "word", action: "inspect", operation: "inspectReferences", filePath, params: { instanceId },
  }), name, text);
}

async function expectMissing(filePath: string): Promise<void> {
  try { await access(filePath); } catch { return; }
  throw new Error(`事务撤销后产物仍存在: ${filePath}`);
}

async function openFixtures(excelPaths: string[], wordPaths: string[], presentationPaths: string[]): Promise<OwnedProcesses> {
  const output = await executePowerShell(`
${psVar("_excelPath1", excelPaths[0])}
${psVar("_excelPath2", excelPaths[1])}
${psVar("_wordPath1", wordPaths[0])}
${psVar("_wordPath2", wordPaths[1])}
${psVar("_presentationPath1", presentationPaths[0])}
${psVar("_presentationPath2", presentationPaths[1])}
$beforeExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$beforeWord = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$beforePresentation = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
if (-not ('WenggeReliabilityWindow' -as [type])) { Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class WenggeReliabilityWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }' }
function Get-AppProcessId($app) {
  $window = [int64]0
  try { $window = [int64]$app.Hwnd } catch {}
  if ($window -eq 0) { try { $window = [int64]$app.HWND } catch {} }
  if ($window -eq 0) { try { $window = [int64]$app.ActiveWindow.Hwnd } catch {} }
  if ($window -eq 0) { try { $window = [int64]$app.ActiveWindow.HWND } catch {} }
  if ($window -eq 0) { return 0 }
  $pidValue = [uint32]0
  [void][WenggeReliabilityWindow]::GetWindowThreadProcessId([IntPtr]$window, [ref]$pidValue)
  return [int]$pidValue
}
function New-TestComApplication([string]$progId) {
  $lastError = $null
  foreach ($attempt in 1..3) {
    $candidate = $null
    try {
      $candidate = New-Object -ComObject $progId
      $null = [string]$candidate.Version
      return $candidate
    } catch {
      $lastError = $_
      if ($null -ne $candidate) {
        try { $candidate.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) } catch {}
      }
      Start-Sleep -Milliseconds (300 * $attempt)
    }
  }
  throw $lastError
}
$apps = @()
try {
  $excel1 = New-TestComApplication 'Excel.Application'; $apps += $excel1; $excel1.Visible = $true; [void]$excel1.Workbooks.Open($_excelPath1)
  $excel2 = New-TestComApplication 'Excel.Application'; $apps += $excel2; $excel2.Visible = $true; [void]$excel2.Workbooks.Open($_excelPath2)
  $word1 = New-TestComApplication 'Word.Application'; $apps += $word1; $word1.Visible = $true; $word1.DisplayAlerts = 0; [void]$word1.Documents.Open($_wordPath1)
  $word2 = New-TestComApplication 'Word.Application'; $apps += $word2; $word2.Visible = $true; $word2.DisplayAlerts = 0; [void]$word2.Documents.Open($_wordPath2)
  $presentation1 = New-TestComApplication 'PowerPoint.Application'; $apps += $presentation1; $presentation1.Visible = -1; [void]$presentation1.Presentations.Open($_presentationPath1)
  $presentation2 = New-TestComApplication 'PowerPoint.Application'; $apps += $presentation2; $presentation2.Visible = -1; [void]$presentation2.Presentations.Open($_presentationPath2)
  $excelPids = @()
  $excelPids += Get-AppProcessId $excel1
  $excelPids += Get-AppProcessId $excel2
  $excelPids = @($excelPids | Where-Object { $_ -notin $beforeExcel } | Select-Object -Unique)
  $wordPids = @()
  $wordPids += Get-AppProcessId $word1
  $wordPids += Get-AppProcessId $word2
  $wordPids = @($wordPids | Where-Object { $_ -notin $beforeWord } | Select-Object -Unique)
  $presentationPids = @()
  $presentationPids += Get-AppProcessId $presentation1
  $presentationPids += Get-AppProcessId $presentation2
  $presentationPids += @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $beforePresentation } | ForEach-Object { [int]$_.Id })
  $presentationPids = @($presentationPids | Where-Object { $_ -notin $beforePresentation } | Select-Object -Unique)
  [pscustomobject]@{ excel = @($excelPids); word = @($wordPids); presentation = @($presentationPids) } | ConvertTo-Json -Depth 4 -Compress
} catch {
  foreach ($app in $apps) { try { $app.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch {} }
  throw
}
`, 120000);
  const parsed = JSON.parse(output) as Record<keyof OwnedProcesses, number | number[]>;
  return { excel: toProcessIds(parsed.excel), word: toProcessIds(parsed.word), presentation: toProcessIds(parsed.presentation) };
}

async function closeFixtures(processes: OwnedProcesses): Promise<void> {
  await executePowerShell(`
${psVar("_processesJson", JSON.stringify(processes))}
$owned = ConvertFrom-Json $_processesJson
$ownedIds = @(@($owned.excel) + @($owned.word) + @($owned.presentation) | ForEach-Object { [int]$_ } | Where-Object { $_ -gt 0 } | Select-Object -Unique)
foreach ($processId in $ownedIds) { try { Stop-Process -Id $processId -Force -ErrorAction Stop } catch {} }
`);
}

async function listMicrosoftOfficeProcessIds(): Promise<number[]> {
  const output = await executePowerShell("@(Get-Process -Name EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id }) | ConvertTo-Json -Compress");
  if (!output) return [];
  const parsed = JSON.parse(output) as number | number[];
  return Array.isArray(parsed) ? parsed : [parsed];
}

async function listVisibleWpsProcessIds(): Promise<number[]> {
  const output = await executePowerShell("@(Get-Process -Name wps, et, wpp -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | ForEach-Object { [int]$_.Id }) | ConvertTo-Json -Compress");
  return output ? toProcessIds(JSON.parse(output) as number | number[]) : [];
}

async function assertProcessesStillRunning(processIds: number[], label: string): Promise<void> {
  if (processIds.length === 0) return;
  const running = await executePowerShell(`${psVar("_idsJson", JSON.stringify(processIds))}\n@((ConvertFrom-Json $_idsJson) | Where-Object { Get-Process -Id ([int]$_) -ErrorAction SilentlyContinue }) | ConvertTo-Json -Compress`);
  const current = running ? toProcessIds(JSON.parse(running) as number | number[]) : [];
  const missing = processIds.filter((id) => !current.includes(id));
  if (missing.length > 0) throw new Error(`${label}被意外关闭: ${missing.join(", ")}`);
}

async function assertOwnedMicrosoftOfficeProcessesStopped(owned: OwnedProcesses): Promise<void> {
  const ownedIds = [...owned.excel, ...owned.word, ...owned.presentation];
  const running = await waitForOfficeProcesses((ids) => ownedIds.filter((id) => ids.includes(id)));
  if (running.length > 0) throw new Error(`Office 冒烟测试登记的进程未退出: ${running.join(", ")}`);
}

async function assertNoUnexpectedMicrosoftOfficeProcesses(before: number[]): Promise<void> {
  const leaked = await waitForOfficeProcesses((ids) => ids.filter((id) => !before.includes(id)));
  if (leaked.length > 0) throw new Error(`Office 冒烟测试遗留了未登记进程: ${leaked.join(", ")}`);
}

async function waitForOfficeProcesses(selectPending: (processIds: number[]) => number[]): Promise<number[]> {
  let pending: number[] = [];
  for (let attempt = 0; attempt < 20; attempt += 1) {
    pending = selectPending(await listMicrosoftOfficeProcessIds());
    if (pending.length === 0) return [];
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  return pending;
}

function samePath(expected: string, actual?: string): boolean {
  return Boolean(actual) && path.resolve(expected).toLowerCase() === path.resolve(actual!).toLowerCase();
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
}

function toProcessIds(value: number | number[]): number[] {
  return (Array.isArray(value) ? value : [value]).map(Number).filter((id) => Number.isSafeInteger(id) && id > 0);
}

function routeToMicrosoftOffice(input: OfficeActionInput): OfficeActionInput {
  const params = { ...input.params };
  if (["exportRangeToWord", "exportRangeToPresentation", "buildReportPackage"].includes(input.operation)) {
    params.sourceHost = "excel";
    params.wordHost = "word";
    params.presentationHost = "powerpoint";
  } else {
    params.host = input.app === "excel" ? "excel" : input.app === "word" ? "word" : "powerpoint";
  }
  return { ...input, params };
}

function assertIndependentInstances(documents: Array<{ processId?: number; instanceId: string }>, label: string): void {
  const instances = new Set(documents.map((document) => document.instanceId));
  const knownProcessIds = documents.map((document) => document.processId).filter((id): id is number => Boolean(id));
  if (instances.size !== documents.length || new Set(knownProcessIds).size !== knownProcessIds.length) {
    throw new Error(`${label} 未识别为独立进程实例: ${JSON.stringify(documents)}`);
  }
}

void main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exitCode = 1;
});
