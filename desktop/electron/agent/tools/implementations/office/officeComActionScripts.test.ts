import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";

import type { OfficeActionInput } from "../../officeCore/types";
import { buildComScript } from "./officeComActionScripts";

const ACTIONS: OfficeActionInput[] = [
  { app: "excel", action: "insert", operation: "createPivotTable", filePath: "C:\\tmp\\book.xlsx", target: "range:Sheet1!A1:D20", params: { rowFields: ["部门"], dataFields: [{ name: "金额", function: "sum" }] } },
  { app: "excel", action: "insert", operation: "addSlicer", filePath: "C:\\tmp\\book.xlsx", params: { field: "部门" } },
  { app: "excel", action: "edit", operation: "createPowerQuery", filePath: "C:\\tmp\\book.xlsx", params: { name: "查询1", mFormula: "let Source = #table({}, {}) in Source" } },
  { app: "excel", action: "inspect", operation: "inspectPowerQueries", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "edit", operation: "managePowerQuery", filePath: "C:\\tmp\\book.xlsx", params: { command: "refresh", name: "查询1" } },
  { app: "excel", action: "inspect", operation: "inspectCharts", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "style", operation: "formatChart", filePath: "C:\\tmp\\book.xlsx", params: { title: "趋势" } },
  { app: "excel", action: "inspect", operation: "inspectWorkbookObjects", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "edit", operation: "manageWorkbookObject", filePath: "C:\\tmp\\book.xlsx", params: { objectType: "name", command: "upsert", name: "Total", refersTo: "=Sheet1!$B$2:$B$4" } },
  { app: "excel", action: "inspect", operation: "captureWorkbookTemplate", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "inspect", operation: "inspectWorkbookFormatting", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "style", operation: "applyWorkbookTemplate", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "edit", operation: "configurePrint", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "inspect", operation: "inspectPrintSettings", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "edit", operation: "exportSheetsToPdf", filePath: "C:\\tmp\\book.xlsx", params: { sheetNames: ["Sheet1"] } },
  { app: "excel", action: "inspect", operation: "traceFormulaDependencies", filePath: "C:\\tmp\\book.xlsx", target: "range:Sheet1!D2" },
  { app: "excel", action: "inspect", operation: "inspectFormulaDependencies", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "edit", operation: "repairFormulaReferences", filePath: "C:\\tmp\\book.xlsx", params: { replacements: [{ find: "#REF!", replace: "A1" }] } },
  { app: "excel", action: "edit", operation: "convertFormulasToValues", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "inspect", operation: "inspectFormulaBackups", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "edit", operation: "restoreFormulas", filePath: "C:\\tmp\\book.xlsx", params: { backupId: "backup-1" } },
  { app: "excel", action: "inspect", operation: "inspectFormulaProtection", filePath: "C:\\tmp\\book.xlsx" },
  { app: "excel", action: "edit", operation: "manageFormulaProtection", filePath: "C:\\tmp\\book.xlsx", params: { command: "lock" } },
  { app: "excel", action: "snapshot", operation: "snapshot", filePath: "C:\\tmp\\book.xlsx", target: "range:Sheet1!A1:D10" },
  { app: "word", action: "style", operation: "formatLongDocument", filePath: "C:\\tmp\\report.docx" },
  { app: "word", action: "inspect", operation: "inspectDocumentFormatting", filePath: "C:\\tmp\\report.docx" },
  { app: "word", action: "inspect", operation: "inspectReferences", filePath: "C:\\tmp\\report.docx" },
  { app: "word", action: "insert", operation: "manageReferences", filePath: "C:\\tmp\\report.docx", params: { command: "addFootnote", text: "来源" } },
  { app: "word", action: "inspect", operation: "inspectRevisions", filePath: "C:\\tmp\\report.docx" },
  { app: "word", action: "edit", operation: "manageRevisions", filePath: "C:\\tmp\\report.docx", params: { command: "acceptAll" } },
  { app: "word", action: "edit", operation: "applyTrackedChanges", filePath: "C:\\tmp\\report.docx", params: { edits: [{ command: "replace", find: "旧", replace: "新" }] } },
  { app: "word", action: "edit", operation: "compareDocuments", filePath: "C:\\tmp\\report.docx", params: { revisedFilePath: "C:\\tmp\\revised.docx" } },
  { app: "word", action: "edit", operation: "prepareMailMergeTemplate", filePath: "C:\\tmp\\template.docx", params: { fields: [{ placeholder: "{{name}}", field: "Name" }] } },
  { app: "word", action: "edit", operation: "mailMerge", filePath: "C:\\tmp\\template.docx", params: { dataSourcePath: "C:\\tmp\\data.xlsx" } },
  { app: "word", action: "edit", operation: "batchMailMerge", filePath: "C:\\tmp\\template.docx", params: { dataSourcePath: "C:\\tmp\\data.xlsx" } },
  { app: "word", action: "inspect", operation: "inspectContentControls", filePath: "C:\\tmp\\template.docx" },
  { app: "word", action: "edit", operation: "populateContentControls", filePath: "C:\\tmp\\template.docx", params: { values: { name: "测试" } } },
  { app: "word", action: "edit", operation: "manageContentControls", filePath: "C:\\tmp\\template.docx", params: { command: "add", type: "text", tag: "name" } },
  { app: "presentation", action: "style", operation: "applyMasterBranding", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "insert", operation: "insertTable", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "style", operation: "layoutElements", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "inspect", operation: "inspectPresentationTheme", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "inspect", operation: "inspectSlideElements", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "inspect", operation: "inspectAnimations", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "inspect", operation: "inspectSpeakerNotes", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "edit", operation: "configureAnimations", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "edit", operation: "configureSlideShow", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "edit", operation: "setSpeakerNotes", filePath: "C:\\tmp\\slides.pptx", params: { text: "讲稿" } },
  { app: "presentation", action: "edit", operation: "exportHandouts", filePath: "C:\\tmp\\slides.pptx" },
  { app: "word", action: "inspect", operation: "inspectLinkedOfficeContent", filePath: "C:\\tmp\\report.docx" },
  { app: "word", action: "edit", operation: "refreshLinkedOfficeContent", filePath: "C:\\tmp\\report.docx" },
  { app: "word", action: "edit", operation: "relinkLinkedOfficeContent", filePath: "C:\\tmp\\report.docx", params: { linkId: "sales", sourcePath: "C:\\tmp\\moved.xlsx" } },
  { app: "presentation", action: "inspect", operation: "inspectLinkedOfficeContent", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "edit", operation: "refreshLinkedOfficeContent", filePath: "C:\\tmp\\slides.pptx" },
  { app: "presentation", action: "edit", operation: "relinkLinkedOfficeContent", filePath: "C:\\tmp\\slides.pptx", params: { linkId: "sales", sourcePath: "C:\\tmp\\moved.xlsx" } },
  { app: "excel", action: "insert", operation: "exportRangeToWord", filePath: "C:\\tmp\\book.xlsx", target: "range:Sheet1!A1:D10", params: { linked: true } },
  { app: "excel", action: "insert", operation: "exportRangeToPresentation", filePath: "C:\\tmp\\book.xlsx", params: { linked: true, sourceType: "chart", chartName: "销售趋势" } },
  { app: "excel", action: "insert", operation: "buildReportPackage", filePath: "C:\\tmp\\book.xlsx", target: "range:Sheet1!A1:D10" },
];

describe("Office COM action scripts", () => {
  it("does not register or quit a pre-existing shared WPS process", () => {
    for (const input of [
      { app: "excel", action: "inspect", operation: "inspectWorkbookFormatting", filePath: "C:\\tmp\\book.xlsx" },
      { app: "word", action: "inspect", operation: "inspectDocumentFormatting", filePath: "C:\\tmp\\report.docx" },
      { app: "presentation", action: "inspect", operation: "inspectPresentationTheme", filePath: "C:\\tmp\\slides.pptx" },
    ] satisfies OfficeActionInput[]) {
      const script = buildComScript(input);
      expect(script).toContain("-in $officeProcessIdsBefore");
      expect(script).toContain("$createdOfficeProcessId = [uint32]0");
      expect(script).toContain("$ownsOfficeProcess = $createdOfficeProcessId -gt 0");
      expect(script).toContain("if ($createdApp) {");
      expect(script).toContain("if ($ownsOfficeProcess) {");
    }
  });

  it("allows all file-scoped actions to target Microsoft Office or WPS explicitly", () => {
    const excel = buildComScript({ app: "excel", action: "inspect", operation: "inspectCharts", filePath: "C:\\tmp\\book.xlsx", params: { host: "wps" } });
    const word = buildComScript({ app: "word", action: "inspect", operation: "inspectReferences", filePath: "C:\\tmp\\report.docx", params: { host: "word" } });
    const presentation = buildComScript({ app: "presentation", action: "inspect", operation: "inspectPresentationTheme", filePath: "C:\\tmp\\slides.pptx", params: { host: "wps" } });

    expect(excel).toContain("$progIds = @('Ket.Application')");
    expect(word).toContain("$progIds = @('Word.Application')");
    expect(presentation).toContain("$progIds = @('Wpp.Application', 'Kwpp.Application')");
    expect(presentation).not.toContain("$progIds = @('PowerPoint.Application', 'Wpp.Application', 'Kwpp.Application')");
    expect(presentation).toContain("$operationData.progId = $progId");
  });

  it("isolates unopened cross-Office targets and honors per-application hosts", () => {
    const script = buildComScript({
      app: "excel",
      action: "insert",
      operation: "buildReportPackage",
      filePath: "C:\\tmp\\book.xlsx",
      params: { sourceHost: "excel", wordHost: "word", presentationHost: "powerpoint" },
    });

    expect(script).toContain("$excel = New-OfficeComObject @('Excel.Application')");
    expect(script).toContain("$word = New-OfficeComObject @('Word.Application')");
    expect(script).toContain("$powerPoint = New-OfficeComObject @('PowerPoint.Application')");
    const creator = script.slice(script.indexOf("function New-OfficeComObject"), script.indexOf("function Find-OfficeDocumentHandle"));
    expect(creator).not.toContain("GetActiveObject");
    expect(creator).toContain("foreach ($attempt in 1..3)");
    expect(creator).toContain("Write-Output -NoEnumerate $candidate");
    expect(script).toContain("if ($excelCreatedApp) { try { $excel.Quit() }");
  });

  it("reuses an already-open source workbook without closing the user document", () => {
    const script = buildComScript({
      app: "excel",
      action: "insert",
      operation: "exportRangeToWord",
      filePath: "C:\\tmp\\book.xlsx",
      params: { linked: true, sourceInstanceId: "excel:100:200" },
    });

    expect(script).toContain("$openedWorkbook = $false");
    expect(script).toContain("if ($sourceInstanceId) { Find-OfficeDocumentHandle");
    expect(script).toContain("[System.IO.Path]::GetFullPath([string]$candidateWorkbook.FullName)");
    expect(script).toContain("if ($null -eq $workbook)");
    expect(script).toContain("if ($openedWorkbook) { try { $workbook.Close($false) }");
    expect(script).not.toContain("if ($null -ne $workbook) { try { $workbook.Close($false) }");
  });

  it("creates linked OLE objects and refreshes them in place", () => {
    const wordExport = buildComScript({ app: "excel", action: "insert", operation: "exportRangeToWord", filePath: "C:\\tmp\\book.xlsx", params: { linked: true } });
    const pptExport = buildComScript({ app: "excel", action: "insert", operation: "exportRangeToPresentation", filePath: "C:\\tmp\\book.xlsx", params: { linked: true, sourceType: "chart", chartName: "销售趋势" } });
    const wordRefresh = buildComScript({ app: "word", action: "edit", operation: "refreshLinkedOfficeContent", filePath: "C:\\tmp\\report.docx" });
    const pptRefresh = buildComScript({ app: "presentation", action: "edit", operation: "refreshLinkedOfficeContent", filePath: "C:\\tmp\\slides.pptx" });

    expect(wordExport).toContain("$insert.PasteSpecial(0, $true, 0, $false, 0)");
    expect(wordExport).toContain("$_.Exception.HResult -ne -2147418111");
    expect(wordExport).toContain("$attempt -le 5");
    expect(wordExport).toContain("Find-WordInlineShapeAt $document $contentStart");
    expect(wordExport).not.toContain("$document.InlineShapes.Item($document.InlineShapes.Count)");
    expect(wordExport).toContain("$document.Bookmarks.Add([string]$destination.bookmark, $managedRange)");
    expect(wordExport).toContain("Set-WordLinkMetadata $document");
    expect(wordExport).toContain("ConvertTo-Json -InputObject ([object[]]$uniqueIds)");
    expect(wordExport).toContain("$existingWordIds = @(Get-WordManagedIds $wordDoc)");
    expect(wordExport).toContain("[void]$insert.Delete()");
    expect(wordExport).toContain("$document.InlineShapes.Item($shapeIndex)");
    expect(wordExport).toContain("$candidateShape.Range.Start -ge $start");
    expect(wordExport).toContain("$operationData.wordReplacement");
    expect(pptExport).toContain("$chartObject.Chart.ChartArea.Copy()");
    expect(pptExport).toContain("$slide.Shapes.PasteSpecial(10, 0, '', 0, '', -1)");
    expect(pptExport).toContain("[int]$candidateShape.Id -notin $beforeShapeIds");
    expect(pptExport).toContain("$shape = if ($createdShapeCount -eq 1) { $shapeHolder.value }");
    expect(pptExport).toContain("foreach ($candidateShape in $slide.Shapes)");
    expect(pptExport).toContain("[int]$candidateShape.Id -notin $beforeShapeIds");
    expect(pptExport).not.toContain("$slide.Shapes.Item($slide.Shapes.Count)");
    expect(pptExport).toContain("Set-ObjectTag $shape 'WENGGE_LINK_ID' $linkId");
    for (const refreshScript of [wordRefresh, pptRefresh]) {
      expect(refreshScript).toContain("$item.LinkFormat.Update()");
      expect(refreshScript).not.toContain("Shapes.Add");
      expect(refreshScript).not.toContain("InlineShapes.Add");
    }
  });

  it("requires stable link ids for incremental report updates", () => {
    expect(() => buildComScript({
      app: "excel",
      action: "insert",
      operation: "exportRangeToWord",
      filePath: "C:\\tmp\\book.xlsx",
      params: { updateExisting: true },
    })).toThrow("params.linkId");
    expect(() => buildComScript({
      app: "excel",
      action: "insert",
      operation: "buildReportPackage",
      filePath: "C:\\tmp\\book.xlsx",
      params: { updateExisting: true, sections: [{ linkId: "sales" }, { range: "A1:D10" }] },
    })).toThrow("第 2 个 section");
  });

  it("updates managed report regions and supports targeted refresh and relink", () => {
    const report = buildComScript({
      app: "excel",
      action: "insert",
      operation: "buildReportPackage",
      filePath: "C:\\tmp\\book.xlsx",
      params: {
        updateExisting: true,
        sections: [{ linkId: "sales", sheetName: "Sheet1", range: "A1:D10" }],
      },
    });
    const wordRelink = buildComScript({ app: "word", action: "edit", operation: "relinkLinkedOfficeContent", filePath: "C:\\tmp\\report.docx", params: { linkId: "sales_word", sourcePath: "C:\\tmp\\moved.xlsx" } });
    const pptRelink = buildComScript({ app: "presentation", action: "edit", operation: "relinkLinkedOfficeContent", filePath: "C:\\tmp\\slides.pptx", params: { linkId: "sales_ppt", sourcePath: "C:\\tmp\\moved.xlsx" } });

    expect(report).toContain("$updateExisting = $actionParams.updateExisting -eq $true");
    expect(report).toContain("Get-WordManagedIds");
    expect(report).toContain("WENGGE_MANAGED_SLIDE");
    expect(report).toContain("$existingShape.Delete()");
    expect(report).toContain("$shape.Left = $geometry.left");
    expect(report).toContain("$shape.Rotation = $geometry.rotation");
    expect(report).toContain("$shape.LockAspectRatio = $geometry.lockAspectRatio");
    expect(report).toContain("$shape.ZOrder(3)");
    expect(report).toContain("$operationData.manifest");
    for (const script of [wordRelink, pptRelink]) {
      expect(script).toContain("$linkIdFilter");
      expect(script).toContain("LinkFormat.SourceFullName = Get-RelinkSource");
      expect(script).toContain("Open-LinkedExcelSource $newSourcePath");
      expect(script).toContain("linked-content-relink");
    }
    expect(wordRelink).toContain("Find-ContainingWordBookmark");
    expect(wordRelink).toContain("WENGGE_META_");
    expect(wordRelink).toContain("Get-WordLinkKey");
    expect(wordRelink).toContain("$canMapManagedIdsByOrder");
    expect(wordRelink).toContain("$managedIds[$managedLinkOrdinal - 1]");
    expect(wordRelink).toContain("$managedIds = @()\n  try { $managedIds = @(");
    expect(wordRelink).toContain("$_.MainWindowHandle -eq 0");
    expect(pptRelink).toContain("$suffix.Replace('[' + $oldName + ']', '[' + $newName + ']')");
  });

  it("reports unsupported precise PowerPoint shape properties without aborting other edits", () => {
    const script = buildComScript({
      app: "presentation",
      action: "style",
      operation: "layoutElements",
      filePath: "C:\\tmp\\linked.pptx",
      params: { mode: "precise", edits: [{ shapeName: "Linked", left: 20, rotation: 5 }] },
    });

    expect(script).toContain("$operationData.editFailures = $script:layoutEditFailures");
    expect(script).toContain("property = 'rotation'");
    expect(script).toContain("if ($shapeEdited) { $count++ }");
  });

  it("updates only Word reference fields and does not refresh external links after adding a bookmark", () => {
    const script = buildComScript({
      app: "word",
      action: "insert",
      operation: "manageReferences",
      filePath: "C:\\tmp\\linked-report.docx",
      params: { command: "addBookmark", name: "ManualKeep" },
    });

    expect(script).toContain("$referenceFieldTypes = @(3, 5, 9, 10, 11, 12, 13, 37, 72)");
    expect(script).toContain("if ([int]$field.Type -in $referenceFieldTypes)");
    expect(script).toContain("if ($command -in @('addCaption', 'addCrossReference', 'addTableOfFigures'))");
    expect(script).not.toContain("$current.Fields.Update()");
    expect(script).toContain("text = ([string]$bookmark.Range.Text).Trim()");
  });

  it.each(ACTIONS)("builds a concrete script for $app/$operation", (input) => {
    const script = buildComScript(input);
    expect(script).not.toContain('throw "不支持的 Excel COM 操作: $_operation"');
    expect(script).not.toContain('throw "不支持的 Word COM 操作: $_operation"');
    expect(script).not.toContain('throw "不支持的 PowerPoint COM 操作: $_operation"');
    expect(script).toContain("$_filePath = [System.Text.Encoding]::Unicode.GetString");
    expect(script).toContain("ConvertTo-Json");
  });

  it.skipIf(process.platform !== "win32")("produces PowerShell scripts without parser errors", () => {
    const tempDir = mkdtempSync(path.join(os.tmpdir(), "office-com-scripts-"));
    try {
      for (const [index, input] of ACTIONS.entries()) {
        const scriptPath = path.join(tempDir, `${index}-${input.app}-${input.operation}.ps1`);
        writeFileSync(scriptPath, buildComScript(input), "utf8");
      }
      const escapedDir = tempDir.replace(/'/g, "''");
      const command = [
        "$failed = $false",
        `Get-ChildItem -LiteralPath '${escapedDir}' -Filter '*.ps1' | ForEach-Object {`,
        "  $fileName = $_.Name",
        "  $source = Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8",
        "  $tokens = $null; $errors = $null",
        "  [void][System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$errors)",
        "  if ($errors.Count -gt 0) { $failed = $true; $errors | ForEach-Object { '{0}:{1}: {2}' -f $fileName, $_.Extent.StartLineNumber, $_.Message } }",
        "}",
        "if ($failed) { exit 1 }",
      ].join("; ");
      const parsed = spawnSync("powershell.exe", ["-NoProfile", "-Command", command], { encoding: "utf8" });
      expect(parsed.status, `${parsed.stdout}${parsed.stderr}`).toBe(0);
    } finally {
      rmSync(tempDir, { recursive: true, force: true });
    }
  }, 30_000);
});
