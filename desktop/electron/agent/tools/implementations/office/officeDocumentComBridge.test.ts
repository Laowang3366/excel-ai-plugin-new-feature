import { beforeEach, describe, expect, it, vi } from "vitest";

import { executePowerShell } from "../../../automation/powershell";
import { OfficeDocumentComBridge } from "./officeDocumentComBridge";

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return { ...actual, executePowerShell: vi.fn() };
});

const executePowerShellMock = vi.mocked(executePowerShell);

describe("OfficeDocumentComBridge", () => {
  beforeEach(() => executePowerShellMock.mockReset());

  it("lists open documents across Office applications", async () => {
    executePowerShellMock.mockResolvedValue(
      JSON.stringify([
        {
          app: "excel",
          name: "book.xlsx",
          fullName: "C:\\tmp\\book.xlsx",
          index: 1,
          active: true,
          progId: "Excel.Application",
        },
        {
          app: "word",
          name: "report.docx",
          fullName: "C:\\tmp\\report.docx",
          index: 1,
          active: true,
          progId: "Word.Application",
        },
      ]),
    );

    const result = await new OfficeDocumentComBridge().listDocuments();

    expect(result).toHaveLength(2);
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("[WenggeRotDiscovery]::Enumerate()");
    expect(script).toContain("Get-AllOfficeDocumentHandles");
    expect(script).toContain("instanceId = [string]$handle.instanceId");
    expect(script).toContain("if ($null -eq $saved) { return '' }");
    expect(script).toContain("$document.ActiveWindow.Hwnd");
    expect(script).toContain("$document.Windows.Item(1).HWND");
    expect(script).toContain("$officeHost = if");
    expect(script).toContain("$seenPaths.ContainsKey($pathKey)");
    expect(script).toContain("if (-not $documentName -or -not $applicationName) { return $null }");
    expect(script).toContain("$applicationName -match '(?i)Microsoft|Excel|Word|PowerPoint'");
    expect(script).not.toContain("$host = if ($processName");
    expect(script).toContain("@('Excel.Application', 'Ket.Application')");
  });

  it("activates a document using a stable full-path selector", async () => {
    executePowerShellMock.mockResolvedValue(
      JSON.stringify({
        app: "excel",
        name: "book.xlsx",
        fullName: "C:\\tmp\\book.xlsx",
        index: 1,
        active: true,
        progId: "Excel.Application",
      }),
    );

    const result = await new OfficeDocumentComBridge().activateDocument({
      app: "excel",
      filePath: "C:\\tmp\\book.xlsx",
    });

    expect(result.active).toBe(true);
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$handle.document.Activate()");
    expect(script).toContain("([string]$selector.instanceId)");
    expect(script).toContain("if ($candidates.Count -gt 1)");
    expect(script).toContain("GetFullPath");
  });

  it("lists document objects with encoded stable locators", async () => {
    executePowerShellMock.mockResolvedValue(
      JSON.stringify([
        {
          app: "excel",
          documentPath: "C:\\tmp\\book.xlsx",
          kind: "chart",
          name: "销售 趋势/年度",
          locator:
            "chart:%E6%95%B0%E6%8D%AE/%E9%94%80%E5%94%AE%20%E8%B6%8B%E5%8A%BF%2F%E5%B9%B4%E5%BA%A6",
          parent: "数据",
          index: 2,
        },
      ]),
    );

    const result = await new OfficeDocumentComBridge().listObjects({
      app: "excel",
      filePath: "C:\\tmp\\book.xlsx",
      kind: "chart",
    });

    expect(result[0].locator).toContain("%2F");
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("[Uri]::EscapeDataString");
    expect(script).toContain("$target.Worksheets");
    expect(script).toContain("$_kind -ne $kind");
  });

  it("activates an object by full document path and original locator", async () => {
    const locator = "shape:2/%E6%8A%A5%E8%A1%A8%2F%E5%9B%BE%E8%A1%A8";
    executePowerShellMock.mockResolvedValue(
      JSON.stringify({
        app: "presentation",
        documentPath: "C:\\tmp\\slides.pptx",
        kind: "shape",
        name: "报表/图表",
        locator,
        parent: "slide:2",
        index: 5,
        selected: true,
      }),
    );

    const result = await new OfficeDocumentComBridge().activateObject({
      app: "presentation",
      filePath: "C:\\tmp\\slides.pptx",
      locator,
    });

    expect(result).toMatchObject({ locator, selected: true, parent: "slide:2" });
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("[Uri]::UnescapeDataString");
    expect(script).toContain("function Get-SlideById");
    expect(script).toContain("function Get-ShapeByPath");
    expect(script).toContain("$shape.GroupItems.Item");
    expect(script).toContain("[IO.Path]::GetFullPath");
  });

  it("saves only requested dirty open documents before a transaction snapshot", async () => {
    executePowerShellMock.mockResolvedValue(
      JSON.stringify([
        {
          app: "word",
          filePath: "C:\\tmp\\dirty.docx",
          wasDirty: true,
          saved: true,
        },
      ]),
    );

    const result = await new OfficeDocumentComBridge().prepareTransaction([
      "C:\\tmp\\dirty.docx",
      "C:\\tmp\\dirty.docx",
    ]);

    expect(result).toEqual([expect.objectContaining({ wasDirty: true, saved: true })]);
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$wantedPaths = [string[]](ConvertFrom-Json $_pathsJson)");
    expect(script).toContain("$fullName -notin $wantedPaths");
    expect(script).toContain("if ($wasDirty) { try { $handle.document.Save() }");
    expect(script).toContain("$seen.ContainsKey($key)");
  });

  it("restores snapshots while temporarily closing and reopening matching Office documents", async () => {
    executePowerShellMock.mockResolvedValue(
      JSON.stringify([
        {
          app: "excel",
          filePath: "C:\\tmp\\book.xlsx",
          instanceId: "excel:100:200",
          reopened: true,
        },
      ]),
    );

    const result = await new OfficeDocumentComBridge().restoreTransactionFiles([
      {
        filePath: "C:\\tmp\\book.xlsx",
        existed: true,
        snapshotPath: "C:\\journal\\before.xlsx",
      },
    ]);

    expect(result).toEqual([
      expect.objectContaining({ reopened: true, instanceId: "excel:100:200" }),
    ]);
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$handle.document.Close($false)");
    expect(script).toContain("[IO.File]::Copy($source, $stagedPath, $true)");
    expect(script).toContain("$committedFiles += $entry");
    expect(script).toContain("for ($index = $committedFiles.Count - 1; $index -ge 0; $index--)");
    expect(script).toContain("$session.application.Workbooks.Open");
    expect(script).toContain("$session.application.Documents.Open");
    expect(script).toContain("$session.application.Presentations.Open");
  });

  it("covers advanced selectable objects for Excel, Word, and PowerPoint", async () => {
    executePowerShellMock.mockResolvedValue("[]");
    const bridge = new OfficeDocumentComBridge();
    await bridge.listObjects({
      app: "excel",
      filePath: "C:\\tmp\\book.xlsx",
      instanceId: "excel:1:2",
    });
    await bridge.listObjects({
      app: "word",
      filePath: "C:\\tmp\\report.docx",
      instanceId: "word:3:4",
    });
    await bridge.listObjects({
      app: "presentation",
      filePath: "C:\\tmp\\slides.pptx",
      instanceId: "presentation:5:6",
    });

    const [excelScript, wordScript, presentationScript] = executePowerShellMock.mock.calls.map(
      ([script]) => script,
    );
    for (const token of ["'cell'", "'pivotTable'", "'slicer'", "'query'", "'connection'"])
      expect(excelScript).toContain(token);
    for (const token of [
      "'section'",
      "'heading'",
      "'paragraph'",
      "'comment'",
      "'revision'",
      "'footnote'",
      "'endnote'",
      "'caption'",
    ])
      expect(wordScript).toContain(token);
    for (const token of [
      "'master'",
      "'layout'",
      "'notesPage'",
      "Add-PresentationShapeObject",
      "GroupItems.Item",
    ])
      expect(presentationScript).toContain(token);
    expect(presentationScript).toContain("[int]$slide.SlideID");
  });
});
