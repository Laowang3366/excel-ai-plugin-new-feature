import { beforeEach, describe, expect, test, vi } from "vitest";

import { executePowerShell } from "../../../automation/powershell";
import { OfficeComActionBridge } from "./officeComActionBridge";

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return {
    ...actual,
    executePowerShell: vi.fn(),
  };
});

const executePowerShellMock = vi.mocked(executePowerShell);

describe("OfficeComActionBridge", () => {
  beforeEach(() => {
    executePowerShellMock.mockReset();
  });

  test("runs Excel insertChart through COM and cleans up the application", async () => {
    executePowerShellMock.mockResolvedValue(JSON.stringify({
      outputPath: "C:\\tmp\\book-chart.xlsx",
      changes: [{ kind: "chart", target: "range:Sheet1!A1:B5", detail: "已创建图表" }],
    }));

    const bridge = new OfficeComActionBridge();
    const result = await bridge.executeAction({
      app: "excel",
      action: "insert",
      operation: "insertChart",
      filePath: "C:\\tmp\\book.xlsx",
      outputPath: "C:\\tmp\\book-chart.xlsx",
      target: "range:Sheet1!A1:B5",
      params: { chartType: "line" },
    });

    const script = executePowerShellMock.mock.calls[0][0];
    expect(result).toMatchObject({ status: "done", engine: "com", summary: "已通过 COM 执行 Excel insertChart" });
    expect(script).toContain("$chartObjects.Add");
    expect(script).toContain("$chart.ChartType");
    expect(script).toContain("$workbook.SaveAs($_outputPath)");
    expect(script).toContain("$app.Quit()");
    expect(script).toContain("ReleaseComObject($workbook)");
  });

  test("runs Word insertOrUpdateToc through COM with field update", async () => {
    executePowerShellMock.mockResolvedValue(JSON.stringify({
      outputPath: "C:\\tmp\\report-toc.docx",
      changes: [{ kind: "toc", detail: "已插入或更新目录" }],
    }));

    const bridge = new OfficeComActionBridge();
    await bridge.executeAction({
      app: "word",
      action: "insert",
      operation: "insertOrUpdateToc",
      filePath: "C:\\tmp\\report.docx",
      outputPath: "C:\\tmp\\report-toc.docx",
    });

    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$doc.TablesOfContents.Add");
    expect(script).toContain("$toc.Update()");
    expect(script).toContain("$doc.Fields.Update()");
    expect(script).toContain("$doc.SaveAs2($_outputPath)");
    expect(script).toContain("$app.Quit()");
  });

  test("runs PowerPoint snapshot through COM slide export", async () => {
    executePowerShellMock.mockResolvedValue(JSON.stringify({
      outputPath: "C:\\tmp\\slide.png",
      changes: [{ kind: "snapshot", target: "slide:1", detail: "已导出幻灯片快照" }],
    }));

    const bridge = new OfficeComActionBridge();
    await bridge.executeAction({
      app: "presentation",
      action: "snapshot",
      operation: "snapshot",
      filePath: "C:\\tmp\\slides.pptx",
      outputPath: "C:\\tmp\\slide.png",
      target: "slide:1",
    });

    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$slide.Export($_outputPath, 'PNG')");
    expect(script).toContain("$pres.Close()");
    expect(script).toContain("$app.Quit()");
  });

  test("runs PowerPoint slide deletion from high to low through COM", async () => {
    executePowerShellMock.mockResolvedValue(JSON.stringify({
      outputPath: "C:\\tmp\\slides-trimmed.pptx",
      changes: [{ kind: "slide-delete", target: "slide:2-6", detail: "已删除幻灯片" }],
    }));

    const bridge = new OfficeComActionBridge();
    await bridge.executeAction({
      app: "presentation",
      action: "edit",
      operation: "deleteSlides",
      filePath: "C:\\tmp\\slides.pptx",
      outputPath: "C:\\tmp\\slides-trimmed.pptx",
      target: "slide:2-6",
      params: { from: 2, to: 6 },
    });

    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$deleteSlideIndexes = @(2, 3, 4, 5, 6)");
    expect(script).toContain("Sort-Object -Descending");
    expect(script).toContain("$pres.Slides.Item($idx).Delete()");
    expect(script).toContain("$pres.SaveAs($_outputPath)");
  });

  test("rejects unsupported COM action before executing PowerShell", async () => {
    const bridge = new OfficeComActionBridge();

    const result = await bridge.executeAction({
      app: "word",
      action: "edit",
      operation: "unknownOperation",
      filePath: "C:\\tmp\\report.docx",
    });

    expect(result.status).toBe("unsupported");
    expect(executePowerShellMock).not.toHaveBeenCalled();
  });
});
