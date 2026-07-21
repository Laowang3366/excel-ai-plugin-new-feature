import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

describe("capability matrix", () => {
  it("documents first-batch and unsupported families with evidence", () => {
    const text = readFileSync(path.join(root, "docs/capability-matrix.md"), "utf8");
    for (const token of [
      "selection.get",
      "range.read",
      "formula.write",
      "sheet.list",
      "range.format",
      "table.list",
      "table.unlist",
      "convertToRange",
      "Phase28",
      "Phase29",
      "Phase30",
      "Phase31",
      "Phase32",
      "Phase33",
      "clearPageBreaks",
      "horizontalPageBreaks",
      "defaultForAllPages",
      "headerMargin",
      "footerMargin",
      "draft",
      "pageOrder",
      "firstPageNumber",
      "paperSize",
      "fitToPages",
      "ExcelApi 1.9",
      "printArea",
      "printTitleRows",
      "fitToOnePageWide",
      "repeatRows",
      "chart.list/create/delete",
      "chart.series.list",
      "chart.source.update",
      "chart.axes.update",
      "chart.series.dataLabels.update",
      "chart.series.axisGroup.update",
      "chart.series.delete",
      "chart.series.add",
      "chart.series.values.update",
      "ExcelApi 1.15",
      "chart.series.bubbleSizes.update",
      "BubbleSizes",
      "chart.series.trendlines",
      "chart.series.markers.update",
      "markerStyle",
      "chart.series.trendlines.format.update",
      "Phase47",
      "Phase48",
      "Phase49",
      "chart.image.get",
      "range.image.get",
      "range.insert",
      "range.delete",
      "range.autofit",
      "Phase38",
      "RangeFormat.autofitRows",
      "RangeFormat.autofitColumns",
      "rowHeight",
      "columnWidth",
      "ExcelApi 1.1",
      "ExcelApi 1.2",
      "workbook.inspect",
      "workbook.objects.inspect",
      "workbook.save",
      "pivot.list",
      "pivot.create",
      "pivot.refresh",
      "slicer.list",
      "slicer.create",
      "slicer.update",
      "slicer.delete",
      "slicer.filter.get",
      "slicer.filter.apply",
      "slicer.filter.clear",
      "ExcelApi 1.10",
      "selectItems",
      "ExcelApi 1.8",
      "Power Query",
      "COM / .NET / Electron",
      "unsupported",
      "office-js",
      "wps-jsa",
      "data labels",
      "source replacement",
      "export",
      "complex layout",
      "name/chartType/smooth",
      "delete/format/leaderLines",
      "showPercentage/showBubbleSize/showLegendKey/separator/position",
      "Phase52",
      "hasDataLabels",
      "ExcelApi 1.7",
      "ExcelApi 1.8",
      "Phase26",
      "Phase27",
      "doughnut",
      "linemarkers",
      // Task A: WPS implemented* vs still-unsupported split at HEAD a8378293
      "range.read expand currentRegion",
      "range.read expand spill / currentArray",
      "wpsJsaRangeRead",
      "wpsJsaSheetOps",
      "wpsJsaFormat",
      "wpsAutofitRange",
      "implemented*",
      "not official JSA",
      "not real sideload",
      "WPS still typed unsupported",
      "expand **spill**",
      "formula.protection.*",
    ]) {
      expect(text).toContain(token);
    }
  });

  it("marks WPS currentRegion/format/copy-move/autofit/structure/governance as implemented* not blanket unsupported", () => {
    const text = readFileSync(path.join(root, "docs/capability-matrix.md"), "utf8");
    // Positive implemented* rows
    expect(text).toMatch(/range\.read expand currentRegion[\s\S]*?implemented\*/);
    expect(text).toMatch(/sheet\.operation copy[\s\S]*?implemented\*/);
    expect(text).toMatch(/sheet\.operation move[\s\S]*?implemented\*/);
    expect(text).toMatch(/range\.format\.read\/write[\s\S]*?implemented\*/);
    expect(text).toMatch(/\| range \| autofit \|[\s\S]*?implemented\*/);
    expect(text).toMatch(/\| range \| insert \/ delete \|[\s\S]*?implemented\*/);
    expect(text).toMatch(/visibility get\/set[\s\S]*?implemented\*/);
    expect(text).toMatch(/protection get\/protect\/unprotect[\s\S]*?implemented\*/);
    expect(text).toMatch(/named range \| list\/create\/update\/delete[\s\S]*?implemented\*/);
    expect(text).toMatch(/dependencies\.inspect[\s\S]*?implemented\*/);
    expect(text).toMatch(/conditional format[\s\S]*?implemented\*/);
    expect(text).toMatch(/data validation[\s\S]*?implemented\*/);
    expect(text).toMatch(/FormatConditions 1-based index ids/);
    expect(text).toMatch(/chart\.source\.update[\s\S]*?cross-sheet/);
    expect(text).toMatch(/source update \(setData\)[\s\S]*?cross-sheet A1/);
    expect(text).not.toMatch(/source update \(setData\)[\s\S]*?\*\*unsupported\*\*: cross-sheet source/);
    // Still unsupported
    expect(text).toMatch(/range\.read expand spill \/ currentArray[\s\S]*?\*\*unsupported\*\*/);
    expect(text).toMatch(/protection inspect\/manage[\s\S]*?\*\*unsupported\*\*/);
    // Assumptions must not claim format/copy-move all unverified
    expect(text).not.toMatch(
      /Format \/ ListObjects \/ ChartObjects \/ expand \/ sheet copy-move are \*\*not\*\* verified/,
    );
  });


  it("Phase60 WPS real-device evidence: Ribbon cold-start + selection.get G17 closed; not blanket pass", () => {
    const matrix = readFileSync(path.join(root, "docs/capability-matrix.md"), "utf8");
    const readme = readFileSync(path.join(root, "README.md"), "utf8");
    for (const doc of [matrix, readme]) {
      expect(doc).toContain("c46362f8");
      expect(doc).toContain("12.1.0.26885");
      expect(doc).toContain("selection.get");
      expect(doc).toContain("G17");
      expect(doc).toContain('address:"G17"');
    }
    expect(matrix).toContain("device-verified");
    expect(matrix).toMatch(/冷启动/);
    expect(matrix).toMatch(/瞬态|transient/i);
    expect(readme).toMatch(/冷启动后/);
    expect(readme).toMatch(/瞬态/);
    expect(readme).toContain("values:[[null]]");
    // Must not reintroduce pending re-verify for the closed G17 path
    expect(matrix).not.toContain("host re-verify pending when Ribbon available");
    expect(readme).not.toContain("待 Ribbon 可点后复验");
    // Honesty boundaries still required
    expect(readme).toMatch(/不得扩大为其它 WPS 工具全部真机通过/);
    expect(readme).toMatch(/member-probe/);
    expect(matrix).toMatch(/不得扩大为全部 WPS 工具真机通过/);
    expect(matrix).toContain("其它工具仍 member-probe");
    expect(matrix).toMatch(/not real sideload/i);
    // Task pane: open+load OK; layout completeness fails — no 完整渲染 claim
    expect(matrix).not.toContain("完整渲染");
    expect(readme).not.toContain("完整渲染");
    expect(matrix).toMatch(/布局完整性未过|右侧已测裁剪|right-side clip/);
    expect(readme).toMatch(/成功打开并加载 UI/);
  });

});
