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
      "showPercentage/showBubbleSize/delete/position/format/leaderLines",
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
    // Still unsupported
    expect(text).toMatch(/range\.read expand spill \/ currentArray[\s\S]*?\*\*unsupported\*\*/);
    expect(text).toMatch(/protection inspect\/manage[\s\S]*?\*\*unsupported\*\*/);
    // Assumptions must not claim format/copy-move all unverified
    expect(text).not.toMatch(
      /Format \/ ListObjects \/ ChartObjects \/ expand \/ sheet copy-move are \*\*not\*\* verified/,
    );
  });

});
