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
    ]) {
      expect(text).toContain(token);
    }
  });
});
