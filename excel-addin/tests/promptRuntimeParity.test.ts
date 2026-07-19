import { describe, expect, it } from "vitest";
import {
  buildAdvancedExcelBoundary,
  composeExcelSystemPrompt,
  formatOfficeConnectionContext,
  formatRuntimeDateTime,
  resolveOfficeAdvancedIntents,
} from "../shared/prompts";

describe("prompt runtime parity", () => {
  it("formats Asia/Shanghai date/time like desktop", () => {
    const parts = formatRuntimeDateTime(new Date("2026-07-19T04:00:00.000Z"));
    expect(parts.CURRENT_DATE).toMatch(/2026/);
    expect(parts.CURRENT_DATE).toMatch(/星期/);
    expect(parts.CURRENT_TIME).toMatch(/^\d{2}:\d{2}$/);
  });

  it("uses desktop office connection context prefix", () => {
    expect(formatOfficeConnectionContext("connected (office-js)")).toBe(
      "- Office 应用连接状态：connected (office-js)",
    );
  });

  it("states value/formula/format/table/chart tool boundaries and WPS unsupported", () => {
    const boundary = buildAdvancedExcelBoundary({});
    expect(boundary).toContain("`range.write`");
    expect(boundary).toContain("`formula.write`");
    expect(boundary).toContain("`formula.context`");
    expect(boundary).toContain("`range.read`");
    expect(boundary).toContain("spill");
    expect(boundary).toContain("`sheet.operation`");
    expect(boundary).toContain("1-based");
    expect(boundary).toContain("`range.format.write`");
    expect(boundary).toContain("`table.list/create/delete/update/unlist`");
    expect(boundary).toContain("convertToRange");
    expect(boundary).toContain("table.unlist");
    expect(boundary).toContain("`chart.list/create/delete/update`");
    expect(boundary).toContain("`chart.series.list/update`");
    expect(boundary).toContain("name/chartType/smooth");
    expect(boundary).toContain("`chart.source.update`");
    expect(boundary).toContain("`chart.axes.update`");
    expect(boundary).toContain("`chart.series.dataLabels.update`");
    expect(boundary).toContain("showValue/showCategoryName/showSeriesName/numberFormat");
    expect(boundary).toContain("hasDataLabels");
    expect(boundary).toMatch(/ExcelApi 1\.7/);
    expect(boundary).toMatch(/ExcelApi 1\.8|show fields ExcelApi 1\.8/);
    expect(boundary).toContain(
      "showPercentage/showBubbleSize/delete/position/format/leaderLines",
    );
    expect(boundary).toContain("`chart.series.axisGroup.update`");
    expect(boundary).toContain("`chart.series.delete`");
    expect(boundary).toContain("`chart.series.add`");
    expect(boundary).toContain("`chart.series.values.update`");
    expect(boundary).toContain("`chart.series.bubbleSizes.update`");
    expect(boundary).toContain("BubbleSizes");
    expect(boundary).toContain("dataBound");
    expect(boundary).toContain("ExcelApi 1.15");
    expect(boundary).toContain("`chart.image.get`");
    expect(boundary).toContain("`range.image.get`");
    expect(boundary).toContain("Base64");
    expect(boundary).toMatch(/WPS JSA[\s\S]*chart\.image\.get|chart\.image\.get[\s\S]*WPS JSA/);
    expect(boundary).toContain("WPS JSA");
    expect(boundary).toContain("unsupported");
    expect(boundary).not.toMatch(/值、公式、格式、固定汇总用 `range\.write`/);
  });

  it("detects advanced intents but keeps PQ/Pivot unsupported in boundary text", () => {
    const pq = resolveOfficeAdvancedIntents({ content: "创建 Power Query 可刷新 ETL" });
    expect(pq.has("refreshable-etl")).toBe(true);
    const boundary = buildAdvancedExcelBoundary({
      content: "创建数据透视表和切片器",
    });
    expect(boundary).toContain("unsupported");
    expect(boundary).toMatch(/透视|Power Query|交互/);
  });

  it("injects boundary and connection context into composed prompt", () => {
    const prompt = composeExcelSystemPrompt({
      routing: { content: "清洗 excel 工作表并创建 power query" },
      officeConnectionStatus: "connected (office-js)",
      now: new Date("2026-07-19T04:00:00.000Z"),
      dynamicArrayEnabled: true,
    });
    expect(prompt).toContain("Office 应用连接状态：connected (office-js)");
    expect(prompt).toContain("unsupported");
    expect(prompt).toContain("`range.format.write`");
    expect(prompt).toMatch(/2026/);
  });

  it("hard boundary covers macro/office-tools/general-office routes", () => {
    const routes = [
      { content: "编写 vba 宏并 macro.write 运行" },
      { content: "清洗 excel 工作表并汇总报告" },
      { content: "excel 条件格式与数据验证" },
    ];
    for (const routing of routes) {
      const prompt = composeExcelSystemPrompt({
        routing,
        officeConnectionStatus: "connected (office-js)",
        now: new Date("2026-07-19T04:00:00.000Z"),
      });
      expect(prompt).toContain("本加载项运行时能力边界");
      expect(prompt).toMatch(/macro\.(write|run|detect)/i);
      expect(prompt).toMatch(/宏[\s\S]{0,80}unsupported/i);
      expect(prompt).toMatch(/Open XML[\s\S]{0,40}unsupported/i);
      expect(prompt).toMatch(/COM[\s\S]{0,40}unsupported|COM \/ \.NET/i);
      expect(prompt).toMatch(/UserForm|菜单/);
      expect(prompt).toMatch(/Power Query[\s\S]{0,40}unsupported|透视表[\s\S]{0,40}unsupported/i);
      // Hard boundary must appear after synced scenario so it overrides desktop macro narrative
      const hardIdx = prompt.indexOf("本加载项运行时能力边界");
      const macroScenarioIdx = prompt.indexOf("macro.write");
      if (macroScenarioIdx >= 0 && prompt.includes("Excel/WPS 内部宏执行规则")) {
        expect(hardIdx).toBeGreaterThan(prompt.indexOf("Excel/WPS 内部宏执行规则"));
      }
    }
  });

  it("does not present macro or Open XML as available add-in tools in runtime prompt", () => {
    const prompt = composeExcelSystemPrompt({
      routing: { content: "创建 vba 宏并写入 open xml 工作簿" },
      officeConnectionStatus: "connected (wps-jsa)",
      now: new Date("2026-07-19T04:00:00.000Z"),
    });
    expect(prompt).toContain("macro.write");
    expect(prompt).toMatch(/macro\.write[\s\S]{0,200}unsupported|均为 \*\*unsupported\*\*/i);
    expect(prompt).toMatch(/Open XML[\s\S]{0,80}\*\*unsupported\*\*|Open XML[\s\S]{0,80}unsupported/i);
    expect(prompt).not.toMatch(/本加载项.*已实现.*macro\.write/);
    expect(prompt).not.toMatch(/请使用 C# Open XML/);
  });
});
