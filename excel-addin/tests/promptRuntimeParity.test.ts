import { describe, expect, it } from "vitest";
import {
  buildAddInHardBoundary,
  buildAdvancedExcelBoundary,
  composeExcelSystemPrompt,
  getPromptText,
  resolveOfficeAdvancedIntents,
} from "../shared/prompts";

/**
 * Patterns that imply a desktop-only capability is still available to the model.
 * Mentioning the same names only inside an explicit unsupported/forbid list is OK.
 */
const AVAILABLE_CAPABILITY_LEAKS: RegExp[] = [
  /preferEngine\s*:\s*["']com["']/i,
  /请使用\s*C#\s*Open\s*XML/i,
  /直接使用 Open XML 文件级/i,
  /用 `office\.action\.(?:inspect|apply|validate)`/i,
  /用 `office\.workflow\.run`/i,
  /office\.documents\.(?:list|activate)/i,
  /Excel 用 `createWorkbook`/i,
  /Word 用 `createDocument`/i,
  /PPT 用 `createPresentation`/i,
  /调用 `knowledge\.search`/i,
  /写入 `memory\.write`/i,
  /用 `web\.search`/i,
  /用 `ocr\.parseDocument`/i,
  /Word\/PPT 用对应 `word\.\*`/i,
];

function expectNoDesktopLeaks(prompt: string): void {
  for (const pattern of AVAILABLE_CAPABILITY_LEAKS) {
    expect(prompt, `available-capability leak ${pattern}`).not.toMatch(pattern);
  }
  // Stack isolation markers must appear only as forbidden/unsupported, not as how-to.
  expect(prompt).not.toMatch(/需要 COM 兜底/i);
  expect(prompt).toMatch(/COM \/ \.NET Worker \/ Electron IPC[\s\S]{0,40}unsupported|禁止且 unsupported/i);
}

describe("add-in prompt runtime parity", () => {
  it("reuses desktop Excel formula methodology with host.status (not office.connection)", () => {
    const formula = getPromptText("scenarios/formula.zh-CN.md");
    expect(formula).toContain("动态数组锚点");
    expect(formula).toContain("range.write");
    expect(formula).toContain("range.read");
    expect(formula).toMatch(/FILTER|UNIQUE|SORT|SEQUENCE|LET|XLOOKUP/);
    expect(formula).toContain("host.status");
    expect(formula).not.toContain("office.connection.status");
    expect(formula).not.toContain("knowledge.search");
  });

  it("reuses desktop dynamic-array runtime capability prompts", () => {
    const enabled = getPromptText("runtime/dynamic-array-enabled.zh-CN.md");
    const disabled = getPromptText("runtime/dynamic-array-disabled.zh-CN.md");
    expect(enabled).toMatch(/已开启|spill|动态数组/);
    expect(disabled).toMatch(/已关闭|尚未确认|动态数组/);
  });

  it("advanced boundary lists in-workbook tools and marks PQ/Pivot unsupported", () => {
    const boundary = buildAdvancedExcelBoundary({ content: "清洗表格并做图表" });
    expect(boundary).toContain("`range.write`");
    expect(boundary).toContain("`range.format.write`");
    expect(boundary).toContain("`table.list");
    expect(boundary).toContain("`chart.");
    expect(boundary).toContain("unsupported");
    expect(boundary).toMatch(/Power Query|透视/);
    expect(boundary).toContain("WPS JSA");
    expect(boundary).not.toMatch(/值、公式、格式、固定汇总用 `range\.write`/);
  });

  it("advanced boundary allows WPS implemented* subset and keeps spill/protection/table unsupported", () => {
    const boundary = buildAdvancedExcelBoundary({ content: "清洗表格并做图表" });
    // Implemented* on WPS — model must not blanket-reject these
    expect(boundary).toContain("workbook.objects.inspect");
    expect(boundary).toContain("currentRegion");
    expect(boundary).toContain("range.format.read/write");
    expect(boundary).toContain("range.autofit");
    expect(boundary).toContain("range.insert");
    expect(boundary).toContain("range.delete");
    expect(boundary).toContain("sheet.visibility.get/set");
    expect(boundary).toContain("sheet.protection.get/protect/unprotect");
    expect(boundary).toContain("namedRange.list/create/update/delete");
    expect(boundary).toContain("copy/move");
    expect(boundary).toContain("formula.dependencies.inspect");
    expect(boundary).toContain("formula.backups.inspect|restore");
    expect(boundary).toContain("conditionalFormat.list/add/delete");
    expect(boundary).toContain("dataValidation.read/write/clear");
    expect(boundary).toContain("跨表");
    expect(boundary).toMatch(/成员缺失/);
    expect(boundary).toContain("COM/.NET/Shell");
    // Still unsupported on WPS
    expect(boundary).toContain("spill|currentArray");
    expect(boundary).toContain("formula.protection.*");
    expect(boundary).toContain("仍 typed unsupported");
    expect(boundary).toContain("table/filter/sort");
    expect(boundary).toContain("chart 全系");
    // Must not claim the whole batch is WPS unsupported for format/copy/autofit
    expect(boundary).not.toMatch(
      /WPS JSA 对本批 expand、format、range\.insert、range\.delete、range\.autofit/,
    );
    // insert/delete/visibility/protection/namedRange are no longer in the WPS unsupported list
    expect(boundary).not.toMatch(
      /仍 typed unsupported[\s\S]*`range\.insert`\/`range\.delete`[\s\S]*visibility/,
    );
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

  it("injects boundary, connection context, and dynamic-array by host capability", () => {
    const officeJs = composeExcelSystemPrompt({
      routing: { content: "清洗 excel 工作表并创建 power query" },
      officeConnectionStatus: "connected (office-js)",
      now: new Date("2026-07-19T04:00:00.000Z"),
      dynamicArrayEnabled: true,
    });
    expect(officeJs).toContain("Office 应用连接状态：connected (office-js)");
    expect(officeJs).toContain("unsupported");
    expect(officeJs).toContain("`range.format.write`");
    expect(officeJs).toMatch(/2026/);
    expect(officeJs).toMatch(/已开启|spill/);
    expectNoDesktopLeaks(officeJs);

    const wps = composeExcelSystemPrompt({
      routing: { content: "写一个动态数组公式" },
      officeConnectionStatus: "connected (wps-jsa)",
      now: new Date("2026-07-19T04:00:00.000Z"),
      dynamicArrayEnabled: false,
    });
    expect(wps).toContain("connected (wps-jsa)");
    expect(wps).toMatch(/已关闭|尚未确认/);
    expect(wps).toContain("host.status");
    expectNoDesktopLeaks(wps);
  });

  it("E-class hard boundary covers macro/OpenXML/COM/path/workflow/export", () => {
    const hard = buildAddInHardBoundary();
    expect(hard).toMatch(/macro\.(write|run|detect)/i);
    expect(hard).toMatch(/Open XML/i);
    expect(hard).toMatch(/COM|Electron|\.NET/i);
    expect(hard).toMatch(/Power Query|透视表|切片器/);
    expect(hard).toMatch(/打开\/创建\/保存\/切换|任意磁盘路径/);
    expect(hard).toMatch(/office\.workflow|事务备份/);
    expect(hard).toMatch(/Word|PPT|PDF/);
  });

  it("hard boundary covers macro/office-tools/general-office routes without desktop leaks", () => {
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
      expect(prompt).toMatch(/unsupported/i);
      expectNoDesktopLeaks(prompt);
      const hardIdx = prompt.indexOf("本加载项运行时能力边界");
      expect(hardIdx).toBeGreaterThan(0);
    }
  });

  it("does not present macro or Open XML as available add-in tools", () => {
    const prompt = composeExcelSystemPrompt({
      routing: { content: "创建 vba 宏并写入 open xml 工作簿" },
      officeConnectionStatus: "connected (wps-jsa)",
      now: new Date("2026-07-19T04:00:00.000Z"),
    });
    expect(prompt).toMatch(/macro/i);
    expect(prompt).toMatch(/unsupported/i);
    expect(prompt).not.toMatch(/本加载项.*已实现.*macro\.write/);
    expect(prompt).not.toMatch(/请使用 C# Open XML/);
    expectNoDesktopLeaks(prompt);
  });

  it("base prompt describes task-pane Excel scope only", () => {
    const base = getPromptText("system/base.zh-CN.md");
    expect(base).toMatch(/任务窗格|当前活动工作簿/);
    expect(base).toContain("host.status");
    expect(base).toContain("range.read");
    expect(base).not.toMatch(/Word 文档和 PowerPoint/);
    expect(base).not.toContain("createWorkbook");
  });
});
