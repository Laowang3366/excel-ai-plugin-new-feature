import { describe, expect, it } from "vitest";

import generalOfficePrompt from "./templates/scenarios/general-office.zh-CN.md?raw";
import officeToolsPrompt from "./templates/scenarios/office-tools.zh-CN.md?raw";

describe("Office tool semantic boundaries", () => {
  it("keeps direct cell edits on range tools regardless of data size", () => {
    expect(officeToolsPrompt).toContain("直接写值、公式、格式或固定汇总结果");
    expect(officeToolsPrompt).toContain("range.write");
    expect(officeToolsPrompt).toContain("数据量大本身不构成调用复杂工具的理由");
    expect(officeToolsPrompt).toContain("禁止为写值先建 Power Query/透视表");
  });

  it("requires ETL semantics before using Power Query", () => {
    expect(officeToolsPrompt).toContain("createPowerQuery/managePowerQuery");
    expect(officeToolsPrompt).toContain("多来源可刷新 ETL");
    expect(officeToolsPrompt).toMatch(/明确源、转换、加载.*filePath/);
  });

  it("reserves pivot tables and slicers for interactive object behavior", () => {
    expect(officeToolsPrompt).toContain("createPivotTable/refreshPivotTables");
    expect(officeToolsPrompt).toContain("交互多维布局");
    expect(officeToolsPrompt).toContain("addSlicer");
    expect(officeToolsPrompt).toContain("已有透视表/结构化表");
  });

  it("requires chart and pivot readback instead of trusting summaries", () => {
    expect(officeToolsPrompt).toContain("data.verification.ok");
    expect(officeToolsPrompt).toContain("inspectCharts");
    expect(officeToolsPrompt).toContain("data.readback.verification.ok");
    expect(officeToolsPrompt).toContain("inspectWorkbookObjects");
    expect(officeToolsPrompt).toContain("失败不得声称成功");
  });

  it("keeps file-level actions behind an explicit file path", () => {
    expect(officeToolsPrompt).toMatch(/文件级修改.*`filePath`/);
    expect(officeToolsPrompt).toContain("明确磁盘 `filePath` 才用 `office.action.*`");
  });

  it("keeps general cleaning on typed Office tools without external scripts", () => {
    expect(generalOfficePrompt).toContain("直接改单元格值、公式或格式时默认走 `range.write`");
    expect(generalOfficePrompt).toContain("严禁");
    expect(generalOfficePrompt).toContain("外部脚本");
    expect(generalOfficePrompt).toMatch(/可刷新、多来源 ETL/);
  });

});
