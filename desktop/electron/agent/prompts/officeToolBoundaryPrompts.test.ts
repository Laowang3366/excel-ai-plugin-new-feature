import { describe, expect, it } from "vitest";

import generalOfficePrompt from "./templates/scenarios/general-office.zh-CN.md?raw";
import { buildContextualPromptSections } from "./systemPrompt";

describe("Office tool semantic boundaries", () => {
  it("keeps direct cell edits on range tools regardless of data size", () => {
    const officeToolsPrompt = buildContextualPromptSections({
      content: "把 Excel 单元格写入数据并汇总",
    });
    expect(officeToolsPrompt).toContain("值、公式、格式、固定汇总");
    expect(officeToolsPrompt).toContain("range.write");
    expect(officeToolsPrompt).toContain("数据量不是升级理由");
    expect(officeToolsPrompt).toContain("相关 operation 不向模型开放");
    expect(officeToolsPrompt).not.toContain("createPowerQuery");
    expect(officeToolsPrompt).not.toContain("createPivotTable");
  });

  it("requires ETL semantics before using Power Query", () => {
    const officeToolsPrompt = buildContextualPromptSections({
      content: "用 Power Query 合并外部数据源并保持可刷新",
    });
    expect(officeToolsPrompt).toContain("createPowerQuery/managePowerQuery");
    expect(officeToolsPrompt).toContain("多来源可刷新 ETL");
    expect(officeToolsPrompt).toContain("准确参数格式以本轮工具定义为准");
    expect(officeToolsPrompt).not.toContain('advancedIntent:"refreshable-etl"');
    expect(officeToolsPrompt).not.toContain('sourceKind:"external"|"multi-source"');
  });

  it("reserves pivot tables and slicers for interactive object behavior", () => {
    const officeToolsPrompt = buildContextualPromptSections({
      content: "创建交互式数据透视表并添加切片器",
    });
    expect(officeToolsPrompt).toContain("createPivotTable/refreshPivotTables");
    expect(officeToolsPrompt).toContain("明确要求交互式透视");
    expect(officeToolsPrompt).toContain("addSlicer");
    expect(officeToolsPrompt).toContain("准确参数格式以本轮工具定义为准");
    expect(officeToolsPrompt).not.toContain('advancedIntent:"interactive-pivot"');
  });

  it("requires chart and pivot readback instead of trusting summaries", () => {
    const officeToolsPrompt = buildContextualPromptSections({
      content: "在 Excel 创建交互式数据透视表和图表",
    });
    expect(officeToolsPrompt).toContain("回读关键结果");
    expect(officeToolsPrompt).toContain("失败或未验证不得声称成功");
  });

  it("keeps file-level actions behind an explicit file path", () => {
    const officeToolsPrompt = buildContextualPromptSections({
      content: "美化 Excel 文件",
    });
    expect(officeToolsPrompt).toMatch(/文件级修改.*`filePath`/);
    expect(officeToolsPrompt).toContain("明确磁盘 `filePath` 才用 `office.action.*`");
  });

  it("creates standalone Office files without requiring a connected desktop app", () => {
    const officeToolsPrompt = buildContextualPromptSections({
      content: "在桌面创建新的 Excel、Word 和约 15 页的防溺水 PPT",
    });

    expect(officeToolsPrompt).toContain("createWorkbook");
    expect(officeToolsPrompt).toContain("createDocument");
    expect(officeToolsPrompt).toContain("createPresentation");
    expect(officeToolsPrompt).toContain("addSlides");
    expect(officeToolsPrompt).toContain("独立新建不查连接");
    expect(officeToolsPrompt).toContain("禁止先编辑不存在文件");
  });

  it("keeps general cleaning on typed Office tools without external scripts", () => {
    expect(generalOfficePrompt).toContain("直接改单元格值、公式或格式时默认走 `range.write`");
    expect(generalOfficePrompt).toContain("严禁");
    expect(generalOfficePrompt).toContain("外部脚本");
    expect(generalOfficePrompt).toMatch(/可刷新、多来源 ETL/);
  });
});
