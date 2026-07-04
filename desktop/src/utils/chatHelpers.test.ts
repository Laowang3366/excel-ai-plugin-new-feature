import { beforeEach, describe, expect, test, vi } from "vitest";
import type { TurnItem } from "../electronApi";

const ipcMocks = vi.hoisted(() => ({
  getSelection: vi.fn(),
  getSelectionAddress: vi.fn(),
}));

vi.mock("../services/ipcApi", () => ({
  ipcApi: {
    excel: {
      getSelection: ipcMocks.getSelection,
      getSelectionAddress: ipcMocks.getSelectionAddress,
    },
  },
}));

import {
  getChatTitleSummary,
  getLiveTurnDurationSeconds,
  getUserFacingMessageContent,
  pickExcelRange,
} from "./chatHelpers";

beforeEach(() => {
  vi.clearAllMocks();
});

describe("getLiveTurnDurationSeconds", () => {
  test("adds the running segment after the last completed assistant item", () => {
    const items: TurnItem[] = [
      {
        type: "reasoning",
        id: "reasoning-1",
        summaryText: [],
        rawContent: ["thinking"],
        timestamp: 4_000,
      },
      {
        type: "tool_call",
        id: "tool-1",
        toolName: "inspect",
        arguments: {},
        status: "completed",
        timestamp: 8_000,
      },
    ];

    expect(getLiveTurnDurationSeconds(items, 1_000, 13_000)).toBe(12);
  });

  test("uses the previous user timestamp while the first streaming segment has no completed item", () => {
    expect(getLiveTurnDurationSeconds([], 1_000, 11_000)).toBe(10);
  });
});

describe("user-facing chat content", () => {
  test("removes hidden module instruction lines from displayed user content", () => {
    expect(getUserFacingMessageContent([
      "【功能模块：生成公式】",
      "模块指令：本轮必须按公式助手模式处理。",
      "任务说明：提取字段",
    ].join("\n"))).toBe("【功能模块：生成公式】\n任务说明：提取字段");
  });

  test("builds chat title from visible user fields only", () => {
    const messages: TurnItem[] = [{
      type: "user_message",
      id: "u1",
      content: [
        "【功能模块：公式助手】",
        "模块指令：这段不应该进入标题",
        "任务说明：提取需要的内容",
      ].join("\n"),
      timestamp: 1,
    }];

    expect(getChatTitleSummary(messages, "新会话")).toBe("【功能模块：公式助手】 任务说明：提取需要的内容");
  });

  test("hides formula module delivery guardrails while keeping panel fields", () => {
    const visible = getUserFacingMessageContent([
      "【功能模块：生成公式】",
      "任务说明：按通道累加",
      "当前连接环境：WPS",
      "交付要求：必须使用 Excel/WPS 函数公式完成，不要改写为 VBA、JS、Python 或手工值。",
      "数据源选区：Sheet1!D4:G7",
      "答案参考样例：Sheet1!I4:L7",
      "答案参考样例类型：完整样例",
      "答案填入锚点/选区：由 Agent 选择空白区域",
      "是否支持动态数组：是",
    ].join("\n"));

    expect(visible).toContain("【功能模块：生成公式】");
    expect(visible).toContain("任务说明：按通道累加");
    expect(visible).toContain("当前连接环境：WPS");
    expect(visible).toContain("数据源选区：Sheet1!D4:G7");
    expect(visible).toContain("答案参考样例类型：完整样例");
    expect(visible).toContain("是否支持动态数组：是");
    expect(visible).not.toContain("交付要求：");
    expect(visible).not.toContain("由 Agent 选择空白区域");
  });

  test("hides empty/default code module fields from displayed content", () => {
    const visible = getUserFacingMessageContent([
      "【功能模块：代码生成】",
      "代码需求：生成录入窗体",
      "运行环境：Microsoft Excel",
      "首选语言：自动",
      "数据源选区：未指定，请读取工作簿快照后自主判断。",
      "答案参考样例：未指定。",
      "输出/操作锚点：未指定。",
    ].join("\n"));

    expect(visible).toBe([
      "【功能模块：代码生成】",
      "代码需求：生成录入窗体",
      "运行环境：Microsoft Excel",
      "首选语言：自动",
    ].join("\n"));
  });

  test("hides report module delivery method but keeps selected output fields", () => {
    const visible = getUserFacingMessageContent([
      "【功能模块：报告生成】",
      "报告类型：Word 文档",
      "数据源选区：Sheet1!A1:F20",
      "需求说明：输出经营报告",
      "存储路径：C:\\Users\\wfq\\Desktop",
      "交付方式：在上述路径创建 Word 文档，写入报告内容后用系统默认应用打开。",
    ].join("\n"));

    expect(visible).toContain("报告类型：Word 文档");
    expect(visible).toContain("数据源选区：Sheet1!A1:F20");
    expect(visible).toContain("存储路径：C:\\Users\\wfq\\Desktop");
    expect(visible).not.toContain("交付方式：");
  });

  test("keeps user-written delivery requirements in ordinary messages", () => {
    const visible = getUserFacingMessageContent([
      "请帮我写一份方案",
      "交付要求：用表格列出风险",
    ].join("\n"));

    expect(visible).toContain("交付要求：用表格列出风险");
  });
});

describe("pickExcelRange", () => {
  test("uses the fast address-only selection API", async () => {
    ipcMocks.getSelectionAddress.mockResolvedValue({ address: "B2:H9", sheetName: "Sheet2" });

    await expect(pickExcelRange()).resolves.toBe("Sheet2!B2:H9");
    expect(ipcMocks.getSelectionAddress).toHaveBeenCalledTimes(1);
    expect(ipcMocks.getSelection).not.toHaveBeenCalled();
  });
});
