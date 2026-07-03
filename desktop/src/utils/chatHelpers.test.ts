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
});

describe("pickExcelRange", () => {
  test("uses the fast address-only selection API", async () => {
    ipcMocks.getSelectionAddress.mockResolvedValue({ address: "B2:H9", sheetName: "Sheet2" });

    await expect(pickExcelRange()).resolves.toBe("Sheet2!B2:H9");
    expect(ipcMocks.getSelectionAddress).toHaveBeenCalledTimes(1);
    expect(ipcMocks.getSelection).not.toHaveBeenCalled();
  });
});
