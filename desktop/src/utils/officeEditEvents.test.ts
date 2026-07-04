import { describe, expect, it } from "vitest";
import type { TurnItem } from "../electronApi";
import {
  collectOfficeEditEvents,
  getOfficePreviewToggleLocation,
} from "./officeEditEvents";

describe("collectOfficeEditEvents", () => {
  it("extracts unified Office action status events", () => {
    const events = collectOfficeEditEvents([
      {
        type: "tool_result",
        id: "result-action",
        toolCallId: "call-action",
        toolName: "office.action.apply",
        isError: false,
        timestamp: 8000,
        result: {
          status: "needsCom",
          engine: "openxml",
          app: "word",
          action: "insert",
          operation: "insertOrUpdateToc",
          filePath: "D:\\docs\\report.docx",
          summary: "目录字段需要 Word 刷新",
          changes: [],
        },
      },
    ]);

    expect(events[0].summary).toBe("Office action word/insertOrUpdateToc：needsCom");
    expect(events[0].detail).toEqual({
      summary: "目录字段需要 Word 刷新",
      changes: [],
      validation: undefined,
      error: undefined,
    });
  });

  it("ignores legacy direct Open XML tool results", () => {
    const events = collectOfficeEditEvents([
      {
        type: "tool_result",
        id: "result-legacy",
        toolCallId: "call-legacy",
        toolName: "office.file.inspect",
        isError: false,
        timestamp: 1000,
        result: {
          engine: "openxml",
          operation: "inspect",
          documentType: "word",
          filePath: "D:\\docs\\report.docx",
        },
      },
    ]);

    expect(events).toEqual([]);
  });
});

describe("getOfficePreviewToggleLocation", () => {
  it("moves the monitor toggle between chat header and panel header", () => {
    expect(getOfficePreviewToggleLocation(false)).toBe("chat-header");
    expect(getOfficePreviewToggleLocation(true)).toBe("panel-header");
  });
});
