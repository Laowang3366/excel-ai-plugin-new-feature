import { beforeEach, describe, expect, it, vi } from "vitest";
import { ALL_TOOL_DEFINITIONS } from "../../tools/registry/toolDefinitions";

import {
  canAlwaysAllowTool,
  clearAlwaysAllowedTools,
  getAlwaysAllowedTools,
  markToolAlwaysAllowed,
  requestToolApproval,
  shouldRequireApproval,
} from "./toolApproval";

describe("toolApproval", () => {
  beforeEach(() => {
    clearAlwaysAllowedTools();
  });

  it("keeps dangerous, destructive, egress and unknown tools fail-closed", () => {
    expect(shouldRequireApproval("macro.write", "confirm_all")).toBe(true);
    expect(shouldRequireApproval("macro.run", "confirm_all")).toBe(true);
    expect(shouldRequireApproval("range.clear", "confirm_all")).toBe(true);
    expect(shouldRequireApproval("web.search", "confirm_all")).toBe(true);
    expect(shouldRequireApproval("ocr.parseDocument", "confirm_all")).toBe(true);
    expect(shouldRequireApproval("memory.write", "confirm_all")).toBe(true);
    expect(shouldRequireApproval("unknown.tool", "confirm_all")).toBe(true);
  });

  it("enforces mandatory approval metadata for every model-visible tool", () => {
    for (const tool of ALL_TOOL_DEFINITIONS) {
      if (
        tool.riskLevel !== "dangerous" &&
        !tool.isFileDeletion &&
        !tool.isDataEgress &&
        !tool.requiresExplicitApproval
      )
        continue;

      expect(
        shouldRequireApproval(tool.name, "confirm_all"),
        `${tool.name} must stay fail-closed`,
      ).toBe(true);
    }
  });

  it("distinguishes destructive and non-destructive operations on a shared tool", () => {
    const addScope = {
      threadId: "thread-1",
      arguments: { operation: "add", sheetName: "Sheet2" },
    };
    const deleteScope = {
      threadId: "thread-1",
      arguments: { operation: "delete", sheetName: "Sheet2" },
    };

    expect(shouldRequireApproval("sheet.operation", "confirm_all", addScope)).toBe(false);
    expect(shouldRequireApproval("sheet.operation", "confirm_all", deleteScope)).toBe(true);
  });

  it("denies approval when no callback is configured", async () => {
    await expect(
      requestToolApproval(
        {
          toolCallId: "call-1",
          toolName: "range.write",
          arguments: { range: "A1" },
          riskLevel: "moderate",
        },
        { permissionMode: "normal" },
      ),
    ).resolves.toEqual({ approved: false });
  });

  it("scopes temporary grants by thread, operation and target file", () => {
    const scope = {
      threadId: "thread-1",
      arguments: { operation: "format", filePath: "C:\\books\\a.xlsx" },
    };

    expect(canAlwaysAllowTool("office.action.apply", scope)).toBe(true);
    expect(markToolAlwaysAllowed("office.action.apply", scope)).toBe(true);
    expect(getAlwaysAllowedTools().has("office.action.apply")).toBe(true);
    expect(shouldRequireApproval("office.action.apply", "normal", scope)).toBe(false);
    expect(
      shouldRequireApproval("office.action.apply", "normal", {
        ...scope,
        threadId: "thread-2",
      }),
    ).toBe(true);
    expect(
      shouldRequireApproval("office.action.apply", "normal", {
        threadId: "thread-1",
        arguments: { operation: "format", filePath: "C:\\books\\b.xlsx" },
      }),
    ).toBe(true);
    expect(
      shouldRequireApproval("office.action.apply", "normal", {
        threadId: "thread-1",
        arguments: { operation: "delete", filePath: "C:\\books\\a.xlsx" },
      }),
    ).toBe(true);
  });

  it("does not persist grants without a stable workbook or file identity", () => {
    const scope = {
      threadId: "thread-1",
      arguments: { sheetName: "Sheet1", range: "A1", values: [[1]] },
    };

    expect(canAlwaysAllowTool("range.write", scope)).toBe(false);
    expect(markToolAlwaysAllowed("range.write", scope)).toBe(false);
    expect(shouldRequireApproval("range.write", "normal", scope)).toBe(true);
  });

  it("delegates approval requests to the configured callback", async () => {
    const callback = vi.fn(async () => ({ approved: true, alwaysAllow: true }));
    const request = {
      toolCallId: "call-2",
      toolName: "office.action.apply",
      arguments: { operation: "format", filePath: "C:\\books\\a.xlsx" },
      riskLevel: "moderate" as const,
      canAlwaysAllow: true,
    };

    await expect(
      requestToolApproval(request, {
        permissionMode: "normal",
        requestToolApproval: callback,
      }),
    ).resolves.toEqual({ approved: true, alwaysAllow: true });
    expect(callback).toHaveBeenCalledWith(request);
  });
});
