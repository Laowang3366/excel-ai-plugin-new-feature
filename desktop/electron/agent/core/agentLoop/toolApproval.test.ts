import { beforeEach, describe, expect, it, vi } from "vitest";

import {
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

  it("honors permission modes and always-allowed overrides", () => {
    expect(shouldRequireApproval("range.read", "normal")).toBe(true);
    expect(shouldRequireApproval("range.read", "auto_approve_safe")).toBe(false);
    expect(shouldRequireApproval("range.clear", "auto_approve_safe")).toBe(true);
    expect(shouldRequireApproval("range.clear", "confirm_all")).toBe(false);

    markToolAlwaysAllowed("range.clear");
    expect(getAlwaysAllowedTools().has("range.clear")).toBe(true);
    expect(shouldRequireApproval("range.clear", "normal")).toBe(false);
  });

  it("approves by default when no approval callback is configured", async () => {
    await expect(requestToolApproval({
      toolCallId: "call-1",
      toolName: "range.write",
      arguments: { range: "A1" },
      riskLevel: "moderate",
    }, { permissionMode: "normal" })).resolves.toEqual({ approved: true });
  });

  it("delegates approval requests to the configured callback", async () => {
    const requestToolApprovalCallback = vi.fn(async () => ({ approved: true, alwaysAllow: true }));

    const result = await requestToolApproval({
      toolCallId: "call-2",
      toolName: "shell.execute",
      arguments: { command: "git status" },
      riskLevel: "dangerous",
      sandboxJustification: "Needs approval",
    }, {
      permissionMode: "confirm_all",
      requestToolApproval: requestToolApprovalCallback,
    });

    expect(result).toEqual({ approved: true, alwaysAllow: true });
    expect(requestToolApprovalCallback).toHaveBeenCalledWith(expect.objectContaining({
      toolCallId: "call-2",
      toolName: "shell.execute",
      sandboxJustification: "Needs approval",
    }));
  });
});
