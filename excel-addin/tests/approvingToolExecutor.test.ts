import { describe, expect, it, vi } from "vitest";
import { ApprovalGate } from "../shared/agentChat/approvalGate";
import { ApprovingToolExecutor } from "../shared/agentChat/approvingToolExecutor";
import {
  CHAT_APPROVAL_REJECT_PREFIX,
  deniedToolError,
} from "../shared/agentChat/approvalPolicy";
import { isAbortError } from "../shared/agent/streamProvider";
import type { ToolCall, ToolResult } from "../shared/tools/types";

function okResult(name: string): ToolResult {
  return { ok: true, tool: name as never, data: { wrote: true } };
}

describe("ApprovingToolExecutor", () => {
  it("safe read bypasses gate; host once", async () => {
    const gate = new ApprovalGate();
    const request = vi.spyOn(gate, "request");
    const inner = vi.fn(async (call: ToolCall) => okResult(call.name));
    const ex = new ApprovingToolExecutor({ execute: inner }, gate);
    const r = await ex.execute({ name: "range.read", arguments: { sheetName: "S", range: "A1" } });
    expect(r.ok).toBe(true);
    expect(inner).toHaveBeenCalledTimes(1);
    expect(request).not.toHaveBeenCalled();
  });

  it("moderate approve passes raw args to host", async () => {
    const gate = new ApprovalGate();
    const inner = vi.fn(async (call: ToolCall) => {
      expect(call.arguments).toEqual({
        sheetName: "S",
        range: "A1",
        values: [["secret-raw"]],
      });
      return okResult(call.name);
    });
    const ex = new ApprovingToolExecutor({ execute: inner }, gate, () => ({
      toolCallId: "c1",
      round: 2,
    }));
    const p = ex.execute({
      name: "range.write",
      arguments: { sheetName: "S", range: "A1", values: [["secret-raw"]] },
    });
    // wait microtask for request
    await Promise.resolve();
    const pending = gate.getPending();
    expect(pending?.toolCallId).toBe("c1");
    expect(pending?.round).toBe(2);
    const pub = JSON.stringify(pending);
    expect(pub).not.toContain("secret-raw");
    expect(pub).toContain("[grid");
    expect(gate.approve(pending!.requestId)).toBe(true);
    await expect(p).resolves.toMatchObject({ ok: true });
    expect(inner).toHaveBeenCalledTimes(1);
  });

  it("reject returns fixed failure without host", async () => {
    const gate = new ApprovalGate();
    const inner = vi.fn(async (call: ToolCall) => okResult(call.name));
    const ex = new ApprovingToolExecutor({ execute: inner }, gate);
    const p = ex.execute({
      name: "range.write",
      arguments: { sheetName: "S", range: "A1", values: [["x"]] },
    });
    await Promise.resolve();
    gate.reject(gate.getPending()!.requestId);
    const r = await p;
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error).toContain(CHAT_APPROVAL_REJECT_PREFIX);
    expect(inner).not.toHaveBeenCalled();
  });

  it("cancel throws AbortError; unknown denies", async () => {
    const gate = new ApprovalGate();
    const inner = vi.fn(async (call: ToolCall) => okResult(call.name));
    const ex = new ApprovingToolExecutor({ execute: inner }, gate);
    const p = ex.execute({
      name: "sheet.delete",
      arguments: { sheetName: "S" },
    });
    await Promise.resolve();
    gate.cancelAll("stop");
    await expect(p).rejects.toSatisfy((e) => isAbortError(e));
    expect(inner).not.toHaveBeenCalled();

    const denied = await ex.execute({
      name: "not.a.tool" as never,
      arguments: {},
    });
    expect(denied.ok).toBe(false);
    if (!denied.ok) expect(denied.error).toBe(deniedToolError("not.a.tool"));
  });
});
