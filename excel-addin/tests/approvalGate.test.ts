import { describe, expect, it } from "vitest";
import { ApprovalGate } from "../shared/agentChat/approvalGate";
import { isAbortError } from "../shared/agent/streamProvider";

describe("ApprovalGate", () => {
  it("approve / reject settle single pending; wrong id no-op", async () => {
    const gate = new ApprovalGate();
    const events: string[] = [];
    gate.subscribe((e) => events.push(e.type + ":" + ("decision" in e ? e.decision : "req")));

    const p = gate.request({
      name: "range.write",
      riskLevel: "moderate",
      destructive: false,
      argsPreview: { sheetName: "S" },
      impactHint: "write",
    });
    const pending = gate.getPending();
    expect(pending?.name).toBe("range.write");
    expect(JSON.stringify(pending)).not.toContain("password");
    expect(gate.approve("wrong")).toBe(false);
    expect(gate.approve(pending!.requestId)).toBe(true);
    await expect(p).resolves.toBe("approved");
    expect(gate.getPending()).toBeNull();

    const p2 = gate.request({
      name: "sheet.delete",
      riskLevel: "moderate",
      destructive: true,
      argsPreview: {},
      impactHint: "del",
    });
    const id2 = gate.getPending()!.requestId;
    expect(gate.reject(id2)).toBe(true);
    await expect(p2).resolves.toBe("rejected");
    expect(events.some((e) => e.startsWith("requested"))).toBe(true);
    expect(events.some((e) => e.includes("approved"))).toBe(true);
    expect(events.some((e) => e.includes("rejected"))).toBe(true);
  });

  it("cancelAll aborts pending and arms cancelled for later requests", async () => {
    const gate = new ApprovalGate();
    const p = gate.request({
      name: "range.write",
      riskLevel: "moderate",
      destructive: false,
      argsPreview: {},
      impactHint: "x",
    });
    gate.cancelAll("stop");
    await expect(p).rejects.toSatisfy((e) => isAbortError(e));
    await expect(
      gate.request({
        name: "range.write",
        riskLevel: "moderate",
        destructive: false,
        argsPreview: {},
        impactHint: "x",
      }),
    ).rejects.toSatisfy((e) => isAbortError(e));
    expect(gate.approve("any")).toBe(false);
  });

  it("only one pending at a time", async () => {
    const gate = new ApprovalGate();
    void gate.request({
      name: "a",
      riskLevel: "moderate",
      destructive: false,
      argsPreview: {},
      impactHint: "x",
    });
    await expect(
      gate.request({
        name: "b",
        riskLevel: "moderate",
        destructive: false,
        argsPreview: {},
        impactHint: "x",
      }),
    ).rejects.toThrow(/already has a pending/);
  });
});
