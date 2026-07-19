import { describe, expect, it } from "vitest";
import {
  formatToolArgs,
  formatToolOutcome,
  mapChatError,
  projectTraceEvent,
  summarizePayload,
  truncateDisplay,
} from "../src/chat/chatPresentation";
import type { AgentToolOutcome } from "../shared/agent/types";

describe("chatPresentation", () => {
  it("truncates long text and collapses base64", () => {
    expect(truncateDisplay("a".repeat(200)).endsWith("…")).toBe(true);
    const b64 = "A".repeat(400);
    expect(summarizePayload(b64)).toMatch(/binary\/base64/);
    expect(formatToolArgs('{"x":"' + "y".repeat(300) + '"}').length).toBeLessThan(200);
  });

  it("maps preflight / cors / max_rounds / aborted errors in Chinese", () => {
    expect(mapChatError({ message: "API key 未设置", kind: "missing_key" })).toMatch(
      /模型供应商/,
    );
    expect(
      mapChatError({ message: "no active provider configured", kind: "parse" }),
    ).toMatch(/模型供应商/);
    expect(mapChatError({ message: "cors", kind: "cors" })).toMatch(/CORS/);
    expect(mapChatError(undefined, "max_rounds")).toMatch(/最大/);
    expect(mapChatError(undefined, "aborted")).toMatch(/停止/);
    expect(mapChatError({ message: "sk-secret-xyz", kind: "http", status: 401 })).not.toBe(
      "sk-secret-xyz",
    );
  });

  it("formats tool outcomes and projects trace events", () => {
    const ok: AgentToolOutcome = {
      kind: "host",
      toolName: "range.read",
      result: { ok: true, tool: "range.read", data: { imageBase64: "A".repeat(500) } },
    };
    const formatted = formatToolOutcome(ok);
    expect(formatted.tone).toBe("ok");
    expect(formatted.text).toMatch(/binary\/base64|\[binary/);
    expect(formatted.text).not.toContain("A".repeat(100));
    expect(formatted.text.length).toBeLessThan(220);

    const fail: AgentToolOutcome = {
      kind: "host",
      toolName: "range.write",
      result: {
        ok: false,
        tool: "range.write",
        error: "chat readonly: tool not allowed: range.write",
      },
    };
    expect(formatToolOutcome(fail).tone).toBe("fail");

    const parsed = projectTraceEvent(
      {
        type: "tool_call_parsed",
        round: 1,
        call: {
          id: "c1",
          name: "range.read",
          argumentsJson: '{"sheetName":"S","range":"A1"}',
        },
      },
      1,
    );
    expect(parsed?.text).toContain("range.read");
    expect(parsed?.text).toContain("sheetName");
  });
});
