import { describe, expect, it } from "vitest";
import {
  formatToolArgs,
  formatToolOutcome,
  mapChatError,
  projectTraceEvent,
  safeJson,
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
    expect(formatted.text).toMatch(/omitted binary|binary\/base64|\[binary/);
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

describe("budgeted safeJson", () => {
  it("omits known binary keys without reading full payload into output", () => {
    const text = safeJson({
      imageBase64: "A".repeat(5000),
      base64: "B".repeat(2000),
      ok: true,
    });
    expect(text).toMatch(/omitted binary|binary\/base64/);
    expect(text).not.toContain("A".repeat(40));
    expect(text).not.toContain("B".repeat(40));
  });

  it("limits depth, keys, array items; skips budget-exceeded getters", () => {
    let deep: Record<string, unknown> = { v: 1 };
    for (let i = 0; i < 10; i += 1) deep = { nested: deep };
    const deepText = safeJson(deep);
    expect(deepText).toContain("max depth");

    const manyKeys: Record<string, unknown> = {};
    for (let i = 0; i < 40; i += 1) manyKeys[`k${i}`] = i;
    const keysText = safeJson(manyKeys);
    expect(keysText).toMatch(/\+/);

    const arrText = safeJson(Array.from({ length: 30 }, (_, i) => i));
    expect(arrText).toMatch(/more/);

    let accessed = 0;
    const obj: Record<string, unknown> = {};
    for (let i = 0; i < 20; i += 1) {
      Object.defineProperty(obj, `p${i}`, {
        enumerable: true,
        get() {
          accessed += 1;
          if (i >= 12) throw new Error("should not access beyond budget");
          return i;
        },
      });
    }
    // Only first MAX_OBJECT_KEYS keys are accessed.
    expect(() => safeJson(obj)).not.toThrow();
    expect(accessed).toBeLessThanOrEqual(12);
  });

  it("handles circular refs without throwing", () => {
    const a: Record<string, unknown> = { x: 1 };
    a.self = a;
    expect(() => safeJson(a)).not.toThrow();
    expect(safeJson(a)).toContain("circular");
  });
});
