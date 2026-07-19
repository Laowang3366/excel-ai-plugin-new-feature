import { describe, expect, it, vi } from "vitest";
import {
  formatToolArgs,
  formatToolOutcome,
  mapChatError,
  MAX_PARSEABLE_TOOL_ARGS_CHARS,
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

  it("redacts secret keys in tool args and safeJson", () => {
    const shown = formatToolArgs('{"password":"secret","sheetName":"S"}');
    expect(shown).toContain("[REDACTED]");
    expect(shown).not.toContain('"secret"');
    expect(safeJson({ apiKey: "sk-test", values: [["a"]] })).toContain("[REDACTED]");
    expect(safeJson({ apiKey: "sk-test", values: [["a"]] })).not.toContain("sk-test");
  });

  it("omits oversized tool args without parsing or leaking secrets", () => {
    const secret = "super-secret-password-value";
    const b64 = "A".repeat(500);
    const values = JSON.stringify([["cell-secret-grid"]]);
    // Build valid JSON larger than parse threshold with secrets embedded.
    const pad = "x".repeat(
      MAX_PARSEABLE_TOOL_ARGS_CHARS + 200 - (secret.length + b64.length + values.length),
    );
    const huge = JSON.stringify({
      password: secret,
      imageBase64: b64,
      values: [["cell-secret-grid"]],
      pad,
    });
    expect(huge.length).toBeGreaterThan(MAX_PARSEABLE_TOOL_ARGS_CHARS);

    const parseSpy = vi.spyOn(JSON, "parse");
    const shown = formatToolArgs(huge);
    expect(shown).toBe(`[arguments JSON omitted: ${huge.length} chars]`);
    expect(shown).not.toContain(secret);
    expect(shown).not.toContain("password");
    expect(shown).not.toContain(b64.slice(0, 40));
    expect(shown).not.toContain("cell-secret-grid");
    expect(parseSpy).not.toHaveBeenCalled();
    parseSpy.mockRestore();

    // Boundary: length == threshold still parses + redacts; over-by-one omits.
    const base = '{"password":"secret-at-limit","pad":"';
    const suffix = '"}';
    const padLen = MAX_PARSEABLE_TOOL_ARGS_CHARS - base.length - suffix.length;
    expect(padLen).toBeGreaterThan(0);
    const atLimit = `${base}${"y".repeat(padLen)}${suffix}`;
    expect(atLimit.length).toBe(MAX_PARSEABLE_TOOL_ARGS_CHARS);
    const boundary = formatToolArgs(atLimit);
    expect(boundary).toContain("[REDACTED]");
    expect(boundary).not.toContain("secret-at-limit");
    const overByOne = `${base}${"y".repeat(padLen + 1)}${suffix}`;
    expect(overByOne.length).toBe(MAX_PARSEABLE_TOOL_ARGS_CHARS + 1);
    expect(formatToolArgs(overByOne)).toBe(
      `[arguments JSON omitted: ${overByOne.length} chars]`,
    );

    // Short ordinary args remain readable.
    expect(formatToolArgs('{"sheetName":"Sheet1","range":"A1"}')).toContain("Sheet1");

    // Invalid short JSON: length summary only, no raw prefix.
    const badShort = '{"password":"leaked"';
    const badShortOut = formatToolArgs(badShort);
    expect(badShortOut).toBe(`[invalid arguments JSON: ${badShort.length} chars]`);
    expect(badShortOut).not.toContain("leaked");
    expect(badShortOut).not.toContain("password");

    // Invalid long: omit-by-length without parse/leak.
    const badLong = `{"password":"leaked-long","pad":"${"z".repeat(MAX_PARSEABLE_TOOL_ARGS_CHARS)}"`;
    const parseSpy2 = vi.spyOn(JSON, "parse");
    const badLongOut = formatToolArgs(badLong);
    expect(badLongOut).toBe(`[arguments JSON omitted: ${badLong.length} chars]`);
    expect(badLongOut).not.toContain("leaked-long");
    expect(parseSpy2).not.toHaveBeenCalled();
    parseSpy2.mockRestore();
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

describe("approval trace projection", () => {
  it("projects needed/resolved without raw args", () => {
    const needed = projectTraceEvent(
      {
        type: "approval_needed",
        request: {
          requestId: "r1",
          name: "sheet.delete",
          riskLevel: "moderate",
          destructive: true,
          argsPreview: { sheetName: "S" },
          impactHint: "将删除工作表",
          createdAt: 1,
        },
      },
      9,
    );
    expect(needed?.kind).toBe("approval");
    expect(needed?.tone).toBe("warn");
    expect(needed?.text).toContain("待审批");
    expect(needed?.text).toContain("sheet.delete");
    expect(needed?.text).not.toContain("password");

    const approved = projectTraceEvent(
      {
        type: "approval_resolved",
        requestId: "r1",
        decision: "approved",
        request: {
          requestId: "r1",
          name: "sheet.delete",
          riskLevel: "moderate",
          destructive: true,
          argsPreview: {},
          impactHint: "x",
          createdAt: 1,
        },
      },
      10,
    );
    expect(approved?.text).toContain("已批准");
    expect(approved?.tone).toBe("ok");
  });
});
