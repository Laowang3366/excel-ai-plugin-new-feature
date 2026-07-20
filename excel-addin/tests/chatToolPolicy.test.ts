import { describe, expect, it } from "vitest";
import { TOOL_DEFINITIONS } from "../shared/tools";
import {
  classifyChatTool,
  listChatTools,
} from "../shared/agentChat/chatToolPolicy";
import {
  buildArgsPreview,
  isDestructiveTool,
  buildImpactHint,
} from "../shared/agentChat/approvalPreview";

describe("chat tool policy", () => {
  it("listChatTools equals registry set with fresh copies", () => {
    const listed = listChatTools();
    expect(listed).toHaveLength(TOOL_DEFINITIONS.length);
    expect(listed).toHaveLength(88);
    expect(listed.map((t) => t.name)).toEqual(TOOL_DEFINITIONS.map((t) => t.name));
    expect(listed).not.toBe(TOOL_DEFINITIONS);
    expect(listed[0]).not.toBe(TOOL_DEFINITIONS[0]);
    listed.pop();
    expect(TOOL_DEFINITIONS).toHaveLength(88);

    let direct = 0;
    let approval = 0;
    let deny = 0;
    for (const t of listChatTools()) {
      const d = classifyChatTool(t.name).disposition;
      if (d === "direct") direct += 1;
      else if (d === "approval") approval += 1;
      else deny += 1;
    }
    expect(direct + approval).toBe(88);
    expect(deny).toBe(0);
    expect(direct).toBe(listChatTools().filter((t) => t.riskLevel === "safe").length);
    expect(approval).toBe(
      listChatTools().filter(
        (t) => t.riskLevel === "moderate" || t.riskLevel === "dangerous",
      ).length,
    );
    expect(classifyChatTool("nope").disposition).toBe("deny");
  });
});

describe("approval preview", () => {
  it("redacts secrets/base64; limits structure; destructive hints", () => {
    const preview = buildArgsPreview({
      password: "p@ss",
      apiKey: "sk-secret",
      imageBase64: "A".repeat(2000),
      values: Array.from({ length: 20 }, () => ["x".repeat(100)]),
      nested: { a: { b: { c: { d: { e: 1 } } } } },
    });
    const text = JSON.stringify(preview);
    expect(text).toContain("[REDACTED]");
    expect(text).not.toContain("p@ss");
    expect(text).not.toContain("sk-secret");
    expect(text).not.toContain("A".repeat(40));
    expect(isDestructiveTool("sheet.delete", {})).toBe(true);
    expect(
      isDestructiveTool("sheet.operation", { operation: "delete" }),
    ).toBe(true);
    expect(isDestructiveTool("range.write", {})).toBe(false);
    expect(buildImpactHint("sheet.delete", {}, true)).toMatch(/删除/);
  });

  it("cycle-safe and budget getters", () => {
    const a: Record<string, unknown> = { x: 1 };
    a.self = a;
    expect(() => buildArgsPreview(a)).not.toThrow();
    let accessed = 0;
    const obj: Record<string, unknown> = {};
    for (let i = 0; i < 20; i += 1) {
      Object.defineProperty(obj, `k${i}`, {
        enumerable: true,
        get() {
          accessed += 1;
          if (i >= 12) throw new Error("over budget");
          return i;
        },
      });
    }
    expect(() => buildArgsPreview(obj)).not.toThrow();
    expect(accessed).toBeLessThanOrEqual(12);
  });
});
