import { describe, expect, it } from "vitest";
import {
  buildChatApprovalBoundaryLines,
  CHAT_APPROVAL_PROMPT_MARKER,
  composeChatApprovalSystemPrompt,
} from "../shared/agentChat/chatApprovalPrompt";
import {
  DEFAULT_PERMISSION_MODE,
  PERMISSION_MODES,
  type PermissionMode,
} from "../shared/agentChat/approvalPolicy";
import { listChatTools } from "../shared/agentChat/chatToolPolicy";
import type { ToolDefinition, ToolName } from "../shared/tools/types";
import { ChatController } from "../shared/agentChat/chatController";
import { ProviderStore } from "../shared/provider";
import { MockHostAdapter } from "./mockHost";

function fakeTools(safeCount: number, approvalCount: number): ToolDefinition[] {
  const tools: ToolDefinition[] = [];
  for (let i = 0; i < safeCount; i += 1) {
    tools.push({
      name: `safe.tool_${i}` as ToolName,
      description: "safe",
      riskLevel: "safe",
      parameters: { type: "object", properties: {} },
    });
  }
  for (let i = 0; i < approvalCount; i += 1) {
    tools.push({
      name: `write.tool_${i}` as ToolName,
      description: "write",
      riskLevel: i % 2 === 0 ? "moderate" : "dangerous",
      parameters: { type: "object", properties: {} },
    });
  }
  return tools;
}

const MODE_EXPECTATIONS: Record<
  PermissionMode,
  {
    mustInclude: RegExp[];
    mustNotInclude: RegExp[];
  }
> = {
  normal: {
    mustInclude: [
      /当前审批模式：normal/,
      /所有工具调用（含 safe/,
      /均需用户确认/,
      /未在工具注册表中的未知工具一律拒绝/,
    ],
    mustNotInclude: [
      /可直接执行（safe）/,
      /自动执行/,
      /用户已.*逐项批准/,
    ],
  },
  auto_approve_safe: {
    mustInclude: [
      /当前审批模式：auto_approve_safe/,
      /safe 工具可直接执行/,
      /moderate\/dangerous/,
      /需用户确认/,
      /可直接执行（safe）/,
      /需确认后执行（moderate\/dangerous）/,
    ],
    mustNotInclude: [
      /所有工具调用（含 safe/,
      /用户已.*逐项批准/,
    ],
  },
  confirm_all: {
    mustInclude: [
      /当前审批模式：confirm_all/,
      /已注册的已知工具将自动执行/,
      /不等于用户已对每次调用做过逐项批准/,
      /未知或未注册工具仍会被拒绝/,
      /可自动执行的已知工具/,
    ],
    mustNotInclude: [
      /均需用户确认后才会执行/,
      /可直接执行（safe）/,
      /用户已逐项批准/,
    ],
  },
};

describe("chatApprovalPrompt mode boundaries", () => {
  it("default permission mode is auto_approve_safe", () => {
    expect(DEFAULT_PERMISSION_MODE).toBe("auto_approve_safe");
  });

  for (const mode of PERMISSION_MODES) {
    it(`table: ${mode} boundary matches executor gate semantics`, () => {
      const tools = fakeTools(2, 3);
      const text = buildChatApprovalBoundaryLines(mode, tools).join("\n");
      expect(text).toContain(CHAT_APPROVAL_PROMPT_MARKER);
      for (const re of MODE_EXPECTATIONS[mode].mustInclude) {
        expect(text, String(re)).toMatch(re);
      }
      for (const re of MODE_EXPECTATIONS[mode].mustNotInclude) {
        expect(text, String(re)).not.toMatch(re);
      }
      // Dynamic tool names from provided registry — no hardcoded counts.
      expect(text).toContain("safe.tool_0");
      expect(text).toContain("safe.tool_1");
      expect(text).toContain("write.tool_0");
      expect(text).toContain("write.tool_2");
    });
  }

  it("tool registry size changes do not break assertions (count-agnostic)", () => {
    const small = buildChatApprovalBoundaryLines("auto_approve_safe", fakeTools(1, 1)).join(
      "\n",
    );
    const large = buildChatApprovalBoundaryLines("auto_approve_safe", fakeTools(5, 7)).join(
      "\n",
    );
    expect(small).toMatch(/可直接执行（safe）：safe\.tool_0/);
    expect(large).toContain("safe.tool_4");
    expect(large).toContain("write.tool_6");
    // Never hardcode absolute tool counts in the boundary prose.
    expect(small).not.toMatch(/\b71\b/);
    expect(large).not.toMatch(/\b71\b/);
  });

  it("live listChatTools names appear without hardcoding length", () => {
    const tools = listChatTools();
    expect(tools.length).toBeGreaterThan(0);
    const text = buildChatApprovalBoundaryLines("normal", tools).join("\n");
    const safe = tools.filter((t) => t.riskLevel === "safe").map((t) => t.name);
    const risky = tools
      .filter((t) => t.riskLevel === "moderate" || t.riskLevel === "dangerous")
      .map((t) => t.name);
    for (const name of safe.slice(0, 3)) expect(text).toContain(name);
    for (const name of risky.slice(0, 3)) expect(text).toContain(name);
    expect(text).toContain(`需确认后执行的工具：${tools.map((t) => t.name).join(", ")}`);
  });

  it("unknown-tool defense is present in every mode", () => {
    for (const mode of PERMISSION_MODES) {
      const text = buildChatApprovalBoundaryLines(mode, fakeTools(1, 1)).join("\n");
      expect(text).toMatch(/未知|未注册/);
      expect(text).toMatch(/拒绝/);
    }
  });

  it("composeChatApprovalSystemPrompt defaults to auto_approve_safe and embeds marker", () => {
    const prompt = composeChatApprovalSystemPrompt({
      routing: { content: "hello" },
      tools: fakeTools(1, 1),
    });
    expect(prompt).toContain(CHAT_APPROVAL_PROMPT_MARKER);
    expect(prompt).toMatch(/当前审批模式：auto_approve_safe/);
  });

  it("invalid permissionMode falls back to default mapping text", () => {
    const text = buildChatApprovalBoundaryLines("bogus", fakeTools(1, 0)).join("\n");
    expect(text).toMatch(/当前审批模式：auto_approve_safe/);
  });

  it("confirm_all never claims per-call user pre-approval", () => {
    const text = buildChatApprovalBoundaryLines("confirm_all", fakeTools(2, 2)).join("\n");
    expect(text).toMatch(/不等于用户已对每次调用做过逐项批准/);
    expect(text).not.toMatch(/用户已经批准了所有工具/);
    expect(text).not.toMatch(/用户已逐项批准/);
  });
});

describe("ChatController composes approval prompt with live permission mode", () => {
  function storeOk() {
    const store = new ProviderStore();
    store.add({
      name: "o",
      provider: "openai",
      apiKey: "sk-test",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "openai",
    });
    return store;
  }

  it("re-reads mode each turn (same getter as executor)", async () => {
    let mode: PermissionMode = "normal";
    const systemPrompts: string[] = [];
    const recording = {
      async *streamChat(req: { systemPrompt?: string; messages?: { role: string; content: string }[] }) {
        const sys =
          typeof req.systemPrompt === "string"
            ? req.systemPrompt
            : (req.messages ?? [])
                .filter((m) => m.role === "system")
                .map((m) => m.content)
                .join("\n");
        systemPrompts.push(sys);
        yield { type: "text_delta" as const, delta: "ok" };
        yield { type: "finish" as const, reason: "stop" as const };
      },
    };

    const controller = new ChatController({
      store: storeOk(),
      host: new MockHostAdapter(),
      getPermissionMode: () => mode,
      createProvider: () => ({ ok: true, provider: recording }),
    });

    await controller.send("turn1");
    mode = "confirm_all";
    await controller.send("turn2");
    mode = "auto_approve_safe";
    await controller.send("turn3");

    expect(systemPrompts.length).toBe(3);
    expect(systemPrompts[0]).toMatch(/当前审批模式：normal/);
    expect(systemPrompts[1]).toMatch(/当前审批模式：confirm_all/);
    expect(systemPrompts[2]).toMatch(/当前审批模式：auto_approve_safe/);
    for (const p of systemPrompts) {
      expect(p).toContain(CHAT_APPROVAL_PROMPT_MARKER);
    }
  });
});
