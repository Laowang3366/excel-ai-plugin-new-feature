import { describe, expect, it } from "vitest";
import { AgentLoop } from "../shared/agent/agentLoop";
import {
  estimateRequestTokens,
  estimateTokens,
  groupMessageAtoms,
  resolveMessageTokenBudget,
  resolveOutputReserve,
  trimMessagesForRequest,
} from "../shared/agent/historyBudget";
import {
  ScriptedStreamProvider,
  textThenStop,
  toolCallThenFinish,
} from "../shared/agent/scriptedProvider";
import type { AgentMessage, StreamChatRequest } from "../shared/agent/types";
import type { ToolDefinition } from "../shared/tools/types";
import { ToolExecutor } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";

function user(content: string): AgentMessage {
  return { role: "user", content };
}

function assistant(
  content: string,
  toolCalls?: AgentMessage["toolCalls"],
): AgentMessage {
  return toolCalls?.length
    ? { role: "assistant", content, toolCalls }
    : { role: "assistant", content };
}

function tool(toolCallId: string, content: string): AgentMessage {
  return { role: "tool", content, toolCallId };
}

const tinyTools: ToolDefinition[] = [
  {
    name: "host.status",
    description: "status",
    riskLevel: "safe",
    parameters: { type: "object", properties: {} },
  },
];

/** English chars ≈ 0.25 tokens each; size content to force trimming under small windows. */
function big(tokens: number): string {
  return "x".repeat(tokens * 4);
}

function hardLimit(ctx: number): number {
  return ctx - resolveOutputReserve(ctx);
}

/** Tool message must sit after an assistant that lists its toolCallId. */
function assertNoOrphanTools(msgs: AgentMessage[]): void {
  for (let i = 0; i < msgs.length; i += 1) {
    const m = msgs[i];
    if (m.role !== "tool") continue;
    let j = i - 1;
    while (j >= 0 && msgs[j].role === "tool") j -= 1;
    expect(msgs[j]?.role).toBe("assistant");
    expect(msgs[j]?.toolCalls?.some((c) => c.id === m.toolCallId)).toBe(true);
  }
}

/** Assistant (except leading malformed) must follow a user in the same turn. */
function assertNoOrphanAssistant(msgs: AgentMessage[]): void {
  for (let i = 0; i < msgs.length; i += 1) {
    if (msgs[i].role !== "assistant") continue;
    let j = i - 1;
    while (j >= 0 && msgs[j].role !== "user") {
      // Stay within turn: stop if we hit another user... we only walk back.
      j -= 1;
    }
    expect(j).toBeGreaterThanOrEqual(0);
    expect(msgs[j].role).toBe("user");
  }
}

describe("estimateTokens / estimateRequestTokens", () => {
  it("matches desktop-style zh/en heuristics", () => {
    expect(estimateTokens("")).toBe(0);
    expect(estimateTokens("hello world")).toBe(Math.ceil(11 / 4));
    expect(estimateTokens("你好世界")).toBe(Math.ceil(4 / 1.5));
  });

  it("includes system, tools, and toolCalls overhead", () => {
    const base = estimateRequestTokens({ messages: [user("hi")] });
    const withPrompt = estimateRequestTokens({
      systemPrompt: "system prompt text",
      messages: [user("hi")],
    });
    const withTools = estimateRequestTokens({
      systemPrompt: "system prompt text",
      messages: [user("hi")],
      tools: tinyTools,
    });
    const withToolCalls = estimateRequestTokens({
      messages: [
        assistant("call", [
          { id: "c1", name: "host.status", argumentsJson: "{}" },
        ]),
      ],
    });
    expect(withPrompt).toBeGreaterThan(base);
    expect(withTools).toBeGreaterThan(withPrompt);
    expect(withToolCalls).toBeGreaterThan(
      estimateRequestTokens({ messages: [assistant("call")] }),
    );
  });
});

describe("groupMessageAtoms", () => {
  it("groups user-initiated turns (user + following assistant/tool until next user)", () => {
    const messages: AgentMessage[] = [
      user("u1"),
      assistant("a1", [{ id: "c1", name: "host.status", argumentsJson: "{}" }]),
      tool("c1", '{"ok":true}'),
      tool("c1b", '{"ok":true}'),
      assistant("after-tools"),
      user("u2"),
      assistant("done"),
    ];
    const atoms = groupMessageAtoms(messages);
    expect(atoms.map((a) => a.map((m) => m.role))).toEqual([
      ["user", "assistant", "tool", "tool", "assistant"],
      ["user", "assistant"],
    ]);
  });

  it("bundles leading non-user messages until the next user", () => {
    const messages: AgentMessage[] = [
      assistant("orphan-lead", [
        { id: "c0", name: "host.status", argumentsJson: "{}" },
      ]),
      tool("c0", '{"ok":true}'),
      user("u1"),
      assistant("a1"),
    ];
    const atoms = groupMessageAtoms(messages);
    expect(atoms.map((a) => a.map((m) => m.role))).toEqual([
      ["assistant", "tool"],
      ["user", "assistant"],
    ]);
  });
});


describe("resolveOutputReserve / resolveMessageTokenBudget", () => {
  it("reserves desktop off-mode output budget from context window", () => {
    // floor(128000*0.06)=7680 clamped to [4096,16384], then min(7680, 10240)=7680
    expect(resolveOutputReserve(128_000)).toBe(7_680);
    // small window: min(4096, ctx*0.08)
    expect(resolveOutputReserve(5_000)).toBe(400);
    expect(resolveOutputReserve(0)).toBe(resolveOutputReserve(128_000));
  });

  it("subtracts system/tools overhead and output reserve from context", () => {
    const budget = resolveMessageTokenBudget({
      contextWindowSize: 10_000,
      systemPrompt: "sys",
      tools: tinyTools,
    });
    const overhead = estimateRequestTokens({
      systemPrompt: "sys",
      tools: tinyTools,
      messages: [],
    });
    expect(budget).toBe(10_000 - overhead - resolveOutputReserve(10_000));
    expect(budget).toBeGreaterThan(0);
  });
});

describe("trimMessagesForRequest", () => {
  it("returns a full copy when under budget (no trim)", () => {
    const messages = [user("old"), assistant("reply"), user("now")];
    const out = trimMessagesForRequest({
      messages,
      systemPrompt: "sys",
      tools: tinyTools,
      contextWindowSize: 128_000,
      protectFromIndex: 2,
    });
    expect(out).toEqual(messages);
    expect(out).not.toBe(messages);
  });

  it("drops oldest history atoms first", () => {
    const oldUser = user("OLD_USER|" + big(1_500));
    const oldAsst = assistant("OLD_ASST|" + big(1_500));
    const midUser = user("MID_USER|" + big(1_500));
    const midAsst = assistant("MID_ASST|" + big(1_500));
    const current = user("current-turn");
    const messages = [oldUser, oldAsst, midUser, midAsst, current];
    const protectFromIndex = 4;
    const ctx = 5_000;
    const full = estimateRequestTokens({
      systemPrompt: "S",
      tools: tinyTools,
      messages,
    });
    expect(full).toBeGreaterThan(hardLimit(ctx));

    const out = trimMessagesForRequest({
      messages,
      systemPrompt: "S",
      tools: tinyTools,
      contextWindowSize: ctx,
      protectFromIndex,
    });

    expect(out[out.length - 1]).toEqual(current);
    expect(out.some((m) => m.content.startsWith("OLD_USER|"))).toBe(false);
    // Dropping OLD_USER turn also drops its assistant (turn atom).
    expect(out.some((m) => m.content.startsWith("OLD_ASST|"))).toBe(false);
    expect(messages).toHaveLength(5);
    expect(messages[0]).toBe(oldUser);
    assertNoOrphanAssistant(out);
    assertNoOrphanTools(out);
    expect(
      estimateRequestTokens({
        systemPrompt: "S",
        tools: tinyTools,
        messages: out,
      }),
    ).toBeLessThanOrEqual(hardLimit(ctx));
  });

  it("drops whole user turns (never leaves assistant without its user)", () => {
    const turn1User = user("T1_USER|" + big(1_800));
    const turn1Asst = assistant("T1_ASST|" + big(1_800));
    const turn2User = user("T2_USER|" + big(1_800));
    const turn2Asst = assistant("T2_ASST|" + big(1_800));
    const current = user("now");
    const messages = [turn1User, turn1Asst, turn2User, turn2Asst, current];

    const out = trimMessagesForRequest({
      messages,
      systemPrompt: "sys",
      tools: tinyTools,
      contextWindowSize: 4_000,
      protectFromIndex: 4,
    });

    expect(out[out.length - 1]).toEqual(current);
    // Dropping T1 must drop both user and assistant together.
    const hasT1User = out.some((m) => m.content.startsWith("T1_USER|"));
    const hasT1Asst = out.some((m) => m.content.startsWith("T1_ASST|"));
    expect(hasT1User).toBe(hasT1Asst);
    // Must not keep a reply whose question was dropped.
    expect(hasT1Asst && !hasT1User).toBe(false);
    assertNoOrphanAssistant(out);
    assertNoOrphanTools(out);
  });

  it("never splits assistant toolCalls from tool results when dropping turns", () => {
    const oldTurn: AgentMessage[] = [
      user("OLD_U|" + big(2_000)),
      assistant("use tool", [
        { id: "c1", name: "host.status", argumentsJson: "{}" },
      ]),
      tool("c1", "TOOL|" + big(2_000)),
      assistant("after tool|" + big(500)),
    ];
    const midTurn: AgentMessage[] = [
      user("MID_U|" + big(2_000)),
      assistant("mid reply|" + big(2_000)),
    ];
    const current = user("now");
    const messages = [...oldTurn, ...midTurn, current];

    const out = trimMessagesForRequest({
      messages,
      systemPrompt: "sys",
      tools: tinyTools,
      contextWindowSize: 4_000,
      protectFromIndex: messages.length - 1,
    });

    expect(out[out.length - 1]?.content).toBe("now");
    assertNoOrphanTools(out);
    assertNoOrphanAssistant(out);

    const hasOldTool = out.some((m) => m.role === "tool" && m.toolCallId === "c1");
    const hasOldAsst = out.some(
      (m) => m.role === "assistant" && m.toolCalls?.some((c) => c.id === "c1"),
    );
    const hasOldUser = out.some((m) => m.content.startsWith("OLD_U|"));
    // Tool chain and its user question stay or leave together.
    expect(hasOldTool).toBe(hasOldAsst);
    expect(hasOldAsst).toBe(hasOldUser);
  });

  it("never drops protected current-turn suffix even when over budget", () => {
    const current = user(big(5_000));
    const messages = [user("old"), current];
    const out = trimMessagesForRequest({
      messages,
      systemPrompt: "sys",
      tools: tinyTools,
      contextWindowSize: 1_200,
      protectFromIndex: 1,
    });
    expect(out).toEqual([current]);
  });

  it("does not mutate input messages", () => {
    const messages = [user("a"), assistant("b"), user("c")];
    const snapshot = JSON.stringify(messages);
    trimMessagesForRequest({
      messages,
      systemPrompt: "sys",
      contextWindowSize: 100,
      protectFromIndex: 2,
    });
    expect(JSON.stringify(messages)).toBe(snapshot);
  });
});

describe("AgentLoop request history budget", () => {
  it("does not trim when contextWindowSize omitted", async () => {
    const history: AgentMessage[] = [user(big(200)), assistant(big(200))];
    const historySnapshot = JSON.stringify(history);
    const provider = new ScriptedStreamProvider({
      rounds: [textThenStop("ok")],
    });
    const result = await new AgentLoop({
      provider,
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "sys",
    }).run({ userMessage: "now", history });

    expect(provider.lastRequest?.messages.map((m) => m.role)).toEqual([
      "user",
      "assistant",
      "user",
    ]);
    expect(result.messages.map((m) => m.role)).toEqual([
      "user",
      "assistant",
      "user",
      "assistant",
    ]);
    expect(JSON.stringify(history)).toBe(historySnapshot);
  });

  it("trims request copies across rounds but keeps full run messages", async () => {
    const history: AgentMessage[] = [
      user("H0|" + big(1_500)),
      assistant("H1|" + big(1_500)),
      user("H2|" + big(1_500)),
      assistant("H3|" + big(1_500)),
    ];
    const historySnapshot = JSON.stringify(history);
    const seen: AgentMessage[][] = [];

    const provider = new ScriptedStreamProvider({
      rounds: [
        (ctx) => {
          seen.push(ctx.request.messages.slice());
          return toolCallThenFinish("c1", "host.status", "{}");
        },
        (ctx) => {
          seen.push(ctx.request.messages.slice());
          return textThenStop("final");
        },
      ],
    });

    const ctxWindow = 5_000;
    const result = await new AgentLoop({
      provider,
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "S",
      tools: tinyTools,
      contextWindowSize: ctxWindow,
    }).run({ userMessage: "current", history });

    expect(result.status).toBe("completed");
    expect(seen).toHaveLength(2);

    // Full internal/result history retains prior turns + current turn chain.
    expect(result.messages.map((m) => m.role)).toEqual([
      "user",
      "assistant",
      "user",
      "assistant",
      "user",
      "assistant",
      "tool",
      "assistant",
    ]);
    expect(result.messages[0].content).toBe(history[0].content);
    expect(JSON.stringify(history)).toBe(historySnapshot);

    const limit = hardLimit(ctxWindow);
    for (const req of seen) {
      const est = estimateRequestTokens({
        systemPrompt: "S",
        tools: tinyTools,
        messages: req,
      });
      expect(est).toBeLessThanOrEqual(limit);
      expect(req.some((m) => m.role === "user" && m.content === "current")).toBe(
        true,
      );
    }

    // Round 2 request includes current-turn tool chain; no orphan tools.
    const round2 = seen[1];
    const toolIdx = round2.findIndex((m) => m.role === "tool");
    expect(toolIdx).toBeGreaterThan(0);
    expect(round2[toolIdx - 1]?.role).toBe("assistant");
    expect(round2[toolIdx - 1]?.toolCalls?.length).toBeGreaterThan(0);

    // Oldest history should have been dropped from at least one request.
    expect(
      seen.some((req) => !req.some((m) => m.content.startsWith("H0|"))),
    ).toBe(true);
  });

  it("ChatController-style: committed history stays complete after budgeted run", async () => {
    const committed: AgentMessage[] = [
      user("C0|" + big(1_200)),
      assistant("C1|" + big(1_200)),
      user("C2|" + big(1_200)),
      assistant("C3|" + big(1_200)),
    ];
    const requests: StreamChatRequest[] = [];
    const ctxWindow = 3_500;
    const provider = new ScriptedStreamProvider({
      rounds: [
        (ctx) => {
          requests.push({
            systemPrompt: ctx.request.systemPrompt,
            messages: ctx.request.messages.slice(),
            tools: ctx.request.tools.slice(),
          });
          return textThenStop("answer");
        },
      ],
    });

    const history = committed.slice();
    const result = await new AgentLoop({
      provider,
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "sys",
      tools: tinyTools,
      contextWindowSize: ctxWindow,
    }).run({ userMessage: "q", history });

    expect(history).toEqual(committed);
    expect(result.messages.length).toBe(committed.length + 2);
    expect(result.messages.some((m) => m.content.startsWith("C0|"))).toBe(true);
    expect(requests[0].messages.some((m) => m.content === "q")).toBe(true);
    // Request must drop some old history (not merely exclude the new assistant).
    expect(requests[0].messages.length).toBeLessThan(committed.length + 1);
    expect(
      estimateRequestTokens({
        systemPrompt: "sys",
        tools: tinyTools,
        messages: requests[0].messages,
      }),
    ).toBeLessThanOrEqual(hardLimit(ctxWindow));
  });
});
