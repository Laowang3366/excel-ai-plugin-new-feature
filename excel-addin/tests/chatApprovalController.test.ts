import { describe, expect, it, vi } from "vitest";
import {
  ChatController,
  CHAT_APPROVAL_PROMPT_MARKER,
  CHAT_APPROVAL_REJECT_PREFIX,
  listChatTools,
} from "../shared/agentChat";
import { ScriptedStreamProvider, toolCallThenFinish, textThenStop } from "../shared/agent/index";
import { ProviderStore } from "../shared/provider";
import { MockHostAdapter } from "./mockHost";
import type { ChatTraceEvent } from "../shared/agentChat";

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

describe("ChatController approval path", () => {
  it("write awaits approval; approve executes host with raw args", async () => {
    const host = new MockHostAdapter();
    const write = vi.spyOn(host, "writeRange");
    const events: ChatTraceEvent[] = [];
    let controller!: ChatController;

    const provider = new ScriptedStreamProvider({
      rounds: [
        toolCallThenFinish(
          "call_w",
          "range.write",
          '{"sheetName":"Sheet1","range":"A1","values":[["raw-value"]]}',
        ),
        textThenStop("done after approve"),
      ],
    });

    controller = new ChatController({
      store: storeOk(),
      host,
      createProvider: () => ({ ok: true, provider }),
      onEvent: (e) => {
        events.push(e);
        if (e.type === "approval_needed") {
          // approve on next microtask to let status settle
          queueMicrotask(() => {
            expect(controller.getState().status).toBe("awaiting_approval");
            expect(controller.getState().pendingApproval?.name).toBe("range.write");
            const pub = JSON.stringify(controller.getState().pendingApproval);
            expect(pub).not.toContain("raw-value");
            expect(pub).toContain("[grid");
            expect(controller.approve(e.request.requestId)).toBe(true);
          });
        }
      },
    });

    const result = await controller.send("write A1");
    expect(result.turnStatus).toBe("completed");
    expect(write).toHaveBeenCalledTimes(1);
    expect(write.mock.calls[0]?.slice(0,3)).toEqual(["Sheet1","A1",[["raw-value"]]]);
    expect(events.some((e) => e.type === "approval_needed")).toBe(true);
    expect(events.some((e) => e.type === "approval_resolved")).toBe(true);
    expect(result.run?.assistantText).toContain("done after approve");
  });

  it("reject continues loop with tool failure; no host write", async () => {
    const host = new MockHostAdapter();
    const write = vi.spyOn(host, "writeRange");
    let controller!: ChatController;
    const provider = new ScriptedStreamProvider({
      rounds: [
        toolCallThenFinish(
          "call_w",
          "range.write",
          '{"sheetName":"Sheet1","range":"A1","values":[["x"]]}',
        ),
        textThenStop("after reject"),
      ],
    });
    controller = new ChatController({
      store: storeOk(),
      host,
      createProvider: () => ({ ok: true, provider }),
      onEvent: (e) => {
        if (e.type === "approval_needed") {
          queueMicrotask(() => controller.reject(e.request.requestId));
        }
      },
    });
    const result = await controller.send("try write");
    expect(result.turnStatus).toBe("completed");
    expect(write).not.toHaveBeenCalled();
    const toolMsgs = result.run?.messages.filter((m) => m.role === "tool") ?? [];
    expect(toolMsgs.some((m) => m.content.includes(CHAT_APPROVAL_REJECT_PREFIX))).toBe(
      true,
    );
    expect(result.run?.assistantText).toContain("after reject");
  });

  it("stop while pending aborts turn; cancel-then-late-approve host=0", async () => {
    const host = new MockHostAdapter();
    const write = vi.spyOn(host, "writeRange");
    let controller!: ChatController;
    const provider = new ScriptedStreamProvider({
      rounds: [
        toolCallThenFinish(
          "call_w",
          "range.write",
          '{"sheetName":"Sheet1","range":"A1","values":[["x"]]}',
        ),
      ],
      onExhausted: "error",
    });
    controller = new ChatController({
      store: storeOk(),
      host,
      createProvider: () => ({ ok: true, provider }),
      onEvent: (e) => {
        if (e.type === "approval_needed") {
          queueMicrotask(() => {
            controller.stop();
            // late approve must no-op
            expect(controller.approve(e.request.requestId)).toBe(false);
          });
        }
      },
    });
    const result = await controller.send("write then stop");
    expect(result.turnStatus).toBe("aborted");
    expect(write).not.toHaveBeenCalled();
    expect(controller.getState().pendingApproval).toBeNull();
    expect(controller.getState().status).toBe("idle");
  });

  it("exposes full tool list and approval prompt marker; two turns new gates", async () => {
    const host = new MockHostAdapter();
    let seenTools = 0;
    const provider = new ScriptedStreamProvider({
      rounds: [
        (ctx) => {
          seenTools = ctx.request.tools.length;
          return textThenStop("ok1");
        },
        () => textThenStop("ok2"),
      ],
    });
    // Scripted provider API check - use function rounds if supported
    // Fall back: inspect via lastRequest if available after run
    const controller = new ChatController({
      store: storeOk(),
      host,
      createProvider: () => ({ ok: true, provider }),
      composeSystemPrompt: (msg) =>
        `sys ${msg}\n## ${CHAT_APPROVAL_PROMPT_MARKER}\napproval`,
    });
    // If scripted doesn't pass tools on ctx, assert listChatTools length separately
    expect(listChatTools()).toHaveLength(62);

    // Use a provider that records tools from AgentLoop
    const bodies: number[] = [];
    const recording = {
      async *streamChat(req: { tools: { name: string }[] }) {
        bodies.push(req.tools.length);
        yield { type: "text_delta" as const, delta: "hi" };
        yield { type: "finish" as const, reason: "stop" as const };
      },
    };
    const c2 = new ChatController({
      store: storeOk(),
      host,
      createProvider: () => ({ ok: true, provider: recording }),
      composeSystemPrompt: (msg) =>
        `PROMPT ${msg}\n## ${CHAT_APPROVAL_PROMPT_MARKER}\n`,
    });
    const r1 = await c2.send("t1");
    expect(r1.turnStatus).toBe("completed");
    expect(bodies[0]).toBe(62);
    const r2 = await c2.send("t2");
    expect(r2.turnStatus).toBe("completed");
    expect(bodies[1]).toBe(62);
    // prompt marker used
    expect(ChatController.approvalMarker).toBe(CHAT_APPROVAL_PROMPT_MARKER);
    void controller;
    void seenTools;
  });
});
