import { describe, expect, it, vi } from "vitest";
import { ChatController } from "../shared/agentChat";
import type { ChatTraceEvent } from "../shared/agentChat";
import {
  ScriptedStreamProvider,
  toolCallThenFinish,
  textThenStop,
} from "../shared/agent/index";
import { ProviderStore } from "../shared/provider";
import type { HostAdapter } from "../shared/host";
import { MockHostAdapter } from "./mockHost";

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

describe("ChatController permission mode path", () => {
  it("confirm_all auto-executes write without approval_needed", async () => {
    const host = new MockHostAdapter();
    const write = vi.spyOn(host, "writeRange");
    const events: ChatTraceEvent[] = [];
    const provider = new ScriptedStreamProvider({
      rounds: [
        toolCallThenFinish(
          "call_w",
          "range.write",
          '{"sheetName":"Sheet1","range":"A1","values":[["ok"]]}',
        ),
        textThenStop("done"),
      ],
    });
    const controller = new ChatController({
      store: storeOk(),
      host: host as unknown as HostAdapter,
      getPermissionMode: () => "confirm_all",
      createProvider: () => ({ ok: true, provider }),
      onEvent: (e) => events.push(e),
    });
    const result = await controller.send("write");
    expect(result.turnStatus).toBe("completed");
    expect(write).toHaveBeenCalledTimes(1);
    expect(events.some((e) => e.type === "approval_needed")).toBe(false);
    expect(controller.getState().pendingApproval).toBeNull();
  });

  it("normal mode requires approval for safe read", async () => {
    const host = new MockHostAdapter();
    const read = vi.spyOn(host, "readRange");
    let controller!: ChatController;
    const provider = new ScriptedStreamProvider({
      rounds: [
        toolCallThenFinish(
          "call_r",
          "range.read",
          '{"sheetName":"Sheet1","range":"A1"}',
        ),
        textThenStop("done"),
      ],
    });
    controller = new ChatController({
      store: storeOk(),
      host: host as unknown as HostAdapter,
      getPermissionMode: () => "normal",
      createProvider: () => ({ ok: true, provider }),
      onEvent: (e) => {
        if (e.type === "approval_needed") {
          queueMicrotask(() => {
            expect(controller.getState().status).toBe("awaiting_approval");
            expect(e.request.name).toBe("range.read");
            expect(controller.approve(e.request.requestId)).toBe(true);
          });
        }
      },
    });
    const result = await controller.send("read");
    expect(result.turnStatus).toBe("completed");
    expect(read).toHaveBeenCalled();
  });

  it("default auto_approve_safe keeps moderate approval + reject path", async () => {
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
      host: host as unknown as HostAdapter,
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
  });
});
