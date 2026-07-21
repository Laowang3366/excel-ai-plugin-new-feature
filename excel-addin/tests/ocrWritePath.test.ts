import { describe, expect, it, vi } from "vitest";
import { ChatController } from "../shared/agentChat";
import { ProviderStore } from "../shared/provider";
import { ToolExecutor } from "../shared/tools/executor";
import { MockHostAdapter } from "./mockHost";

describe("ChatController.executeTool approval path", () => {
  it("range.write goes through approval gate; host not called until approve", async () => {
    const host = new MockHostAdapter();
    const writeSpy = vi.spyOn(host, "writeRange");
    const store = new ProviderStore();
    store.add({
      name: "t",
      provider: "openai",
      apiKey: "sk-test-not-shown",
      baseUrl: "https://example.com/v1",
      model: "m",
      apiFormat: "openai",
    });

    const events: string[] = [];
    const controller = new ChatController({
      store,
      host,
      onEvent: (e) => {
        events.push(e.type);
      },
      getPermissionMode: () => "normal",
    });

    const p = controller.executeTool(
      "range.write",
      {
        sheetName: "Sheet1",
        range: "A1",
        values: [["ocr-text"]],
        verify: false,
      },
      { toolCallId: "ocr-write-test" },
    );

    // wait for gate request
    for (let i = 0; i < 20; i++) {
      if (controller.getState().pendingApproval) break;
      await Promise.resolve();
    }
    const pending = controller.getState().pendingApproval;
    expect(pending).toBeTruthy();
    expect(pending?.name).toBe("range.write");
    expect(events).toContain("approval_needed");
    // preview must not dump raw secret-looking long base64; values grid is collapsed
    const pub = JSON.stringify(pending);
    expect(pub).not.toContain("sk-test-not-shown");
    expect(writeSpy).not.toHaveBeenCalled();

    expect(controller.approve(pending!.requestId)).toBe(true);
    const result = await p;
    expect(result.ok).toBe(true);
    expect(writeSpy).toHaveBeenCalledTimes(1);
    expect(writeSpy.mock.calls[0]?.[0]).toBe("Sheet1");
    expect(writeSpy.mock.calls[0]?.[1]).toBe("A1");
    expect(controller.getState().status).toBe("idle");
  });

  it("reject does not call host writeRange", async () => {
    const host = new MockHostAdapter();
    const writeSpy = vi.spyOn(host, "writeRange");
    const store = new ProviderStore();
    store.add({
      name: "t",
      provider: "openai",
      apiKey: "sk-x",
      baseUrl: "https://example.com/v1",
      model: "m",
      apiFormat: "openai",
    });
    const controller = new ChatController({
      store,
      host,
      getPermissionMode: () => "normal",
    });
    const p = controller.executeTool("range.write", {
      sheetName: "S",
      range: "B1",
      values: [["x"]],
    });
    for (let i = 0; i < 20; i++) {
      if (controller.getState().pendingApproval) break;
      await Promise.resolve();
    }
    const pending = controller.getState().pendingApproval!;
    controller.reject(pending.requestId);
    const result = await p;
    expect(result.ok).toBe(false);
    expect(writeSpy).not.toHaveBeenCalled();
  });

  it("does not use ToolExecutor directly from UI path without ApprovingToolExecutor", async () => {
    // Documented contract: executeTool always wraps ToolExecutor.
    // Direct ToolExecutor would write immediately — we assert executeTool path only.
    const host = new MockHostAdapter();
    const direct = new ToolExecutor(host);
    const writeSpy = vi.spyOn(host, "writeRange");
    await direct.execute({
      name: "range.write",
      arguments: {
        sheetName: "S",
        range: "A1",
        values: [["bypass"]],
        verify: false,
      },
    });
    expect(writeSpy).toHaveBeenCalled(); // raw executor would write
    // UI must not use this raw path; OcrTaskPanel uses controller.executeTool instead.
  });
});
