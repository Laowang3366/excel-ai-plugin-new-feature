/** @vitest-environment jsdom */
import { act, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { App } from "../src/App";
import { ChatPanel } from "../src/components/ChatPanel";
import {
  ChatController,
  type ChatControllerDeps,
  type ChatSendResult,
  type ChatTraceEvent,
} from "../shared/agentChat";
import { ProviderStore } from "../shared/provider";
import { MockHostAdapter } from "./mockHost";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT =
  true;

function mount(ui: ReactElement): { root: Root; container: HTMLDivElement } {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  act(() => {
    root.render(ui);
  });
  return { root, container };
}


function setTextArea(el: HTMLTextAreaElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value",
  )?.set;
  setter?.call(el, value);
  el.dispatchEvent(new Event("input", { bubbles: true }));
  el.dispatchEvent(new Event("change", { bubbles: true }));
}

function unmount(root: Root, container: HTMLDivElement) {
  act(() => {
    root.unmount();
  });
  container.remove();
}

class FakeController {
  status: "idle" | "running" | "stopping" = "idle";
  onEvent?: (e: ChatTraceEvent) => void;
  stopCalls = 0;
  sendCalls = 0;
  clearCalls = 0;
  messages: Array<{ role: string; content: string }> = [];
  private pendingResolve?: (r: ChatSendResult) => void;
  private mode: "auto" | "hang";

  constructor(deps: ChatControllerDeps, mode: "auto" | "hang" = "auto") {
    this.onEvent = deps.onEvent;
    this.mode = mode;
  }

  getState() {
    return { status: this.status, messages: this.messages as never };
  }

  stop() {
    this.stopCalls += 1;
    this.status = "stopping";
  }

  clear() {
    this.clearCalls += 1;
    if (this.status !== "idle") return { ok: false as const, error: "busy" };
    this.messages = [];
    return { ok: true as const };
  }

  complete(result: ChatSendResult) {
    this.status = "idle";
    this.messages = (result.run?.messages ?? []) as never;
    this.onEvent?.({
      type: "run_end",
      status: result.run?.status ?? "completed",
      rounds: result.run?.rounds ?? 0,
    });
    this.onEvent?.({ type: "turn_end", turnStatus: result.turnStatus });
    this.pendingResolve?.(result);
    this.pendingResolve = undefined;
  }

  async send(userMessage: string): Promise<ChatSendResult> {
    this.sendCalls += 1;
    this.status = "running";
    this.onEvent?.({ type: "round_start", round: 1 });
    this.onEvent?.({ type: "text_delta", delta: "Hel", round: 1 });
    this.onEvent?.({ type: "text_delta", delta: "lo", round: 1 });
    this.onEvent?.({
      type: "tool_call_parsed",
      round: 1,
      call: {
        id: "t1",
        name: "host.status",
        argumentsJson: '{"note":"' + "A".repeat(300) + '"}',
      },
    });
    this.onEvent?.({
      type: "tool_outcome",
      toolCallId: "t1",
      round: 1,
      outcome: {
        kind: "host",
        toolName: "host.status",
        result: {
          ok: true,
          tool: "host.status",
          data: { imageBase64: "B".repeat(400) },
        },
      },
    });
    this.onEvent?.({
      type: "round_end",
      round: 1,
      finishReason: "stop",
      toolCallCount: 1,
    });

    if (this.mode === "hang") {
      return new Promise<ChatSendResult>((resolve) => {
        this.pendingResolve = resolve;
      });
    }

    const result: ChatSendResult = {
      turnStatus: "completed",
      run: {
        status: "completed",
        assistantText: "Hello",
        messages: [
          { role: "user", content: userMessage },
          { role: "assistant", content: "Hello" },
        ],
        rounds: 1,
        usage: { inputTokens: 0, outputTokens: 0 },
      },
    };
    this.status = "idle";
    this.messages = result.run!.messages as never;
    this.onEvent?.({ type: "run_end", status: "completed", rounds: 1 });
    this.onEvent?.({ type: "turn_end", turnStatus: "completed" });
    return result;
  }
}

describe("ChatPanel UI", () => {
  let root: Root;
  let container: HTMLDivElement;

  afterEach(() => {
    if (root && container) unmount(root, container);
  });

  it("renders readonly banner, streams text, tool trace truncated, final bubble", async () => {
    const store = new ProviderStore();
    store.add({
      name: "o",
      provider: "openai",
      apiKey: "sk-test-secret-key",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "openai",
    });
    const host = new MockHostAdapter();
    let fake: FakeController | undefined;

    const m = mount(
      <ChatPanel
        store={store}
        adapter={host}
        createController={(deps) => {
          fake = new FakeController(deps, "auto");
          return fake as unknown as ChatController;
        }}
      />,
    );
    root = m.root;
    container = m.container;

    expect(container.textContent).toContain("变更操作会在执行前等待你的批准");
    const textarea = container.querySelector("textarea") as HTMLTextAreaElement;
    const sendBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "发送",
    ) as HTMLButtonElement;

    await act(async () => {
      setTextArea(textarea, "你好");
    });
    await act(async () => {
      sendBtn.click();
      await Promise.resolve();
      await Promise.resolve();
    });

    expect(container.textContent).toContain("Hello");
    expect(container.textContent).toContain("host.status");
    // long args / base64 collapsed — full AAA/BBB runs should not dump
    expect(container.textContent ?? "").not.toContain("A".repeat(200));
    expect(container.textContent ?? "").not.toContain("B".repeat(200));
    expect(container.innerHTML).not.toContain("sk-test-secret-key");
    expect(fake!.sendCalls).toBe(1);
  });

  it("stop calls controller.stop; busy prevents double send", async () => {
    const store = new ProviderStore();
    store.add({
      name: "o",
      provider: "openai",
      apiKey: "sk-x",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "openai",
    });
    const host = new MockHostAdapter();
    let fake: FakeController | undefined;

    const m = mount(
      <ChatPanel
        store={store}
        adapter={host}
        createController={(deps) => {
          fake = new FakeController(deps, "hang");
          return fake as unknown as ChatController;
        }}
      />,
    );
    root = m.root;
    container = m.container;

    const textarea = container.querySelector("textarea") as HTMLTextAreaElement;
    const sendBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "发送",
    ) as HTMLButtonElement;

    await act(async () => {
      setTextArea(textarea, "q");
    });
    await act(async () => {
      sendBtn.click();
      await Promise.resolve();
    });
    // second click while busy
    await act(async () => {
      sendBtn.click();
      await Promise.resolve();
    });
    expect(fake!.sendCalls).toBe(1);

    const stopBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "停止",
    ) as HTMLButtonElement;
    expect(stopBtn).toBeTruthy();
    await act(async () => {
      stopBtn.click();
    });
    expect(fake!.stopCalls).toBe(1);
    expect(container.textContent).toMatch(/停止|进行中的表格操作/);

    const clearBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "清空",
    ) as HTMLButtonElement;
    expect(clearBtn.disabled).toBe(true);

    await act(async () => {
      fake?.complete({
        turnStatus: "aborted",
        run: {
          status: "aborted",
          assistantText: "Hel",
          messages: [{ role: "user", content: "q" }],
          rounds: 0,
          usage: { inputTokens: 0, outputTokens: 0 },
        },
      });
      await Promise.resolve();
      await Promise.resolve();
    });
  });

  it("preflight missing key surfaces Chinese config hint", async () => {
    const store = new ProviderStore();
    // no provider / no key
    const host = new MockHostAdapter();
    // Real controller path (no fake)
    const m = mount(<ChatPanel store={store} adapter={host} />);
    root = m.root;
    container = m.container;
    const textarea = container.querySelector("textarea") as HTMLTextAreaElement;
    const sendBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "发送",
    ) as HTMLButtonElement;
    await act(async () => {
      setTextArea(textarea, "hello");
    });
    await act(async () => {
      sendBtn.click();
      await Promise.resolve();
      await Promise.resolve();
    });
    expect(container.textContent).toMatch(/模型供应商/);
  });
});

describe("App chat tab", () => {
  it("defaults to chat; tools and providers tabs remain", async () => {
    const m = mount(<App />);
    const { root, container } = m;
    const labels = Array.from(container.querySelectorAll("nav.tabs button")).map(
      (b) => b.textContent,
    );
    expect(labels).toEqual(["聊天", "宿主", "工具", "模型供应商"]);
    expect(
      container.textContent?.includes("等待你的批准") ||
        container.textContent?.includes("检测宿主"),
    ).toBe(true);
    const toolsBtn = Array.from(container.querySelectorAll("nav.tabs button")).find(
      (b) => b.textContent === "工具",
    ) as HTMLButtonElement;
    await act(async () => {
      toolsBtn.click();
    });
    // adapter may still be detecting
    expect(container.textContent).toMatch(/工具|检测宿主|Excel/);
    unmount(root, container);
  });
});
