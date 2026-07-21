/** @vitest-environment jsdom */
import { act, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { ChatPanel } from "../src/components/ChatPanel";
import { ChatSessionProvider } from "../src/chat/ChatSessionContext";
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

function unmount(root: Root, container: HTMLDivElement) {
  act(() => {
    root.unmount();
  });
  container.remove();
}

function setTextArea(el: HTMLTextAreaElement, value: string) {
  Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")?.set?.call(
    el,
    value,
  );
  el.dispatchEvent(new Event("input", { bubbles: true }));
}

class HangFake {
  status: "idle" | "running" | "stopping" = "idle";
  onEvent?: (e: ChatTraceEvent) => void;
  stopCalls = 0;
  sendCalls = 0;
  private resolveSend?: (r: ChatSendResult) => void;

  constructor(deps: ChatControllerDeps) {
    this.onEvent = deps.onEvent;
  }

  getState() {
    return { status: this.status, messages: [] as never };
  }

  stop() {
    this.stopCalls += 1;
    this.status = "stopping";
  }

  clear() {
    return this.status === "idle"
      ? ({ ok: true } as const)
      : ({ ok: false, error: "busy" } as const);
  }

  emit(event: ChatTraceEvent) {
    this.onEvent?.(event);
  }

  finish(result: ChatSendResult) {
    this.status = "idle";
    this.resolveSend?.(result);
    this.resolveSend = undefined;
  }

  async send(_userMessage: string): Promise<ChatSendResult> {
    this.sendCalls += 1;
    this.status = "running";
    this.onEvent?.({ type: "text_delta", delta: "old-", round: 1 });
    return new Promise<ChatSendResult>((resolve) => {
      this.resolveSend = resolve;
    });
  }
}

describe("useChatController lifecycle isolation", () => {
  let root: Root;
  let container: HTMLDivElement;

  afterEach(() => {
    if (root && container) unmount(root, container);
  });

  it("adapter switch stops old controller and ignores late events", async () => {
    const store = new ProviderStore();
    store.add({
      name: "o",
      provider: "openai",
      apiKey: "sk-x",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "openai",
    });
    const hostA = new MockHostAdapter();
    const hostB = new MockHostAdapter();
    const fakes: HangFake[] = [];

    function Panel({ adapter }: { adapter: MockHostAdapter }) {
      return (
        <ChatSessionProvider store={store} adapter={adapter} createController={(deps) => {
            const f = new HangFake(deps);
            fakes.push(f);
            return f as unknown as ChatController;
          }}><ChatPanel /></ChatSessionProvider>
      );
    }

    const m = mount(<Panel adapter={hostA} />);
    root = m.root;
    container = m.container;

    const textarea = container.querySelector("textarea") as HTMLTextAreaElement;
    const sendBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "发送",
    ) as HTMLButtonElement;

    await act(async () => {
      setTextArea(textarea, "from-a");
    });
    await act(async () => {
      sendBtn.click();
      await Promise.resolve();
    });
    expect(fakes[0]?.sendCalls).toBe(1);
    expect(container.textContent).toContain("old-");

    // Switch adapter → cleanup old, create new
    await act(async () => {
      root.render(<Panel adapter={hostB} />);
    });
    expect(fakes[0]?.stopCalls).toBe(1);
    expect(fakes.length).toBeGreaterThanOrEqual(2);

    // Late emissions from disposed controller must not pollute
    await act(async () => {
      fakes[0]!.emit({ type: "text_delta", delta: "POISON", round: 1 });
      fakes[0]!.emit({
        type: "tool_call_parsed",
        round: 1,
        call: { id: "x", name: "range.read", argumentsJson: "{}" },
      });
      fakes[0]!.emit({ type: "turn_end", turnStatus: "completed" });
      fakes[0]!.finish({
        turnStatus: "completed",
        run: {
          status: "completed",
          assistantText: "should-not-appear",
          messages: [],
          rounds: 1,
          usage: { inputTokens: 0, outputTokens: 0 },
        },
      });
      await Promise.resolve();
    });
    expect(container.textContent).not.toContain("POISON");
    expect(container.textContent).not.toContain("should-not-appear");

    // New instance can still send once UI is idle again
    await act(async () => {
      await Promise.resolve();
    });
    const ta2 = container.querySelector("textarea") as HTMLTextAreaElement;
    const send2 = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "发送",
    ) as HTMLButtonElement;
    await act(async () => {
      setTextArea(ta2, "from-b");
    });
    // draft non-empty + idle controller => can send
    expect(send2.disabled).toBe(false);
    const before = fakes[fakes.length - 1]?.sendCalls ?? 0;
    await act(async () => {
      send2.click();
      await Promise.resolve();
      await Promise.resolve();
    });
    expect(fakes[fakes.length - 1]?.sendCalls).toBe(before + 1);
  });

  it("unmount stops controller; late events do not throw", async () => {
    const store = new ProviderStore();
    store.add({
      name: "o",
      provider: "openai",
      apiKey: "sk-x",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "openai",
    });
    let fake: HangFake | undefined;
    const m = mount(
      <ChatSessionProvider store={store} adapter={new MockHostAdapter()} createController={(deps) => {
          fake = new HangFake(deps);
          return fake as unknown as ChatController;
        }}><ChatPanel /></ChatSessionProvider>,
    );
    root = m.root;
    container = m.container;
    await act(async () => {
      setTextArea(container.querySelector("textarea") as HTMLTextAreaElement, "x");
      (
        Array.from(container.querySelectorAll("button")).find(
          (b) => b.textContent === "发送",
        ) as HTMLButtonElement
      ).click();
      await Promise.resolve();
    });
    unmount(root, container);
    // prevent afterEach double unmount
    root = undefined as unknown as Root;
    container = undefined as unknown as HTMLDivElement;
    expect(fake?.stopCalls).toBe(1);
    // late emit should be ignored
    expect(() => {
      fake?.emit({ type: "text_delta", delta: "after-unmount", round: 1 });
      fake?.finish({
        turnStatus: "aborted",
        run: {
          status: "aborted",
          assistantText: "x",
          messages: [],
          rounds: 0,
          usage: { inputTokens: 0, outputTokens: 0 },
        },
      });
    }).not.toThrow();
  });

  it("preflight_failed does not keep user bubble; banner shown", async () => {
    const store = new ProviderStore();
    // no active provider
    const m = mount(
      <ChatSessionProvider store={store} adapter={new MockHostAdapter()}><ChatPanel /></ChatSessionProvider>,
    );
    root = m.root;
    container = m.container;
    await act(async () => {
      setTextArea(container.querySelector("textarea") as HTMLTextAreaElement, "hello-preflight");
      (
        Array.from(container.querySelectorAll("button")).find(
          (b) => b.textContent === "发送",
        ) as HTMLButtonElement
      ).click();
      await Promise.resolve();
      await Promise.resolve();
    });
    expect(container.textContent).toMatch(/模型供应商/);
    expect(container.querySelectorAll(".chat-bubble.user").length).toBe(0);
    const ta = container.querySelector("textarea") as HTMLTextAreaElement;
    expect(ta.value).toBe("hello-preflight");
  });
});
