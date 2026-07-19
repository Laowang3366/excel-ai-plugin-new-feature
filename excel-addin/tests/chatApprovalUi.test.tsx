/** @vitest-environment jsdom */
import { act, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { ChatPanel } from "../src/components/ChatPanel";
import {
  ChatController,
  type ApprovalRequest,
  type ChatControllerDeps,
  type ChatSendResult,
  type ChatTraceEvent,
} from "../shared/agentChat";
import { ProviderStore } from "../shared/provider";
import { MockHostAdapter } from "./mockHost";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT =
  true;

function mount(ui: ReactElement) {
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

function sampleRequest(over: Partial<ApprovalRequest> = {}): ApprovalRequest {
  return {
    requestId: "apr-1",
    toolCallId: "call_1",
    round: 1,
    name: "range.write",
    riskLevel: "moderate",
    destructive: true,
    argsPreview: {
      sheetName: "S",
      password: "[REDACTED]",
      imageBase64: "[omitted binary 500 chars]",
      values: "[grid 1x1]",
    },
    impactHint: "将修改工作簿内容，需你确认后才会执行。",
    createdAt: Date.now(),
    ...over,
  };
}

class ApprovalFake {
  status: "idle" | "running" | "awaiting_approval" | "stopping" = "idle";
  onEvent?: (e: ChatTraceEvent) => void;
  stopCalls = 0;
  sendCalls = 0;
  approveCalls: string[] = [];
  rejectCalls: string[] = [];
  private resolveSend?: (r: ChatSendResult) => void;
  private lastRequest: ApprovalRequest | null = null;

  constructor(deps: ChatControllerDeps) {
    this.onEvent = deps.onEvent;
  }

  getState() {
    return {
      status: this.status,
      messages: [] as never,
      pendingApproval: this.lastRequest,
    };
  }

  stop() {
    this.stopCalls += 1;
    this.status = "stopping";
    if (this.lastRequest) {
      const req = this.lastRequest;
      this.lastRequest = null;
      this.onEvent?.({
        type: "approval_resolved",
        requestId: req.requestId,
        decision: "cancelled",
        request: req,
      });
    }
  }

  clear() {
    return this.status === "idle"
      ? ({ ok: true } as const)
      : ({ ok: false, error: "busy" } as const);
  }

  approve(id?: string) {
    if (!this.lastRequest) return false;
    if (id && id !== this.lastRequest.requestId) return false;
    this.approveCalls.push(id ?? this.lastRequest.requestId);
    const req = this.lastRequest;
    this.lastRequest = null;
    this.status = "running";
    this.onEvent?.({
      type: "approval_resolved",
      requestId: req.requestId,
      decision: "approved",
      request: req,
    });
    this.onEvent?.({ type: "text_delta", delta: "Wrote ok", round: 1 });
    this.finish({
      turnStatus: "completed",
      run: {
        status: "completed",
        assistantText: "Wrote ok",
        messages: [
          { role: "user", content: "write" },
          { role: "assistant", content: "Wrote ok" },
        ],
        rounds: 1,
        usage: { inputTokens: 0, outputTokens: 0 },
      },
    });
    return true;
  }

  reject(id?: string) {
    if (!this.lastRequest) return false;
    if (id && id !== this.lastRequest.requestId) return false;
    this.rejectCalls.push(id ?? this.lastRequest.requestId);
    const req = this.lastRequest;
    this.lastRequest = null;
    this.status = "running";
    this.onEvent?.({
      type: "approval_resolved",
      requestId: req.requestId,
      decision: "rejected",
      request: req,
    });
    this.onEvent?.({ type: "text_delta", delta: "Skipped write", round: 1 });
    this.finish({
      turnStatus: "completed",
      run: {
        status: "completed",
        assistantText: "Skipped write",
        messages: [
          { role: "user", content: "write" },
          { role: "assistant", content: "Skipped write" },
        ],
        rounds: 1,
        usage: { inputTokens: 0, outputTokens: 0 },
      },
    });
    return true;
  }

  finish(result: ChatSendResult) {
    this.status = "idle";
    this.resolveSend?.(result);
    this.resolveSend = undefined;
  }

  async send(_msg: string): Promise<ChatSendResult> {
    this.sendCalls += 1;
    this.status = "running";
    const req = sampleRequest();
    this.lastRequest = req;
    this.onEvent?.({ type: "round_start", round: 1 });
    this.onEvent?.({
      type: "tool_call_parsed",
      round: 1,
      call: {
        id: "call_1",
        name: "range.write",
        argumentsJson: '{"password":"secret","values":[["x"]]}',
      },
    });
    this.onEvent?.({ type: "approval_needed", request: req });
    this.status = "awaiting_approval";
    return new Promise((resolve) => {
      this.resolveSend = resolve;
    });
  }
}

describe("Chat approval UI", () => {
  let root: Root;
  let container: HTMLDivElement;
  let fake: ApprovalFake;

  afterEach(() => {
    if (root && container) unmount(root, container);
  });

  function renderPanel() {
    const store = new ProviderStore();
    store.add({
      name: "o",
      provider: "openai",
      apiKey: "sk-ui-secret",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "openai",
    });
    const m = mount(
      <ChatPanel
        store={store}
        adapter={new MockHostAdapter()}
        createController={(deps) => {
          fake = new ApprovalFake(deps);
          return fake as unknown as ChatController;
        }}
      />,
    );
    root = m.root;
    container = m.container;
  }

  async function sendMessage(text = "write cell") {
    const ta = container.querySelector("textarea") as HTMLTextAreaElement;
    const sendBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "发送",
    ) as HTMLButtonElement;
    await act(async () => {
      setTextArea(ta, text);
    });
    await act(async () => {
      sendBtn.click();
      await Promise.resolve();
      await Promise.resolve();
    });
  }

  it("renders approval card with sanitized args; no secrets/base64 in DOM", async () => {
    renderPanel();
    expect(container.textContent).toContain("等待你的批准");
    expect(container.textContent).not.toContain("只读模式");
    await sendMessage();
    const card = container.querySelector('[role="alertdialog"]') as HTMLElement;
    expect(card).toBeTruthy();
    expect(card.getAttribute("aria-modal")).toBe("true");
    expect(card.textContent).toContain("range.write");
    expect(card.textContent).toMatch(/需批准|破坏性/);
    expect(card.textContent).toContain("将修改工作簿");
    expect(container.innerHTML).not.toContain("sk-ui-secret");
    expect(container.innerHTML).not.toContain("secret");
    expect(container.innerHTML).not.toMatch(/AAAA{10,}/);
    // focus on reject by default
    const rejectBtn = Array.from(card.querySelectorAll("button")).find(
      (b) => b.textContent === "拒绝",
    ) as HTMLButtonElement;
    expect(document.activeElement).toBe(rejectBtn);
    // send disabled while pending
    const sendBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "发送",
    ) as HTMLButtonElement;
    expect(sendBtn.disabled).toBe(true);
    const stopBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "停止",
    ) as HTMLButtonElement;
    expect(stopBtn).toBeTruthy();
    expect(stopBtn.disabled).toBe(false);
  });

  it("Escape rejects; Approve passes requestId; assistant bubble after resolve", async () => {
    renderPanel();
    await sendMessage();
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
      await Promise.resolve();
      await Promise.resolve();
    });
    expect(fake.rejectCalls).toEqual(["apr-1"]);
    expect(container.querySelector('[role="alertdialog"]')).toBeNull();
    expect(container.textContent).toContain("Skipped write");
    expect(container.textContent).toMatch(/已拒绝|拒绝/);

    // second turn approve path
    await sendMessage("again");
    const card = container.querySelector('[role="alertdialog"]') as HTMLElement;
    const approveBtn = Array.from(card.querySelectorAll("button")).find(
      (b) => b.textContent === "批准",
    ) as HTMLButtonElement;
    await act(async () => {
      approveBtn.click();
      await Promise.resolve();
      await Promise.resolve();
    });
    expect(fake.approveCalls.at(-1)).toBe("apr-1");
    expect(container.textContent).toContain("Wrote ok");
  });

  it("Stop cancels pending; card disappears; no mis-approve after cancel", async () => {
    renderPanel();
    await sendMessage();
    const stopBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "停止",
    ) as HTMLButtonElement;
    await act(async () => {
      stopBtn.click();
      await Promise.resolve();
    });
    expect(fake.stopCalls).toBe(1);
    expect(container.querySelector('[role="alertdialog"]')).toBeNull();
    // late wrong approve id
    expect(fake.approve("nope")).toBe(false);
  });
});
