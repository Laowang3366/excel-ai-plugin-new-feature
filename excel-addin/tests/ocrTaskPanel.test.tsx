/** @vitest-environment jsdom */
import { act, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ChatSessionProvider } from "../src/chat/ChatSessionContext";
import { OcrTaskPanel } from "../src/components/tasks/OcrTaskPanel";
import {
  ChatController,
  type ChatControllerDeps,
  type ChatSendResult,
  type ChatTraceEvent,
  type ChatToolExecuteResult,
} from "../shared/agentChat";
import { ProviderStore } from "../shared/provider";
import { MockHostAdapter } from "./mockHost";
import {
  OCR_RESULT_MARKER_CLOSE,
  OCR_RESULT_MARKER_OPEN,
} from "../shared/tasks";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT =
  true;

function makeImageFile(name: string, bytes: number[] = [1, 2, 3]): File {
  const file = new File([new Uint8Array(bytes)], name, { type: "image/png" });
  if (typeof (file as File).arrayBuffer !== "function") {
    Object.defineProperty(file, "arrayBuffer", {
      value: async () => new Uint8Array(bytes).buffer,
    });
  }
  return file;
}

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

class OcrFake {
  status: "idle" | "running" | "awaiting_approval" | "stopping" = "idle";
  onEvent?: (e: ChatTraceEvent) => void;
  executeCalls: Array<{ name: string; args: Record<string, unknown> }> = [];
  lastUserMessage = "";

  constructor(deps: ChatControllerDeps) {
    this.onEvent = deps.onEvent;
  }

  getState() {
    return { status: this.status, messages: [] as never };
  }

  stop() {}
  clear() {
    return { ok: true as const };
  }
  approve() {
    return false;
  }
  reject() {
    return false;
  }

  async send(userMessage: string): Promise<ChatSendResult> {
    this.lastUserMessage = userMessage;
    this.status = "running";
    const assistant = [
      "识别完成",
      OCR_RESULT_MARKER_OPEN,
      JSON.stringify({
        kind: "invoice",
        text: "发票摘要",
        fields: {},
        rows: [],
        invoices: [
          {
            filename: "inv.png",
            text: "",
            fields: { 发票号码: "NO1", 金额: "9.9" },
            rows: [],
          },
        ],
        errors: [],
      }),
      OCR_RESULT_MARKER_CLOSE,
    ].join("\n");
    this.status = "idle";
    return {
      turnStatus: "completed",
      run: {
        status: "completed",
        assistantText: assistant,
        messages: [
          { role: "user", content: userMessage },
          { role: "assistant", content: assistant },
        ],
        rounds: 1,
        usage: { inputTokens: 1, outputTokens: 1 },
      },
    };
  }

  async executeTool(
    toolName: string,
    args: Record<string, unknown>,
  ): Promise<ChatToolExecuteResult> {
    this.executeCalls.push({ name: toolName, args });
    // Simulate approval-gated path: never write without going through controller API
    return {
      ok: true,
      tool: toolName,
      requiredApproval: true,
      result: { ok: true, tool: toolName as never, data: {} },
    };
  }
}

describe("OcrTaskPanel preview + write via executeTool", () => {
  let root: Root;
  let container: HTMLDivElement;

  afterEach(() => {
    if (root && container) unmount(root, container);
  });

  it("parses structured result, selects fields, writes via executeTool not host", async () => {
    const store = new ProviderStore();
    store.add({
      name: "o",
      provider: "openai",
      apiKey: "sk-should-not-appear-in-ui",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "openai",
    });
    const host = new MockHostAdapter();
    const writeSpy = vi.spyOn(host, "writeRange");
    let fake: OcrFake | undefined;

    const m = mount(
      <ChatSessionProvider
        store={store}
        adapter={host}
        createController={(deps) => {
          fake = new OcrFake(deps);
          return fake as unknown as ChatController;
        }}
      >
        <OcrTaskPanel />
      </ChatSessionProvider>,
    );
    root = m.root;
    container = m.container;

    // Upload fake image
    const fileInput = container.querySelector(
      'input[type="file"]',
    ) as HTMLInputElement;
    const file = makeImageFile("inv.png");
    await act(async () => {
      Object.defineProperty(fileInput, "files", {
        value: [file],
        configurable: true,
      });
      fileInput.dispatchEvent(new Event("change", { bubbles: true }));
      await Promise.resolve();
    });

    const rangeInput = container.querySelector(
      'input[aria-label="OCR 写入锚点"]',
    ) as HTMLInputElement;
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        "value",
      )?.set;
      setter?.call(rangeInput, "Sheet1!A1");
      rangeInput.dispatchEvent(new Event("input", { bubbles: true }));
    });

    const recognize = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent?.includes("识别"),
    ) as HTMLButtonElement;
    await act(async () => {
      recognize.click();
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();
    });
    // flush microtasks from base64 + send
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0));
    });

    expect(container.textContent).toMatch(/识别结果预览|发票号码|发票摘要/);
    expect(container.textContent).not.toContain("sk-should-not-appear-in-ui");
    expect(container.textContent).not.toMatch(/data:image\/png;base64/);

    const writeFields = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "写入所选字段",
    ) as HTMLButtonElement;
    expect(writeFields).toBeTruthy();
    await act(async () => {
      writeFields.click();
      await Promise.resolve();
      await Promise.resolve();
    });

    expect(fake!.executeCalls.length).toBe(1);
    expect(fake!.executeCalls[0]?.name).toBe("range.write");
    expect(fake!.executeCalls[0]?.args.sheetName).toBe("Sheet1");
    expect(fake!.executeCalls[0]?.args.range).toBe("A1");
    // Host not called directly from panel
    expect(writeSpy).not.toHaveBeenCalled();
  });

  it("empty write address errors without executeTool", async () => {
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
    let fake: OcrFake | undefined;
    const m = mount(
      <ChatSessionProvider
        store={store}
        adapter={host}
        createController={(deps) => {
          fake = new OcrFake(deps);
          return fake as unknown as ChatController;
        }}
      >
        <OcrTaskPanel />
      </ChatSessionProvider>,
    );
    root = m.root;
    container = m.container;

    // Seed preview by recognizing first
    const fileInput = container.querySelector(
      'input[type="file"]',
    ) as HTMLInputElement;
    const file = makeImageFile("a.png", [9]);
    await act(async () => {
      Object.defineProperty(fileInput, "files", {
        value: [file],
        configurable: true,
      });
      fileInput.dispatchEvent(new Event("change", { bubbles: true }));
      await Promise.resolve();
    });
    const recognize = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent?.includes("识别"),
    ) as HTMLButtonElement;
    await act(async () => {
      recognize.click();
      await Promise.resolve();
      await Promise.resolve();
    });
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0));
    });

    const writeText = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "写入整段文本",
    ) as HTMLButtonElement;
    await act(async () => {
      writeText.click();
      await Promise.resolve();
    });
    expect(container.textContent).toMatch(/写入锚点|读取选区/);
    expect(fake!.executeCalls.length).toBe(0);
  });
});
