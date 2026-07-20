/** @vitest-environment jsdom */
import { act, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { ProviderSettingsPanel } from "../src/components/ProviderSettingsPanel";
import { ProviderStore } from "../shared/provider";

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

describe("ProviderSettingsPanel", () => {
  let root: Root;
  let container: HTMLDivElement;

  afterEach(() => {
    if (root && container) unmount(root, container);
  });

  it("documents memory-only key and re-enter after refresh", () => {
    const store = new ProviderStore();
    const m = mount(<ProviderSettingsPanel store={store} />);
    root = m.root;
    container = m.container;
    expect(container.textContent).toMatch(/仅保存在当前页面内存|仅保存在/);
    expect(container.textContent).toMatch(/刷新页面或重开任务窗格后需重新输入 API Key/);
  });

  it("clear API key for direct; rejects invalid context window", async () => {
    const store = new ProviderStore();
    const created = store.add({
      name: "P",
      provider: "openai",
      apiKey: "sk-secret",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "openai",
      contextWindowSize: 128_000,
    });
    expect(store.list()[0]?.hasApiKey).toBe(true);

    const m = mount(<ProviderSettingsPanel store={store} />);
    root = m.root;
    container = m.container;
    expect(container.textContent).toMatch(/key 已设/i);

    const clearBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "清除 API Key",
    ) as HTMLButtonElement;
    expect(clearBtn).toBeTruthy();
    await act(async () => {
      clearBtn.click();
    });
    expect(store.getWithSecret(created.id)?.apiKey).toBe("");
    expect(store.list().find((p) => p.id === created.id)?.hasApiKey).toBe(false);

    // edit invalid ctx
    const editBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "编辑",
    ) as HTMLButtonElement;
    await act(async () => {
      editBtn.click();
    });
    const ctxInput = Array.from(container.querySelectorAll("input")).find(
      (el) => el.getAttribute("aria-label") === "Context Window",
    ) as HTMLInputElement;
    expect(ctxInput).toBeTruthy();
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        "value",
      )?.set;
      setter?.call(ctxInput, "500");
      ctxInput.dispatchEvent(new Event("input", { bubbles: true }));
      ctxInput.dispatchEvent(new Event("change", { bubbles: true }));
    });
    const saveBtn = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === "保存",
    ) as HTMLButtonElement;
    await act(async () => {
      saveBtn.click();
    });
    expect(container.textContent).toMatch(/Context Window/);
    expect(store.list().find((p) => p.id === created.id)?.contextWindowSize).toBe(
      128_000,
    );
  });
});
