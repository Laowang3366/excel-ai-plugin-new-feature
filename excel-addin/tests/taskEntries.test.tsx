/** @vitest-environment jsdom */
import { act, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { describe, expect, it } from "vitest";
import { App } from "../src/App";

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

describe("task entry tabs and deep links", () => {
  it("renders five task tabs plus core tabs", () => {
    const { root, container } = mount(<App />);
    const labels = Array.from(container.querySelectorAll("nav.tabs button")).map(
      (b) => b.textContent,
    );
    for (const need of [
      "聊天",
      "公式助手",
      "数据清洗",
      "OCR识别",
      "图表制作",
      "报告生成",
      "宿主",
      "工具",
      "模型供应商",
    ]) {
      expect(labels).toContain(need);
    }
    unmount(root, container);
  });

  it("deep-link page=formula opens formula panel", async () => {
    const prev = window.location.search;
    window.history.replaceState({}, "", "?page=formula");
    const { root, container } = mount(<App />);
    await act(async () => {
      await Promise.resolve();
    });
    expect(container.textContent).toMatch(/公式助手|生成 Excel/);
    window.history.replaceState({}, "", prev || "?");
    unmount(root, container);
  });
});
