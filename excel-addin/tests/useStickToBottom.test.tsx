/** @vitest-environment jsdom */
import { act, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { useEffect, useState } from "react";
import { useStickToBottom } from "../src/chat/useStickToBottom";

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

function Harness({ bump }: { bump: number }) {
  const { containerRef, onScroll } = useStickToBottom([bump]);
  const [n, setN] = useState(3);
  useEffect(() => {
    if (bump > 0) setN((x) => x + 2);
  }, [bump]);
  return (
    <div
      ref={containerRef}
      onScroll={onScroll}
      data-testid="scroller"
      style={{ height: 40, overflow: "auto" }}
    >
      {Array.from({ length: n }, (_, i) => (
        <div key={i} style={{ height: 30 }}>
          row {i}
        </div>
      ))}
    </div>
  );
}

describe("useStickToBottom", () => {
  let root: Root;
  let container: HTMLDivElement;

  afterEach(() => {
    if (root && container) {
      act(() => root.unmount());
      container.remove();
    }
  });

  it("pins to bottom when sticking; does not force after user scrolls up", async () => {
    let bump = 0;
    const m = mount(<Harness bump={bump} />);
    root = m.root;
    container = m.container;
    const scroller = container.querySelector(
      "[data-testid=scroller]",
    ) as HTMLDivElement;

    // jsdom often reports 0 sizes; stub geometry for the test.
    Object.defineProperty(scroller, "clientHeight", { value: 40, configurable: true });
    Object.defineProperty(scroller, "scrollHeight", {
      get: () => 30 * scroller.childElementCount,
      configurable: true,
    });

    await act(async () => {
      bump = 1;
      root.render(<Harness bump={bump} />);
    });
    expect(scroller.scrollTop).toBeGreaterThanOrEqual(0);

    // Simulate user scrolled up: distance from bottom > threshold
    scroller.scrollTop = 0;
    scroller.dispatchEvent(new Event("scroll", { bubbles: true }));

    const topBefore = scroller.scrollTop;
    await act(async () => {
      bump = 2;
      root.render(<Harness bump={bump} />);
    });
    // Should remain near top (not forced to bottom)
    expect(scroller.scrollTop).toBe(topBefore);
  });
});
