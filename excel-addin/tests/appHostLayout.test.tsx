/** @vitest-environment jsdom */
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { act, type ReactElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { App } from "../src/App";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT =
  true;

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

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

type Win = Window &
  typeof globalThis & {
    Application?: { Name?: string; ActiveWorkbook?: unknown };
    Excel?: { run?: (...args: unknown[]) => Promise<unknown> };
    Office?: unknown;
  };

function clearHostGlobals() {
  const w = window as Win;
  delete w.Application;
  delete w.Excel;
  delete w.Office;
}

afterEach(() => {
  clearHostGlobals();
});

describe("App host layout markers (Phase61)", () => {
  it("sets data-host=wps-jsa and app--wps-jsa when Application is present", async () => {
    const w = window as Win;
    w.Application = { Name: "WPS" };
    const { root, container } = mount(<App />);
    await act(async () => {
      await Promise.resolve();
      await Promise.resolve();
    });
    const app = container.querySelector(".app");
    expect(app).not.toBeNull();
    expect(app?.getAttribute("data-host")).toBe("wps-jsa");
    expect(app?.classList.contains("app--wps-jsa")).toBe(true);
    expect(app?.classList.contains("app")).toBe(true);
    unmount(root, container);
  });

  it("sets data-host=office-js without WPS class when Excel.run exists", async () => {
    const w = window as Win;
    // Presence of Excel.run is the host signal; signature is host-global.
    w.Excel = { run: (async () => null) as never };
    const { root, container } = mount(<App />);
    await act(async () => {
      await Promise.resolve();
      await Promise.resolve();
    });
    const app = container.querySelector(".app");
    expect(app?.getAttribute("data-host")).toBe("office-js");
    expect(app?.classList.contains("app--wps-jsa")).toBe(false);
    unmount(root, container);
  });

  it("keeps data-host=unknown without WPS class when no host APIs", async () => {
    clearHostGlobals();
    const { root, container } = mount(<App />);
    await act(async () => {
      await Promise.resolve();
      await Promise.resolve();
    });
    // waitForOfficeReady may settle unknown quickly without Office.onReady
    const app = container.querySelector(".app");
    expect(app?.getAttribute("data-host")).toBe("unknown");
    expect(app?.classList.contains("app--wps-jsa")).toBe(false);
    unmount(root, container);
  });
});

describe("WPS layout CSS contract (Phase61)", () => {
  const css = readFileSync(path.join(rootDir, "src/styles.css"), "utf8");

  it("keeps default .app centered 720 and adds WPS-only left-align 520", () => {
    expect(css).toMatch(/\.app\s*\{[^}]*max-width:\s*720px/s);
    expect(css).toMatch(/\.app\s*\{[^}]*margin:\s*0 auto/s);
    expect(css).toMatch(/\[data-host="wps-jsa"\]/);
    expect(css).toMatch(/\.app--wps-jsa/);
    expect(css).toMatch(/max-width:\s*520px/);
    expect(css).toMatch(/margin-left:\s*0/);
    expect(css).toMatch(/margin-right:\s*0/);
    // must not hide overflow on body as the layout "fix"
    expect(css).not.toMatch(/body\s*\{[^}]*overflow\s*:\s*hidden/s);
  });

  it("allows tabs wrap and sets min-width 0 for WPS flex/form children", () => {
    expect(css).toMatch(/\.tabs\s*\{[^}]*flex-wrap:\s*wrap/s);
    expect(css).toMatch(/min-width:\s*0/);
    expect(css).toMatch(
      /\[data-host="wps-jsa"\][\s\S]*?(input|select|textarea|pre)/,
    );
    expect(css).toMatch(/max-width:\s*100%/);
  });
});
