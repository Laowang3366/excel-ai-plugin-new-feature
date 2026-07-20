import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import { describe, expect, it } from "vitest";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const entrySource = readFileSync(path.join(root, "manifest/wps-jsa/wps-entry.js"), "utf8");

type Pane = {
  ID?: string | number;
  Id?: string | number;
  id?: string | number;
  Visible?: boolean;
  url?: string;
};

function loadEntry(opts: {
  locationHref?: string;
  locationThrows?: boolean;
  api?: Record<string, unknown> | null;
  apiOn?: "wps" | "window.wps" | "window.Wps" | "none";
  openImpl?: ((url: string) => unknown) | null;
  alertImpl?: ((msg: string) => void) | null;
}) {
  const calls = {
    create: [] as string[],
    get: [] as Array<string | number>,
    open: [] as string[],
    alert: [] as string[],
    workbook: 0,
    focus: 0,
  };

  const panes = new Map<string | number, Pane>();
  let nextId = 1;
  const storage = new Map<string, string>();

  const defaultApi = {
    PluginStorage: {
      getItem(key: string) {
        return storage.has(key) ? storage.get(key)! : null;
      },
      setItem(key: string, value: string) {
        storage.set(key, String(value));
      },
    },
    CreateTaskPane(url: string) {
      calls.create.push(url);
      const id = nextId++;
      const pane: Pane = { ID: id, Visible: false, url };
      panes.set(id, pane);
      return pane;
    },
    GetTaskPane(id: string | number) {
      calls.get.push(id);
      if (panes.has(id)) return panes.get(id)!;
      const n = Number(id);
      if (!Number.isNaN(n) && panes.has(n)) return panes.get(n)!;
      return panes.get(String(id)) ?? null;
    },
    Application: {
      ActiveWorkbook: {
        get Name() {
          calls.workbook += 1;
          return "Book1";
        },
      },
    },
  };

  const api = opts.api === null ? null : { ...defaultApi, ...(opts.api || {}) };
  const on = opts.apiOn ?? "wps";

  const defaultHref =
    opts.locationHref ?? "file:///tmp/jsaddons/wengge-excel-ai-addin/index.html";
  const windowObj: Record<string, unknown> = {
    location: {
      get href() {
        if (opts.locationThrows) throw new Error("location boom");
        return defaultHref;
      },
    },
    focus() {
      calls.focus += 1;
    },
    open(url: string) {
      calls.open.push(url);
      if (opts.openImpl) return opts.openImpl(url);
      return { ok: true };
    },
    alert(msg: string) {
      calls.alert.push(String(msg));
      if (opts.alertImpl) opts.alertImpl(msg);
    },
  };

  if (on === "window.wps") windowObj.wps = api;
  if (on === "window.Wps") windowObj.Wps = api;

  const sandbox: Record<string, unknown> = {
    window: windowObj,
    URL,
    calls,
    storage,
    panes,
  };
  if (on === "wps") sandbox.wps = api;
  // bare identifier access
  vm.createContext(sandbox);
  vm.runInContext(entrySource, sandbox, { filename: "wps-entry.js" });

  return {
    window: windowObj as typeof windowObj & {
      WenggeExcelAiOnLoad: (ui: unknown) => void;
      WenggeExcelAiTabVisible: () => boolean;
      WenggeExcelAiOpenChat: () => boolean;
      WenggeExcelAiOpenProviders: () => boolean;
      WenggeExcelAiOpenHost: () => boolean;
      WenggeExcelAiGetImage: (control?: unknown) => string;
      __WenggeExcelAiEntryTest: {
        buildTaskPaneUrl: (r: string) => string;
        openTaskPaneForRoute: (r: string) => boolean;
        getRibbonUI: () => unknown;
      };
    },
    calls,
    storage,
    panes,
    api,
  };
}

describe("WPS entry task pane lifecycle", () => {
  it("loads without opening panes or touching workbook", () => {
    const env = loadEntry({});
    expect(env.calls.create).toEqual([]);
    expect(env.calls.workbook).toBe(0);
    expect(typeof env.window.WenggeExcelAiOpenChat).toBe("function");
    expect(env.window.WenggeExcelAiTabVisible()).toBe(true);
  });

  it("onLoad stores ribbonUI", () => {
    const env = loadEntry({});
    const ui = { id: "ribbon" };
    env.window.WenggeExcelAiOnLoad(ui);
    expect(env.window.__WenggeExcelAiEntryTest.getRibbonUI()).toBe(ui);
  });

  it("builds chat/provider/host URLs from current file location only", () => {
    const env = loadEntry({
      locationHref: "file:///data/wengge-excel-ai-addin/index.html?old=1#hash",
    });
    const chat = env.window.__WenggeExcelAiEntryTest.buildTaskPaneUrl("chat");
    const providers = env.window.__WenggeExcelAiEntryTest.buildTaskPaneUrl("providers");
    const host = env.window.__WenggeExcelAiEntryTest.buildTaskPaneUrl("host");
    expect(chat).toBe("file:///data/wengge-excel-ai-addin/index.html?page=chat");
    expect(providers).toBe("file:///data/wengge-excel-ai-addin/index.html?page=providers");
    expect(host).toBe("file:///data/wengge-excel-ai-addin/index.html?page=host");
    expect(chat).not.toMatch(/localhost|cdn\.|https:\/\/[^f]/);
  });

  it("creates via api.CreateTaskPane and reuses same route", () => {
    const env = loadEntry({});
    expect(env.window.WenggeExcelAiOpenChat()).toBe(true);
    expect(env.calls.create).toHaveLength(1);
    expect(env.calls.create[0]).toContain("page=chat");
    const firstId = [...env.panes.keys()][0];
    expect(env.panes.get(firstId!)!.Visible).toBe(true);

    expect(env.window.WenggeExcelAiOpenChat()).toBe(true);
    expect(env.calls.create).toHaveLength(1);
    expect(env.calls.get).toContain(firstId);
    expect(env.calls.workbook).toBe(0);
  });

  it("hides old pane and creates on route change", () => {
    const env = loadEntry({});
    env.window.WenggeExcelAiOpenChat();
    const firstId = [...env.panes.keys()][0]!;
    env.window.WenggeExcelAiOpenProviders();
    expect(env.calls.create).toHaveLength(2);
    expect(env.panes.get(firstId)!.Visible).toBe(false);
    const secondId = [...env.panes.keys()][1]!;
    expect(env.panes.get(secondId)!.Visible).toBe(true);
    expect(env.calls.create[1]).toContain("page=providers");
  });

  it("falls back to Application.CreateTaskPane", () => {
    const createApp: string[] = [];
    const env = loadEntry({
      api: {
        CreateTaskPane: undefined,
        Application: {
          CreateTaskPane(url: string) {
            createApp.push(url);
            return { ID: 99, Visible: false, url };
          },
          GetTaskPane() {
            return null;
          },
        },
        GetTaskPane: undefined,
        PluginStorage: {
          getItem: () => null,
          setItem: () => undefined,
        },
      },
    });
    expect(env.window.WenggeExcelAiOpenHost()).toBe(true);
    expect(createApp[0]).toContain("page=host");
  });

  it("resolves window.wps and window.Wps", () => {
    for (const apiOn of ["window.wps", "window.Wps"] as const) {
      const env = loadEntry({ apiOn });
      expect(env.window.WenggeExcelAiOpenChat()).toBe(true);
      expect(env.calls.create.length).toBeGreaterThan(0);
    }
  });

  it("PluginStorage failure still opens a pane", () => {
    const env = loadEntry({
      api: {
        PluginStorage: {
          getItem() {
            throw new Error("storage down");
          },
          setItem() {
            throw new Error("storage down");
          },
        },
      },
    });
    expect(env.window.WenggeExcelAiOpenChat()).toBe(true);
    expect(env.calls.create).toHaveLength(1);
  });

  it("GetTaskPane throw / stale pane recreates", () => {
    let shouldThrow = true;
    const env = loadEntry({
      api: {
        GetTaskPane() {
          if (shouldThrow) throw new Error("gone");
          return null;
        },
      },
    });
    env.window.WenggeExcelAiOpenChat();
    // force stored id without valid pane
    env.storage.set("WenggeExcelAi.taskPaneId", "stale");
    env.storage.set("WenggeExcelAi.taskPaneRoute", "chat");
    shouldThrow = true;
    expect(env.window.WenggeExcelAiOpenChat()).toBe(true);
    expect(env.calls.create.length).toBeGreaterThanOrEqual(2);
  });

  it("Visible setter failure uses window.open fallback, not focus-only success", () => {
    const env = loadEntry({
      api: {
        CreateTaskPane(url: string) {
          const pane: Record<string, unknown> = { ID: 7, url };
          Object.defineProperty(pane, "Visible", {
            configurable: true,
            set() {
              throw new Error("visible fail");
            },
            get() {
              return false;
            },
          });
          return pane;
        },
      },
    });
    const ok = env.window.WenggeExcelAiOpenChat();
    expect(ok).toBe(true);
    expect(env.calls.open[0]).toContain("page=chat");
    expect(env.calls.focus).toBe(0);
  });

  it("missing API uses window.open then alert", () => {
    const env = loadEntry({
      apiOn: "none",
      api: null,
      openImpl: () => null,
    });
    const ok = env.window.WenggeExcelAiOpenChat();
    expect(ok).toBe(false);
    expect(env.calls.open.length).toBe(1);
    expect(env.calls.alert.length).toBe(1);
  });

  it("never hardcodes remote CDN task pane hosts", () => {
    expect(entrySource).not.toMatch(/["']https:\/\/cdn\./i);
    expect(entrySource).not.toMatch(/["']https:\/\/localhost/i);
    expect(entrySource).not.toMatch(/\brequire\s*\(|node:child_process|\belectron\b/);
  });

  it("getImage maps control ids to package-relative icons only", () => {
    const env = loadEntry({});
    const getImage = env.window.WenggeExcelAiGetImage;
    expect(typeof getImage).toBe("function");
    expect(getImage({ Id: "wenggeExcelAiOpenChatButton" })).toBe("assets/icon-32.png");
    expect(getImage({ ID: "wenggeExcelAiOpenProvidersButton" })).toBe("assets/icon-32.png");
    expect(getImage({ id: "wenggeExcelAiOpenHostButton" })).toBe("assets/icon-16.png");
    expect(getImage("wenggeExcelAiOpenHostButton")).toBe("assets/icon-16.png");
    expect(getImage(null)).toBe("assets/icon-32.png");
    expect(getImage({ Id: "unknown" })).toBe("assets/icon-32.png");
    expect(getImage({ Id: "../evil.png" })).toBe("assets/icon-32.png");
    expect(getImage({ Id: "https://evil.example/x.png" })).toBe("assets/icon-32.png");
    // must not echo arbitrary path from control fields
    expect(getImage({ Id: "wenggeExcelAiOpenChatButton", path: "/tmp/x" })).toBe(
      "assets/icon-32.png",
    );
  });

  it("rejects javascript/data/blob locations without open or create", () => {
    for (const href of [
      "javascript:alert(1)",
      "data:text/html,hi",
      "blob:https://example.com/uuid",
      "about:blank",
    ]) {
      const env = loadEntry({ locationHref: href });
      const ok = env.window.WenggeExcelAiOpenChat();
      expect(ok).toBe(false);
      expect(env.calls.create).toEqual([]);
      expect(env.calls.open).toEqual([]);
      expect(env.calls.alert.length).toBeGreaterThan(0);
      expect(() => env.window.__WenggeExcelAiEntryTest.buildTaskPaneUrl("chat")).toThrow(
        /protocol|forbidden|unsupported|unsafe|location/i,
      );
    }
  });

  it("location getter throw fails closed without open/create", () => {
    const env = loadEntry({ locationThrows: true });
    expect(env.window.WenggeExcelAiOpenChat()).toBe(false);
    expect(env.calls.create).toEqual([]);
    expect(env.calls.open).toEqual([]);
    expect(env.calls.alert.length).toBeGreaterThan(0);
  });

  it("control characters in location fail closed", () => {
    const env = loadEntry({
      locationHref: "file:///tmp/index.html\u0000evil",
    });
    expect(env.window.WenggeExcelAiOpenChat()).toBe(false);
    expect(env.calls.create).toEqual([]);
    expect(env.calls.open).toEqual([]);
  });

  it("CreateTaskPane throw falls back to Application.CreateTaskPane", () => {
    const appCreates: string[] = [];
    const env = loadEntry({
      api: {
        CreateTaskPane() {
          throw new Error("create boom");
        },
        Application: {
          CreateTaskPane(url: string) {
            appCreates.push(url);
            return { ID: 42, Visible: false, url };
          },
          GetTaskPane() {
            return null;
          },
        },
        GetTaskPane: undefined,
        PluginStorage: {
          getItem: () => null,
          setItem: () => undefined,
        },
      },
    });
    expect(() => env.window.WenggeExcelAiOpenChat()).not.toThrow();
    expect(env.window.WenggeExcelAiOpenChat()).toBe(true);
    expect(appCreates.length).toBeGreaterThan(0);
    expect(appCreates[0]).toContain("page=chat");
  });

  it("both CreateTaskPane paths throw: callback does not throw; uses open fallback", () => {
    const env = loadEntry({
      api: {
        CreateTaskPane() {
          throw new Error("create boom");
        },
        Application: {
          CreateTaskPane() {
            throw new Error("app create boom");
          },
        },
        PluginStorage: {
          getItem: () => null,
          setItem: () => undefined,
        },
      },
    });
    let threw = false;
    let ok = false;
    try {
      ok = env.window.WenggeExcelAiOpenChat();
    } catch {
      threw = true;
    }
    expect(threw).toBe(false);
    expect(ok).toBe(true);
    expect(env.calls.open[0]).toContain("page=chat");
  });

  it("pane ID getter throw after Visible success still returns true", () => {
    let createCount = 0;
    const env = loadEntry({
      api: {
        CreateTaskPane(url: string) {
          createCount += 1;
          const pane: Record<string, unknown> = { url, Visible: false };
          Object.defineProperty(pane, "ID", {
            get() {
              throw new Error("id boom");
            },
          });
          Object.defineProperty(pane, "Id", {
            get() {
              throw new Error("id boom");
            },
          });
          Object.defineProperty(pane, "id", {
            get() {
              throw new Error("id boom");
            },
          });
          return pane;
        },
      },
    });
    expect(() => env.window.WenggeExcelAiOpenChat()).not.toThrow();
    expect(env.window.WenggeExcelAiOpenChat()).toBe(true);
    expect(createCount).toBeGreaterThan(0);
    expect(env.calls.open).toEqual([]);
  });

});
