import { existsSync, mkdtempSync, readFileSync, rmSync } from "fs";
import os from "os";
import path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const powershellMocks = vi.hoisted(() => ({
  executePowerShell: vi.fn(async (_script: string) => JSON.stringify({ changed: true })),
}));

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return { ...actual, executePowerShell: powershellMocks.executePowerShell };
});

import { WpsJsaBridge } from "./wpsJsaBridge";

function createBridge(host: "excel" | "wps"): WpsJsaBridge {
  return new WpsJsaBridge({ host } as never);
}

describe("WpsJsaBridge", () => {
  let originalAppData: string | undefined;
  let tempAppData = "";

  beforeEach(() => {
    originalAppData = process.env.APPDATA;
    tempAppData = mkdtempSync(path.join(os.tmpdir(), "wengge-jsa-test-"));
    process.env.APPDATA = tempAppData;
    powershellMocks.executePowerShell.mockClear();
  });

  afterEach(() => {
    if (originalAppData === undefined) delete process.env.APPDATA;
    else process.env.APPDATA = originalAppData;
    rmSync(tempAppData, { recursive: true, force: true });
  });

  it("does not advertise desktop JavaScript as an Excel macro language", async () => {
    await expect(createBridge("excel").detectCapabilities()).resolves.toEqual({
      language: "javascript",
      supported: false,
      ready: false,
      internal: true,
      engine: "WPS JSA",
      reason: "WPS JSA 仅在 WPS 表格中可用",
    });
  });

  it("rejects JavaScript writes outside the WPS internal JSA host", async () => {
    await expect(createBridge("excel").writeCode("function main() {}", {
      entryPoint: "main",
    })).rejects.toThrow("Microsoft Excel 请使用 VBA");
  });

  it("validates the JSA entry point before installing or changing the add-in", async () => {
    await expect(createBridge("wps").writeCode("function other() {}", {
      entryPoint: "main",
    })).rejects.toThrow("找不到入口函数: main");
  });

  it("installs the narrow JSA bridge on first write and requires a WPS restart", async () => {
    await expect(createBridge("wps").writeCode("function main() {}", {
      entryPoint: "main",
    })).rejects.toThrow("内部桥接已安装");

    const addonDir = path.join(tempAppData, "kingsoft", "wps", "jsaddons", "WenggeJsaBridge_");
    expect(existsSync(path.join(addonDir, "index.html"))).toBe(true);
    expect(readFileSync(path.join(addonDir, "bridge-config.js"), "utf8")).toContain("WENGGE_JSA_BRIDGE");
    expect(powershellMocks.executePowerShell).toHaveBeenCalledOnce();
    expect(powershellMocks.executePowerShell.mock.calls[0][0]).toContain("System.Xml.XmlDocument");
  });
});
