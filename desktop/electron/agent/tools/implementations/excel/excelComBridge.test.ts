import { describe, expect, test, vi } from "vitest";
import { ExcelComBridge, normalizeWorkbookInspectMetadata, resolveSpreadsheetHost } from "./excelComBridge";
import {
  normalizeWorkbookInspectMetadata as normalizeWorkbookInspectMetadataFromConnection,
  resolveSpreadsheetHost as resolveSpreadsheetHostFromConnection,
} from "./connectionMetadata";

describe("resolveSpreadsheetHost", () => {
  test("requires an explicit host when Excel and WPS are both available", () => {
    expect(resolveSpreadsheetHost(["excel", "wps"], null)).toBeNull();
  });

  test("uses the selected host when it is available", () => {
    expect(resolveSpreadsheetHost(["excel", "wps"], "wps")).toBe("wps");
  });

  test("keeps the current host when it is still available", () => {
    expect(resolveSpreadsheetHost(["excel", "wps"], null, "excel")).toBe("excel");
  });

  test("auto-selects the only available host", () => {
    expect(resolveSpreadsheetHost(["wps"], null)).toBe("wps");
  });
});

describe("normalizeWorkbookInspectMetadata", () => {
  test("uses the connected WPS host instead of COM-compatible Excel metadata", () => {
    const raw = {
      name: "Microsoft Excel",
      version: "12.0",
      workbooks: [{ name: "demo.xlsx", sheets: [] }],
    };

    expect(normalizeWorkbookInspectMetadata(raw, "wps", "12.1.0")).toEqual({
      host: "wps",
      name: "WPS 表格",
      version: "12.1.0",
      workbooks: [{ name: "demo.xlsx", sheets: [] }],
    });
  });
});

describe("ExcelComBridge.connect", () => {
  test("uses the shared zero-retry detection path", async () => {
    const bridge = new ExcelComBridge();
    const status = {
      connected: false,
      host: "unknown",
      availableHosts: ["excel", "wps"],
    };
    const detectAndConnect = vi.fn().mockResolvedValue(status);

    (bridge as unknown as { detectAndConnect: typeof detectAndConnect }).detectAndConnect = detectAndConnect;

    await expect(bridge.connect()).resolves.toBe(status);
    expect(detectAndConnect).toHaveBeenCalledWith(0, false);
  });
});

describe("connection metadata helpers", () => {
  test("exports host resolution helpers from the capability module", () => {
    expect(resolveSpreadsheetHostFromConnection(["excel", "wps"], "excel")).toBe("excel");
    expect(normalizeWorkbookInspectMetadataFromConnection({ workbooks: [] }, "wps")).toMatchObject({
      host: "wps",
      name: "WPS 表格",
    });
  });
});
