import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { CHAT_READONLY_TOOL_ALLOWLIST } from "../shared/agentChat/chatReadOnlyTools";
import { MockHostAdapter } from "./mockHost";
import {
  installWorkbookSaveExcel,
  uninstallWorkbookSaveExcel,
} from "./fakes/officeJsWorkbookSaveFake";

describe("workbook.save registry", () => {
  it("registers moderate tool with empty schema; not chat-readonly", () => {
    const def = TOOL_DEFINITIONS.find((t) => t.name === "workbook.save");
    expect(def).toBeTruthy();
    expect(def!.riskLevel).toBe("moderate");
    expect(def!.parameters).toMatchObject({
      type: "object",
      required: [],
      additionalProperties: false,
    });
    expect(Object.keys((def!.parameters as { properties: object }).properties)).toEqual([]);
    expect(CHAT_READONLY_TOOL_ALLOWLIST).not.toContain("workbook.save");
  });

  it("executor success + unknown field reject + forced fail", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const ok = await executor.execute({ name: "workbook.save", arguments: {} });
    expect(ok.ok).toBe(true);
    if (ok.ok) expect(ok.data).toEqual({ workbookName: "Book1.xlsx", saved: true });

    const unknown = await executor.execute({
      name: "workbook.save",
      arguments: { path: "C:\\x.xlsx" },
    });
    expect(unknown.ok).toBe(false);
    if (!unknown.ok) expect(unknown.error).toMatch(/unknown field: path/);

    host.failCapability = "workbook.save";
    const failed = await executor.execute({ name: "workbook.save", arguments: {} });
    expect(failed.ok).toBe(false);
  });
});

describe("workbook.save Office.js", () => {
  afterEach(() => {
    uninstallWorkbookSaveExcel();
  });

  it("queues save + sync and returns name", async () => {
    const fake = installWorkbookSaveExcel({ workbookName: "Sales.xlsx" });
    const result = await new OfficeJsAdapter().saveWorkbook();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.data).toEqual({ workbookName: "Sales.xlsx", saved: true });
    }
    expect(fake.saveCallCount()).toBe(1);
    expect(fake.syncCount()).toBe(1);
  });

  it("ExcelApi 1.1 missing → unsupported; isSetSupported throw → unsupported", async () => {
    installWorkbookSaveExcel({ excelApi11: false });
    const missing = await new OfficeJsAdapter().saveWorkbook();
    expect(missing.ok).toBe(false);
    if (!missing.ok) {
      expect(missing.unsupported).toBe(true);
      expect(missing.reason).toMatch(/ExcelApi 1\.1/);
    }

    uninstallWorkbookSaveExcel();
    installWorkbookSaveExcel({ isSetSupportedThrows: true });
    const throws = await new OfficeJsAdapter().saveWorkbook();
    expect(throws.ok).toBe(false);
    if (!throws.ok) expect(throws.unsupported).toBe(true);

    uninstallWorkbookSaveExcel();
    installWorkbookSaveExcel({ missingIsSetSupported: true });
    const noFn = await new OfficeJsAdapter().saveWorkbook();
    expect(noFn.ok).toBe(false);
    if (!noFn.ok) expect(noFn.unsupported).toBe(true);
  });

  it("sync error after save → ordinary fail (not unsupported)", async () => {
    installWorkbookSaveExcel({ syncError: "user cancelled save dialog" });
    const result = await new OfficeJsAdapter().saveWorkbook();
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).not.toBe(true);
      expect(result.reason).toMatch(/cancelled|save dialog/i);
    }
  });
});

describe("workbook.save WPS JSA", () => {
  afterEach(() => {
    delete (globalThis as unknown as { window?: { Application?: unknown } }).window;
    delete (globalThis as unknown as { Application?: unknown }).Application;
  });

  function installWps(opts: { save?: (() => void) | "missing" | "throw"; name?: string } = {}) {
    const { save = () => {}, name = "BookWps.xlsx" } = opts;
    const workbook: Record<string, unknown> = { Name: name };
    if (save === "throw") {
      workbook.Save = () => {
        throw new Error("disk full");
      };
    } else if (save !== "missing") {
      workbook.Save = save;
    }
    const app = { ActiveWorkbook: workbook, Name: "WPS 表格" };
    (globalThis as unknown as { window: unknown }).window = globalThis;
    (globalThis as unknown as { Application: unknown }).Application = app;
    return workbook;
  }

  it("Save present → ok with name", async () => {
    let called = 0;
    installWps({
      save: () => {
        called += 1;
      },
      name: "W.xlsx",
    });
    const result = await new WpsJsaAdapter().saveWorkbook();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.data).toEqual({ workbookName: "W.xlsx", saved: true });
    expect(called).toBe(1);
  });

  it("Save missing → unsupported with bridge evidence", async () => {
    installWps({ save: "missing" });
    const result = await new WpsJsaAdapter().saveWorkbook();
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      expect(result.reason).toMatch(/Save is not a function/);
      expect(result.evidence).toMatch(/wps-jsa-bridge/);
    }
  });

  it("Save throws → fail", async () => {
    installWps({ save: "throw" });
    const result = await new WpsJsaAdapter().saveWorkbook();
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).not.toBe(true);
      expect(result.reason).toMatch(/disk full/);
    }
  });
});
