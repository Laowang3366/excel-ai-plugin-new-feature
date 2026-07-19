import { afterEach, describe, expect, it } from "vitest";
import { withExcel } from "../shared/host/officeJsRuntime";
import { fail, ok, unsupported } from "../shared/host/types";
import { mapHostResultToToolResult } from "../shared/tools/hostResultMapping";
import { ToolExecutor } from "../shared/tools";
import { executeChartTool } from "../shared/tools/chartExecutor";
import type { HostAdapter } from "../shared/host/types";
import { MockHostAdapter } from "./mockHost";

describe("phase35 failure semantics", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
    delete (globalThis as { window?: unknown }).window;
  });

  it("withExcel missing Excel.run is typed unsupported", async () => {
    (globalThis as unknown as { window: unknown }).window = globalThis;
    delete (globalThis as { Excel?: unknown }).Excel;
    const result = await withExcel("range.read", async () => ({ ok: true }));
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      expect(result.capability).toBe("range.read");
      expect(result.host).toBe("office-js");
      expect(result.reason).toMatch(/Excel\.run is not available/);
    }
  });

  it("withExcel existing run with callback throw is ordinary fail", async () => {
    (globalThis as unknown as { window: unknown }).window = globalThis;
    (globalThis as unknown as { Excel: { run: Function } }).Excel = {
      run: async <T>(fn: (ctx: { sync: () => Promise<void> }) => Promise<T>) =>
        fn({
          async sync() {
            throw new Error("sync blew up");
          },
        }),
    };
    const result = await withExcel("range.read", async (ctx) => {
      await ctx.sync();
      return { sheetName: "S" };
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).not.toBe(true);
      expect(result.capability).toBe("range.read");
      expect(result.host).toBe("office-js");
      expect(result.reason).toMatch(/sync blew up/);
    }
  });

  it("withExcel run wrapper reject is ordinary fail", async () => {
    (globalThis as unknown as { window: unknown }).window = globalThis;
    (globalThis as unknown as { Excel: { run: Function } }).Excel = {
      run: async () => {
        throw new Error("Excel.run rejected batch");
      },
    };
    const result = await withExcel("range.read", async () => 1);
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).not.toBe(true);
      expect(result.reason).toMatch(/Excel\.run rejected batch/);
    }
  });

  it("mapHostResultToToolResult preserves ordinary fail without unsupported flag", () => {
    const hostFail = fail("range.read", "office-js", "sheet missing");
    const mapped = mapHostResultToToolResult("range.read", hostFail);
    expect(mapped.ok).toBe(false);
    if (!mapped.ok) {
      expect(mapped.unsupported).not.toBe(true);
      expect(mapped.error).toMatch(/sheet missing/);
      expect(mapped.detail).toMatchObject({
        capability: "range.read",
        host: "office-js",
        reason: "sheet missing",
      });
    }
  });

  it("mapHostResultToToolResult keeps true unsupported", () => {
    const hostUn = unsupported(
      "range.read",
      "office-js",
      "no Excel.run",
      "Requires Microsoft Office Excel with Office.js loaded",
    );
    const mapped = mapHostResultToToolResult("range.read", hostUn);
    expect(mapped.ok).toBe(false);
    if (!mapped.ok) {
      expect(mapped.unsupported).toBe(true);
      expect(mapped.detail).toMatchObject({
        capability: "range.read",
        unsupported: true,
        evidence: expect.stringMatching(/Office\.js/),
      });
    }
  });

  it("ToolExecutor range.read maps Host fail() without unsupported true", async () => {
    const host = new MockHostAdapter();
    host.readRange = async () => fail("range.read", "office-js", "forced ordinary fail");
    const result = await new ToolExecutor(host).execute({
      name: "range.read",
      arguments: { sheetName: "Sheet1", range: "A1" },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).not.toBe(true);
      const detail = result.detail as { capability?: string; host?: string; reason?: string };
      expect(detail.capability).toBe("range.read");
      expect(detail.host).toBe("office-js");
      expect(detail.reason).toMatch(/forced ordinary fail/);
    }
  });

  it("chart.create maps Host fail() without unsupported true", async () => {
    const host = {
      kind: "office-js" as const,
      async createChart() {
        return fail("chart.create", "office-js", "chart create failed");
      },
    } as unknown as HostAdapter;
    const result = await executeChartTool(host, {
      name: "chart.create",
      arguments: { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: "column" },
    });
    expect(result).not.toBeNull();
    expect(result!.ok).toBe(false);
    if (result && !result.ok) {
      expect(result.unsupported).not.toBe(true);
      const detail = result.detail as { capability?: string; reason?: string };
      expect(detail.capability).toBe("chart.create");
      expect(detail.reason).toMatch(/chart create failed/);
    }
  });

  it("true unsupported() through chart mapper stays unsupported", async () => {
    const host = {
      kind: "office-js" as const,
      async createChart() {
        return unsupported("chart.create", "office-js", "no charts", "evidence");
      },
    } as unknown as HostAdapter;
    const result = await executeChartTool(host, {
      name: "chart.create",
      arguments: { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: "column" },
    });
    expect(result!.ok).toBe(false);
    if (result && !result.ok) {
      expect(result.unsupported).toBe(true);
      expect((result.detail as { evidence?: string }).evidence).toBe("evidence");
    }
  });

  it("ok host maps to success", () => {
    const mapped = mapHostResultToToolResult("host.status", ok({ connected: true }));
    expect(mapped).toEqual({ ok: true, tool: "host.status", data: { connected: true } });
  });
});
