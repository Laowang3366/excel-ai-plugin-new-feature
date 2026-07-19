import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools";
import { installObjectUpdateExcel } from "./fakes/officeJsObjectUpdateFake";
import { installSyncGatedExcel } from "./fakes/officeJsSyncGated";
import { MockHostAdapter } from "./mockHost";

describe("phase14 chart style + legend", () => {
  describe("Office.js list", () => {
    beforeEach(() => {
      installSyncGatedExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("list defaults include style and legendVisible", async () => {
      const adapter = new OfficeJsAdapter();
      await adapter.createChart({
        sheetName: "Sheet1",
        sourceRange: "A1:B2",
        name: "CList",
      });
      const listed = await adapter.listCharts("Sheet1");
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        const row = listed.data.find((c) => c.name === "CList");
        expect(row?.style).toBe(2);
        expect(row?.legendVisible).toBe(true);
      }
    });
  });

  describe("Office.js update", () => {
    beforeEach(() => {
      installObjectUpdateExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("updates style alone with writeback", async () => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChart({
        sheetName: "Sheet1",
        chartName: "C1",
        style: 10,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.style).toBe(10);
        expect(updated.data.legendVisible).toBe(true);
      }
    });

    it("updates showLegend alone with writeback", async () => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChart({
        sheetName: "Sheet1",
        chartName: "C1",
        showLegend: false,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.legendVisible).toBe(false);
        expect(updated.data.style).toBe(2);
      }
    });

    it("updates style and showLegend together", async () => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChart({
        sheetName: "Sheet1",
        chartName: "C1",
        style: 8,
        showLegend: false,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.style).toBe(8);
        expect(updated.data.legendVisible).toBe(false);
      }
    });
  });

  describe("executor", () => {
    it("accepts style/showLegend and rejects invalid", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);

      expect(
        (
          await executor.execute({
            name: "chart.update",
            arguments: { sheetName: "Sheet1", chartName: "C1", style: 5 },
          })
        ).ok,
      ).toBe(true);

      expect(
        (
          await executor.execute({
            name: "chart.update",
            arguments: { sheetName: "Sheet1", chartName: "C1", showLegend: false },
          })
        ).ok,
      ).toBe(true);

      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", style: 0 },
        { sheetName: "Sheet1", chartName: "C1", style: 1.5 },
        { sheetName: "Sheet1", chartName: "C1", style: -1 },
        { sheetName: "Sheet1", chartName: "C1", showLegend: "yes" },
        { sheetName: "Sheet1", chartName: "C1", style: null },
        { sheetName: "Sheet1", chartName: "C1", showLegend: null },
        { sheetName: "Sheet1", chartName: "C1", style: 5, showLegend: null },
      ]) {
        const result = await executor.execute({
          name: "chart.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });
  });

  it("WPS chart.update remains unsupported", async () => {
    const executor = new ToolExecutor(new WpsJsaAdapter());
    const result = await executor.execute({
      name: "chart.update",
      arguments: { sheetName: "Sheet1", chartName: "C1", style: 3, showLegend: true },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.unsupported).toBe(true);
  });
});
