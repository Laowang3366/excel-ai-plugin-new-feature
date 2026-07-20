import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools/executor";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { installChartAxesExcel } from "./fakes/officeJsChartAxesFake";
import { MockHostAdapter } from "./mockHost";

describe("phase46 chart axes displayUnit/scaleType/gridlines", () => {
  describe("Office.js sync-gated advanced fields", () => {
    let fake: ReturnType<typeof installChartAxesExcel>;
    beforeEach(() => {
      fake = installChartAxesExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("updates displayUnit and showDisplayUnitLabel with host readback", async () => {
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        displayUnit: "thousands",
        showDisplayUnitLabel: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.displayUnit).toBe("thousands");
        expect(result.data.showDisplayUnitLabel).toBe(true);
      }
      expect(fake.getCommitted("Value", "Primary")?.displayUnit).toBe("Thousands");
      expect(fake.getCommitted("Value", "Primary")?.showDisplayUnitLabel).toBe(true);
    });

    it("sets custom display unit via setCustomDisplayUnit", async () => {
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        displayUnit: "custom",
        customDisplayUnit: 2500,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.displayUnit).toBe("custom");
        expect(result.data.customDisplayUnit).toBe(2500);
      }
      expect(fake.getCommitted("Value", "Primary")?.displayUnit).toBe("Custom");
      expect(fake.getCommitted("Value", "Primary")?.customDisplayUnit).toBe(2500);
    });

    it("updates scaleType logarithmic + logBase with readback", async () => {
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        scaleType: "logarithmic",
        logBase: 2,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.scaleType).toBe("logarithmic");
        expect(result.data.logBase).toBe(2);
      }
      expect(fake.getCommitted("Value", "Primary")?.scaleType).toBe("Logarithmic");
      expect(fake.getCommitted("Value", "Primary")?.logBase).toBe(2);
    });

    it("updates major/minor gridlines visibility", async () => {
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        majorGridlinesVisible: false,
        minorGridlinesVisible: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.majorGridlinesVisible).toBe(false);
        expect(result.data.minorGridlinesVisible).toBe(true);
      }
      expect(fake.getCommitted("Value", "Primary")?.majorGridlinesVisible).toBe(false);
      expect(fake.getCommitted("Value", "Primary")?.minorGridlinesVisible).toBe(true);
    });

    it("does not echo request when host tampers displayUnit after write", async () => {
      const adapter = new OfficeJsAdapter();
      // First write queues; poison after would need mid-path — poison committed before call
      // so load snapshot returns wrong displayUnit while write pending merges then load.
      // Simulate: write succeeds, then poison and re-read path via second call with reverse only
      // fails if poison makes type invalid — instead poison after write by checking:
      const ok = await adapter.updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        displayUnit: "millions",
      });
      expect(ok.ok).toBe(true);
      fake.poisonCommitted("Value", "Primary", { displayUnit: "Hundreds" });
      const second = await adapter.updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        reverse: true,
      });
      expect(second.ok).toBe(true);
      if (second.ok) {
        // reverse write does not set displayUnit; host still returns poisoned Hundreds
        expect(second.data.displayUnit).toBe("hundreds");
        expect(second.data.reverse).toBe(true);
      }
    });

    it("returns unsupported when ExcelApi 1.7 missing for displayUnit", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartAxesExcel({ excelApi17: false });
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        displayUnit: "thousands",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.7/);
      }
    });

    it("gridlines-only still works when ExcelApi 1.7 missing", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartAxesExcel({ excelApi17: false });
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        majorGridlinesVisible: false,
      });
      // Always loads 1.7 fields for full snapshot — on 1.7-missing hosts this may still run
      // through Excel.run; gate only applies to 1.7 write fields. Gridlines write is 1.1.
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.majorGridlinesVisible).toBe(false);
      }
    });

    it("fails closed when host tampers scaleType type after write", async () => {
      fake.poisonCommitted("Value", "Primary", {
        scaleType: { bad: true } as unknown as string,
      });
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        reverse: true,
      });
      expect(result.ok).toBe(false);
    });
  });

  describe("executor + schema", () => {
    it("MockHost advanced fields round-trip", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "chart.axes.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          kind: "value",
          displayUnit: "millions",
          showDisplayUnitLabel: true,
          scaleType: "logarithmic",
          logBase: 10,
          majorGridlinesVisible: false,
          minorGridlinesVisible: true,
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({
          displayUnit: "millions",
          showDisplayUnitLabel: true,
          scaleType: "logarithmic",
          logBase: 10,
          majorGridlinesVisible: false,
          minorGridlinesVisible: true,
        });
      }
    });

    it("rejects invalid advanced args before host", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      const calls = 0;
      const original = host.updateChartAxis.bind(host);
      let hostCalls = 0;
      host.updateChartAxis = async (input) => {
        hostCalls += 1;
        return original(input);
      };
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", kind: "value", displayUnit: "kilo" },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", displayUnit: "custom" },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          kind: "value",
          displayUnit: "millions",
          customDisplayUnit: 100,
        },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", scaleType: "log" },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", logBase: 0 },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", logBase: null },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          kind: "value",
          majorGridlinesVisible: "yes",
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          kind: "value",
          displayUnit: "none",
          extra: true,
        },
      ]) {
        const result = await executor.execute({
          name: "chart.axes.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
      expect(hostCalls).toBe(0);
      void calls;
    });

    it("schema lists advanced enums and stays moderate", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.axes.update");
      expect(def?.riskLevel).toBe("moderate");
      const props = (def?.parameters as { properties: Record<string, { enum?: string[] }> })
        .properties;
      expect(props.displayUnit.enum).toContain("thousands");
      expect(props.displayUnit.enum).toContain("custom");
      expect(props.scaleType.enum).toEqual(["linear", "logarithmic"]);
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
    });
  });

  it("WPS remains typed unsupported for advanced axes fields", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.axes.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        displayUnit: "thousands",
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.axes.update");
      expect(detail.evidence).toMatch(/ChartAxis/);
    }
  });
});
