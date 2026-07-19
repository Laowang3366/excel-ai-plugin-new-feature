import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartAxesExcel } from "./fakes/officeJsChartAxesFake";
import { MockHostAdapter } from "./mockHost";

describe("phase18 chart axes update", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartAxesExcel>;
    beforeEach(() => {
      fake = installChartAxesExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("updates value primary min/max with real readback", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        minimum: 10,
        maximum: 200,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({
          sheetName: "Sheet1",
          chartName: "C1",
          kind: "value",
          group: "primary",
          minimum: 10,
          maximum: 200,
          majorUnit: 10,
          reverse: false,
        });
      }
      expect(fake.getCommitted("Value", "Primary")?.minimum).toBe(10);
    });

    it("sets title and clears with empty string", async () => {
      const adapter = new OfficeJsAdapter();
      const set = await adapter.updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "category",
        group: "primary",
        title: "X Axis",
      });
      expect(set.ok).toBe(true);
      if (set.ok) {
        expect(set.data.title).toBe("X Axis");
        expect(set.data.titleVisible).toBe(true);
      }
      const cleared = await adapter.updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "category",
        title: "",
      });
      expect(cleared.ok).toBe(true);
      if (cleared.ok) {
        expect(cleared.data.title).toBe("");
        expect(cleared.data.titleVisible).toBe(false);
      }
    });

    it("updates secondary value reverse and numberFormat", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        group: "secondary",
        reverse: true,
        numberFormat: "0.0",
        majorUnit: 5,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.group).toBe("secondary");
        expect(result.data.reverse).toBe(true);
        expect(result.data.numberFormat).toBe("0.0");
        expect(result.data.majorUnit).toBe(5);
      }
    });

    it("skipping first sync after write yields old axis snapshot", async () => {
      const broken = await fake.brokenUpdateSkipFirstSync();
      expect(broken.minimum).toBe(0);
      expect(broken.maximum).toBe(100);
      expect(fake.getCommitted("Value", "Primary")?.minimum).toBe(5);
      expect(fake.getCommitted("Value", "Primary")?.maximum).toBe(50);

      const adapter = new OfficeJsAdapter();
      const ok = await adapter.updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        minimum: 1,
        maximum: 9,
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) {
        expect(ok.data.minimum).toBe(1);
        expect(ok.data.maximum).toBe(9);
      }
    });

    it("rejects NaN/object/undefined scale and non-string chart.name after load", async () => {
      fake.poisonCommitted("Value", "Primary", { minimum: Number.NaN });
      expect(
        (
          await new OfficeJsAdapter().updateChartAxis({
            sheetName: "Sheet1",
            chartName: "C1",
            kind: "value",
            reverse: true,
          })
        ).ok,
      ).toBe(false);

      const fake2 = installChartAxesExcel();
      fake2.poisonCommitted("Value", "Primary", {
        minimum: { x: 1 } as unknown as number,
      });
      expect(
        (
          await new OfficeJsAdapter().updateChartAxis({
            sheetName: "Sheet1",
            chartName: "C1",
            kind: "value",
            reverse: true,
          })
        ).ok,
      ).toBe(false);

      const fake3 = installChartAxesExcel();
      fake3.poisonCommitted("Value", "Primary", {
        reverse: undefined as unknown as boolean,
      });
      expect(
        (
          await new OfficeJsAdapter().updateChartAxis({
            sheetName: "Sheet1",
            chartName: "C1",
            kind: "value",
            minimum: 1,
          })
        ).ok,
      ).toBe(false);

      const fake4 = installChartAxesExcel();
      fake4.setChartName(42);
      expect(
        (
          await new OfficeJsAdapter().updateChartAxis({
            sheetName: "Sheet1",
            chartName: "C1",
            kind: "value",
            reverse: true,
          })
        ).ok,
      ).toBe(false);
    });
  });

  describe("executor + schema", () => {
    it("MockHost category and value primary/secondary", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);

      const cat = await executor.execute({
        name: "chart.axes.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          kind: "category",
          title: "Cat",
        },
      });
      expect(cat.ok).toBe(true);
      if (cat.ok) {
        expect(cat.data).toMatchObject({
          sheetName: "Sheet1",
          chartName: "C1",
          kind: "category",
          group: "primary",
          title: "Cat",
          titleVisible: true,
        });
      }

      const val = await executor.execute({
        name: "chart.axes.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          kind: "value",
          group: "secondary",
          minimum: 2,
          maximum: 20,
          reverse: true,
        },
      });
      expect(val.ok).toBe(true);
      if (val.ok) {
        expect(val.data).toMatchObject({
          kind: "value",
          group: "secondary",
          minimum: 2,
          maximum: 20,
          reverse: true,
        });
      }
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", kind: "value" },
        { sheetName: "Sheet1", chartName: "C1", kind: "series", minimum: 1 },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", group: "tertiary", minimum: 1 },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", minimum: null },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", minimum: undefined },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", majorUnit: -1 },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", numberFormat: "" },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", numberFormat: "   " },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", title: null },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", title: "   " },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", reverse: "yes" },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", minimum: 1, extra: true },
        { sheetName: "", chartName: "C1", kind: "value", minimum: 1 },
      ]) {
        const result = await executor.execute({
          name: "chart.axes.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema additionalProperties false and moderate risk", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.axes.update");
      expect(def?.riskLevel).toBe("moderate");
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
    });
  });

  it("WPS chart.axes.update is typed unsupported with ChartAxis evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.axes.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        minimum: 0,
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
