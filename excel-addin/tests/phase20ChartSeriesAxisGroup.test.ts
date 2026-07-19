import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { mapAxisGroupLabel } from "../shared/host/officeJsChartSeriesAxisGroup";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartSeriesAxisGroupExcel } from "./fakes/officeJsChartSeriesAxisGroupFake";
import { MockHostAdapter } from "./mockHost";

describe("phase20 chart series axisGroup", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSeriesAxisGroupExcel>;
    beforeEach(() => {
      fake = installChartSeriesAxisGroupExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("updates primary to secondary with real readback", async () => {
      const result = await new OfficeJsAdapter().updateChartSeriesAxisGroup({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        axisGroup: "secondary",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          axisGroup: "secondary",
        });
      }
      expect(fake.getCommitted(0)?.axisGroup).toBe("Secondary");
    });

    it("updates series 2 to primary", async () => {
      const result = await new OfficeJsAdapter().updateChartSeriesAxisGroup({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 2,
        axisGroup: "primary",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.seriesIndex).toBe(2);
        expect(result.data.axisGroup).toBe("primary");
      }
    });

    it("skipping first sync yields old axisGroup snapshot", async () => {
      const broken = await fake.brokenUpdateSkipFirstSync();
      expect(broken.axisGroup).toBe("Primary");
      expect(fake.getCommitted(0)?.axisGroup).toBe("Secondary");

      const ok = await new OfficeJsAdapter().updateChartSeriesAxisGroup({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        axisGroup: "primary",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.axisGroup).toBe("primary");
    });

    it("unknown host axisGroup string is passed through on load readback", async () => {
      expect(mapAxisGroupLabel("Primary")).toBe("primary");
      expect(mapAxisGroupLabel("Secondary")).toBe("secondary");
      expect(mapAxisGroupLabel("CustomGroup")).toBe("CustomGroup");
      expect(() => mapAxisGroupLabel(null)).toThrow(/loaded string/);

      // write secondary → sync, then load captures forced unknown host value
      fake.setLoadOverride(0, "CustomGroup");
      const result = await new OfficeJsAdapter().updateChartSeriesAxisGroup({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        axisGroup: "secondary",
      });
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.data.axisGroup).toBe("CustomGroup");
      expect(fake.getCommitted(0)?.axisGroup).toBe("Secondary");
    });

    it("out of range and bad chart.name fail", async () => {
      expect(
        (
          await new OfficeJsAdapter().updateChartSeriesAxisGroup({
            sheetName: "Sheet1",
            chartName: "C1",
            seriesIndex: 99,
            axisGroup: "primary",
          })
        ).ok,
      ).toBe(false);

      const f2 = installChartSeriesAxisGroupExcel();
      f2.setChartName(42);
      expect(
        (
          await new OfficeJsAdapter().updateChartSeriesAxisGroup({
            sheetName: "Sheet1",
            chartName: "C1",
            seriesIndex: 1,
            axisGroup: "secondary",
          })
        ).ok,
      ).toBe(false);
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity primary and secondary", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);

      const sec = await executor.execute({
        name: "chart.series.axisGroup.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          axisGroup: "secondary",
        },
      });
      expect(sec.ok).toBe(true);
      if (sec.ok) {
        expect(sec.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          axisGroup: "secondary",
        });
      }

      const pri = await executor.execute({
        name: "chart.series.axisGroup.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 2,
          axisGroup: "primary",
        },
      });
      expect(pri.ok).toBe(true);
      if (pri.ok) expect(pri.data).toMatchObject({ seriesIndex: 2, axisGroup: "primary" });
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 0, axisGroup: "primary" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1.5, axisGroup: "primary" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: null, axisGroup: "primary" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, axisGroup: "tertiary" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, axisGroup: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, axisGroup: undefined },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, axisGroup: "primary", extra: 1 },
        { sheetName: "", chartName: "C1", seriesIndex: 1, axisGroup: "primary" },
        { sheetName: undefined, chartName: "C1", seriesIndex: 1, axisGroup: "primary" },
        { sheetName: "Sheet1", chartName: undefined, seriesIndex: 1, axisGroup: "primary" },
        { sheetName: null, chartName: "C1", seriesIndex: 1, axisGroup: "primary" },
        { sheetName: "Sheet1", chartName: null, seriesIndex: 1, axisGroup: "primary" },
      ]) {
        const result = await executor.execute({
          name: "chart.series.axisGroup.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema additionalProperties false and moderate risk", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.axisGroup.update");
      expect(def?.riskLevel).toBe("moderate");
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
    });
  });

  it("WPS is typed unsupported with ChartSeries.axisGroup evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.series.axisGroup.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        axisGroup: "secondary",
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.series.axisGroup.update");
      expect(detail.evidence).toMatch(/ChartSeries\.axisGroup/);
    }
  });
});
