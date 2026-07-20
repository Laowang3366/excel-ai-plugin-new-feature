import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor, TOOL_DEFINITIONS } from "../shared/tools";
import { installObjectUpdateExcel } from "./fakes/officeJsObjectUpdateFake";
import { MockHostAdapter } from "./mockHost";

function installOfficeRequirements(
  isSetSupported: (name: string, minVersion?: string) => boolean = () => true,
): void {
  (globalThis as unknown as { Office: unknown }).Office = {
    context: { requirements: { isSetSupported } },
  };
}

describe("phase7 object update", () => {
  it("registers table.update and chart.update", () => {
    const names = TOOL_DEFINITIONS.map((tool) => tool.name);
    expect(names).toContain("table.update");
    expect(names).toContain("chart.update");
    const tableUpdate = TOOL_DEFINITIONS.find((tool) => tool.name === "table.update");
    const properties = (tableUpdate?.parameters as { properties?: Record<string, unknown> })
      .properties;
    expect(Object.keys(properties ?? {})).toEqual(
      expect.arrayContaining(["resizeAddress", "showBandedRows", "showBandedColumns"]),
    );
  });

  describe("Office.js", () => {
    beforeEach(() => {
      installObjectUpdateExcel();
      installOfficeRequirements();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("updates table range/name/style/toggles and reads back", async () => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateTable({
        sheetName: "Sheet1",
        tableName: "T1",
        resizeAddress: "Sheet1!a1:d5",
        newName: "SalesTable",
        style: "TableStyleMedium9",
        showHeaders: false,
        showTotals: true,
        showFilterButton: false,
        showBandedRows: false,
        showBandedColumns: true,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.name).toBe("SalesTable");
        expect(updated.data.address).toBe("Sheet1!A1:D5");
        expect(updated.data.style).toBe("TableStyleMedium9");
        expect(updated.data.hasHeaders).toBe(false);
        expect(updated.data.showTotals).toBe(true);
        expect(updated.data.showFilter).toBe(false);
        expect(updated.data.showBandedRows).toBe(false);
        expect(updated.data.showBandedColumns).toBe(true);
      }
    });

    it("rejects a cross-sheet table resize as an ordinary host failure", async () => {
      const result = await new OfficeJsAdapter().updateTable({
        sheetName: "Sheet1",
        tableName: "T1",
        resizeAddress: "Other!A1:D5",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(false);
        expect(result.reason).toMatch(/resizeAddress.*same worksheet.*table/);
        expect(result.reason).not.toMatch(/sourceRange|chart/);
      }
    });

    it("classifies unsupported table API sets before running the update", async () => {
      for (const [input, version] of [
        [{ resizeAddress: "A1:D5" }, "1.13"],
        [{ showBandedRows: false }, "1.3"],
      ] as const) {
        installOfficeRequirements((_name, minVersion) => minVersion !== version);
        const result = await new OfficeJsAdapter().updateTable({
          sheetName: "Sheet1",
          tableName: "T1",
          ...input,
        });
        expect(result.ok).toBe(false);
        if (!result.ok) {
          expect(result.unsupported).toBe(true);
          expect(result.reason).toContain(version);
        }
      }
    });

    it("keeps shallow table updates working without a higher API precheck", async () => {
      delete (globalThis as { Office?: unknown }).Office;
      const result = await new OfficeJsAdapter().updateTable({
        sheetName: "Sheet1",
        tableName: "T1",
        style: "TableStyleLight1",
      });
      expect(result.ok).toBe(true);
    });

    it("returns an ordinary failure when the host rejects the resize geometry", async () => {
      const result = await new OfficeJsAdapter().updateTable({
        sheetName: "Sheet1",
        tableName: "T1",
        resizeAddress: "A2:D5",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(false);
        expect(result.reason).toMatch(/header row.*same row/);
      }
    });

    it("updates chart rename/type/title/position/size and reads back", async () => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChart({
        sheetName: "Sheet1",
        chartName: "C1",
        newName: "Trend",
        chartType: "line",
        title: "Revenue",
        left: 40,
        top: 50,
        width: 400,
        height: 250,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.name).toBe("Trend");
        expect(updated.data.chartType).toBe("line");
        expect(updated.data.title).toBe("Revenue");
        expect(updated.data.titleVisible).toBe(true);
        expect(updated.data.left).toBe(40);
        expect(updated.data.width).toBe(400);
      }
      const hide = await adapter.updateChart({
        sheetName: "Sheet1",
        chartName: "Trend",
        showTitle: false,
      });
      expect(hide.ok).toBe(true);
      if (hide.ok) expect(hide.data.titleVisible).toBe(false);

      const clearTitle = await adapter.updateChart({
        sheetName: "Sheet1",
        chartName: "Trend",
        title: "",
      });
      expect(clearTitle.ok).toBe(true);
      if (clearTitle.ok) {
        expect(clearTitle.data.title).toBe("");
        expect(clearTitle.data.titleVisible).toBe(true);
      }

      const clearAndHide = await adapter.updateChart({
        sheetName: "Sheet1",
        chartName: "Trend",
        title: "",
        showTitle: false,
      });
      expect(clearAndHide.ok).toBe(true);
      if (clearAndHide.ok) {
        expect(clearAndHide.data.title).toBe("");
        expect(clearAndHide.data.titleVisible).toBe(false);
      }
    });
  });

  it("executor success and validation branches", async () => {
    const host = new MockHostAdapter();
    await host.createTable({ sheetName: "Sheet1", address: "A1:C2", name: "T1" });
    await host.createChart({
      sheetName: "Sheet1",
      sourceRange: "A1:B2",
      name: "C1",
    });
    const executor = new ToolExecutor(host);

    const tableUpdated = await executor.execute({
      name: "table.update",
      arguments: {
        sheetName: "Sheet1",
        tableName: "T1",
        resizeAddress: "A1:D4",
        showBandedRows: false,
        showBandedColumns: true,
      },
    });
    expect(tableUpdated.ok).toBe(true);
    if (tableUpdated.ok) {
      expect(tableUpdated.data).toMatchObject({
        address: "Sheet1!A1:D4",
        showBandedRows: false,
        showBandedColumns: true,
      });
    }

    expect(
      (
        await executor.execute({
          name: "chart.update",
          arguments: { sheetName: "Sheet1", chartName: "C1", chartType: "line" },
        })
      ).ok,
    ).toBe(true);

    expect(
      (
        await executor.execute({
          name: "table.update",
          arguments: { sheetName: "Sheet1", tableName: "T1" },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "table.update",
          arguments: {
            sheetName: "Sheet1",
            tableName: "T1",
            resizeAddress: null,
            showTotals: true,
          },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "table.update",
          arguments: {
            sheetName: "Sheet1",
            tableName: "T1",
            resizeAddress: "Other!A1:D4",
          },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "table.update",
          arguments: {
            sheetName: "Sheet1",
            tableName: "T1",
            resizeAddress: "A0:D4",
          },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "table.update",
          arguments: {
            sheetName: "Sheet1",
            tableName: "T1",
            showBandedRows: "yes",
          },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "table.update",
          arguments: {
            sheetName: "Sheet1",
            tableName: "T1",
            showTotals: true,
            unknown: true,
          },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "table.update",
          arguments: { sheetName: "Sheet1", tableName: "T1", newName: "   " },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "chart.update",
          arguments: { sheetName: "Sheet1", chartName: "C1", chartType: "stock" },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "chart.update",
          arguments: { sheetName: "Sheet1", chartName: "C1", width: -1 },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "chart.update",
          arguments: { sheetName: "Sheet1", chartName: "C1", left: Number.NaN },
        })
      ).ok,
    ).toBe(false);

    const emptyTitle = await executor.execute({
      name: "chart.update",
      arguments: { sheetName: "Sheet1", chartName: "C1", title: "" },
    });
    expect(emptyTitle.ok).toBe(true);
    if (emptyTitle.ok) {
      expect((emptyTitle.data as { title?: string; titleVisible?: boolean }).title).toBe("");
      expect((emptyTitle.data as { titleVisible?: boolean }).titleVisible).toBe(true);
    }

    const emptyHidden = await executor.execute({
      name: "chart.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        title: "",
        showTitle: false,
      },
    });
    expect(emptyHidden.ok).toBe(true);
    if (emptyHidden.ok) {
      expect((emptyHidden.data as { titleVisible?: boolean }).titleVisible).toBe(false);
    }
  });

  it("WPS returns unsupported for table.update and chart.update", async () => {
    const executor = new ToolExecutor(new WpsJsaAdapter());
    for (const call of [
      {
        name: "table.update" as const,
        arguments: {
          sheetName: "Sheet1",
          tableName: "T1",
          resizeAddress: "A1:D4",
          showBandedRows: false,
        },
      },
      {
        name: "chart.update" as const,
        arguments: { sheetName: "Sheet1", chartName: "C1", chartType: "line" },
      },
    ]) {
      const result = await executor.execute(call);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    }
  });
});
