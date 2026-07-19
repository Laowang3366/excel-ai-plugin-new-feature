import { describe, expect, it } from "vitest";
import { ToolExecutor, TOOL_DEFINITIONS } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";

describe("phase3 tools", () => {
  it("registers format/table/chart/workbook tools", () => {
    const names = TOOL_DEFINITIONS.map((tool) => tool.name);
    for (const name of [
      "range.format.read",
      "range.format.write",
      "table.list",
      "table.create",
      "table.delete",
      "chart.list",
      "chart.create",
      "chart.delete",
      "workbook.inspect",
    ]) {
      expect(names).toContain(name);
    }
  });

  it("writes format and inspects workbook via host", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const written = await executor.execute({
      name: "range.format.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        format: { fontBold: true, numberFormat: "0.00", wrapText: true },
      },
    });
    expect(written.ok).toBe(true);

    const inspect = await executor.execute({ name: "workbook.inspect", arguments: {} });
    expect(inspect.ok).toBe(true);
    if (inspect.ok) {
      expect(inspect.data).toMatchObject({ workbookName: "Book1.xlsx" });
    }
  });

  it("creates and deletes table/chart", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const table = await executor.execute({
      name: "table.create",
      arguments: {
        sheetName: "Sheet1",
        range: "A1:B3",
        name: "Sales",
        hasHeaders: true,
      },
    });
    expect(table.ok).toBe(true);

    const chart = await executor.execute({
      name: "chart.create",
      arguments: {
        sheetName: "Sheet1",
        sourceRange: "A1:B3",
        chartType: "line",
        name: "Trend",
        title: "Sales",
      },
    });
    expect(chart.ok).toBe(true);

    const listed = await executor.execute({ name: "chart.list", arguments: {} });
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      expect((listed.data as { name: string }[]).some((item) => item.name === "Trend")).toBe(
        true,
      );
    }

    expect(
      (
        await executor.execute({
          name: "table.delete",
          arguments: { sheetName: "Sheet1", tableName: "Sales" },
        })
      ).ok,
    ).toBe(true);
    expect(
      (
        await executor.execute({
          name: "chart.delete",
          arguments: { sheetName: "Sheet1", chartName: "Trend" },
        })
      ).ok,
    ).toBe(true);
  });

  it("rejects invalid chartType and non-boolean hasHeaders", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const badChart = await executor.execute({
      name: "chart.create",
      arguments: {
        sheetName: "Sheet1",
        sourceRange: "A1:B2",
        chartType: "stock",
      },
    });
    expect(badChart.ok).toBe(false);
    if (!badChart.ok) expect(badChart.error).toContain("chartType");

    const badHeaders = await executor.execute({
      name: "table.create",
      arguments: {
        sheetName: "Sheet1",
        range: "A1:B2",
        hasHeaders: "yes",
      },
    });
    expect(badHeaders.ok).toBe(false);
    if (!badHeaders.ok) expect(badHeaders.error).toContain("hasHeaders");
  });
});
