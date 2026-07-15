import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function definition(name: string) {
  const value = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!value) throw new Error(`missing tool definition: ${name}`);
  return value;
}

describe("Excel chart operation parameter schemas", () => {
  it("accepts the Worker-supported formatChart parameter structure", () => {
    const args = {
      app: "excel",
      action: "style",
      operation: "formatChart",
      filePath: "C:/book.xlsx",
      target: "range:Sheet1!A1:D20",
      params: {
        host: "excel",
        chartIndex: 1,
        chartType: "linemarkers",
        showTitle: true,
        title: "Quarterly revenue",
        showLegend: true,
        width: 640,
        height: 360,
        replaceSeries: true,
        series: [
          {
            command: "add",
            name: "Revenue",
            values: "Sheet1!B2:B5",
            categories: ["Q1", "Q2", "Q3", "Q4"],
            axisGroup: "primary",
            dataLabels: { enabled: true, showValue: true },
          },
        ],
        axes: [
          {
            kind: "value",
            group: "primary",
            title: "CNY",
            minimum: 0,
            majorUnit: 100,
            numberFormat: "0.00",
          },
        ],
      },
    };

    expect(
      parseAndValidateToolArguments(
        JSON.stringify(args),
        definition("office.action.apply").parameters,
      ).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ steps: [args] }),
        definition("office.workflow.run").parameters,
      ).error,
    ).toBeUndefined();
  });

  it("rejects unknown or malformed nested formatChart parameters", () => {
    const base = {
      app: "excel",
      action: "style",
      operation: "formatChart",
      filePath: "C:/book.xlsx",
    };
    const apply = definition("office.action.apply").parameters;

    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { chartType: "surface" } }),
        apply,
      ).error,
    ).toContain("chartType");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...base,
          params: { series: [{ command: "add", shellCommand: "whoami" }] },
        }),
        apply,
      ).error,
    ).toContain("shellCommand");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...base,
          params: { axes: [{ kind: "value", dataLabels: true }] },
        }),
        apply,
      ).error,
    ).toContain("dataLabels");
  });

  it("uses strict inspectCharts params for inspect and validate tools", () => {
    const valid = {
      app: "excel",
      operation: "inspectCharts",
      filePath: "C:/book.xlsx",
      params: { host: "wps", chartName: "RevenueChart" },
    };

    for (const name of ["office.action.inspect", "office.action.validate"]) {
      const schema = definition(name).parameters;
      expect(parseAndValidateToolArguments(JSON.stringify(valid), schema).error).toBeUndefined();
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ ...valid, params: { ...valid.params, unknown: true } }),
          schema,
        ).error,
      ).toContain("unknown");
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ ...valid, params: { ...valid.params, host: "powerpoint" } }),
          schema,
        ).error,
      ).toContain("host");
    }
  });
});
