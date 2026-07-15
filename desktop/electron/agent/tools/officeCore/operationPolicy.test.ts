import { describe, expect, it } from "vitest";
import { officeAdvancedOperationError } from "./operationPolicy";

describe("officeAdvancedOperationError", () => {
  const powerQueryBase = {
    app: "excel" as const,
    operation: "createPowerQuery",
    params: {
      advancedIntent: "refreshable-etl",
      sourceKind: "external",
      name: "SalesImport",
      mFormula: "let Source = Csv.Document(File.Contents(\"C:/sales.csv\")) in Source",
      loadMode: "worksheet",
      destination: "QueryOutput!A1",
    },
  };

  it.each([
    ["name", { ...powerQueryBase.params, name: " " }, "params.name"],
    ["source kind", { ...powerQueryBase.params, sourceKind: "worksheet" }, "sourceKind"],
    ["M formula", { ...powerQueryBase.params, mFormula: "" }, "mFormula"],
    ["load mode", { ...powerQueryBase.params, loadMode: "" }, "loadMode"],
    ["worksheet destination", { ...powerQueryBase.params, destination: "" }, "destination"],
  ])("rejects Power Query creation without a valid %s", (_label, params, expected) => {
    expect(officeAdvancedOperationError({ ...powerQueryBase, params })).toContain(expected);
  });

  it("allows Power Query lifecycle commands without create-only fields", () => {
    expect(officeAdvancedOperationError({
      app: "excel",
      operation: "managePowerQuery",
      params: { advancedIntent: "refreshable-etl", command: "refresh", name: "SalesImport" },
    })).toBeUndefined();
  });

  it.each([
    [undefined, { rowFields: ["Department"] }, "range:"],
    ["range:", { rowFields: ["Department"] }, "range:"],
    ["range:Sheet1!A1:B10", { rowFields: [" "], dataFields: [{ name: "" }] }, "至少需要一个"],
  ])("rejects an invalid pivot source or empty field declaration", (target, params, expected) => {
    expect(officeAdvancedOperationError({
      app: "excel",
      operation: "createPivotTable",
      target,
      params: { advancedIntent: "interactive-pivot", ...params },
    })).toContain(expected);
  });

  it("accepts a named data field object for a pivot table", () => {
    expect(officeAdvancedOperationError({
      app: "excel",
      operation: "createPivotTable",
      target: "range:Sheet1!A1:B10",
      params: { advancedIntent: "interactive-pivot", dataFields: [{ name: "Amount", function: "sum" }] },
    })).toBeUndefined();
  });

  it.each([
    [{ advancedIntent: "interactive-pivot", field: "Department" }, "pivotName"],
    [{ advancedIntent: "interactive-pivot", pivotName: "SalesPivot" }, "params.field"],
  ])("requires the slicer pivot and field", (params, expected) => {
    expect(officeAdvancedOperationError({ app: "excel", operation: "addSlicer", params })).toContain(expected);
  });

  it("does not apply Excel boundaries to other Office apps", () => {
    expect(officeAdvancedOperationError({ app: "word", operation: "createPowerQuery", params: {} })).toBeUndefined();
  });
});
