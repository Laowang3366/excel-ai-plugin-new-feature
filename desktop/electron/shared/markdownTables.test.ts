import { describe, expect, it } from "vitest";

import { extractMarkdownTables } from "./markdownTables";

describe("markdownTables", () => {
  it("extracts github-style markdown table rows", () => {
    const markdown = [
      "invoice details",
      "",
      "| field | value |",
      "| --- | ---: |",
      "| invoice no | 001 |",
      "| total | 100.00 |",
    ].join("\n");

    expect(extractMarkdownTables(markdown)).toEqual([
      ["field", "value"],
      ["invoice no", "001"],
      ["total", "100.00"],
    ]);
  });
});
