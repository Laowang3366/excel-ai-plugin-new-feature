import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";

import { getProviderModelSelectorKind, ProviderModelSelector } from "./ProviderModelSelector";

describe("ProviderModelSelector", () => {
  it("chooses the selector kind from provider model capabilities", () => {
    expect(getProviderModelSelectorKind({ isAggregation: true, modelOptions: [] })).toBe(
      "aggregation",
    );
    expect(getProviderModelSelectorKind({ isAggregation: false, modelOptions: ["gpt-5"] })).toBe(
      "select",
    );
    expect(getProviderModelSelectorKind({ isAggregation: false, modelOptions: [] })).toBe("input");
  });

  it("renders aggregation model configs with an optional empty choice", () => {
    const html = renderToStaticMarkup(
      React.createElement(ProviderModelSelector, {
        value: "",
        onChange: vi.fn(),
        isAggregation: true,
        modelConfigs: [{ name: "deepseek-r1" }, { name: "qwen-max" }],
        noModelLabel: "No model",
        showEmptyOption: true,
      }),
    );

    expect(html).toContain("-- No model --");
    expect(html).toContain("deepseek-r1");
    expect(html).toContain("qwen-max");
  });

  it("preserves the current model when it is not in preset options", () => {
    const html = renderToStaticMarkup(
      React.createElement(ProviderModelSelector, {
        value: "legacy-model",
        onChange: vi.fn(),
        isAggregation: false,
        modelOptions: ["gpt-5"],
        noModelLabel: "No model",
        preserveCurrentValue: true,
      }),
    );

    expect(html).toContain("gpt-5");
    expect(html).toContain("legacy-model");
  });
});
