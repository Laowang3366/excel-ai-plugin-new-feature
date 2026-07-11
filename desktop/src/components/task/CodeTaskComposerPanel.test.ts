import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";

import { CodeTaskComposerPanel } from "./CodeTaskComposerPanel";

describe("CodeTaskComposerPanel", () => {
  it("stacks environment and language controls when embedded in the feature sidebar", () => {
    const html = renderToStaticMarkup(
      React.createElement(CodeTaskComposerPanel, {
        embedded: true,
        onSubmit: vi.fn(),
        onClose: vi.fn(),
      }),
    );

    expect(html).toContain('class="task-field-row task-field-row--stacked"');
    expect(html).toContain("Microsoft Excel");
    expect(html).toContain("Python");
  });
});
