import { describe, expect, it } from "vitest";

import { buildOcrToolResult } from "./ocrExecutorResult";

describe("buildOcrToolResult", () => {
  it("reports per-document text and row truncation in the top-level result", () => {
    const result = buildOcrToolResult({
      fallbacks: [{ provider: "local", success: true }],
      filePaths: ["C:\\docs\\report.xlsx"],
      maxTableRows: 2,
      maxTextChars: 10,
      mode: "style",
      selected: [
        {
          document: {
            filename: "report.xlsx",
            text: "1234567890more text",
            rows: [
              ["A", "1"],
              ["B", "2"],
              ["C", "3"],
            ],
            provider: "local",
            sourceType: "xlsx",
            warnings: [],
          },
          provider: "local",
        },
      ],
      warnings: ["existing warning", "existing warning"],
    });

    expect(result).toMatchObject({
      success: true,
      data: {
        provider: "local",
        textTruncated: true,
        rowsTruncated: true,
        rows: [
          ["A", "1"],
          ["B", "2"],
        ],
        warnings: expect.arrayContaining([
          "existing warning",
          expect.stringContaining("本地免费兜底解析"),
          expect.stringContaining("office.action.inspect"),
        ]),
        nextTools: expect.arrayContaining([
          { tool: "office.action.inspect", useWhen: expect.any(String) },
          { tool: "office.action.validate", useWhen: expect.any(String) },
          { tool: "file.getPaths", useWhen: expect.any(String) },
        ]),
      },
    });
  });
});
