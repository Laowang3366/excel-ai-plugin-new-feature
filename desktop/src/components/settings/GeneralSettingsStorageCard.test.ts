import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { GeneralSettingsStorageCard } from "./GeneralSettingsStorageCard";

describe("GeneralSettingsStorageCard", () => {
  it("shows a bounded privacy export action and credential exclusion notice", () => {
    const html = renderToStaticMarkup(
      React.createElement(GeneralSettingsStorageCard, {
        language: "zh-CN",
        dataPath: "C:\\data",
        pathError: "",
        copied: false,
        isMigrating: false,
        isExporting: false,
        exportMessage: "",
        onOpenDataPath: () => undefined,
        onCopyDataPath: () => undefined,
        onChangeDataPath: () => undefined,
        onExportUserData: () => undefined,
      }),
    );

    expect(html).toContain("导出本地数据");
    expect(html).toContain("不会导出 API Key");
    expect(html).toContain("不会复制数据目录外的原始文档");
  });
});
