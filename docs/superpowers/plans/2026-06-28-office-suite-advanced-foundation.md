# Office Suite Advanced Foundation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 建立 Excel、Word、PowerPoint 三件套统一高级操作基础层，并接入文档生产增强包的首批高级能力。

**Architecture:** 新增 `officeCore` 作为跨三件套 action、locator、capability、result、validation 协议层；`tools/registry` 和 `tools/executors` 只暴露并路由统一工具；具体执行优先落到 `officeOpenXml`，COM 只在明确需要当前窗口或 Open XML 返回 `needsCom` 时使用。

**Tech Stack:** TypeScript、Vitest、Electron main process、JSZip、现有 Excel/Word/PPT COM bridge、现有 Office Open XML engine。

---

## 执行规范

- 一次只执行一个阶段。
- 禁止过度设计：只实现本阶段测试覆盖的协议和能力。
- 禁止过度兜底：Open XML 失败或不支持时返回 `unsupported` / `needsCom`，不隐藏调用 COM。
- 禁止过度约束边界导致既有 Excel/Word/PPT 工具异常。
- 代码风格保持与 `desktop/electron/agent` 当前模块一致。
- 注释必须说明模块作用和关联模块，不写空泛注释。
- 遵循单一职责，但不要单纯把代码拆碎。
- 每个阶段完成后运行聚焦测试、review，通过后提交。
- 所有非必须、非必要的测试临时文件验证后必须清理。
- 不触碰执行前已有无关脏文件：`screenshot.png`、`desktop/dev-server.err.log`、`desktop/dev-server.log`。

## 计划文件结构

- 新建 `desktop/electron/agent/tools/officeCore/types.ts`  
  作用：定义 `OfficeAction`、locator、capability、result、validation 的基础类型。

- 新建 `desktop/electron/agent/tools/officeCore/locator.ts`  
  作用：解析 `sheet:Sheet1`、`range:Sheet1!A1:D10`、`slide:1`、`table:1` 等定位字符串。

- 新建 `desktop/electron/agent/tools/officeCore/results.ts`  
  作用：生成统一 `done / unsupported / needsCom / failed` 结果。

- 新建 `desktop/electron/agent/tools/officeCore/capabilities.ts`  
  作用：声明首批 Excel/Word/PPT 高级操作能力及首选 engine。

- 新建 `desktop/electron/agent/tools/officeCore/officeActionAdapter.ts`  
  作用：把统一 action 路由到 Open XML file bridge 或 COM bridge。

- 新建 `desktop/electron/agent/tools/officeCore/*.test.ts`  
  作用：覆盖 locator、result、capability、adapter 路由。

- 修改 `desktop/electron/agent/tools/contracts/office.ts`  
  作用：增加 `OfficeActionBridge` 契约。

- 修改 `desktop/electron/agent/tools/registry/office.ts`  
  作用：新增 `office.action.inspect`、`office.action.apply`、`office.action.validate`。

- 修改 `desktop/electron/agent/tools/executors/officeExecutors.ts`  
  作用：新增统一 action executor，仅校验参数和转发。

- 修改 `desktop/electron/agent/tools/registry/officeTools.test.ts`  
  作用：锁定工具注册和 executor 路由。

- 修改 `desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts`  
  作用：适配统一 action 的 Open XML 优先路径。

- 新建 `desktop/electron/agent/tools/implementations/officeOpenXml/advancedExcel.ts`  
  作用：实现 Excel 首批高级能力的 Open XML 文件级操作。

- 新建 `desktop/electron/agent/tools/implementations/officeOpenXml/advancedWord.ts`  
  作用：实现 Word 首批高级能力的 Open XML 文件级操作。

- 新建 `desktop/electron/agent/tools/implementations/officeOpenXml/advancedPresentation.ts`  
  作用：实现 PPT 首批高级能力的 Open XML 文件级操作。

- 修改 `desktop/electron/agent/prompts/sections/officeToolsPrompt.ts`  
  作用：引导模型优先使用统一 action 工具和 Open XML 路线。

- 修改 `desktop/src/utils/officeEditEvents.ts`  
  作用：侧边栏识别统一 action 结果状态和 validation。

---

### 任务 1：建立 `officeCore` 协议层

**文件：**
- 新建：`desktop/electron/agent/tools/officeCore/types.ts`
- 新建：`desktop/electron/agent/tools/officeCore/locator.ts`
- 新建：`desktop/electron/agent/tools/officeCore/results.ts`
- 新建：`desktop/electron/agent/tools/officeCore/capabilities.ts`
- 测试：`desktop/electron/agent/tools/officeCore/locator.test.ts`
- 测试：`desktop/electron/agent/tools/officeCore/results.test.ts`
- 测试：`desktop/electron/agent/tools/officeCore/capabilities.test.ts`

- [ ] **步骤 1：编写 locator 红灯测试**

创建 `desktop/electron/agent/tools/officeCore/locator.test.ts`：

```ts
import { describe, expect, it } from "vitest";
import { parseOfficeLocator } from "./locator";

describe("parseOfficeLocator", () => {
  it("parses app-neutral Office locators", () => {
    expect(parseOfficeLocator("range:Sheet1!A1:D10")).toEqual({
      kind: "range",
      value: "Sheet1!A1:D10",
      sheetName: "Sheet1",
      address: "A1:D10",
    });
    expect(parseOfficeLocator("slide:3")).toEqual({ kind: "slide", value: "3", index: 3 });
    expect(parseOfficeLocator("table:1")).toEqual({ kind: "table", value: "1", index: 1 });
  });
});
```

运行：

```powershell
npm test -- electron/agent/tools/officeCore/locator.test.ts
```

预期：失败，提示 `./locator` 不存在。

- [ ] **步骤 2：实现基础类型和 locator**

创建 `types.ts`，包含：

```ts
export type OfficeActionApp = "excel" | "word" | "presentation";
export type OfficeActionKind = "inspect" | "edit" | "style" | "insert" | "snapshot" | "validate";
export type OfficeActionStatus = "done" | "unsupported" | "needsCom" | "failed";
export type OfficeActionEngine = "openxml" | "com";

export interface OfficeLocator {
  kind: string;
  value: string;
  sheetName?: string;
  address?: string;
  index?: number;
}

export interface OfficeActionInput {
  app: OfficeActionApp;
  action: OfficeActionKind;
  operation: string;
  filePath?: string;
  outputPath?: string;
  target?: string;
  preferEngine?: OfficeActionEngine;
  params?: Record<string, unknown>;
}

export interface OfficeActionValidation {
  ok: boolean;
  checks: Array<{ name: string; ok: boolean; message: string }>;
}

export interface OfficeActionResult {
  status: OfficeActionStatus;
  engine: OfficeActionEngine;
  app: OfficeActionApp;
  action: OfficeActionKind;
  operation: string;
  filePath?: string;
  outputPath?: string;
  target?: string;
  summary: string;
  changes: Array<{ kind: string; target?: string; detail: string }>;
  validation?: OfficeActionValidation;
  error?: string;
  data?: unknown;
}
```

创建 `locator.ts`，实现 `parseOfficeLocator(locator: string): OfficeLocator`。只处理 `range`、`sheet`、`slide`、`table`、`chart`、`shape`、`pictureSlot`、`header`、`footer`，未知 kind 返回 `{ kind, value }`。

- [ ] **步骤 3：运行 locator 测试**

```powershell
npm test -- electron/agent/tools/officeCore/locator.test.ts
```

预期：通过。

- [ ] **步骤 4：编写 result 红灯测试**

创建 `desktop/electron/agent/tools/officeCore/results.test.ts`：

```ts
import { describe, expect, it } from "vitest";
import { doneResult, needsComResult, unsupportedResult } from "./results";

describe("office action results", () => {
  it("builds consistent status results", () => {
    expect(doneResult({
      engine: "openxml",
      app: "word",
      action: "style",
      operation: "styleTables",
      summary: "已美化 Word 表格",
    }).status).toBe("done");

    expect(needsComResult({
      app: "word",
      action: "insert",
      operation: "insertOrUpdateToc",
      summary: "目录字段需要 Word 刷新",
    }).engine).toBe("openxml");

    expect(unsupportedResult({
      app: "presentation",
      action: "insert",
      operation: "editAnimationTimeline",
      summary: "首阶段不支持动画时间轴",
    }).status).toBe("unsupported");
  });
});
```

运行：

```powershell
npm test -- electron/agent/tools/officeCore/results.test.ts
```

预期：失败，提示 `./results` 不存在。

- [ ] **步骤 5：实现 result helpers**

创建 `results.ts`，实现：

```ts
import type { OfficeActionInput, OfficeActionResult } from "./types";

type ResultBase = Pick<OfficeActionResult, "engine" | "app" | "action" | "operation" | "summary"> &
  Partial<Pick<OfficeActionResult, "filePath" | "outputPath" | "target" | "changes" | "validation" | "error" | "data">>;

export function doneResult(input: ResultBase): OfficeActionResult {
  return { status: "done", changes: [], ...input };
}

export function unsupportedResult(input: Omit<ResultBase, "engine">): OfficeActionResult {
  return { status: "unsupported", engine: "openxml", changes: [], ...input };
}

export function needsComResult(input: Omit<ResultBase, "engine">): OfficeActionResult {
  return { status: "needsCom", engine: "openxml", changes: [], ...input };
}

export function failedResult(action: OfficeActionInput, error: unknown): OfficeActionResult {
  return {
    status: "failed",
    engine: action.preferEngine || "openxml",
    app: action.app,
    action: action.action,
    operation: action.operation,
    filePath: action.filePath,
    outputPath: action.outputPath,
    target: action.target,
    summary: "Office action 执行失败",
    changes: [],
    error: error instanceof Error ? error.message : String(error),
  };
}
```

- [ ] **步骤 6：编写 capability 红灯测试**

创建 `desktop/electron/agent/tools/officeCore/capabilities.test.ts`：

```ts
import { describe, expect, it } from "vitest";
import { findOfficeCapability } from "./capabilities";

describe("office capabilities", () => {
  it("declares first-stage document production capabilities", () => {
    expect(findOfficeCapability("excel", "insertChart")?.preferredEngine).toBe("openxml");
    expect(findOfficeCapability("word", "insertOrUpdateToc")?.fallback).toBe("needsCom");
    expect(findOfficeCapability("presentation", "replacePictureSlot")?.writesFile).toBe(true);
  });
});
```

运行：

```powershell
npm test -- electron/agent/tools/officeCore/capabilities.test.ts
```

预期：失败，提示 `./capabilities` 不存在。

- [ ] **步骤 7：实现 capability 声明**

创建 `capabilities.ts`，定义：

```ts
import type { OfficeActionApp, OfficeActionEngine } from "./types";

export interface OfficeCapability {
  app: OfficeActionApp;
  operation: string;
  preferredEngine: OfficeActionEngine;
  writesFile: boolean;
  fallback: "none" | "needsCom";
}

export const OFFICE_CAPABILITIES: OfficeCapability[] = [
  { app: "excel", operation: "insertChart", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "applyConditionalFormatting", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "setDataValidation", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "excel", operation: "styleTable", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "word", operation: "applyHeadingStyles", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "insertOrUpdateToc", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "styleTables", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "word", operation: "insertOrReplaceImage", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "word", operation: "setHeaderFooter", preferredEngine: "openxml", writesFile: true, fallback: "none" },
  { app: "presentation", operation: "applyTheme", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "normalizeLayouts", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "insertChart", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "replacePictureSlot", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
  { app: "presentation", operation: "alignShapes", preferredEngine: "openxml", writesFile: true, fallback: "needsCom" },
];

export function findOfficeCapability(app: OfficeActionApp, operation: string): OfficeCapability | undefined {
  return OFFICE_CAPABILITIES.find((capability) => capability.app === app && capability.operation === operation);
}
```

- [ ] **步骤 8：运行任务 1 测试并提交**

运行：

```powershell
npm test -- electron/agent/tools/officeCore/locator.test.ts electron/agent/tools/officeCore/results.test.ts electron/agent/tools/officeCore/capabilities.test.ts
npm run typecheck
git diff --check
```

预期：全部通过。

提交：

```powershell
git add desktop/electron/agent/tools/officeCore
git commit -m "feat: add office action core"
```

---

### 任务 2：注册统一 action 工具并接入 executor

**文件：**
- 修改：`desktop/electron/agent/tools/contracts/office.ts`
- 修改：`desktop/electron/agent/tools/registry/office.ts`
- 修改：`desktop/electron/agent/tools/executors/officeExecutors.ts`
- 修改：`desktop/electron/agent/tools/registry/officeTools.test.ts`

- [ ] **步骤 1：编写工具注册红灯测试**

在 `officeTools.test.ts` 的完整工具目录断言中加入：

```ts
"office.action.inspect",
"office.action.apply",
"office.action.validate",
```

并在 `arrayContaining` 断言中加入这三个工具名。

运行：

```powershell
npm test -- electron/agent/tools/registry/officeTools.test.ts
```

预期：失败，提示工具目录缺少三个新工具。

- [ ] **步骤 2：增加契约**

在 `desktop/electron/agent/tools/contracts/office.ts` 增加：

```ts
import type { OfficeActionInput, OfficeActionResult } from "../officeCore/types";

export interface OfficeActionBridge {
  executeAction(input: OfficeActionInput): Promise<OfficeActionResult>;
}
```

如果现有 import 风格需要保持 type-only，将 import 放在文件顶部并只引用类型。

- [ ] **步骤 3：新增工具定义**

在 `office.ts` 新增三个 `ToolDefinition`：

```ts
const OFFICE_ACTION_APPLY_DEF: ToolDefinition = {
  name: "office.action.apply",
  description: "统一 Office 高级操作入口。Open XML 优先执行 Excel/Word/PPT 文档生产增强动作；COM 仅在返回 needsCom 后作为显式兜底",
  parameters: {
    type: "object",
    properties: {
      app: { type: "string", enum: ["excel", "word", "presentation"], description: "目标应用类型" },
      action: { type: "string", enum: ["inspect", "edit", "style", "insert", "snapshot", "validate"], description: "动作类型" },
      operation: { type: "string", description: "具体操作，如 insertChart、styleTables、replacePictureSlot" },
      filePath: { type: "string", description: "Office 文件绝对路径" },
      outputPath: { type: "string", description: "输出文件路径；未指定时由实现生成副本" },
      target: { type: "string", description: "对象定位，如 range:Sheet1!A1:D10、table:1、slide:1" },
      preferEngine: { type: "string", enum: ["openxml", "com"], description: "首选引擎，默认 openxml" },
      params: { type: "object", description: "操作参数" },
    },
    required: ["app", "action", "operation"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};
```

`office.action.inspect` 使用相同 schema，但 `riskLevel: "safe"`、`requiresApproval: false`。  
`office.action.validate` 使用相同 schema，但 `riskLevel: "safe"`、`requiresApproval: false`。

- [ ] **步骤 4：编写 executor 红灯测试**

在 `officeTools.test.ts` 的 executor 测试中增加 mock：

```ts
const officeActionBridge: OfficeActionBridge = {
  executeAction: vi.fn(async (input) => ({
    status: "done",
    engine: "openxml",
    app: input.app,
    action: input.action,
    operation: input.operation,
    summary: "ok",
    changes: [],
  })),
};
```

调用 `createToolExecutors(..., officeFileBridge, officeActionBridge)` 之前，需要先在 `createToolExecutors` 签名中规划新增参数；测试会先失败。

断言：

```ts
expect([...executors.keys()]).toEqual(expect.arrayContaining([
  "office.action.inspect",
  "office.action.apply",
  "office.action.validate",
]));
```

并验证：

```ts
const result = await executors.get("office.action.apply")!.execute({
  app: "excel",
  action: "insert",
  operation: "insertChart",
  filePath: "D:\\docs\\book.xlsx",
  target: "range:Sheet1!A1:B5",
  params: { chartType: "column" },
});

expect(result.success).toBe(true);
expect(officeActionBridge.executeAction).toHaveBeenCalledWith({
  app: "excel",
  action: "insert",
  operation: "insertChart",
  filePath: "D:\\docs\\book.xlsx",
  target: "range:Sheet1!A1:B5",
  params: { chartType: "column" },
});
```

- [ ] **步骤 5：实现 executor 路由**

修改 `createToolExecutors.ts`，新增可选参数 `officeActionBridge?: OfficeActionBridge`，并传入 `addOfficeExecutors`。

修改 `OfficeExecutorDeps`：

```ts
officeActionBridge?: OfficeActionBridge;
```

在 `addOfficeExecutors` 中新增 `office.action.inspect/apply/validate` 三条路由，只做必填参数校验和转发。

- [ ] **步骤 6：运行任务 2 验证并提交**

运行：

```powershell
npm test -- electron/agent/tools/registry/officeTools.test.ts
npm run typecheck
git diff --check
```

预期：全部通过。

提交：

```powershell
git add desktop/electron/agent/tools/contracts/office.ts desktop/electron/agent/tools/registry/office.ts desktop/electron/agent/tools/executors/officeExecutors.ts desktop/electron/agent/tools/executors/createToolExecutors.ts desktop/electron/agent/tools/registry/officeTools.test.ts
git commit -m "feat: register office action tools"
```

---

### 任务 3：实现统一 action adapter，复用现有 Open XML 能力

**文件：**
- 新建：`desktop/electron/agent/tools/officeCore/officeActionAdapter.ts`
- 测试：`desktop/electron/agent/tools/officeCore/officeActionAdapter.test.ts`
- 修改：`desktop/electron/agent/runtime/agentRuntime.ts`

- [ ] **步骤 1：编写 adapter 红灯测试**

创建 `officeActionAdapter.test.ts`：

```ts
import { describe, expect, it, vi } from "vitest";
import { createOfficeActionBridge } from "./officeActionAdapter";
import type { OfficeFileBridge } from "../contracts/office";

describe("createOfficeActionBridge", () => {
  it("routes table style actions to the Open XML file bridge", async () => {
    const officeFileBridge: OfficeFileBridge = {
      inspectFile: vi.fn(),
      replaceText: vi.fn(),
      inspectLayout: vi.fn(),
      inspectTable: vi.fn(),
      applyTableStyle: vi.fn(async () => ({
        engine: "openxml",
        operation: "applyTableStyle",
        documentType: "spreadsheet",
        filePath: "D:\\docs\\book.xlsx",
        outputPath: "D:\\docs\\book-styled.xlsx",
        changedParts: ["xl/worksheets/sheet1.xml"],
      })),
      snapshot: vi.fn(),
    };

    const bridge = createOfficeActionBridge({ officeFileBridge });
    const result = await bridge.executeAction({
      app: "excel",
      action: "style",
      operation: "styleTable",
      filePath: "D:\\docs\\book.xlsx",
      target: "table:1",
      params: { style: "professional" },
    });

    expect(result.status).toBe("done");
    expect(result.engine).toBe("openxml");
    expect(result.validation?.ok).toBe(true);
    expect(officeFileBridge.applyTableStyle).toHaveBeenCalledWith({
      filePath: "D:\\docs\\book.xlsx",
      target: "table:1",
      style: "professional",
      outputPath: undefined,
    });
  });
});
```

运行：

```powershell
npm test -- electron/agent/tools/officeCore/officeActionAdapter.test.ts
```

预期：失败，提示 `officeActionAdapter` 不存在。

- [ ] **步骤 2：实现 adapter 首批路由**

创建 `officeActionAdapter.ts`：

- `inspect` + `operation: "layout"` → `officeFileBridge.inspectLayout`
- `inspect` + `operation: "tables"` → `officeFileBridge.inspectTable`
- `style` + `operation: "styleTable"` → `officeFileBridge.applyTableStyle`
- `snapshot` → `officeFileBridge.snapshot`
- 未支持 operation → `unsupportedResult`
- capability 存在且 `fallback: "needsCom"` → 返回 `needsComResult`

`styleTable` 的默认 style 从 `params.style` 读取；不是 `professional/compact/financial` 时默认 `professional`。

- [ ] **步骤 3：接入 runtime**

在 `agentRuntime.ts` 创建 `officeActionBridge`：

```ts
const officeActionBridge = createOfficeActionBridge({ officeFileBridge });
```

传给 `createToolExecutors` 的新增参数。

- [ ] **步骤 4：运行任务 3 验证并提交**

运行：

```powershell
npm test -- electron/agent/tools/officeCore/officeActionAdapter.test.ts electron/agent/tools/registry/officeTools.test.ts
npm run typecheck
git diff --check
```

预期：全部通过。

提交：

```powershell
git add desktop/electron/agent/tools/officeCore/officeActionAdapter.ts desktop/electron/agent/tools/officeCore/officeActionAdapter.test.ts desktop/electron/agent/runtime/agentRuntime.ts
git commit -m "feat: route office actions through openxml"
```

---

### 任务 4：接入 Excel 首批高级能力

**文件：**
- 新建：`desktop/electron/agent/tools/implementations/officeOpenXml/advancedExcel.ts`
- 测试：`desktop/electron/agent/tools/implementations/officeOpenXml/advancedExcel.test.ts`
- 修改：`desktop/electron/agent/tools/officeCore/officeActionAdapter.ts`
- 修改：`desktop/electron/agent/tools/implementations/officeOpenXml/types.ts`

- [ ] **步骤 1：编写 Excel 高级能力红灯测试**

创建 `advancedExcel.test.ts`，用最小 `.xlsx` zip 验证：

```ts
import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { applyExcelAdvancedAction } from "./advancedExcel";

async function writeZip(filePath: string, files: Record<string, string>): Promise<void> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) zip.file(name, content);
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}

async function readZipText(filePath: string, partName: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(filePath));
  const part = zip.file(partName);
  if (!part) throw new Error(`missing ${partName}`);
  return part.async("text");
}

describe("applyExcelAdvancedAction", () => {
  it("adds data validation to an Excel worksheet copy", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-excel-advanced-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      const outputPath = path.join(tempDir, "book-edited.xlsx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "xl/worksheets/sheet1.xml": "<worksheet><sheetData /></worksheet>",
      });

      const result = await applyExcelAdvancedAction({
        operation: "setDataValidation",
        filePath,
        outputPath,
        target: "range:Sheet1!A2:A10",
        params: { type: "list", values: ["通过", "失败"] },
      });

      const sheetXml = await readZipText(outputPath, "xl/worksheets/sheet1.xml");
      expect(result.status).toBe("done");
      expect(sheetXml).toContain("<dataValidations");
      expect(sheetXml).toContain("通过,失败");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
```

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/advancedExcel.test.ts
```

预期：失败，提示模块不存在。

- [ ] **步骤 2：实现 Excel 最小能力**

创建 `advancedExcel.ts`，实现：

- `setDataValidation`：写入 `<dataValidations>`。
- `applyConditionalFormatting`：写入 `<conditionalFormatting>`。
- `styleTable`：调用现有 `applyOfficeOpenXmlTableStyle`。
- `insertChart`：第一阶段返回 `needsCom`，summary 说明 Open XML 图表包生成将在后续阶段深化。

写文件默认输出 `<base>-advanced.xlsx`。

- [ ] **步骤 3：接入 adapter**

在 `officeActionAdapter.ts` 中，当 `app === "excel"` 且 operation 属于 Excel 首批能力时，调用 `applyExcelAdvancedAction`。

- [ ] **步骤 4：运行任务 4 验证并提交**

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/advancedExcel.test.ts electron/agent/tools/officeCore/officeActionAdapter.test.ts
npm run typecheck
git diff --check
```

预期：全部通过。

提交：

```powershell
git add desktop/electron/agent/tools/implementations/officeOpenXml/advancedExcel.ts desktop/electron/agent/tools/implementations/officeOpenXml/advancedExcel.test.ts desktop/electron/agent/tools/officeCore/officeActionAdapter.ts desktop/electron/agent/tools/implementations/officeOpenXml/types.ts
git commit -m "feat: add excel advanced openxml actions"
```

---

### 任务 5：接入 Word 首批高级能力

**文件：**
- 新建：`desktop/electron/agent/tools/implementations/officeOpenXml/advancedWord.ts`
- 测试：`desktop/electron/agent/tools/implementations/officeOpenXml/advancedWord.test.ts`
- 修改：`desktop/electron/agent/tools/officeCore/officeActionAdapter.ts`

- [ ] **步骤 1：编写 Word 高级能力红灯测试**

创建 `advancedWord.test.ts`，验证标题样式和页眉页脚：

```ts
import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { applyWordAdvancedAction } from "./advancedWord";

async function writeZip(filePath: string, files: Record<string, string>): Promise<void> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) zip.file(name, content);
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}

async function readZipText(filePath: string, partName: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(filePath));
  const part = zip.file(partName);
  if (!part) throw new Error(`missing ${partName}`);
  return part.async("text");
}

describe("applyWordAdvancedAction", () => {
  it("applies heading style to matching Word paragraphs", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-word-advanced-"));
    try {
      const filePath = path.join(tempDir, "report.docx");
      const outputPath = path.join(tempDir, "report-edited.docx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "word/document.xml": "<w:document><w:body><w:p><w:r><w:t>一、概览</w:t></w:r></w:p></w:body></w:document>",
      });

      const result = await applyWordAdvancedAction({
        operation: "applyHeadingStyles",
        filePath,
        outputPath,
        params: { startsWith: "一、", level: 1 },
      });

      const xml = await readZipText(outputPath, "word/document.xml");
      expect(result.status).toBe("done");
      expect(xml).toContain('w:val="Heading1"');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
```

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/advancedWord.test.ts
```

预期：失败，提示模块不存在。

- [ ] **步骤 2：实现 Word 最小能力**

创建 `advancedWord.ts`，实现：

- `applyHeadingStyles`：对匹配段落插入 `<w:pPr><w:pStyle w:val="HeadingN" /></w:pPr>`。
- `styleTables`：调用现有 `applyOfficeOpenXmlTableStyle`。
- `setHeaderFooter`：写入 `word/header1.xml` 或 `word/footer1.xml`，并在结果中返回 changedParts。
- `insertOrUpdateToc`：返回 `needsCom`，因为目录字段刷新需要 Word。
- `insertOrReplaceImage`：第一阶段没有关系图和 media 写入时返回 `needsCom`。

- [ ] **步骤 3：接入 adapter**

在 `officeActionAdapter.ts` 中，当 `app === "word"` 且 operation 属于 Word 首批能力时，调用 `applyWordAdvancedAction`。

- [ ] **步骤 4：运行任务 5 验证并提交**

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/advancedWord.test.ts electron/agent/tools/officeCore/officeActionAdapter.test.ts
npm run typecheck
git diff --check
```

预期：全部通过。

提交：

```powershell
git add desktop/electron/agent/tools/implementations/officeOpenXml/advancedWord.ts desktop/electron/agent/tools/implementations/officeOpenXml/advancedWord.test.ts desktop/electron/agent/tools/officeCore/officeActionAdapter.ts
git commit -m "feat: add word advanced openxml actions"
```

---

### 任务 6：接入 PPT 首批高级能力

**文件：**
- 新建：`desktop/electron/agent/tools/implementations/officeOpenXml/advancedPresentation.ts`
- 测试：`desktop/electron/agent/tools/implementations/officeOpenXml/advancedPresentation.test.ts`
- 修改：`desktop/electron/agent/tools/officeCore/officeActionAdapter.ts`

- [ ] **步骤 1：编写 PPT 高级能力红灯测试**

创建 `advancedPresentation.test.ts`：

```ts
import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { applyPresentationAdvancedAction } from "./advancedPresentation";

async function writeZip(filePath: string, files: Record<string, string>): Promise<void> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) zip.file(name, content);
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}

async function readZipText(filePath: string, partName: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(filePath));
  const part = zip.file(partName);
  if (!part) throw new Error(`missing ${partName}`);
  return part.async("text");
}

describe("applyPresentationAdvancedAction", () => {
  it("applies theme colors to slide text runs", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-ppt-advanced-"));
    try {
      const filePath = path.join(tempDir, "slides.pptx");
      const outputPath = path.join(tempDir, "slides-edited.pptx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "ppt/slides/slide1.xml": '<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>标题</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
      });

      const result = await applyPresentationAdvancedAction({
        operation: "applyTheme",
        filePath,
        outputPath,
        params: { accentColor: "1F4E79" },
      });

      const xml = await readZipText(outputPath, "ppt/slides/slide1.xml");
      expect(result.status).toBe("done");
      expect(xml).toContain('val="1F4E79"');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
```

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/advancedPresentation.test.ts
```

预期：失败，提示模块不存在。

- [ ] **步骤 2：实现 PPT 最小能力**

创建 `advancedPresentation.ts`，实现：

- `applyTheme`：向文本 run 属性写入 `a:solidFill/a:srgbClr`。
- `normalizeLayouts`：第一阶段只返回 `needsCom`，避免伪造复杂坐标重排。
- `alignShapes`：第一阶段只返回 `needsCom`。
- `insertChart`：第一阶段返回 `needsCom`。
- `replacePictureSlot`：第一阶段返回 `needsCom`，因为需要 media part 和 relationships 完整写入。

- [ ] **步骤 3：接入 adapter**

在 `officeActionAdapter.ts` 中，当 `app === "presentation"` 且 operation 属于 PPT 首批能力时，调用 `applyPresentationAdvancedAction`。

- [ ] **步骤 4：运行任务 6 验证并提交**

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/advancedPresentation.test.ts electron/agent/tools/officeCore/officeActionAdapter.test.ts
npm run typecheck
git diff --check
```

预期：全部通过。

提交：

```powershell
git add desktop/electron/agent/tools/implementations/officeOpenXml/advancedPresentation.ts desktop/electron/agent/tools/implementations/officeOpenXml/advancedPresentation.test.ts desktop/electron/agent/tools/officeCore/officeActionAdapter.ts
git commit -m "feat: add presentation advanced openxml actions"
```

---

### 任务 7：更新提示词与侧边栏监控

**文件：**
- 修改：`desktop/electron/agent/prompts/sections/officeToolsPrompt.ts`
- 修改：`desktop/electron/agent/prompts/systemPrompt.test.ts`
- 修改：`desktop/src/utils/officeEditEvents.ts`
- 修改：`desktop/src/utils/officeEditEvents.test.ts`

- [ ] **步骤 1：编写提示词红灯测试**

在 `systemPrompt.test.ts` 增加：

```ts
test("guides advanced Office work through unified actions", () => {
  const prompt = buildSystemPrompt();

  expect(prompt).toContain("office.action.apply");
  expect(prompt).toContain("done / unsupported / needsCom / failed");
  expect(prompt).toContain("Excel/Word/PPT 高级操作优先使用统一 Office action");
});
```

运行：

```powershell
npm test -- electron/agent/prompts/systemPrompt.test.ts
```

预期：失败，提示缺少新文案。

- [ ] **步骤 2：更新提示词**

在 `officeToolsPrompt.ts` 的 Office 文档段落增加：

```md
### Excel/Word/PPT 高级操作
Excel/Word/PPT 高级操作优先使用统一 Office action：office.action.inspect / office.action.apply / office.action.validate。
执行后必须阅读 status：done 表示完成，unsupported 表示本阶段不支持，needsCom 表示需要显式 COM 兜底，failed 表示执行失败。
状态集合为 done / unsupported / needsCom / failed。不要把 unsupported 或 needsCom 当作成功。
```

- [ ] **步骤 3：编写侧边栏红灯测试**

在 `officeEditEvents.test.ts` 增加 tool result：

```ts
{
  type: "tool_result",
  id: "result-action",
  toolCallId: "call-action",
  toolName: "office.action.apply",
  isError: false,
  timestamp: 8000,
  result: {
    status: "needsCom",
    engine: "openxml",
    app: "word",
    action: "insert",
    operation: "insertOrUpdateToc",
    filePath: "D:\\docs\\report.docx",
    summary: "目录字段需要 Word 刷新",
    changes: [],
  },
}
```

断言摘要包含：

```ts
expect(events[0].summary).toBe("Office action word/insertOrUpdateToc：needsCom");
```

运行：

```powershell
npm test -- src/utils/officeEditEvents.test.ts
```

预期：失败，因为 `office.action.apply` 未被监控识别。

- [ ] **步骤 4：扩展侧边栏事件**

在 `officeEditEvents.ts`：

- `TRACKED_OFFICE_TOOLS` 增加 `office.action.inspect/apply/validate`
- 支持 `status` 字段
- 对统一 action 生成摘要：`Office action ${app}/${operation}：${status}`
- `detail` 包含 `summary`、`changes`、`validation`、`error`

- [ ] **步骤 5：运行任务 7 验证并提交**

```powershell
npm test -- electron/agent/prompts/systemPrompt.test.ts src/utils/officeEditEvents.test.ts
npm run typecheck
git diff --check
```

预期：全部通过。

提交：

```powershell
git add desktop/electron/agent/prompts/sections/officeToolsPrompt.ts desktop/electron/agent/prompts/systemPrompt.test.ts desktop/src/utils/officeEditEvents.ts desktop/src/utils/officeEditEvents.test.ts
git commit -m "feat: guide and monitor office actions"
```

---

### 任务 8：最终验证与清理

**文件：**
- 不新增生产文件，除非验证暴露小问题。

- [ ] **步骤 1：运行聚焦测试**

```powershell
npm test -- electron/agent/tools/officeCore/locator.test.ts electron/agent/tools/officeCore/results.test.ts electron/agent/tools/officeCore/capabilities.test.ts electron/agent/tools/officeCore/officeActionAdapter.test.ts
npm test -- electron/agent/tools/registry/officeTools.test.ts
npm test -- electron/agent/tools/implementations/officeOpenXml/advancedExcel.test.ts electron/agent/tools/implementations/officeOpenXml/advancedWord.test.ts electron/agent/tools/implementations/officeOpenXml/advancedPresentation.test.ts
npm test -- electron/agent/prompts/systemPrompt.test.ts src/utils/officeEditEvents.test.ts
```

预期：全部通过。

- [ ] **步骤 2：运行完整验证**

```powershell
npm run typecheck
npm test
npm run build
git diff --check
```

预期：

- typecheck 通过
- 全量测试通过
- build 通过，已有 Vite CJS/chunk size 警告可接受
- diff check 无错误

- [ ] **步骤 3：检查临时文件**

```powershell
git status --short --branch
```

预期：没有本次实现遗留的未提交临时文件；只允许执行前已有的无关脏文件仍存在。

- [ ] **步骤 4：必要时提交最终小修**

如果最终验证暴露小问题，修复后提交：

```powershell
git add <changed-files>
git commit -m "fix: stabilize office action foundation"
```

如果没有修复变更，不创建空提交。

---

## 自检

- 需求覆盖：统一基础层、统一 action 工具、Open XML 优先 adapter、Excel/Word/PPT 首批文档生产增强能力、提示词、侧边栏监控、最终验证均有任务覆盖。
- 范围控制：复杂图表包、PPT 坐标重排、图片关系写入、目录刷新等不稳定能力第一阶段明确返回 `needsCom`，不伪造成功。
- 类型一致性：`OfficeActionInput`、`OfficeActionResult`、`status`、`engine`、`app`、`operation` 在 tasks 中保持同一命名。
- 模块职责：`officeCore` 只做协议和路由，`officeOpenXml` 做文件级实现，executor 只校验和转发。
