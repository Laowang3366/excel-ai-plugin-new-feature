# Office 可视化与表格优化实施计划

> **给执行代理的要求：** 必须使用 `superpowers:subagent-driven-development`（推荐）或 `superpowers:executing-plans` 按任务逐项执行本计划。步骤使用复选框（`- [ ]`）跟踪进度。

**目标：** 建立 Word/PPT 排版优化与 Excel/Word/PPT 表格美化的 Open XML 优先闭环，COM 只作为兜底方案。

**架构：** 在现有 Office Open XML 实现上扩展布局检查、表格检查、样式应用和视觉快照模块。模型可见工具仍放在 `tools/registry` 和 `tools/executors`，文件解析与编辑能力放在 `tools/implementations/officeOpenXml`；COM 桥接只保留为兜底实现。

**技术栈：** TypeScript、Electron 主进程、JSZip、现有 Open XML 引擎、Vitest、现有 COM bridge、运行时探测到的可选本地渲染器。

---

## 执行规范

- 一次只执行一个阶段。
- 禁止过度设计：当前阶段只增加必要接口。
- 禁止过度兜底：只有 Open XML 明确失败或返回不支持能力时，才进入兜底路径。
- 禁止过度约束边界导致既有工具行为异常。
- 代码风格保持与当前 `desktop/electron/agent` 模块一致。
- 注释保持简短清晰，重点说明模块边界与非显而易见逻辑。
- 遵循单一职责，但不要为了拆分而把文件切成过小碎片。
- 每个阶段完成后：运行聚焦测试，进行 review，通过后提交，再进入下一阶段。
- 验证通过后清理非必须、非必要的临时测试文件。
- 不要触碰无关的脏工作区文件，例如 `screenshot.png` 或 dev-server 日志。

## 计划文件结构

- 新建 `desktop/electron/agent/tools/implementations/officeOpenXml/layoutInspector.ts`
  作用：把 `.docx/.pptx/.xlsx` 中的布局对象解析成模型易读结构。

- 新建 `desktop/electron/agent/tools/implementations/officeOpenXml/tableInspector.ts`
  作用：从 Open XML parts 中解析表格区域、单元格、表头、填充色、边框、对齐方式和尺寸信息。

- 新建 `desktop/electron/agent/tools/implementations/officeOpenXml/tableStyler.ts`
  作用：向 Open XML 包应用保守的表格样式预设。

- 新建 `desktop/electron/agent/tools/implementations/officeOpenXml/visualSnapshot.ts`
  作用：将文档页、幻灯片、工作表或表格区域渲染/导出为 PNG/PDF 快照，优先使用 Open XML/headless，COM 作为第二选择。

- 修改 `desktop/electron/agent/tools/implementations/officeOpenXml/types.ts`
  作用：增加布局、表格、视觉工具共享的输入和结果类型。

- 修改 `desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts`
  作用：通过现有 file bridge 暴露 Open XML 布局、表格、快照方法。

- 修改 `desktop/electron/agent/tools/contracts/office.ts`
  作用：增加与具体实现无关的文件级布局、表格、视觉操作契约。

- 修改 `desktop/electron/agent/tools/registry/office.ts`
  作用：新增模型可见工具及描述。

- 修改 `desktop/electron/agent/tools/executors/officeExecutors.ts`
  作用：校验参数并把工具调用路由到 file bridge。

- 修改 `desktop/electron/agent/prompts/sections/officeToolsPrompt.ts`
  作用：要求布局优化、视觉快照和表格美化优先使用 Open XML。

- 修改 `desktop/electron/agent/tools/registry/officeTools.test.ts`
  作用：锁定工具注册和 executor 路由行为。

- 新建或扩展 `desktop/electron/agent/tools/implementations/officeOpenXml/*.test.ts`
  作用：用最小生成的 Office 包验证 Open XML 解析器和样式器行为。

---

### 任务 1：增加工具契约和注册项

**文件：**
- 修改：`desktop/electron/agent/tools/contracts/office.ts`
- 修改：`desktop/electron/agent/tools/registry/office.ts`
- 修改：`desktop/electron/agent/tools/registry/officeTools.test.ts`

- [ ] **步骤 1：编写失败的注册测试**

在 `officeTools.test.ts` 中加入这些预期工具名：

```ts
expect(names).toEqual(expect.arrayContaining([
  "office.layout.inspect",
  "office.table.inspect",
  "office.table.applyStyle",
  "office.visual.snapshot",
]));
```

运行：

```powershell
npm test -- electron/agent/tools/registry/officeTools.test.ts
```

预期：失败，因为这些工具尚未注册。

- [ ] **步骤 2：扩展 OfficeFileBridge 契约**

向 `OfficeFileBridge` 增加方法：

```ts
inspectLayout(input: { filePath: string; target?: string }): Promise<unknown>;
inspectTable(input: { filePath: string; target?: string }): Promise<unknown>;
applyTableStyle(input: {
  filePath: string;
  style: "professional" | "compact" | "financial";
  outputPath?: string;
  target?: string;
}): Promise<unknown>;
snapshot(input: {
  filePath: string;
  target?: string;
  outputPath?: string;
  preferEngine?: "openxml" | "com";
}): Promise<unknown>;
```

- [ ] **步骤 3：增加工具定义**

在 `office.ts` 中增加四个 `ToolDefinition`：

```ts
{
  name: "office.layout.inspect",
  description: "Open XML 优先检查 .docx/.pptx/.xlsx 的页面、幻灯片、工作表对象和基础样式；COM 仅作兜底",
  parameters: {
    type: "object",
    properties: {
      filePath: { type: "string", description: "Office 文件绝对路径" },
      target: { type: "string", description: "可选目标，如 slide:1、page:1、sheet:Sheet1" },
    },
    required: ["filePath"],
  },
  riskLevel: "safe",
  requiresApproval: false,
}
```

按同样 schema 风格补齐：

- `office.table.inspect`：安全工具，参数为 `filePath + target`
- `office.table.applyStyle`：中等风险工具，参数为 `filePath + style + outputPath + target`
- `office.visual.snapshot`：安全工具，参数为 `filePath + target + outputPath + preferEngine`

- [ ] **步骤 4：运行注册测试**

运行：

```powershell
npm test -- electron/agent/tools/registry/officeTools.test.ts
```

预期：注册名断言通过；executor 路由相关断言可能要等任务 2 后才通过。

- [ ] **步骤 5：审查并提交**

审查：

```powershell
git diff -- desktop/electron/agent/tools/contracts/office.ts desktop/electron/agent/tools/registry/office.ts desktop/electron/agent/tools/registry/officeTools.test.ts
```

提交：

```powershell
git add desktop/electron/agent/tools/contracts/office.ts desktop/electron/agent/tools/registry/office.ts desktop/electron/agent/tools/registry/officeTools.test.ts
git commit -m "feat: register office visual table tools"
```

---

### 任务 2：把 executor 路由到 file bridge

**文件：**
- 修改：`desktop/electron/agent/tools/executors/officeExecutors.ts`
- 修改：`desktop/electron/agent/tools/registry/officeTools.test.ts`

- [ ] **步骤 1：编写失败的 executor 测试**

增加 mock bridge 方法：

```ts
const officeFileBridge: OfficeFileBridge = {
  inspectFile: vi.fn(),
  replaceText: vi.fn(),
  inspectLayout: vi.fn(async () => ({ engine: "openxml", operation: "inspectLayout" })),
  inspectTable: vi.fn(async () => ({ engine: "openxml", operation: "inspectTable" })),
  applyTableStyle: vi.fn(async () => ({ engine: "openxml", operation: "applyTableStyle" })),
  snapshot: vi.fn(async () => ({ engine: "openxml", operation: "snapshot", outputPath: "D:\\docs\\snapshot.png" })),
};
```

断言：

```ts
expect([...executors.keys()]).toEqual(expect.arrayContaining([
  "office.layout.inspect",
  "office.table.inspect",
  "office.table.applyStyle",
  "office.visual.snapshot",
]));
```

运行：

```powershell
npm test -- electron/agent/tools/registry/officeTools.test.ts
```

预期：失败，因为 executor 还没有接线。

- [ ] **步骤 2：增加 executor 路由**

在 `if (officeFileBridge)` 内增加路由：

```ts
target.set("office.layout.inspect", {
  name: "office.layout.inspect",
  execute: async (args) => {
    const err = validateArgs(args, { filePath: "string" });
    if (err) return { success: false, error: err };
    const result = await officeFileBridge.inspectLayout({
      filePath: args.filePath as string,
      target: args.target as string | undefined,
    });
    return { success: true, data: result };
  },
});
```

按同样模式补齐：

- `office.table.inspect`
- `office.table.applyStyle`，要求 `filePath` 和 `style`
- `office.visual.snapshot`

- [ ] **步骤 3：运行聚焦测试**

运行：

```powershell
npm test -- electron/agent/tools/registry/officeTools.test.ts
```

预期：通过。

- [ ] **步骤 4：审查并提交**

审查：

```powershell
git diff -- desktop/electron/agent/tools/executors/officeExecutors.ts desktop/electron/agent/tools/registry/officeTools.test.ts
```

提交：

```powershell
git add desktop/electron/agent/tools/executors/officeExecutors.ts desktop/electron/agent/tools/registry/officeTools.test.ts
git commit -m "feat: route office visual table executors"
```

---

### 任务 3：实现 Open XML 布局检查

**文件：**
- 新建：`desktop/electron/agent/tools/implementations/officeOpenXml/layoutInspector.ts`
- 修改：`desktop/electron/agent/tools/implementations/officeOpenXml/types.ts`
- 修改：`desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts`
- 测试：`desktop/electron/agent/tools/implementations/officeOpenXml/layoutInspector.test.ts`

- [ ] **步骤 1：编写 PPTX 布局失败测试**

创建最小 PPTX zip，其中 `ppt/slides/slide1.xml` 包含一个文本形状和一个图片关系引用。断言：

```ts
expect(result.documentType).toBe("presentation");
expect(result.objects[0]).toMatchObject({
  type: "text",
  partName: "ppt/slides/slide1.xml",
});
```

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/layoutInspector.test.ts
```

预期：失败，因为 `inspectOfficeOpenXmlLayout` 尚不存在。

- [ ] **步骤 2：实现最小 PPTX 布局解析器**

实现：

```ts
export async function inspectOfficeOpenXmlLayout(input: OfficeOpenXmlLayoutInspectInput): Promise<OfficeOpenXmlLayoutInspectResult>
```

首轮范围：

- 根据扩展名识别文档类型
- 读取当前引擎规则已经支持的文本 parts
- 返回对象列表，字段包含 `type`、`partName`、`text`、`textLength`
- 遇到暂不支持的对象时跳过，不猜测结构

- [ ] **步骤 3：增加 bridge 方法**

在 `OfficeOpenXmlFileBridge` 中增加：

```ts
inspectLayout(input: OfficeOpenXmlLayoutInspectInput): Promise<unknown> {
  return inspectOfficeOpenXmlLayout(input);
}
```

- [ ] **步骤 4：运行聚焦测试**

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/layoutInspector.test.ts
```

预期：通过。

- [ ] **步骤 5：审查并提交**

运行：

```powershell
npm test -- electron/agent/tools/registry/officeTools.test.ts electron/agent/tools/implementations/officeOpenXml/layoutInspector.test.ts
```

提交：

```powershell
git add desktop/electron/agent/tools/implementations/officeOpenXml/layoutInspector.ts desktop/electron/agent/tools/implementations/officeOpenXml/layoutInspector.test.ts desktop/electron/agent/tools/implementations/officeOpenXml/types.ts desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts
git commit -m "feat: inspect office openxml layout"
```

---

### 任务 4：实现 XLSX、DOCX、PPTX 表格检查

**文件：**
- 新建：`desktop/electron/agent/tools/implementations/officeOpenXml/tableInspector.ts`
- 修改：`desktop/electron/agent/tools/implementations/officeOpenXml/types.ts`
- 修改：`desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts`
- 测试：`desktop/electron/agent/tools/implementations/officeOpenXml/tableInspector.test.ts`

- [ ] **步骤 1：编写失败的表格检查测试**

创建最小 Office 包：

- XLSX：`xl/worksheets/sheet1.xml` 中包含类似表格区域的 `sheetData`
- DOCX：`word/document.xml` 中包含一个 `w:tbl`
- PPTX：`ppt/slides/slide1.xml` 中包含一个 `a:tbl`

断言每个结果都包含：

```ts
expect(result.tables.length).toBe(1);
expect(result.tables[0].rows.length).toBeGreaterThan(0);
expect(result.tables[0].columns).toBeGreaterThan(0);
```

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/tableInspector.test.ts
```

预期：失败，因为表格检查能力还不存在。

- [ ] **步骤 2：实现最小表格解析器**

实现：

```ts
export async function inspectOfficeOpenXmlTables(input: OfficeOpenXmlTableInspectInput): Promise<OfficeOpenXmlTableInspectResult>
```

首轮范围：

- 提取单元格文本
- 返回行数
- 返回列数
- 基础表头行推断：首行大多数单元格为文本
- 基础样式信号：在同一个 XML part 中能直接读到的填充色、加粗标记、对齐方式

- [ ] **步骤 3：接入 bridge 方法**

在 `OfficeOpenXmlFileBridge` 中增加：

```ts
inspectTable(input: OfficeOpenXmlTableInspectInput): Promise<unknown> {
  return inspectOfficeOpenXmlTables(input);
}
```

- [ ] **步骤 4：运行测试**

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/tableInspector.test.ts
```

预期：通过。

- [ ] **步骤 5：审查并提交**

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/tableInspector.test.ts electron/agent/tools/registry/officeTools.test.ts
```

提交：

```powershell
git add desktop/electron/agent/tools/implementations/officeOpenXml/tableInspector.ts desktop/electron/agent/tools/implementations/officeOpenXml/tableInspector.test.ts desktop/electron/agent/tools/implementations/officeOpenXml/types.ts desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts
git commit -m "feat: inspect openxml tables"
```

---

### 任务 5：使用 Open XML 应用保守表格样式

**文件：**
- 新建：`desktop/electron/agent/tools/implementations/officeOpenXml/tableStyler.ts`
- 修改：`desktop/electron/agent/tools/implementations/officeOpenXml/types.ts`
- 修改：`desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts`
- 测试：`desktop/electron/agent/tools/implementations/officeOpenXml/tableStyler.test.ts`

- [ ] **步骤 1：编写失败的样式测试**

验证 `professional` 样式：

- 写出输出文件
- 修改表头行填充色或样式标记
- 提供 `outputPath` 时不修改源文件

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/tableStyler.test.ts
```

预期：失败，因为样式应用能力尚不存在。

- [ ] **步骤 2：实现第一个样式预设**

实现：

```ts
export async function applyOfficeOpenXmlTableStyle(input: OfficeOpenXmlTableStyleInput): Promise<OfficeOpenXmlTableStyleResult>
```

首轮范围：

- 先只让 `style: "professional"` 通过首个绿色测试
- 表头行加粗
- 表头填充色
- 在 XML part 直接支持时添加隔行填充
- 安全时添加简单边框标记
- 默认输出路径为 `<base>-styled<ext>`

- [ ] **步骤 3：增加剩余预设但不扩大行为边界**

增加：

- `compact`：在能直接表示时缩小 padding/行高
- `financial`：右对齐数值单元格，并在能识别时强调汇总行

本阶段不处理高级公式、数据透视表、切片器或图表。

- [ ] **步骤 4：接入 bridge 方法**

在 `OfficeOpenXmlFileBridge` 中增加：

```ts
applyTableStyle(input: OfficeOpenXmlTableStyleInput): Promise<unknown> {
  return applyOfficeOpenXmlTableStyle(input);
}
```

- [ ] **步骤 5：运行测试并提交**

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/tableStyler.test.ts electron/agent/tools/implementations/officeOpenXml/tableInspector.test.ts
```

提交：

```powershell
git add desktop/electron/agent/tools/implementations/officeOpenXml/tableStyler.ts desktop/electron/agent/tools/implementations/officeOpenXml/tableStyler.test.ts desktop/electron/agent/tools/implementations/officeOpenXml/types.ts desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts
git commit -m "feat: apply openxml table styles"
```

---

### 任务 6：增加明确兜底规则的视觉快照

**文件：**
- 新建：`desktop/electron/agent/tools/implementations/officeOpenXml/visualSnapshot.ts`
- 修改：`desktop/electron/agent/tools/implementations/officeOpenXml/types.ts`
- 修改：`desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts`
- 测试：`desktop/electron/agent/tools/implementations/officeOpenXml/visualSnapshot.test.ts`

- [ ] **步骤 1：编写失败的渲染器选择测试**

测试行为不依赖 LibreOffice 或 Office：

```ts
expect(selectSnapshotPlan({ preferEngine: "openxml", hasHeadlessRenderer: true, hasComFallback: true }).engine)
  .toBe("openxml");
expect(selectSnapshotPlan({ preferEngine: "openxml", hasHeadlessRenderer: false, hasComFallback: true }).engine)
  .toBe("com");
expect(selectSnapshotPlan({ preferEngine: "com", hasHeadlessRenderer: true, hasComFallback: true }).engine)
  .toBe("com");
```

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/visualSnapshot.test.ts
```

预期：失败，因为 snapshot planner 尚不存在。

- [ ] **步骤 2：实现 snapshot planner**

实现：

```ts
export function selectSnapshotPlan(input: {
  preferEngine?: "openxml" | "com";
  hasHeadlessRenderer: boolean;
  hasComFallback: boolean;
}): { engine: "openxml" | "com"; reason: string }
```

规则：

- 显式 `preferEngine: "com"` 时使用 COM
- Open XML 在 headless renderer 可用时使用 headless 渲染器
- 只有 headless 渲染器不可用时才使用 COM 兜底
- 没有可用渲染器时抛出清晰错误

- [ ] **步骤 3：实现 snapshot 外壳**

实现：

```ts
export async function createOfficeVisualSnapshot(input: OfficeVisualSnapshotInput): Promise<OfficeVisualSnapshotResult>
```

首轮范围：

- 通过检查已知候选命令探测可用 headless renderer
- renderer 缺失且 COM 兜底不可用时，返回清晰的不支持错误
- 创建输出路径，并确保不留下临时文件

- [ ] **步骤 4：接入 bridge 方法**

在 `OfficeOpenXmlFileBridge` 中增加：

```ts
snapshot(input: OfficeVisualSnapshotInput): Promise<unknown> {
  return createOfficeVisualSnapshot(input);
}
```

- [ ] **步骤 5：运行测试并提交**

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/visualSnapshot.test.ts electron/agent/tools/registry/officeTools.test.ts
```

提交：

```powershell
git add desktop/electron/agent/tools/implementations/officeOpenXml/visualSnapshot.ts desktop/electron/agent/tools/implementations/officeOpenXml/visualSnapshot.test.ts desktop/electron/agent/tools/implementations/officeOpenXml/types.ts desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts
git commit -m "feat: plan office visual snapshots"
```

---

### 任务 7：更新视觉优化的提示词路由

**文件：**
- 修改：`desktop/electron/agent/prompts/sections/officeToolsPrompt.ts`
- 修改：`desktop/electron/agent/prompts/systemPrompt.test.ts`

- [ ] **步骤 1：编写失败的提示词测试**

增加断言：

```ts
expect(prompt).toContain("Open XML 优先");
expect(prompt).toContain("office.visual.snapshot");
expect(prompt).toContain("office.table.applyStyle");
expect(prompt).toContain("COM 仅作为兜底");
```

运行：

```powershell
npm test -- electron/agent/prompts/systemPrompt.test.ts
```

预期：失败，直到提示词更新完成。

- [ ] **步骤 2：更新工具使用指导**

增加段落：

```md
### 视觉排版与表格美化
Open XML 优先：先用 office.layout.inspect / office.table.inspect / office.visual.snapshot 获取结构和截图，再用 office.table.applyStyle 或 Open XML 文件工具修改。
COM 仅作为兜底：只有 Open XML 不支持当前对象、无法渲染截图、或用户明确要求操作当前打开窗口时，才使用 word.* / presentation.* 专用工具。
```

- [ ] **步骤 3：运行提示词测试**

运行：

```powershell
npm test -- electron/agent/prompts/systemPrompt.test.ts
```

预期：通过。

- [ ] **步骤 4：审查并提交**

运行：

```powershell
git diff -- desktop/electron/agent/prompts/sections/officeToolsPrompt.ts desktop/electron/agent/prompts/systemPrompt.test.ts
```

提交：

```powershell
git add desktop/electron/agent/prompts/sections/officeToolsPrompt.ts desktop/electron/agent/prompts/systemPrompt.test.ts
git commit -m "docs: guide openxml visual optimization"
```

---

### 任务 8：增加侧边栏事件摘要

**文件：**
- 修改：`desktop/src/utils/officeEditEvents.ts`
- 修改：`desktop/src/utils/officeEditEvents.test.ts`

- [ ] **步骤 1：编写失败的事件测试**

增加工具结果：

```ts
{
  type: "tool_result",
  id: "result-layout",
  toolCallId: "call-layout",
  toolName: "office.layout.inspect",
  result: { engine: "openxml", operation: "inspectLayout", documentType: "presentation", objectCount: 8 },
  isError: false,
  timestamp: Date.now(),
}
```

断言摘要包含：

```ts
expect(events[0].summary).toContain("Open XML");
expect(events[0].summary).toContain("8");
```

- [ ] **步骤 2：扩展事件提取逻辑**

更新 `officeEditEvents.ts`，识别：

- `office.layout.inspect`
- `office.table.inspect`
- `office.table.applyStyle`
- `office.visual.snapshot`

- [ ] **步骤 3：运行 UI 工具测试**

运行：

```powershell
npm test -- src/utils/officeEditEvents.test.ts
```

预期：通过。

- [ ] **步骤 4：审查并提交**

提交：

```powershell
git add desktop/src/utils/officeEditEvents.ts desktop/src/utils/officeEditEvents.test.ts
git commit -m "feat: summarize office visual events"
```

---

### 任务 9：最终验证与清理

**文件：**
- 除非验证暴露问题，否则不修改生产文件。

- [ ] **步骤 1：运行聚焦测试**

运行：

```powershell
npm test -- electron/agent/tools/implementations/officeOpenXml/layoutInspector.test.ts
npm test -- electron/agent/tools/implementations/officeOpenXml/tableInspector.test.ts
npm test -- electron/agent/tools/implementations/officeOpenXml/tableStyler.test.ts
npm test -- electron/agent/tools/implementations/officeOpenXml/visualSnapshot.test.ts
npm test -- electron/agent/tools/registry/officeTools.test.ts
npm test -- electron/agent/prompts/systemPrompt.test.ts
npm test -- src/utils/officeEditEvents.test.ts
```

预期：全部通过。

- [ ] **步骤 2：运行完整验证**

运行：

```powershell
npm run typecheck
npm test
npm run build
git diff --check
```

预期：

- typecheck 通过
- 全量测试通过
- build 通过，已有 Vite chunk size 警告可接受
- diff check 无空白字符错误

- [ ] **步骤 3：清理临时文件**

只删除实现任务创建的临时文件。不要删除用户文件或无关工作区脏文件。

检查：

```powershell
git status --short --branch
```

预期：只剩意图内的源码/文档/测试变更，以及执行前已经存在的无关脏文件。

- [ ] **步骤 4：必要时提交最终修复**

如果最终验证暴露小问题并已修复：

```powershell
git add <changed-files>
git commit -m "fix: stabilize office visual optimization"
```

如果没有修复变更，不创建空提交。

---

## 自检

- 需求覆盖：Word/PPT 排版优化、表格美化、Open XML 优先路由、视觉快照、COM 兜底、侧边栏监控、分阶段 review 和提交均有对应任务。
- 延迟内容扫描：未发现未落地内容；暂不支持的范围已明确排除在首轮实现之外。
- 类型一致性：工具名在 registry、executor、提示词、事件摘要和测试中保持一致。
- 范围检查：高级对象编辑、完整 PDF 栅格化保真、图表美化、动画、母版页编辑和数据透视表样式不属于首轮实现范围。
