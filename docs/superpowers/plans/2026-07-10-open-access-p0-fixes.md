# 开放使用与 P0 修复实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 彻底移除激活和卡密系统，并修复 Office 安全工具审批绕过及 Excel 样式库覆盖问题。

**Architecture:** 授权系统采用删除式改造，不引入兼容开关。Office 权限由共享纯函数在 executor 和 adapter 两层校验。Excel 样式器保留现有 `styles.xml`，追加样式并把动态索引传给表头单元格。

**Tech Stack:** Electron、React、TypeScript、Vitest、JSZip、Open XML、Express/SQLite（删除）

---

### Task 1: 阻止安全 Office 工具执行写操作

**Files:**
- Create: `desktop/electron/agent/tools/officeCore/operationPolicy.ts`
- Modify: `desktop/electron/agent/tools/executors/officeExecutors.ts`
- Modify: `desktop/electron/agent/tools/officeCore/officeActionAdapter.ts`
- Test: `desktop/electron/agent/tools/executors/officeExecutors.test.ts`
- Test: `desktop/electron/agent/tools/officeCore/officeActionAdapter.test.ts`

- [ ] **Step 1: 编写 executor 失败测试**

为 `office.action.inspect` 和 `office.action.validate` 分别提交写 operation：

```ts
it.each([
  ["office.action.inspect", "excel", "writeRange"],
  ["office.action.validate", "word", "setHeaderFooter"],
  ["office.action.inspect", "presentation", "addSlides"],
])("%s rejects mutation operation %s", async (toolName, app, operation) => {
  const officeActionBridge = { executeAction: vi.fn() } as unknown as OfficeActionBridge;
  const target = createTarget({ officeActionBridge });

  const result = await target.get(toolName)!.execute({
    app,
    operation,
    filePath: "C:/tmp/input.office",
  });

  expect(result.success).toBe(false);
  expect(result.error).toContain("office.action.apply");
  expect(officeActionBridge.executeAction).not.toHaveBeenCalled();
});
```

- [ ] **Step 2: 运行 executor 测试并确认红灯**

Run:

```powershell
npx vitest run electron/agent/tools/executors/officeExecutors.test.ts
```

Expected: 新增测试失败，因为当前 safe executor 仍调用 `executeAction`。

- [ ] **Step 3: 编写 adapter 失败测试**

```ts
it.each([
  ["inspect", "excel", "writeRange"],
  ["validate", "word", "setHeaderFooter"],
  ["inspect", "presentation", "addSlides"],
] as const)("rejects %s action routing mutation %s/%s", async (action, app, operation) => {
  const bridge = createOfficeActionBridge({});
  const result = await bridge.executeAction({
    app,
    action,
    operation,
    filePath: "D:\\docs\\input.office",
  });

  expect(result.status).toBe("failed");
  expect(result.summary).toContain("office.action.apply");
});
```

- [ ] **Step 4: 运行 adapter 测试并确认红灯**

Run:

```powershell
npx vitest run electron/agent/tools/officeCore/officeActionAdapter.test.ts
```

Expected: 写 operation 被实际路由，结果不是权限拒绝。

- [ ] **Step 5: 实现共享 operation 策略**

创建 `operationPolicy.ts`：

```ts
import type { OfficeActionKind } from "./types";

const SAFE_ACTION_OPERATIONS = new Set([
  "inspectFile",
  "layout",
  "tables",
  "snapshot",
]);

export function officeActionOperationError(
  action: OfficeActionKind,
  operation: string
): string | undefined {
  if ((action === "inspect" || action === "validate") &&
      !SAFE_ACTION_OPERATIONS.has(operation)) {
    return `${action} 仅允许只读 Office 操作；修改文件请使用 office.action.apply`;
  }
  return undefined;
}
```

在 executor 构造 `OfficeActionInput` 前调用该函数并返回 `{ success: false, error }`。在 adapter 的 `try` 块最前面再次调用，并通过 `failedResult` 返回。

- [ ] **Step 6: 运行两个目标测试并确认绿灯**

```powershell
npx vitest run electron/agent/tools/executors/officeExecutors.test.ts electron/agent/tools/officeCore/officeActionAdapter.test.ts
```

Expected: 两个文件全部通过。

- [ ] **Step 7: 提交 Office 权限修复**

```powershell
git add desktop/electron/agent/tools/officeCore/operationPolicy.ts desktop/electron/agent/tools/executors/officeExecutors.ts desktop/electron/agent/tools/officeCore/officeActionAdapter.ts desktop/electron/agent/tools/executors/officeExecutors.test.ts desktop/electron/agent/tools/officeCore/officeActionAdapter.test.ts
git commit -m "fix: block Office mutation approval bypass"
```

### Task 2: 保留 Excel 已有样式库

**Files:**
- Modify: `desktop/electron/agent/tools/implementations/officeOpenXml/tableStyler.ts`
- Test: `desktop/electron/agent/tools/implementations/officeOpenXml/tableStyler.test.ts`

- [ ] **Step 1: 编写已有样式保留失败测试**

新增工作簿 fixture，其 `styles.xml` 包含自定义 `numFmt`、两个 font、两个 fill、两个 border、三个 cell xf 和命名样式。工作表的表头和数据行使用不同索引。

关键断言：

```ts
expect(stylesXml).toContain('formatCode="yyyy-mm-dd"');
expect(stylesXml).toContain("<name val=\"OriginalFont\" />");
expect(stylesXml).toContain('<border><left style="thin"');
expect(stylesXml).toContain('<cellStyle name="Custom" xfId="1"');
expect(stylesXml).toContain('<fonts count="3">');
expect(stylesXml).toContain('<fills count="3">');
expect(stylesXml).toContain('<cellXfs count="4">');
expect(outputSheetXml).toContain('r="A1" s="3"');
expect(outputSheetXml).toContain('r="A2" s="2"');
```

- [ ] **Step 2: 运行样式测试并确认红灯**

```powershell
npx vitest run electron/agent/tools/implementations/officeOpenXml/tableStyler.test.ts
```

Expected: 原 `styles.xml` 被替换，原格式断言失败，表头仍使用 `s="1"`。

- [ ] **Step 3: 实现样式追加和动态索引**

将表头样式函数改为接收索引：

```ts
function styleSpreadsheetHeaderRow(rowXml: string, styleIndex: number): string
function styleSpreadsheetRows(xml: string, styleIndex: number): { xml: string; changed: boolean }
```

新增集合追加函数，匹配 `fonts`、`fills` 和 `cellXfs`，从实际子节点数量计算旧数量，更新 `count` 后将新节点追加到集合末尾：

```ts
interface AppendedCollection {
  xml: string;
  index: number;
}

function appendStyleCollectionItem(
  xml: string,
  collectionName: "fonts" | "fills" | "cellXfs",
  itemName: "font" | "fill" | "xf",
  itemXml: string
): AppendedCollection
```

`ensureSpreadsheetStyleParts` 返回新增 `cellXfs` 索引：

```ts
async function ensureSpreadsheetStyleParts(
  zip: JSZip,
  color: string,
  changedParts: string[]
): Promise<number>
```

已有样式表依次追加：

```xml
<font><b /></font>
<fill><patternFill patternType="solid"><fgColor rgb="FF{color}" /><bgColor indexed="64" /></patternFill></fill>
<xf numFmtId="0" fontId="{newFontIndex}" fillId="{newFillIndex}" borderId="0" xfId="0" applyFont="1" applyFill="1" />
```

不存在 `styles.xml` 时继续生成最小样式表并返回索引 `1`。

- [ ] **Step 4: 运行样式测试并确认绿灯**

```powershell
npx vitest run electron/agent/tools/implementations/officeOpenXml/tableStyler.test.ts
```

Expected: Word 测试、无样式 Excel 测试和已有样式保留测试全部通过。

- [ ] **Step 5: 提交 Excel 样式修复**

```powershell
git add desktop/electron/agent/tools/implementations/officeOpenXml/tableStyler.ts desktop/electron/agent/tools/implementations/officeOpenXml/tableStyler.test.ts
git commit -m "fix: preserve existing Excel styles"
```

### Task 3: 移除桌面端激活系统

**Files:**
- Delete: `desktop/electron/main-modules/activationManager.ts`
- Delete: `desktop/src/store/activationStore.ts`
- Delete: `desktop/src/components/ActivationDialog.tsx`
- Delete: `desktop/src/components/ActivationAdminView.tsx`
- Delete: `desktop/src/components/DeviceManagementView.tsx`
- Modify: `desktop/electron/main.ts`
- Modify: `desktop/electron/main-modules/ipcHandlers.ts`
- Modify: `desktop/electron/preload.ts`
- Modify: `desktop/electron/shared/ipcSchemas.ts`
- Modify: `desktop/src/App.tsx`
- Modify: `desktop/src/components/SettingsPage.tsx`
- Modify: `desktop/src/electronApi.d.ts`
- Modify: `desktop/src/services/ipcApi.ts`
- Modify: `desktop/src/services/ipcApiTypes.ts`
- Modify: `desktop/src/services/ipcApiMock.ts`

- [ ] **Step 1: 移除主进程激活生命周期**

从 `main.ts` 删除 `activationManager` import、`initActivation()`、`stopHeartbeat()` 以及启动顺序注释中的激活步骤。

从 `ipcHandlers.ts` 删除激活 schema import、manager import 和全部 `activation:*` handler。

从 `ipcSchemas.ts` 删除三个 Activation schema。

- [ ] **Step 2: 移除 preload 和 renderer IPC 契约**

从 `preload.ts` 删除 `activation` namespace。

从 `electronApi.d.ts`、`ipcApiTypes.ts` 删除 activation 类型和 `ActivationInfo`。

从 `ipcApi.ts`、`ipcApiMock.ts` 删除 activation namespace。

- [ ] **Step 3: 移除激活界面和门控**

从 `App.tsx` 删除 activation store/dialog import、状态读取、`loadActivationStatus()` 和激活弹窗分支。

从 `SettingsPage.tsx` 删除 activation section、文案、图标、import 和渲染分支。

删除五个独立激活文件。

- [ ] **Step 4: 运行类型检查**

```powershell
npm run typecheck
```

Expected: 通过，且不再存在 activation 类型或 API 引用。

- [ ] **Step 5: 搜索桌面端残余引用**

```powershell
git grep -n -E "activationManager|useActivationStore|activation:|ActivationDialog|ActivationAdminView|DeviceManagementView|卡密|许可证服务器" -- desktop
```

Expected: 无运行时结果。

- [ ] **Step 6: 提交桌面端开放使用改动**

仅暂存本任务列出的文件，不暂存用户已有的 `desktop/package.json`、`desktop/package-lock.json` 或知识库修改。

```powershell
git add desktop/electron/main.ts desktop/electron/main-modules/ipcHandlers.ts desktop/electron/preload.ts desktop/electron/shared/ipcSchemas.ts desktop/src/App.tsx desktop/src/components/SettingsPage.tsx desktop/src/electronApi.d.ts desktop/src/services/ipcApi.ts desktop/src/services/ipcApiTypes.ts desktop/src/services/ipcApiMock.ts
git add -u -- desktop/electron/main-modules/activationManager.ts desktop/src/store/activationStore.ts desktop/src/components/ActivationDialog.tsx desktop/src/components/ActivationAdminView.tsx desktop/src/components/DeviceManagementView.tsx
git commit -m "feat: remove desktop activation requirements"
```

### Task 4: 删除激活服务及全部数据

**Files:**
- Delete directory: `admin-server`

- [ ] **Step 1: 验证删除目标**

```powershell
$workspace = (Resolve-Path .).Path
$target = (Resolve-Path admin-server).Path
if (-not $target.StartsWith($workspace + [IO.Path]::DirectorySeparatorChar)) {
  throw "拒绝删除工作区外路径: $target"
}
$target
```

Expected: `D:\excel-ai-plugin-new-feature\admin-server`。

- [ ] **Step 2: 删除整个服务目录**

```powershell
Remove-Item -LiteralPath $target -Recurse -Force
```

- [ ] **Step 3: 验证目录和引用均已消失**

```powershell
Test-Path admin-server
git grep -n "admin-server" -- ':!docs/**'
```

Expected: `False`，且源码中无运行时引用。

- [ ] **Step 4: 提交服务删除**

```powershell
git add -u -- admin-server
git commit -m "feat: remove activation server"
```

### Task 5: 完整验证

**Files:**
- Modify only if verification exposes a direct regression from Tasks 1-4.

- [ ] **Step 1: 运行 lint**

```powershell
npm run lint
```

Expected: 0 errors；原有 Hook warning 若相关组件已删除则应减少。

- [ ] **Step 2: 运行类型检查**

```powershell
npm run typecheck
```

Expected: 通过。

- [ ] **Step 3: 运行完整桌面端测试**

```powershell
npm test
```

Expected: 全部测试通过。

- [ ] **Step 4: 运行生产构建**

```powershell
npm run build
```

Expected: 构建通过；允许记录既有 chunk size 警告。

- [ ] **Step 5: 检查最终残余和工作区范围**

```powershell
git grep -n -E "activationManager|useActivationStore|activation:|ActivationDialog|ActivationAdminView|DeviceManagementView|卡密|许可证服务器" -- desktop ':!desktop/public/knowledge/**'
git status --short
```

Expected: 无授权系统运行时引用；状态中只包含用户原有改动和本任务明确产生的提交。
