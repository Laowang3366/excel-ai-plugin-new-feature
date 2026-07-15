# 项目开发规范

> 基于 4 周代码审查修复经验，按六大审查方向编写。
> 生成日期：2026-06-25 | 分支：`feature/new-feature`
> 审查范围：18 项任务（#1–#23），涉及模块拆分、接口精简、持久化优化、测试基础设施、IPC 解耦、Bug 修复。

---

## 一、大文件模块拆分

> 审查发现问题：6 个核心文件超过 600 行，职责混合，可维护性差。
> 涉及任务：#9 toolRegistry、#10 agentLoop 首轮、#11 main.ts、#12 chatStore、#18 agentLoop 深度拆分，以及 ChatPage、Sidebar、CSS 的早期拆分。

### 1.1 文件行数硬性上限

| 文件类型 | 上限 | 超限后果 |
|----------|------|----------|
| TypeScript/TSX 通用模块 | ≤ 400 行 | 必须拆分后才能合入 |
| React 组件 | ≤ 300 行 | 提取子组件 + hooks |
| Zustand Store | ≤ 400 行 | 提取 actions/事件处理为独立模块 |
| CSS 文件 | ≤ 500 行 | 按组件/功能域拆分 |

### 1.2 拆分方法：目录按职责分层，入口按需要保留

当单文件超过上限时，按职责创建目录和子模块。不要为了“统一入口”强行增加 `index.ts`；只有公共 API 稳定、消费方确实需要聚合入口时，才保留轻量入口。拆分后应优先让消费方直接 import 具体职责模块。

```
# 拆分前
electron/agent/agentLoop.ts                    (1,066 行)

# 拆分后
electron/agent/agentLoop/
  agentLoop.ts                                 (450 行) — 编排器，对外 API 不变
  streamCollector.ts                           (140 行) — 流式事件收集
  toolExecutor.ts                              (260 行) — 工具执行 + 审批
  buildStreamParams.ts                         (90 行)  — 参数构建
```

**入口文件规则**：
- 不把 `index.ts` 当作默认要求。
- 不保留旧模块的“兼容出口”作为长期影子层。
- 公共入口只导出当前层明确对外承诺的能力，禁止成为无职责的重导出漏斗。
- 内部消费优先直接引用具体模块，例如 `core/agentLoop/toolExecutor`。

```typescript
// ✅ 直接引用具体职责模块
import { processToolCalls } from "./toolExecutor";
import { collectStreamEvents } from "./streamCollector";
```

### 1.3 拆分原则

1. **按职责域拆分**：每个子模块只负责一个运行时关注点（流式收集、工具执行、压缩管理……）
2. **对外 API 不变**：拆分是内部重构，不改变外部调用方式
3. **子模块可独立测试**：纯函数和工具函数可以单独写单元测试
4. **编排器保留流程控制**：`agentLoop.ts` 保留 `runTurn()` 主流程，调用子模块完成具体工作

### 1.4 按文件类型的拆分模式

| 文件类型 | 拆分方向 | 典型案例 |
|----------|----------|----------|
| 大型 React 页面 | 子组件 + hooks + helpers | ChatPage → 7 个模块 |
| 大型 Store | 事件处理 + actions 提取 | chatStore → agentEventHandler + threadActions |
| 大型 Electron 入口 | settingsManager + windowManager + ipcHandlers + eventForwarder | main.ts → 4 个模块 |
| 大型工具注册表 | interfaces + definitions + executors | toolRegistry → 4 个模块 |
| 大型桥接层 | Electron 保持类型化薄桥，Worker 按应用和操作域拆分 | `officeWorker/*` → `desktop/dotnet/Wengge.OfficeWorker/{Office,Excel,Word,Presentation,OpenXml}` |
| 大型 CSS | 按组件/功能域拆分 | global.css → 23 个模块 |

### 1.5 当前待拆分文件

| 文件 | 行数 | 建议拆分方向 |
|------|------|-------------|
| `ModelSettings.tsx` | 1,186 | ProviderList + ProviderForm + ModelConfig 子组件 |
| `aiClient.ts` | 1,078 | providers/ 目录 + streaming.ts + messageBuilder.ts |
| `systemPrompt.ts` | 714 | sections/ 目录，每段提示词独立文件 |

---

## 二、Props 与接口精简

> 审查发现问题：组件 Props 膨胀（28 个独立 props），调用方代码冗长，接口设计不够内聚。
> 涉及任务：#19 ComposerArea Props 精简、#20 settingsStore 接口设计。

### 2.1 Hook 返回值整体传入

当组件大部分 props 来自同一个 hook 时，将 hook 返回值作为**单个 prop** 传入，而非逐个解构。

```typescript
// ❌ 禁止 — 28 个独立 props，调用方冗长
<ComposerArea
  inputText={inputText}
  setInputText={setInputText}
  handleSend={handleSend}
  showAttachPopover={showAttachPopover}
  setShowAttachPopover={setShowAttachPopover}
  ... (24 more props)
/>

// ✅ 正确 — 7 个 props，职责清晰
const composer = useComposer();
<ComposerArea
  composer={composer}           // ReturnType<typeof useComposer>
  currentFolder={currentFolder}
  currentFolderFiles={currentFolderFiles}
  showWelcomeComposer={showWelcomeComposer}
  onSend={handleSend}
  onInterrupt={handleInterrupt}
  onOpenSettings={handleOpenSettings}
/>
```

### 2.2 类型派生

使用 `ReturnType<typeof hook>` 派生 Props 类型，避免手动维护接口：

```typescript
type ComposerState = ReturnType<typeof useComposer>;

interface ComposerAreaProps {
  composer: ComposerState;
  currentFolder: string | null;
  // ...
}
```

### 2.3 组件内部解构

组件内部通过 prop 对象按需解构，不改变外部接口：

```typescript
export const ComposerArea: React.FC<ComposerAreaProps> = ({ composer, ... }) => {
  const {
    inputText, setInputText, handleSend,
    showAttachPopover, setShowAttachPopover,
    attachedFiles, removeAttachedFile,
    // ...
  } = composer;
  // ...
};
```

### 2.4 Store 接口设计

Zustand Store 的 action 设计原则：
- **单个 action 只改一个关注点**：`setPermissionMode` 只改权限模式
- **持久化由 action 自行触发**：action 内部调用 `savePartial`，不依赖外部调用 `saveSettings`
- **避免"上帝方法"**：`saveSettings()` 全量写入 9 个 key 是反模式，已废弃

---

## 三、持久化优化

> 审查发现问题：每次设置变更全量写入 9 个 IPC key，浪费资源且增加冲突风险。
> 涉及任务：#20 settingsStore 增量持久化。

### 3.1 增量写入原则

**禁止**每次变更全量写入所有 key，**必须**仅写入实际变更的 key。

```typescript
// ❌ 全量写入 — 每次写 9 个 key
saveSettings: async () => {
  await ipcApi.settings.set("aiProviders", providers);
  await ipcApi.settings.set("activeProvider", activeProviderId);
  await ipcApi.settings.set("permissionMode", permissionMode);
  // ... 6 more keys
};

// ✅ 增量写入 — 每次只写 1 个 key
setPermissionMode: (mode) => {
  set({ permissionMode: mode });
  savePartial(["permissionMode"], get);
};
```

### 3.2 KEY_MAP 显式映射

state 字段名与 electron-store key 不一致时，使用 `KEY_MAP` 常量显式映射：

```typescript
const KEY_MAP: Partial<Record<keyof SettingsState, string>> = {
  providers: "aiProviders",        // state 字段 ≠ store key
  activeProviderId: "activeProvider",
  permissionMode: "permissionMode", // state 字段 = store key
};
```

### 3.3 复合字段处理

多个 state 字段映射到同一个 store key（嵌套对象）时，使用特殊处理：

```typescript
const COMPACTION_FIELDS: (keyof SettingsState)[] = [
  "compactionEnabled",
  "autoCompactThresholdPercent",
];

// savePartial 中检测到 COMPACTION_FIELDS 时，组合写入
if (needCompaction) {
  toWrite.push(["compactionConfig", {
    enabled: state.compactionEnabled,
    autoCompactThresholdPercent: state.autoCompactThresholdPercent,
  }]);
}
```

### 3.4 并行写入

多个 key 的写入使用 `Promise.all` 并行执行，不阻塞 UI：

```typescript
await Promise.all(
  toWrite.map(([k, v]) => ipcApi.settings.set(k, v))
);
```

### 3.5 向后兼容

全量 `saveSettings()` 方法保留但不再有调用方，作为应急恢复手段。

---

## 四、测试基础设施

> 审查发现问题：项目无单元测试，代码改动无法验证，回归风险高。
> 涉及任务：#21 vitest 基础设施、#22 OCR Mock 标记。

### 4.1 测试框架

- **框架**：Vitest 3 + .NET 8 测试项目
- **配置**：`desktop/vitest.config.ts`、`desktop/dotnet/Wengge.OfficeWorker.Tests/`
- **当前基线**：最近盘点为 TypeScript 147 个测试文件、751 项测试，.NET Worker 21 项测试；数量以命令实际输出为准。提交前必须保持 `npm run typecheck`、`npm run lint`、`npm test` 和 `npm run office:test` 通过。产品站测试在 `product-site/` 独立执行。
- **运行命令**：

```bash
npm test              # 单次运行
npm run test:watch    # 监听模式
npm run test:coverage # 覆盖率报告
npm run office:test   # .NET Worker 测试
```

### 4.1.1 Office 冒烟时限

- 常规单元测试和 Worker 测试不得启动真实 Office 进程。
- 真实 Office 冒烟只按本次修改涉及的应用和 operation 定向运行，不把所有脚本串成默认提交门禁。
- 生产 Office action 默认超时 120 秒；冒烟环境默认单动作超时 30 秒，可通过 `WENGGE_OFFICE_SMOKE_TIMEOUT_MS` 在 5-600 秒范围内调整。
- 冒烟等待期间每 10 秒输出一次探测信息；超时后必须终止当前检查并报告 operation、宿主和文件，不允许无反馈挂起数小时。

### 4.2 必须有测试的模块

| 模块类型 | 覆盖要求 | 示例 |
|----------|----------|------|
| 纯函数 | 覆盖所有分支和边界条件 | `compaction.ts`（token 估算、压缩判断） |
| 工具执行逻辑 | 覆盖风险等级和审批模式 | `toolApproval.test.ts`、`toolExecutor.test.ts` |
| 状态转换逻辑 | 覆盖主要路径 | `agentEventHandler.ts` |
| IPC Schema | 覆盖合法/非法输入 | `ipcSchemas.ts` |
| Office 可靠性 | 覆盖 Worker 协议、进程归属、完整路径 locator、暂停续跑、组事务撤销/重做；链接 OLE 变更必须跑真实 Office 冒烟 | `WorkerProtocolTests.cs`、`transactionJournal.test.ts`、`npm run test:office-reliability` |

### 4.3 测试文件规范

**命名**：`{源文件名}.test.ts`，放在源文件同级或子目录

```
electron/agent/compaction.ts          → electron/agent/compaction.test.ts
electron/agent/agentLoop/toolExecutor.ts → electron/agent/agentLoop/agentLoop.test.ts
```

**结构**：按被测函数/功能分组

```typescript
import { describe, it, expect, vi, beforeEach } from "vitest";

describe("shouldRequireApproval", () => {
  beforeEach(() => {
    clearAlwaysAllowedTools(); // 每个测试前重置状态
  });

  it("should require approval for moderate tools in confirm_all mode", () => {
    expect(shouldRequireApproval("range.write", "confirm_all")).toBe(true);
  });

  it("should auto-approve safe tools in auto_approve_safe mode", () => {
    expect(shouldRequireApproval("range.read", "auto_approve_safe")).toBe(false);
  });
});
```

### 4.4 Mock 规范

**Mock 数据必须标记**：在 mock 数据处添加 `@MOCK_INTERFACE` 注释，标注原因和接入计划。

```typescript
// @MOCK_INTERFACE — OCR 识别：本地能力尚未接入，以下为 mock 数据。
// 接入真实 OCR 后，ipcApi.ocr.recognize 将代理到 preload 实现，此 mock 自动失效。
const mockResult: OcrResult = { ... };
```

**IPC Mock**：使用 `createMockIpcApi()` 注入 mock IPC 实现：

```typescript
vi.mock("../services/ipcApi", () => ({
  ipcApi: createMockIpcApi({
    settings: {
      getAll: vi.fn().mockResolvedValue({ aiProviders: {}, permissionMode: "normal" }),
    },
  }),
}));
```

### 4.5 测试中修复的典型 Bug

本次审查中测试暴露了两个实际问题：
- **import 路径错误**：`toolExecutor.ts` 从 `../types` 导入 `TOOL_DEFINITIONS_MAP`，但该常量定义在 `../toolRegistry`，导致运行时 `undefined`
- **工具名不匹配**：测试用 `read_range`，实际工具定义名为 `range.read`

**教训**：写测试本身就是一种代码审查，能发现运行时才会暴露的问题。

---

## 五、IPC 依赖注入

> 审查发现问题：前端 71 处直接访问 `window.electronAPI`，测试时无法 mock，SSR 环境崩溃。
> 涉及任务：#23 IPC 抽象层、#5 Zod Schema、#6 结构化日志。

### 5.1 ipcApi 抽象层

**禁止**直接访问 `window.electronAPI`，**必须**通过 `src/services/ipcApi.ts`。

```typescript
// ❌ 禁止
const result = await window.electronAPI.agent.startTurn(input);

// ✅ 正确
import { ipcApi } from "../services/ipcApi";
const result = await ipcApi.agent.startTurn(input);
```

**抽象层特性**：
- 无 Electron 环境时安全降级（返回空值而非崩溃）
- 类型安全（`IIpcApi` 接口与 `electronApi.d.ts` 一致）
- 测试时可通过 `createMockIpcApi(overrides)` 注入 mock

### 5.2 Zod Schema 运行时校验

所有 IPC 通道的输入**必须**使用 zod schema 进行运行时校验：

```typescript
// electron/shared/ipcSchemas.ts
export const startTurnSchema = z.object({
  content: z.string().min(1),
  attachments: z.array(fileAttachmentSchema).optional(),
  clientId: z.string().optional(),
  isResume: z.boolean().optional(),
  resumeContext: z.string().optional(),
});

// electron/main-modules/ipcHandlers.ts
ipcMain.handle("agent:startTurn", async (_event, args) => {
  const parsed = startTurnSchema.safeParse(args);
  if (!parsed.success) {
    return { success: false, error: "Invalid parameters" };
  }
  // 使用 parsed.data 而非 raw args
});
```

#### 5.2.1 模型工具 Schema 也是执行边界

- 模型可见工具必须在 `tools/registry/` 声明完整 JSON Schema，并通过 `toolDefinitions.ts` 聚合；禁止只在 executor 中增加模型不可见的参数规则。
- `tools/registry/toolSchema.ts` 会把已声明对象规范化为 `additionalProperties:false`，同一份 Schema 同时发送给模型，并在用户审批前、executor 调用前各校验一次。
- 必填项、类型、enum、整数/数值范围、数组数量和嵌套对象必须写入 Schema；执行器中的业务校验用于跨字段语义，不能替代结构校验。
- 只有 `params`、`variables` 等确实承载多种 operation 的扩展对象可以显式开放；仍必须受统一 JSON 大小、深度、节点数、字符串和数组预算限制，并尽快拆成 operation 级判别 Schema。
- 新增工具必须进入 `toolSchema.test.ts` 的全量 malformed 校验；参数非法时不得先弹审批框或调用 Worker。

### 5.3 结构化日志

**禁止**在生产代码中使用 `console.log`，**必须**使用 `electron/shared/logger.ts`。

```typescript
import { logger } from "../shared/logger";

// 指定上下文
const log = logger.child({ context: "agentLoop" });
log.info("Turn started", { turnId, threadId });
log.error("Stream failed", { error: err.message });
```

### 5.4 IPC 通道命名

| 格式 | 示例 | 用途 |
|------|------|------|
| `domain:action` | `agent:startTurn` | 主进程操作 |
| `domain:event` | `agent:onEvent` | 事件监听 |
| `domain:query` | `settings:get` | 数据查询 |

### 5.5 迁移进度追踪

| 模块 | `window.electronAPI` 调用 | 状态 |
|------|--------------------------|------|
| `store/chatStore.ts` | 10 → 0 | ✅ 已迁移 |
| `store/threadActions.ts` | 13 → 0 | ✅ 已迁移 |
| `store/settingsStore.ts` | 15 → 0 | ✅ 已迁移 |
| `hooks/useComposer.ts` | 4 → 0 | ✅ 已迁移 |
| `hooks/useExcelConnection.ts` | 4 → 0 | ✅ 已迁移 |
| `components/ChatPage.tsx` | 1 → 0 | ✅ 已迁移 |
| `components/Sidebar.tsx` | 3 → 0 | ✅ 已迁移 |
| `components/settings/*.tsx` | 8 → 0 | ✅ 已迁移 |
| `components/task/*.tsx` | 7 → 0 | ✅ 已迁移（OCR 通过 `ipcApi.ocr` 抽象） |
| `App.tsx` | 2 → 0 | ✅ 已迁移 |
| `utils/chatHelpers.tsx` | 2 → 0 | ✅ 已迁移 |

---

## 六、Bug 修复与防御

> 审查发现问题：多个运行时 Bug 和类型错误影响用户体验和 API 兼容性。
> 涉及任务：#1 TurnItem 顺序、#2 API 400、#3 类型检查、#4 reasoningMode、#7 N+1 查询、#10 agentLoop 首轮拆分。

### 6.1 事件驱动架构的顺序保证

**问题模式**：Agent 事件到达顺序 ≠ UI 渲染顺序，导致消息闪烁或排序混乱。

**防御规则**：
- 流式阶段**不立即发出**中间状态的 `item_started` 事件
- 所有 items 在流结束后按**真实 API 顺序**依次发出
- 前端 `item_completed` handler 按事件到达顺序 `push`，不做"智能插入"
- JSONL 历史恢复时使用 `sortItemsByRound` 作为兜底排序
- 模型流只有在尚未发出正文、推理或工具事件时允许透明重试；首次可见事件后必须保留当前 attempt 并直接报告错误，避免同一 round 重复追加。

```typescript
// ✅ 流结束后按顺序发出
for (const item of streamResult.items) {
  callbacks.onEvent({ type: "item_started", item });
  callbacks.onEvent({ type: "item_completed", item });
}
```

### 6.2 API 消息格式校验

**问题模式**：assistant 消息含 `tool_calls` 但缺少对应 `tool` result，AI API 返回 400。

**防御规则**：
- `turnItemsToChatMessages()` 执行三遍处理：识别孤立 tool_call → 跳过 → 清理空壳
- `buildRequestMessages()` 执行双重校验：ID 匹配 → 连续性检查
- 孤立 tool_call **跳过而非合成假 result**，保持上下文一致性

```typescript
// 三遍处理
const orphanIds = findOrphanToolCalls(items);   // 第一遍：识别
const messages = buildWithoutOrphans(items, orphanIds); // 第二遍：跳过
const cleaned = removeEmptyAssistants(messages); // 第三遍：清理
```

### 6.3 类型安全

**防御规则**：
- 两个 tsconfig 均启用 `strict: true`
- 新增/修改类型后必须运行 `npm run typecheck` 验证
- `electronApi.d.ts`（前端视角）与 `types.ts`（主进程视角）保持同步
- 可选属性使用 `?.` 链式访问，避免 `TypeError: Cannot read properties of undefined`

```typescript
// ✅ 安全访问
const toolDef = TOOL_DEFINITIONS_MAP.get(toolName);
if (toolDef?.riskLevel === "safe") { ... }

// ❌ 不安全访问
const toolDef = TOOL_DEFINITIONS_MAP.get(toolName);
if (toolDef.riskLevel === "safe") { ... } // 可能崩溃
```

### 6.4 性能防御

**N+1 查询**：对列表数据的逐项查询必须改为批量查询。

```typescript
// ❌ N+1 — 每个 turn 一次 IPC
for (const turn of turns) {
  const usage = await ipcApi.settings.get(`usage_${turn.turnId}`);
}

// ✅ 批量查询 — 一次 IPC
const allUsage = await ipcApi.settings.getAll();
```

### 6.5 提交前自查清单

| # | 检查项 | 对应审查方向 |
|---|--------|-------------|
| 1 | 新文件 ≤ 400 行，组件 ≤ 300 行 | 一、模块拆分 |
| 2 | 组件 props ≤ 10 个，或使用 hook 归组 | 二、接口精简 |
| 3 | 设置变更使用 `savePartial` 增量写入 | 三、持久化优化 |
| 4 | 核心逻辑有对应 `.test.ts` 文件 | 四、测试基础设施 |
| 5 | IPC 调用使用 `ipcApi`，非 `window.electronAPI` | 五、IPC 解耦 |
| 6 | 新增 IPC 通道有 zod schema 校验 | 五、IPC 解耦 |
| 7 | mock 数据有 `@MOCK_INTERFACE` 标记 | 四、测试基础设施 |
| 8 | 无 `console.log`，使用 `logger` | 五、IPC 解耦 |
| 9 | `npm run typecheck` 通过 | 六、Bug 防御 |
| 10 | `npm test` 全部通过 | 六、Bug 防御 |
| 11 | 用户可感知变化已更新 `CHANGELOG.md`；架构或流程变化已更新对应当前文档 | 文档维护 |
