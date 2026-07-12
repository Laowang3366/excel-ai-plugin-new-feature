# Agent Module Splitting Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 `desktop/electron/agent` 从“已按目录归类”推进到“模块职责清晰、低耦合、可测试”的结构。

**Architecture:** 先拆工具契约和执行器，切断 `registry` 对具体实现的混合职责；再把 Office/Excel 共用的本机自动化能力抽到 `automation`；随后把 Agent 装配和 Electron 交互适配迁入 `runtime`、`interaction`；最后收敛核心循环与提示词大文件。每个任务都保留兼容出口，完成后用测试和架构守护逐步收紧。

**Tech Stack:** TypeScript, Electron main process, Vitest, PowerShell/COM automation, Office/Excel bridge modules.

---

## Execution Standards

每次执行任务必须遵循以下规范：

- 禁止过度设计：只实现当前阶段需要的结构迁移，不引入未被现有业务使用的新抽象、新配置或新流程。
- 禁止过度兜底：保留现有错误处理语义，不为了“看起来更安全”吞掉异常、改写返回值或隐藏原有失败信号。
- 禁止过度约束边界：模块边界用于降低耦合，不应阻断现有 Excel、Word、PowerPoint、Shell、知识检索等业务能力。
- 必须保持代码风格一致：沿用当前 TypeScript、命名、导出、测试和中文注释风格。
- 必须注释清晰：新模块顶部说明当前模块职责和关联模块；注释解释边界与意图，不重复代码表面含义。
- 必须遵循单一职责：每个模块只承载一个明确职责；拆分依据是业务能力和依赖方向，不是机械地把代码切碎。
- 必须按开发清单分阶段完成：一次只执行一个 Task，当前 Task 完成验证与 review 后才进入下一 Task。
- 每个阶段完成后必须 review：检查是否破坏原有业务、是否引入新问题、是否存在过度设计/过度兜底/过度约束。
- 阶段通过标准：聚焦测试、`npm run typecheck`、必要的构建或 `git diff --check` 通过；review 未发现业务破坏和新增风险。
- 所有非必须、非必要测试文件必须清理：临时烟测、一次性验证、探索性测试在验证通过后删除；仅保留架构守护、业务回归和长期维护价值明确的测试。
- 每个阶段 review 通过后必须提交 Git：提交范围只包含当前已通过阶段及其必要依赖，禁止混入无关文件；提交后再进入下一阶段。

---

## Current Hotspots

| 文件 | 当前行数 | 问题 | 目标归属 |
| --- | ---: | --- | --- |
| `desktop/electron/agent/tools/implementations/excel/excelComBridge.ts` | 1016 | Excel COM 操作集中在单文件 | `tools/implementations/excel/*` 按 workbook/range/table/chart/pivot 分拆 |
| `desktop/electron/main-modules/ipcHandlers.ts` | 692 | IPC、Agent 会话、知识库、文件选择、工具信息混在一起 | `interaction/` + 保留 Electron 通用 IPC |
| `desktop/electron/agent/core/agentLoop/agentLoop.ts` | 696 | Turn 状态、线程恢复、压缩、流式循环仍在一个类内 | `core/agentLoop/*` |
| `desktop/electron/agent/prompts/systemPrompt.ts` | 685 | 系统提示词内容和拼接函数混在一起 | `prompts/sections/*` |
| `desktop/electron/agent/tools/registry/definitions.ts` | 677 | 所有工具定义堆在单文件 | `tools/registry/definitions/*` |
| `desktop/electron/agent/tools/registry/executors.ts` | 664 | 参数校验、Shell、Excel、Word、PPT、知识检索执行器混在一起 | `tools/executors/*` |
| `desktop/electron/agent/tools/implementations/excel/excelBridgeHelpers.ts` | 371 | PowerShell/Python/JScript/JSON helper 挂在 Excel 实现下 | `automation/*` |

## Target File Map

### Tools

- Create `desktop/electron/agent/tools/contracts/excel.ts`
  - Excel workbook、VBA、脚本、UI 桥接接口。
- Create `desktop/electron/agent/tools/contracts/office.ts`
  - Word、PowerPoint、Office script 桥接接口。
- Create `desktop/electron/agent/tools/contracts/scriptEnvironment.ts`
  - `ScriptEnvironment`、`ScriptResult` 等脚本运行契约。
- Create `desktop/electron/agent/tools/contracts/index.ts`
  - 契约层统一出口。
- Modify `desktop/electron/agent/tools/registry/interfaces.ts`
  - 临时兼容出口，只 re-export `../contracts`。
- Create `desktop/electron/agent/tools/executors/validation.ts`
  - 工具参数校验。
- Create `desktop/electron/agent/tools/executors/shellExecutor.ts`
  - Shell 工具执行器和沙箱评估入口。
- Create `desktop/electron/agent/tools/executors/excelExecutors.ts`
  - Workbook/range/formula/chart/table/pivot/VBA/Excel script/UI 执行器。
- Create `desktop/electron/agent/tools/executors/officeExecutors.ts`
  - Word/PPT/Office script 执行器。
- Create `desktop/electron/agent/tools/executors/fileExecutors.ts`
  - `file.getPaths` 等文件上下文工具。
- Create `desktop/electron/agent/tools/executors/knowledgeExecutors.ts`
  - `knowledge.search`。
- Create `desktop/electron/agent/tools/executors/index.ts`
  - `createToolExecutors` 统一装配出口。
- Modify `desktop/electron/agent/tools/registry/executors.ts`
  - 临时兼容出口，只 re-export `../executors`。
- Split `desktop/electron/agent/tools/registry/definitions.ts`
  - `definitions/workbook.ts`
  - `definitions/range.ts`
  - `definitions/formula.ts`
  - `definitions/sheet.ts`
  - `definitions/script.ts`
  - `definitions/ui.ts`
  - `definitions/file.ts`
  - `definitions/shell.ts`
  - `definitions/knowledge.ts`
  - `definitions/office.ts`
  - `definitions/index.ts`

### Automation

- Create `desktop/electron/agent/automation/powershell.ts`
  - `executePowerShell`、`psEscape`、`psVar`。
- Create `desktop/electron/agent/automation/python.ts`
  - embedded Python 检测、`getPythonPath`、`executePythonScript`、`pyVar`。
- Create `desktop/electron/agent/automation/jscript.ts`
  - `executeJScript`、`jsVar`。
- Create `desktop/electron/agent/automation/scriptEngine.ts`
  - `ScriptEngine`、`detectScriptEngine`、`executeSmart`、缓存重置。
- Create `desktop/electron/agent/automation/json.ts`
  - `safeJsonParse`。
- Create `desktop/electron/agent/automation/index.ts`
  - 自动化基础层统一出口。
- Modify `desktop/electron/agent/tools/implementations/excel/excelBridgeHelpers.ts`
  - 临时兼容出口，只保留 Excel 专属 `normalize2D` 或 re-export automation。

### Runtime And Interaction

- Create `desktop/electron/agent/runtime/agentRuntime.ts`
  - 创建 bridge、RAG、tool executors、compaction config、`AgentLoop`。
- Create `desktop/electron/agent/runtime/bridgeRegistry.ts`
  - 保存 Excel/Word/PPT/Office bridge 实例，并提供 getter。
- Create `desktop/electron/agent/runtime/knowledgeRuntime.ts`
  - 初始化 `SqliteStore`、`EmbeddingService`、`KnowledgeIndexer`、`Retriever`。
- Create `desktop/electron/agent/runtime/compactionRuntime.ts`
  - 从设置生成 `CompactionConfig`。
- Create `desktop/electron/agent/runtime/index.ts`
  - runtime 统一出口。
- Create `desktop/electron/agent/interaction/eventForwarder.ts`
  - 迁入 Agent 事件转发和工具审批。
- Create `desktop/electron/agent/interaction/ipcAgentHandlers.ts`
  - 迁入 Agent 会话、消息、工具列表、知识库相关 IPC。
- Create `desktop/electron/agent/interaction/index.ts`
  - interaction 统一出口。
- Modify `desktop/electron/main.ts`
  - 主进程只保留生命周期、窗口创建、runtime 创建、IPC 注册。
- Modify `desktop/electron/main-modules/ipcHandlers.ts`
  - 保留非 Agent 的 Electron 通用 IPC，Agent 相关逻辑迁出。
- Modify `desktop/electron/main-modules/eventForwarder.ts`
  - 临时兼容出口或删除，取决于所有导入迁移是否完成。

### Core And Prompts

- Create `desktop/electron/agent/core/agentLoop/maxTokens.ts`
  - `resolveMaxTokens`。
- Create `desktop/electron/agent/core/agentLoop/threadLifecycle.ts`
  - start/resume/reset thread 相关纯函数或小服务。
- Create `desktop/electron/agent/core/agentLoop/turnRunner.ts`
  - 单轮 turn 主流程。
- Create `desktop/electron/agent/core/agentLoop/turnState.ts`
  - active thread/turn/running/abort 状态管理。
- Modify `desktop/electron/agent/core/agentLoop/agentLoop.ts`
  - 保留公共类 API，委托给小模块。
- Create `desktop/electron/agent/prompts/sections/modelPrompt.ts`
  - 模型基础提示词：身份、工作流程、附件处理、最终回复格式。
- Create `desktop/electron/agent/prompts/sections/formulaAssistantPrompt.ts`
  - 公式助手专用提示词。
- Create `desktop/electron/agent/prompts/sections/officeToolsPrompt.ts`
  - Excel/Word/PowerPoint/Shell 工具选择提示词。
- Create `desktop/electron/agent/prompts/sections/permissionPrompt.ts`
  - 权限与命令策略提示词。
- Create `desktop/electron/agent/prompts/sections/scriptPrompt.ts`
  - 脚本规范提示词。
- Create `desktop/electron/agent/prompts/sections/qualityPrompt.ts`
  - 质量守则提示词。
- Create `desktop/electron/agent/prompts/sections/scenarioPrompt.ts`
  - 通用场景提示词。
- Create `desktop/electron/agent/prompts/sections/folderContextPrompt.ts`
  - 文件夹上下文提示词。
- Create `desktop/electron/agent/prompts/sections/index.ts`
  - 提示词 sections 统一出口。
- Modify `desktop/electron/agent/prompts/systemPrompt.ts`
  - 只保留 `buildSystemPrompt`、`appendFolderContext`、类型出口。

## Task 1: Move Tool Contracts Out Of Registry

**Files:**
- Create: `desktop/electron/agent/tools/contracts/excel.ts`
- Create: `desktop/electron/agent/tools/contracts/office.ts`
- Create: `desktop/electron/agent/tools/contracts/scriptEnvironment.ts`
- Create: `desktop/electron/agent/tools/contracts/index.ts`
- Modify: `desktop/electron/agent/tools/registry/interfaces.ts`
- Modify imports in:
  - `desktop/electron/agent/tools/registry/executors.ts`
  - `desktop/electron/agent/tools/registry/officeTools.test.ts`
  - `desktop/electron/agent/tools/implementations/excel/*.ts`
  - `desktop/electron/agent/tools/implementations/office/*.ts`
- Test: `desktop/electron/agent/tools/registry/officeTools.test.ts`

- [x] **Step 1: Write architecture guard for contracts**

Add assertions to `desktop/electron/agent/architecture.test.ts`:

```ts
expect(fileExists("tools/contracts/index.ts")).toBe(true);
expect(fileExists("tools/contracts/excel.ts")).toBe(true);
expect(fileExists("tools/contracts/office.ts")).toBe(true);
expect(fileExists("tools/contracts/scriptEnvironment.ts")).toBe(true);
```

- [x] **Step 2: Run guard and confirm it fails**

Run:

```bash
cd desktop
npm test -- electron/agent/architecture.test.ts
```

Expected: FAIL because contract files do not exist yet.

- [x] **Step 3: Move interfaces by responsibility**

Move existing interface declarations from `tools/registry/interfaces.ts`:

| Source symbol | New file |
| --- | --- |
| `ExcelWorkbookBridge` | `tools/contracts/excel.ts` |
| `ExcelVbaBridge` | `tools/contracts/excel.ts` |
| `ExcelScriptBridge` | `tools/contracts/excel.ts` |
| `ExcelUiBridge` | `tools/contracts/excel.ts` |
| `WordDocumentBridge` | `tools/contracts/office.ts` |
| `PresentationBridge` | `tools/contracts/office.ts` |
| `OfficeScriptBridge` | `tools/contracts/office.ts` |
| `ScriptEnvironment` | `tools/contracts/scriptEnvironment.ts` |
| `ScriptResult` | `tools/contracts/scriptEnvironment.ts` |

Copy each interface declaration body byte-for-byte from `tools/registry/interfaces.ts` into the listed target file, preserving method names, return types, comments, and optional markers.

```ts
// tools/contracts/index.ts
export type {
  ExcelWorkbookBridge,
  ExcelVbaBridge,
  ExcelScriptBridge,
  ExcelUiBridge,
} from "./excel";
export type {
  WordDocumentBridge,
  PresentationBridge,
  OfficeScriptBridge,
} from "./office";
export type { ScriptEnvironment, ScriptResult } from "./scriptEnvironment";
```

Keep `tools/registry/interfaces.ts` as compatibility:

```ts
export type {
  ExcelWorkbookBridge,
  ExcelVbaBridge,
  ExcelScriptBridge,
  ExcelUiBridge,
  WordDocumentBridge,
  PresentationBridge,
  OfficeScriptBridge,
  ScriptEnvironment,
  ScriptResult,
} from "../contracts";
```

- [x] **Step 4: Update imports**

Replace internal imports from `tools/registry/interfaces` with `tools/contracts` or relative `../contracts`.

- [x] **Step 5: Verify**

Run:

```bash
cd desktop
npm test -- electron/agent/architecture.test.ts electron/agent/tools/registry/officeTools.test.ts
npm run typecheck
```

Expected: all pass.

## Task 2: Split Tool Executors By Domain

**Files:**
- Create: `desktop/electron/agent/tools/executors/validation.ts`
- Create: `desktop/electron/agent/tools/executors/shellExecutor.ts`
- Create: `desktop/electron/agent/tools/executors/excelExecutors.ts`
- Create: `desktop/electron/agent/tools/executors/officeExecutors.ts`
- Create: `desktop/electron/agent/tools/executors/fileExecutors.ts`
- Create: `desktop/electron/agent/tools/executors/knowledgeExecutors.ts`
- Create: `desktop/electron/agent/tools/executors/index.ts`
- Modify: `desktop/electron/agent/tools/registry/executors.ts`
- Modify: `desktop/electron/agent/tools/registry/index.ts`
- Test: `desktop/electron/agent/tools/registry/officeTools.test.ts`

- [x] **Step 1: Add tests for executor parity**

Extend `officeTools.test.ts` or create `tools/executors/executors.test.ts` to assert:

```ts
const executors = createToolExecutors(
  workbookBridge,
  vbaBridge,
  scriptBridge,
  uiBridge,
  "D:\\temp",
  knowledgeRetriever,
  wordBridge,
  presentationBridge,
  officeScriptBridge
);

expect(executors.has("workbook.inspect")).toBe(true);
expect(executors.has("range.write")).toBe(true);
expect(executors.has("knowledge.search")).toBe(true);
expect(executors.has("word.replaceText")).toBe(true);
expect(executors.has("presentation.addSlide")).toBe(true);
```

- [x] **Step 2: Run test before split**

Run:

```bash
cd desktop
npm test -- electron/agent/tools/registry/officeTools.test.ts
```

Expected: PASS. This locks current behavior before moving code.

- [x] **Step 3: Extract validation**

Create `tools/executors/validation.ts`:

```ts
export type RequiredArgType = "string" | "number" | "array" | "object";

export function validateArgs(
  args: Record<string, unknown>,
  required: Record<string, RequiredArgType>
): string | null {
  for (const [key, expectedType] of Object.entries(required)) {
    const val = args[key];
    if (val === undefined || val === null) return `缺少必填参数: ${key}`;
    if (expectedType === "string" && typeof val !== "string") return `参数 ${key} 应为字符串，实际为 ${typeof val}`;
    if (expectedType === "number" && typeof val !== "number") return `参数 ${key} 应为数字，实际为 ${typeof val}`;
    if (expectedType === "array" && !Array.isArray(val)) return `参数 ${key} 应为数组，实际为 ${typeof val}`;
    if (expectedType === "object" && (typeof val !== "object" || Array.isArray(val))) {
      return `参数 ${key} 应为对象，实际为 ${Array.isArray(val) ? "数组" : typeof val}`;
    }
  }
  return null;
}
```

- [x] **Step 4: Extract shell executor**

Move `ShellCommandResult` and `executeShellCommand` into `tools/executors/shellExecutor.ts`. Import sandbox from `../sandbox`.

- [x] **Step 5: Extract domain executor builders**

Use small builder functions:

```ts
export function addExcelExecutors(target: Map<string, ToolExecutor>, deps: ExcelExecutorDeps): void;
export function addOfficeExecutors(target: Map<string, ToolExecutor>, deps: OfficeExecutorDeps): void;
export function addFileExecutors(target: Map<string, ToolExecutor>, deps: FileExecutorDeps): void;
export function addKnowledgeExecutors(target: Map<string, ToolExecutor>, deps: KnowledgeExecutorDeps): void;
```

Each builder mutates the passed map and owns only one domain.

- [x] **Step 6: Rebuild createToolExecutors as composition**

`tools/executors/index.ts` should become:

```ts
export function createToolExecutors(
  workbookBridge: ExcelWorkbookBridge,
  vbaBridge: ExcelVbaBridge,
  scriptBridge: ExcelScriptBridge,
  uiBridge: ExcelUiBridge,
  sessionFolderPath?: string,
  knowledgeRetriever?: Retriever,
  wordBridge?: WordDocumentBridge,
  presentationBridge?: PresentationBridge,
  officeScriptBridge?: OfficeScriptBridge
): Map<string, ToolExecutor> {
  const executors = new Map<string, ToolExecutor>();
  addExcelExecutors(executors, { workbookBridge, vbaBridge, scriptBridge, uiBridge });
  addFileExecutors(executors, { sessionFolderPath });
  addKnowledgeExecutors(executors, { knowledgeRetriever });
  addOfficeExecutors(executors, { wordBridge, presentationBridge, officeScriptBridge });
  return executors;
}
```

Keep `tools/registry/executors.ts` as compatibility:

```ts
export { createToolExecutors, executeShellCommand };
export type { ShellCommandResult };
```

- [x] **Step 7: Verify**

Run:

```bash
cd desktop
npm test -- electron/agent/tools/registry/officeTools.test.ts electron/agent/architecture.test.ts
npm run typecheck
```

Expected: all pass.

## Task 3: Split Tool Definitions By Domain

**Files:**
- Create: `desktop/electron/agent/tools/registry/definitions/*.ts`
- Modify: `desktop/electron/agent/tools/registry/definitions.ts`
- Modify: `desktop/electron/agent/tools/registry/index.ts`
- Test: `desktop/electron/agent/tools/registry/officeTools.test.ts`

- [x] **Step 1: Add definition count guard**

In `officeTools.test.ts`, assert the full ordered tool catalog still includes current Excel, file, shell, knowledge, Word, PPT, and Office script names.

- [x] **Step 2: Create grouped definition modules**

Move contiguous tool definition objects into domain files. Each file exports an array:

For example, `definitions/workbook.ts` receives only tool objects whose `name` starts with `workbook.` and exports them as `WORKBOOK_TOOL_DEFINITIONS`. The same rule applies to the other files: `range.`, `formula.`, `sheet.`, script-related tools, `ui.`, file tools, shell tools, knowledge tools, and Office tools.

```ts
import type { ToolDefinition } from "../../../shared/types";

export const WORKBOOK_TOOL_DEFINITIONS: ToolDefinition[] = [];
```

Replace the empty array with the existing `workbook.*` tool objects copied from `tools/registry/definitions.ts` during the move. Do not change names, descriptions, schemas, or risk levels in this task.

- [x] **Step 3: Rebuild aggregate**

`tools/registry/definitions/index.ts`:

```ts
export const ALL_TOOL_DEFINITIONS: ToolDefinition[] = [
  ...WORKBOOK_TOOL_DEFINITIONS,
  ...RANGE_TOOL_DEFINITIONS,
  ...FORMULA_TOOL_DEFINITIONS,
  ...SHEET_TOOL_DEFINITIONS,
  ...SCRIPT_TOOL_DEFINITIONS,
  ...UI_TOOL_DEFINITIONS,
  ...FILE_TOOL_DEFINITIONS,
  ...SHELL_TOOL_DEFINITIONS,
  ...KNOWLEDGE_TOOL_DEFINITIONS,
  ...OFFICE_TOOL_DEFINITIONS,
];

export const TOOL_DEFINITIONS_MAP = new Map<string, ToolDefinition>(
  ALL_TOOL_DEFINITIONS.map((tool) => [tool.name, tool])
);
```

Keep `tools/registry/definitions.ts` as compatibility re-export.

- [x] **Step 4: Verify**

Run:

```bash
cd desktop
npm test -- electron/agent/tools/registry/officeTools.test.ts electron/agent/systemPrompt.test.ts
npm run typecheck
```

Expected: all pass.

## Task 4: Extract Shared Automation From Excel Helpers

**Files:**
- Create: `desktop/electron/agent/automation/powershell.ts`
- Create: `desktop/electron/agent/automation/python.ts`
- Create: `desktop/electron/agent/automation/jscript.ts`
- Create: `desktop/electron/agent/automation/scriptEngine.ts`
- Create: `desktop/electron/agent/automation/json.ts`
- Create: `desktop/electron/agent/automation/index.ts`
- Modify: `desktop/electron/agent/tools/implementations/excel/excelBridgeHelpers.ts`
- Modify imports in Excel and Office bridge modules.
- Test: `desktop/electron/agent/tools/implementations/excel/excelBridgeHelpers.test.ts`

- [x] **Step 1: Write automation import guard**

Add architecture assertions:

```ts
expect(fileExists("automation/index.ts")).toBe(true);
expect(fileExists("automation/powershell.ts")).toBe(true);
expect(fileExists("automation/python.ts")).toBe(true);
expect(fileExists("automation/jscript.ts")).toBe(true);
expect(fileExists("automation/scriptEngine.ts")).toBe(true);
expect(fileExists("automation/json.ts")).toBe(true);
```

- [x] **Step 2: Move helpers without behavior change**

Move existing helper bodies exactly, then re-export from `excelBridgeHelpers.ts`:

```ts
export {
  executePowerShell,
  psEscape,
  psVar,
  executePythonScript,
  executeJScript,
  detectScriptEngine,
  executeSmart,
  resetEngineCache,
  safeJsonParse,
} from "../../../automation";
```

Keep `normalize2D` in Excel helper unless another domain uses it.

- [x] **Step 3: Update direct imports**

Office bridge modules should import shared execution helpers from `automation`, not from Excel helper files.

- [x] **Step 4: Verify**

Run:

```bash
cd desktop
npm test -- electron/agent/tools/implementations/excel/excelBridgeHelpers.test.ts electron/agent/tools/registry/officeTools.test.ts electron/agent/architecture.test.ts
npm run typecheck
```

Expected: all pass.

## Task 5: Extract Agent Runtime Assembly

**Files:**
- Create: `desktop/electron/agent/runtime/bridgeRegistry.ts`
- Create: `desktop/electron/agent/runtime/knowledgeRuntime.ts`
- Create: `desktop/electron/agent/runtime/compactionRuntime.ts`
- Create: `desktop/electron/agent/runtime/agentRuntime.ts`
- Create: `desktop/electron/agent/runtime/index.ts`
- Modify: `desktop/electron/main.ts`
- Modify: `desktop/electron/main-modules/ipcHandlers.ts`
- Modify: `desktop/electron/main-modules/settingsManager.ts`

- [x] **Step 1: Add runtime smoke test**

Create `desktop/electron/agent/runtime/compactionRuntime.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { buildCompactionConfig } from "./compactionRuntime";

describe("buildCompactionConfig", () => {
  it("derives threshold from context window and percent", () => {
    expect(buildCompactionConfig({
      contextWindowSize: 100_000,
      savedCompaction: { enabled: true, autoCompactThresholdPercent: 75 },
    })).toMatchObject({
      enabled: true,
      autoCompactTokenThreshold: 75_000,
      contextWindowSize: 100_000,
    });
  });
});
```

- [x] **Step 2: Move bridge state**

`bridgeRegistry.ts` owns Excel/Word/PPT bridge creation and getters. `main.ts` should not keep individual bridge globals after this task.

- [x] **Step 3: Move RAG initialization**

`knowledgeRuntime.ts` owns Sqlite path, store/embedder/indexer/retriever creation, and `setKnowledge*` registration.

- [x] **Step 4: Move AgentLoop creation**

`agentRuntime.ts` exports:

```ts
export async function getOrCreateAgentRuntime(deps: AgentRuntimeDeps): Promise<AgentRuntime>;
```

It returns `{ agentLoop, bridges, knowledge }`.

- [x] **Step 5: Keep Electron lifecycle in main**

`main.ts` keeps:

```ts
app.whenReady().then(async () => {
  getSessionStoreInstance();
  await getOrCreateAgentRuntime({
    getActiveAIConfig,
    getSettingsStore,
    getSessionStoreInstance,
    getMainWindow: () => mainWindow,
  });
  registerIpcHandlers();
  applySandboxConfig();
  recreateMainWindow();
});
```

- [x] **Step 6: Verify**

Run:

```bash
cd desktop
npm test -- electron/agent/runtime/compactionRuntime.test.ts electron/agent/architecture.test.ts
npm run typecheck
```

Expected: all pass.

## Task 6: Move Agent Interaction Out Of main-modules

**Files:**
- Create: `desktop/electron/agent/interaction/eventForwarder.ts`
- Create: `desktop/electron/agent/interaction/ipcAgentHandlers.ts`
- Create: `desktop/electron/agent/interaction/index.ts`
- Modify: `desktop/electron/main-modules/eventForwarder.ts`
- Modify: `desktop/electron/main-modules/ipcHandlers.ts`
- Modify: `desktop/electron/main.ts`

- [x] **Step 1: Classify IPC handlers**

Split `ipcHandlers.ts` by responsibility:

```text
Agent-owned:
- chat/send
- chat/interrupt
- chat/newThread
- chat/resumeThread
- agent/toolDefinitions
- knowledge/*
- session/thread operations that directly touch AgentLoop or SessionStore

Electron-owned:
- dialog/open
- shell/openExternal
- clipboard/*
- app/window/settings UI plumbing that does not call AgentLoop
```

- [x] **Step 2: Move event forwarding**

Move `pendingApprovals`, `createEventForwarder`, `requestToolApproval`, `registerToolApprovalHandlers` into `agent/interaction/eventForwarder.ts`.

Keep `main-modules/eventForwarder.ts` temporarily:

```ts
export {
  pendingApprovals,
  createEventForwarder,
  requestToolApproval,
  registerToolApprovalHandlers,
} from "../agent/interaction/eventForwarder";
```

- [x] **Step 3: Move Agent IPC handlers**

Create `registerAgentIpcHandlers(deps)` in `agent/interaction/ipcAgentHandlers.ts`. It receives getters for window/runtime/settings instead of importing `main.ts`.

- [x] **Step 4: Verify**

Run:

```bash
cd desktop
npm test -- electron/agent/architecture.test.ts electron/agent/core/agentLoop/agentLoop.test.ts
npm run typecheck
npm run build
```

Expected: all pass; build may keep existing Vite chunk-size warnings.

## Task 7: Split AgentLoop Internals

**Files:**
- Create: `desktop/electron/agent/core/agentLoop/maxTokens.ts`
- Create: `desktop/electron/agent/core/agentLoop/turnState.ts`
- Create: `desktop/electron/agent/core/agentLoop/threadLifecycle.ts`
- Create: `desktop/electron/agent/core/agentLoop/turnRunner.ts`
- Modify: `desktop/electron/agent/core/agentLoop/agentLoop.ts`
- Modify: `desktop/electron/agent/core/agentLoop/agentLoop.test.ts`

- [x] **Step 1: Extract max token calculation**

Move `resolveMaxTokens` to `maxTokens.ts` and add direct unit tests for explicit max, off/low/high/max reasoning modes.

- [x] **Step 2: Extract state container**

Move active thread/turn/running/abort bookkeeping into `TurnState`. Keep public API unchanged on `AgentLoop`.

- [x] **Step 3: Extract thread lifecycle**

Move `startThread`, `resumeThread`, `resetThread`, and folder pending logic into a small lifecycle helper. `AgentLoop` delegates to it.

- [x] **Step 4: Extract turn runner**

Move Turn data construction into `turnRunner.ts`. Review decision: the streaming/tool/compaction loop remains in `agentLoop.ts` for this stage because moving it now would require passing many mutable callbacks and state references, increasing coupling instead of reducing it.

- [x] **Step 5: Verify**

Run:

```bash
cd desktop
npm test -- electron/agent/core/agentLoop/agentLoop.test.ts electron/agent/core/agentLoop/maxTokens.test.ts
npm run typecheck
```

Expected: all pass.

## Task 8: Split System Prompt Sections

**Files:**
- Create: `desktop/electron/agent/prompts/sections/modelPrompt.ts`
- Create: `desktop/electron/agent/prompts/sections/formulaAssistantPrompt.ts`
- Create: `desktop/electron/agent/prompts/sections/officeToolsPrompt.ts`
- Create: `desktop/electron/agent/prompts/sections/permissionPrompt.ts`
- Create: `desktop/electron/agent/prompts/sections/scriptPrompt.ts`
- Create: `desktop/electron/agent/prompts/sections/qualityPrompt.ts`
- Create: `desktop/electron/agent/prompts/sections/scenarioPrompt.ts`
- Create: `desktop/electron/agent/prompts/sections/folderContextPrompt.ts`
- Create: `desktop/electron/agent/prompts/sections/index.ts`
- Modify: `desktop/electron/agent/prompts/systemPrompt.ts`
- Modify: `desktop/electron/agent/prompts/systemPrompt.test.ts`

- [x] **Step 1: Add prompt snapshot-ish assertions**

Extend `systemPrompt.test.ts` to assert the prompt still contains Excel, Word, PowerPoint, shell approval, and folder context guidance.

- [x] **Step 2: Move static sections**

Each `sections/*.ts` exports one string constant or builder function. 按业务入口分类为模型提示词、公式助手提示词、Office/工具提示词、权限提示词、脚本提示词、质量提示词、通用场景提示词和文件夹上下文提示词。The first migration keeps existing section wording unchanged, so prompt behavior changes are covered only by the assertions from Step 1.

- [x] **Step 3: Keep systemPrompt as composer**

`systemPrompt.ts` should expose:

```ts
export function buildSystemPrompt(): string {
  return [
    ROLE_PROMPT_SECTION,
    WORKBOOK_PROMPT_SECTION,
    OFFICE_PROMPT_SECTION,
    TOOLS_PROMPT_SECTION,
    SAFETY_PROMPT_SECTION,
  ].join("\n\n");
}
```

Move folder context formatting into `sections/folderContextPrompt.ts`.

- [x] **Step 4: Verify**

Run:

```bash
cd desktop
npm test -- electron/agent/prompts/systemPrompt.test.ts electron/agent/architecture.test.ts
npm run typecheck
```

Expected: all pass.

## Task 9: Split ExcelComBridge By Capability

**Files:**
- Create: `desktop/electron/agent/tools/implementations/excel/connectionOperations.ts`
- Create: `desktop/electron/agent/tools/implementations/excel/workbookOperations.ts`
- Create: `desktop/electron/agent/tools/implementations/excel/rangeOperations.ts`
- Create: `desktop/electron/agent/tools/implementations/excel/formulaOperations.ts`
- Create: `desktop/electron/agent/tools/implementations/excel/sheetOperations.ts`
- Modify: `desktop/electron/agent/tools/implementations/excel/excelComBridge.ts`
- Modify: `desktop/electron/agent/tools/implementations/excel/excelComBridge.test.ts`

- [x] **Step 1: Add behavior tests for moved capabilities**

Keep tests at bridge boundary. Add architecture and compatibility assertions for the moved connection/workbook/range/formula/sheet capability modules. Review note: current `ExcelComBridge` has no public chart/table/pivot methods, so no empty modules are created for those names.

- [x] **Step 2: Extract pure script builders first**

Move PowerShell/JScript/Python script builders into operation modules before moving public class methods. This reduces risk because script text remains grouped by capability while `ExcelComBridge` keeps COM state.

- [x] **Step 3: Delegate class methods**

`ExcelComBridge` should remain the public facade that satisfies `ExcelWorkbookBridge`; each method delegates to capability modules.

- [x] **Step 4: Verify**

Run:

```bash
cd desktop
npm test -- electron/agent/tools/implementations/excel/excelComBridge.test.ts electron/agent/tools/registry/officeTools.test.ts
npm run typecheck
```

Expected: all pass.

## Final Verification

Run after all tasks:

```bash
cd desktop
npm test
npm run typecheck
npm run build
cd ..
git diff --check
```

Expected:
- All Vitest tests pass.
- TypeScript typecheck passes.
- Build passes; existing Vite CJS/chunk-size warnings are acceptable unless new errors appear.
- `git diff --check` has no whitespace errors. LF/CRLF warnings are acceptable if they match the existing repository behavior.

## Execution Order

1. Contracts
2. Executors
3. Tool definitions
4. Automation
5. Runtime
6. Interaction
7. AgentLoop
8. Prompts
9. ExcelComBridge

This order keeps behavioral blast radius small: each stage first adds a new home, keeps a compatibility export, updates direct imports, then runs focused tests before moving to the next high-coupling file.
