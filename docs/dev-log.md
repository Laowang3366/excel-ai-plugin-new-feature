# 开发日志

## 模块化重构记录

### 重构目标
将大型文件按「高内聚、低耦合」原则拆分为独立模块，遵守代码规范：
- 单文件 ≤ 400 行
- 单文件单职责
- React 组件 ≤ 300 行
- TypeScript 编译通过 + Vite 构建通过

---

### ChatPage.tsx 拆分 (1,378 → 300 行)

| 拆分出的模块 | 行数 | 职责 |
|---|---|---|
| `utils/chatHelpers.tsx` | 202 | 纯函数：消息分组、时长计算、标题摘要、Excel 选区、格式化、小组件 |
| `components/common/FloatingTaskPanel.tsx` | 178 | 可拖拽、可折叠的通用浮窗容器 |
| `components/chat/AssistantGroupBlock.tsx` | 240 | 助手消息组渲染（折叠/展开/流式） |
| `components/chat/ChatMessageList.tsx` | 141 | 消息列表区域（空状态、分组、中断恢复、错误） |
| `components/chat/ComposerArea.tsx` | 350 | Pill Composer 输入框（附件/权限/思考/模型切换/发送） |
| `hooks/useComposer.ts` | 267 | 输入框状态与交互逻辑 |
| `hooks/useTaskDrafts.ts` | 80 | 任务面板草稿管理 |

### Sidebar.tsx 拆分 (786 → 436 行)

| 拆分出的模块 | 行数 | 职责 |
|---|---|---|
| `utils/sidebarHelpers.ts` | 75 | 意图常量、时间格式化、状态判断纯函数 |
| `hooks/useExcelConnection.ts` | 69 | Excel/WPS 连接状态管理与轮询 |
| `components/sidebar/FolderSection.tsx` | 252 | 文件夹分组渲染 + 未分组会话列表 |
| `components/sidebar/ThreadContextMenu.tsx` | 136 | 会话右键菜单（删除确认/移动到文件夹） |

### chatStore.ts 拆分 (645 → 418 行)

| 拆分出的模块 | 行数 | 职责 |
|---|---|---|
| `store/agentEventHandler.ts` | 144 | Agent 事件 → 状态 patches 的纯函数 |
| `store/threadActions.ts` | 140 | 会话管理（load/switch/create/delete/move） |

### main.ts 拆分 (947 → 124 行)

| 拆分出的模块 | 行数 | 职责 |
|---|---|---|
| `electron/main-modules/settingsManager.ts` | 198 | 持久化配置管理、数据路径迁移、AI 配置 |
| `electron/main-modules/windowManager.ts` | 130 | 窗口创建、托盘管理、主题应用 |
| `electron/main-modules/ipcHandlers.ts` | 496 | 所有 IPC handle 注册（按功能域分类） |
| `electron/main-modules/eventForwarder.ts` | 94 | Agent 事件转发器 + 工具审批 |

### agentLoop.ts 拆分 (1,101 → 1,054 行) [Week 1-3]

| 拆分出的模块 | 行数 | 职责 |
|---|---|---|
| `electron/agent/imageAttachmentResolver.ts` | 54 | 本地图片路径 → base64 data URI |
| `electron/agent/toolExecution.ts` | 120 | 工具审批判断、审批流程、工具执行 |
| `electron/agent/compactionManager.ts` | 134 | 会话压缩、轮次间压缩、摘要生成 |

### agentLoop.ts 深度拆分 (1,066 → 6 子模块) [Week 4]

| 拆分出的模块 | 行数 | 职责 |
|---|---|---|
| `electron/agent/agentLoop/streamCollector.ts` | ~140 | 流式事件收集，text/reasoning/toolCall 分类 |
| `electron/agent/agentLoop/toolExecutor.ts` | ~260 | 工具执行、审批判断、alwaysAllow 管理 |
| `electron/agent/agentLoop/compactionManager.ts` | ~130 | pre-turn / mid-turn 压缩、摘要生成 |
| `electron/agent/agentLoop/buildStreamParams.ts` | ~90 | 推理模式降级链、系统提示词构建 |
| `electron/agent/agentLoop/agentLoop.ts` | ~450 | AgentLoop 类编排器，对外 API 不变 |
| `electron/agent/agentLoop/index.ts` | ~45 | barrel re-export，消费方无感知 |

### toolRegistry.ts 拆分 (1,228 行) [Week 2-3]

| 拆分出的模块 | 行数 | 职责 |
|---|---|---|
| `electron/agent/toolRegistry/interfaces.ts` | ~110 | 桥接接口定义（ExcelWorkbookBridge 等） |
| `electron/agent/toolRegistry/definitions.ts` | ~460 | 工具定义常量 + ALL_TOOL_DEFINITIONS + TOOL_DEFINITIONS_MAP |
| `electron/agent/tools/data/excelFunctionCatalog.ts` | ~380 | Excel 函数目录数据与搜索 |
| `electron/agent/toolRegistry/executors.ts` | ~450 | 参数校验、Shell 执行、工具执行器工厂 |
| `electron/agent/toolRegistry/index.ts` | ~40 | barrel re-export |

### excelBridge.ts 拆分 [Week 2-3]

| 拆分出的模块 | 行数 | 职责 |
|---|---|---|
| `electron/agent/excelBridge/comBridge.ts` | — | COM 自动化桥接（Excel/WPS 连接检测） |
| `electron/agent/excelBridge/vbaBridge.ts` | — | VBA 宏执行桥接 |
| `electron/agent/excelBridge/scriptBridge.ts` | — | JS/TS 脚本执行桥接 |
| `electron/agent/excelBridge/uiBridge.ts` | — | UI 控件操作桥接 |
| `electron/agent/excelBridge/index.ts` | — | barrel re-export |

### CSS 模块化拆分 (4,827 → 23 文件) [Week 1-3]

| 模块 | 行数 | 职责 |
|---|---|---|
| `tokens.css` | — | CSS 变量（颜色/字体/间距/圆角） |
| `base.css` | — | 全局重置 + body/html |
| `animations.css` | — | 关键帧动画 |
| `buttons.css` | — | 按钮组件样式 |
| `app-layout.css` | — | 整体布局 |
| `sidebar.css` | — | 侧边栏 |
| `chat.css` | — | 对话区域 |
| `message-bubble.css` | — | 消息气泡 |
| `reasoning.css` | — | 推理展示 |
| `tool-call.css` | — | 工具调用 |
| `streaming.css` | — | 流式输出 |
| `composer.css` | — | 输入框 |
| `task-panel.css` | — | 任务面板 |
| `settings.css` | — | 设置页 |
| `model-config.css` | — | 模型配置 |
| `dialog.css` | — | 弹窗 |
| `model-quick-switch.css` | — | 模型快速切换 |
| `tool-confirm.css` | — | 工具审批弹窗 |
| `loading-spinner.css` | — | 加载动画 |
| `error.css` | — | 错误展示 |
| `compaction.css` | — | 压缩提示 |
| `resume-hint.css` | — | 恢复提示 |
| `global.css`（入口） | 51 | `@import` 入口文件 |

---

### 构建验证
- TypeScript `tsc --noEmit`: ✓ 通过
- Vite `vite build`: ✓ 通过（renderer + main + preload）
- vitest `npm test`: ✓ 35 tests 全部通过
- 代码行数统计：重构后模块总计 ~3,000 行，新增 ~1,600 行模块代码

---

## 2026-06-24 — 第一周代码审查修复（Week 1）

### 分支
`feature/new-feature`

### 改动概览
4 项 P0 Bug 修复（#1–#4），涉及对话信息显示顺序、API 400 错误、Electron 类型检查、AgentLoopConfig。

---

### #1 TurnItem 显示顺序修复

**问题描述**：对话区消息显示顺序不符合 API 真实返回顺序。流式阶段就立即发出 `tool_call` 的 `item_started`，导致 UI 闪烁和排序混乱。

**根因分析**：
- `agentLoop.ts` 流式处理阶段在收到 tool_call 时立即 emit `item_started`，而非等到工具执行完毕后按顺序发出
- `chatStore.ts` 的 `item_completed` handler 做了"智能插入"（按类型和 ID 匹配），逻辑复杂且有 bug
- `ChatPage.tsx` 的 `sortItemsByRound` 按固定类型排序，忽略了 API 流式响应的真实顺序

**改动详情**：
- `agentLoop.ts`：流式阶段不再立即发出 `tool_call` 的 `item_started`；流结束后按 `reasoning → assistant_message → tool_call` 顺序依次发出所有 items
- `chatStore.ts`：`item_completed` handler 简化为按事件到达顺序 `push`，不再做智能插入
- `ChatPage.tsx`：流式渲染顺序调整为思考在前、正文在后；`sortItemsByRound` 组内排序调整为 `reasoning → assistant_message → tool_call → tool_result`（JSONL 历史恢复兜底）
- 折叠"本轮工作时长"时，`phase=commentary` 的阶段性正文也一起折叠，只显示 `phase=final` 的总结正文

**验证结果**：TypeScript 编译通过，Vite 构建通过

---

### #2 API 400 错误修复（孤立 tool_call）

**问题描述**：assistant 消息含 `tool_calls` 但缺少对应 `tool` result 消息，导致后续请求被 AI API 拒绝（`invalid_request_error`）。

**根因分析**：
- 中断或压缩后恢复时，历史记录中可能存在"孤立 tool_call"——有 call 但无对应的 result
- AI API 要求 assistant 的 `tool_calls` 必须紧跟对应的 tool result 消息
- `turnItemsToChatMessages()` 之前没有处理这种情况

**改动详情**：
- `aiClient.ts` 中 `turnItemsToChatMessages()` 重写为三遍处理：
  - 第一遍：识别孤立 `tool_call`（有 call 无 result），收集其 ID
  - 第二遍：构建消息时跳过孤立项
  - 第三遍：清理空壳 assistant 消息（tool_calls 全被跳过后 content 为空）
- `buildRequestMessages()` 新增双重安全校验：
  - 首次校验：全局 ID 匹配过滤未配对的 `toolCalls`
  - 二次校验：确保连续性（assistant 后紧跟对应 tool 消息）
- 孤立 `tool_call` 采用跳过策略而非合成假 result，保持上下文一致性

**验证结果**：TypeScript 编译通过

---

### #3 Electron 类型检查修复

**问题描述**：Agent loop 和 Excel bridge 存在 TypeScript 类型错误，`tsc --noEmit` 不通过。

**改动详情**：
- 修复 Agent loop 中的类型标注
- 修复 Excel bridge（COM 自动化）的类型签名
- 确保 `electron/agent/` 下所有文件类型检查通过

**验证结果**：`tsc --noEmit` 通过

---

### #4 AgentLoopConfig 缺少 reasoningMode

**问题描述**：`AgentLoopConfig` 接口缺少 `reasoningMode` 属性，编译时类型不匹配。

**改动详情**：
- `AgentLoopConfig` 接口新增可选 `reasoningMode` 属性
- `agentLoop.ts` 从 `aiConfig.reasoningMode` 读取，支持顶层覆盖
- 默认值为 `"off"`（不启用推理）

**验证结果**：TypeScript 编译通过

---

## 2026-06-24 — 第二周代码审查修复（Week 2）

### 分支
`feature/new-feature`

### 改动概览
4 项中高优先级改进（#5–#8），涵盖 IPC 校验、日志系统、N+1 查询优化、excelBridge 拆分。

---

### #5 IPC Zod Schema 输入验证

**问题描述**：所有 IPC 通道直接使用 `args` 参数，无运行时类型校验，恶意或错误输入可能导致崩溃。

**设计方案**：引入 zod schema，对每个 IPC 通道的输入进行运行时校验。

**改动详情**：
- `electron/main-modules/ipcHandlers.ts`：所有 `ipcMain.handle` 回调新增 zod schema 校验
- 覆盖的 IPC 通道：`agent:startTurn`、`agent:continueTurn`、`thread:load`、`thread:delete`、`thread:resume`、`thread:updateMetadata`、`settings:get`、`settings:set`、`tool:confirm`、`tool:cancel` 等
- 校验失败时返回 `{ success: false, error: "Invalid parameters" }` 而非抛异常

**验证结果**：TypeScript 编译通过

---

### #6 结构化日志系统

**问题描述**：项目使用 `console.log` 输出日志，无格式化、无级别控制、无文件输出。

**设计方案**：新建 `electron/agent/logger.ts`，提供结构化日志。

**改动详情**：
- 新建 `electron/agent/logger.ts`：
  - 日志级别：`debug` / `info` / `warn` / `error`
  - 输出格式：JSON（生产）+ 彩色终端（开发）
  - 文件输出：支持写入日志文件
  - 上下文标记：支持 `context` 字段区分来源（如 `agentLoop`、`ipcHandler`）
- `agentLoop.ts` 中的 `console.log` 替换为 logger 调用

**验证结果**：TypeScript 编译通过

---

### #7 UsageStats N+1 查询优化

**问题描述**：UsageStats 获取使用统计时，对每个 turn 逐个查询，导致 N+1 性能问题。

**改动详情**：
- `electron/main-modules/ipcHandlers.ts` 中 `stats:getSummary` 通道重写
- 单次查询替代循环，减少 IPC 调用次数
- 聚合结果在主进程一次性计算完成

**验证结果**：TypeScript 编译通过

---

### #8 excelBridge.ts 模块拆分

**问题描述**：`excelBridge.ts` 包含多个桥接职责（COM 自动化、VBA 执行、脚本执行、UI 操作），文件过大。

**改动详情**：
- 拆分为 `excelBridge/` 目录，包含 5 个子模块：
  - `comBridge.ts`：COM 自动化桥接（Excel/WPS 连接检测、Range 读写）
  - `vbaBridge.ts`：VBA 宏执行桥接
  - `scriptBridge.ts`：JS/TS 脚本执行桥接
  - `uiBridge.ts`：UI 控件操作桥接
  - `index.ts`：barrel re-export，消费方 import 路径不变

**验证结果**：TypeScript 编译通过，Vite 构建通过

---

## 2026-06-24 — 第三周代码审查修复（Week 3）

### 分支
`feature/new-feature`

### 改动概览
4 项中高优先级改进（#9–#12），涵盖 toolRegistry 拆分、agentLoop 首轮拆分、main.ts 拆分、chatStore 拆分。

---

### #9 toolRegistry.ts 拆分（1,228 行 → 5 子模块）

**问题描述**：`toolRegistry.ts` 包含接口定义、工具定义、Excel 函数数据库、执行器工厂，文件超过 1,200 行。

**改动详情**：
- 拆分为 `toolRegistry/` 目录，包含 5 个子模块：
  - `interfaces.ts`（~110 行）：桥接接口定义（ExcelWorkbookBridge、ExcelVbaBridge 等）
  - `definitions.ts`（~460 行）：工具定义常量 + `ALL_TOOL_DEFINITIONS` + `TOOL_DEFINITIONS_MAP`
  - `tools/data/excelFunctionCatalog.ts`（~380 行）：Excel 函数目录数据与搜索（名称、描述、参数、分类）
  - `executors.ts`（~450 行）：参数校验、Shell 执行、工具执行器工厂 `createToolExecutors()`
  - `index.ts`（~40 行）：barrel re-export，消费方 import 路径不变

**验证结果**：TypeScript 编译通过

---

### #10 agentLoop.ts 首轮拆分（1,101 行 → 3 子模块）

**问题描述**：`agentLoop.ts` 包含图片附件解析、工具执行、压缩管理等多种职责，超过 1,100 行。

**改动详情**：
- 新建 3 个子模块：
  - `electron/agent/imageAttachmentResolver.ts`（54 行）：本地图片路径 → base64 data URI 转换
  - `electron/agent/toolExecution.ts`（120 行）：工具审批判断、审批流程、工具执行
  - `electron/agent/compactionManager.ts`（134 行）：会话压缩、轮次间压缩、摘要生成
- `agentLoop.ts` 从 1,101 行降至 1,054 行

**验证结果**：TypeScript 编译通过

---

### #11 main.ts 拆分（947 行 → 124 行 + 4 子模块）

**问题描述**：`main.ts` 包含窗口管理、设置管理、IPC 注册、事件转发，超过 900 行。

**改动详情**：
- 拆分为 `electron/main-modules/` 目录，包含 4 个子模块：
  - `settingsManager.ts`（198 行）：持久化配置管理、数据路径迁移、AI 配置加载
  - `windowManager.ts`（130 行）：窗口创建、托盘管理、主题应用
  - `ipcHandlers.ts`（496 行）：所有 IPC handle 注册（按功能域分类）
  - `eventForwarder.ts`（94 行）：Agent 事件转发器 + 工具审批回调
- `main.ts` 从 947 行降至 124 行（仅保留初始化入口）

**验证结果**：TypeScript 编译通过，Vite 构建通过

---

### #12 chatStore.ts 拆分（645 行 → 418 行 + 2 子模块）

**问题描述**：`chatStore.ts` 包含事件处理和会话管理，超过 600 行。

**改动详情**：
- 新建 2 个子模块：
  - `store/agentEventHandler.ts`（144 行）：Agent 事件 → 状态 patches 的纯函数（`handleAgentEvent()`）
  - `store/threadActions.ts`（140 行）：会话管理（`loadThreads`、`switchThread`、`createNewThread`、`deleteThread`、`moveThreadToFolder`）
- `chatStore.ts` 从 645 行降至 418 行

**验证结果**：TypeScript 编译通过

---

## 2026-06-25 — 第四周代码审查修复

### 分支
`feature/new-feature`

### 改动概览
6 项代码审查问题全部修复完成（#18–#23），涵盖模块拆分、Props 精简、持久化优化、单元测试、Mock 标记、IPC 依赖注入。

---

### #18 agentLoop.ts 深度拆分（1,066 行 → 6 子模块）

**问题描述**：agentLoop.ts 虽在 Week 1-3 已做初步拆分（imageAttachmentResolver / toolExecution / compactionManager），但仍超过 1,000 行，职责混合。

**设计方案**：按运行时职责域进一步拆分为 4 个独立子模块 + 1 个编排器 + 1 个 barrel：

| 子模块 | 文件 | 行数 | 职责 |
|---|---|---|---|
| streamCollector | `agentLoop/streamCollector.ts` | ~140 | 消费 AI streamChat 事件，分类收集 text/reasoning/toolCall |
| toolExecutor | `agentLoop/toolExecutor.ts` | ~260 | 工具执行、审批判断（shell.execute 始终需审批）、alwaysAllow 管理 |
| compactionManager | `agentLoop/compactionManager.ts` | ~130 | pre-turn / mid-turn 压缩、摘要生成 |
| buildStreamParams | `agentLoop/buildStreamParams.ts` | ~90 | 推理模式降级链（max→high→medium→low→off）、系统提示词构建 |
| agentLoop（编排器） | `agentLoop/agentLoop.ts` | ~450 | AgentLoop 类，对外 API 不变，调用子模块 |
| index | `agentLoop/index.ts` | ~45 | barrel re-export，消费方 import 路径不变 |

**改动详情**：
- 新建 `agentLoop/` 目录，包含 6 个文件
- 原 `electron/agent/agentLoop.ts`（1,066 行）删除
- `electron/agent/index.ts` 的 `export * from "./agentLoop"` 自动解析到 `agentLoop/index.ts`
- 消费方（`main.ts`、`ipcHandlers.ts`、`settingsManager.ts`）import 路径不变

**关键设计**：
- `streamCollector.ts`：`collectStreamEvents()` 消费 streamChat 事件，`emitInterruptedProgress()` 在 AbortError 时保存部分进度
- `toolExecutor.ts`：`shouldRequireApproval()` 封装权限模式逻辑，`processToolCalls()` 批量处理工具调用
- `buildStreamParams.ts`：`DOWNGRADE_MAP` 定义推理模式降级链，第 1 轮使用配置值，后续轮次逐步降级
- `compactionManager.ts`：`performAutoCompaction()` / `performMidTurnCompaction()` 分离 pre-turn 和 mid-turn 两种压缩场景

**验证结果**：TypeScript 编译通过

---

### #19 ComposerArea Props 精简（28 → 7）

**问题描述**：ComposerArea 组件接收 28 个独立 props，大部分来自 useComposer hook，导致调用方代码冗长。

**设计方案**：将 useComposer hook 返回值整体作为 `composer` prop 传入。

**改动详情**：
- 新增 `ComposerState` 类型：`type ComposerState = ReturnType<typeof useComposer>`
- `ComposerArea` Props 从 28 个减为 7 个：`composer`、`currentFolder`、`currentFolderFiles`、`showWelcomeComposer`、`onSend`、`onInterrupt`、`onOpenSettings`
- 组件内部通过 `composer.inputText`、`composer.handleSend` 等解构使用
- `ChatPage.tsx` 调用处同步简化

**验证结果**：TypeScript 编译通过

---

### #20 settingsStore 增量持久化（9 个 IPC → 变更 key）

**问题描述**：settingsStore 的 `saveSettings()` 每次变更都写入全部 9 个 settings key，浪费 IPC 调用。

**设计方案**：新增 `savePartial(keys, get)` 函数，仅写入实际变更的 key。

**改动详情**：
- 新增 `KEY_MAP` 常量：`Partial<Record<keyof SettingsState, string>>`，映射 state 字段 → electron-store key
- 新增 `COMPACTION_FIELDS` 数组：`compactionEnabled` 和 `autoCompactThresholdPercent` 组合为一个嵌套 key
- 新增 `savePartial(keys, get)` 异步函数：按需并行写入变更 key（`Promise.all`）
- 13 个 setter action 全部改为 `savePartial([specificKey], get)`
- `saveSettings()` 保留向后兼容（无外部调用方）

**验证结果**：TypeScript 编译通过

---

### #21 单元测试基础设施（vitest）

**问题描述**：项目无单元测试，代码改动无法验证。

**设计方案**：引入 vitest + @vitest/coverage-v8，先覆盖 compaction 和 agentLoop 子模块。

**改动详情**：
- 新建 `vitest.config.ts`：node 环境、v8 覆盖率、globals 启用、10s 超时
- `package.json` 新增 `test` / `test:watch` / `test:coverage` 脚本
- 新建 `electron/agent/compaction.test.ts`（24 tests）：覆盖 `estimateTokens`、`estimateItemsTokens`、`shouldCompact`、`collectUserMessages`、`buildCompactedHistory`、`performCompaction`、`historyToCompactPrompt`、`buildResumeContext`
- 新建 `electron/agent/agentLoop/agentLoop.test.ts`（11 tests）：覆盖 `getEffectiveReasoningMode`（3 tests）、`shouldRequireApproval`（7 tests，含修复）、`alwaysAllowedTools`（2 tests）

**修复的测试 bug**：
- `toolExecutor.ts` import 路径错误：`TOOL_DEFINITIONS_MAP` 和 `ALL_TOOL_DEFINITIONS` 从 `../types` 改为 `../toolRegistry`
- 测试用例工具名错误：`read_range` 改为 `range.read`（与实际工具定义一致）

**验证结果**：35 tests 全部通过（2 文件）

---

### #22 OCR Mock 标记

**问题描述**：OCR 识别功能使用 mock 数据，但未标记为 mock 接口，后续接入真实 OCR 时容易遗漏。

**设计方案**：在 mock 代码处添加 `@MOCK_INTERFACE` 标记注释。

**改动详情**：
- `OCRTaskComposerPanel.tsx`：OCR mock 数据处添加 `@MOCK_INTERFACE` 标记
- `electronAPI.ocr` 预留接口处添加 `@MOCK_INTERFACE` 标记

**验证结果**：TypeScript 编译通过

---

### #23 IPC 依赖注入重构

**问题描述**：前端 71 处直接访问 `window.electronAPI`，测试时无法注入 mock。

**设计方案**：新建 `src/services/ipcApi.ts` 抽象层，所有 IPC 调用通过 `ipcApi` 实例。

**改动详情**：
- 新建 `src/services/ipcApi.ts`（~350 行）：
  - `IIpcApi` 接口：与 `electronApi.d.ts` 的 `ElectronAPI` 一致
  - `ipcApi` 运行时实例：所有方法在 `window.electronAPI` 不可用时安全降级
  - `createMockIpcApi(overrides)` 辅助函数：测试时一键注入 mock
- 已迁移的模块：
  - `store/chatStore.ts`：`agent.startTurn` / `continueTurn` / `interrupt`，`tool.confirm` / `cancel`，`onEvent` / `onStreamDelta`
  - `store/threadActions.ts`：`thread.list` / `load` / `resume` / `delete` / `newThread` / `updateMetadata`
  - `store/settingsStore.ts`：`settings.get` / `set` / `getAll`（含 `savePartial` 和 `loadSettings`）
  - `hooks/useComposer.ts`：`dialog.openFile` / `openImage` / `openFolder`，`folder.listFiles`
  - `hooks/useExcelConnection.ts`：`excel.detectStatus` / `connect`
- 迁移后 `window.electronAPI` 直接调用：71 处 → 0 处（核心模块）

**验证结果**：TypeScript 编译通过，35 tests 全部通过

---

## 2026-06-25 — 项目开发规范文档

### 分支
`feature/new-feature`

### 改动概览
基于 4 周 18 项代码审查修复经验，编写 `docs/development-standards.md` 项目开发规范文档，按六大审查方向组织。

### 文档结构

| 章节 | 审查方向 | 核心内容 |
|------|----------|----------|
| 一、大文件模块拆分 | #9–#12, #18 | 400行上限、目录+barrel re-export 模式、拆分原则、4种文件类型拆分方向、当前4个待拆文件 |
| 二、Props 与接口精简 | #19, #20 | Hook 返回值整体传入、ReturnType 派生、Store action 单关注点原则 |
| 三、持久化优化 | #20 | savePartial 增量写入、KEY_MAP 显式映射、复合字段处理、Promise.all 并行 |
| 四、测试基础设施 | #21, #22 | 必须测试的模块类型、测试文件命名、createMockIpcApi、@MOCK_INTERFACE 标记 |
| 五、IPC 依赖注入 | #23, #5, #6 | ipcApi 抽象层、zod schema 校验、结构化日志、迁移进度追踪表 |
| 六、Bug 修复与防御 | #1–#4, #7 | 事件顺序保证、API 消息格式校验、类型安全、N+1 查询防御、11项自查清单 |

### 验证结果
- 文档已创建：`docs/development-standards.md`
- CHANGELOG.md 已同步更新
- dev-log.md 已同步更新

---

## 2026-06-25 — TypeScript 编译错误批量修复

### 分支
`feature/new-feature`

### 改动概览
修复 7 项 TypeScript 编译错误，前端 (`tsc --noEmit`) 和 Electron 端 (`tsc -p tsconfig.electron.json --noEmit`) 均实现零错误编译。

---

### #24 agentLoop.ts — TurnItem 上访问 .message 类型错误

**问题描述**：`agentLoop.ts:399` 对类型为 `TurnItem` 的变量 `errorItem` 直接访问 `.message`，但 `.message` 仅存在于 `ErrorItem` 子类型上，编译报错 `Property 'message' does not exist on type 'TurnItem'`。

**根因分析**：`streamCollector.ts` 中构造的 `errorItem` 类型标注为 `TurnItem`（联合类型），消费方未做类型窄化就直接访问了 `ErrorItem` 独有的 `.message` 属性。

**改动详情**：
- `agentLoop.ts:399`：`errorItem.message` → `errorItem.type === "error" ? errorItem.message : "Unknown error"`
- 添加 `type === "error"` 窄化守卫，TypeScript 正确推导为 `ErrorItem`

**验证结果**：TypeScript 编译通过

---

### #25 compactionManager.ts — TurnItem[] | null 不可赋值给 TurnItem[]

**问题描述**：`compactionManager.ts:53` 函数返回类型声明为 `TurnItem[]`，但 early return 返回了 `compactedHistory`（类型 `TurnItem[] | null`）。

**改动详情**：
- `compactionManager.ts:53`：`return compactedHistory` → `return compactedHistory ?? []`

**验证结果**：TypeScript 编译通过

---

### #26 modelSettingsI18n.ts — .ts 文件包含 JSX 语法

**问题描述**：`modelSettingsI18n.ts` 第 60 行和 113 行的 `deleteMessage` 函数返回 JSX (`<>...</>`)，但文件扩展名为 `.ts`，TypeScript 不解析 JSX 语法。

**改动详情**：
- 文件重命名：`modelSettingsI18n.ts` → `modelSettingsI18n.tsx`
- 所有消费方（`AddProviderDialog.tsx`、`DeleteConfirmDialog.tsx`、`ModelSettings.tsx`、`ProviderCard.tsx`）的 import 路径不含扩展名，无需修改

**验证结果**：TypeScript 编译通过

---

### #27 agentLoop.test.ts — 缺少 beforeEach 导入

**问题描述**：测试文件使用了 `beforeEach` 但仅导入了 `describe, it, expect, vi`，缺少 `beforeEach`。

**改动详情**：
- `agentLoop.test.ts:5`：`import { describe, it, expect, vi } from "vitest"` → `import { describe, it, expect, vi, beforeEach } from "vitest"`

**验证结果**：TypeScript 编译通过

---

### #28 settingsManager.ts — ElectronStore 类型不兼容

**问题描述**：`getSettingsStore()` 返回类型声明为 `Store`（即 `Store<Record<string, unknown>>`），但实际实例类型为 `Store<typeof DEFAULT_SETTINGS>`，具体类型不可赋值给宽泛类型。

**改动详情**：
- `settingsManager.ts:41`：返回类型从 `Store` 改为 `Store<typeof DEFAULT_SETTINGS>`

**验证结果**：TypeScript 编译通过

---

### #29 ipcSchemas.ts — Zod v4 z.record() 参数数量不匹配

**问题描述**：`ipcSchemas.ts:102` 的 `z.record(z.unknown())` 在 Zod v4 中报错，v4 要求 `z.record()` 传入 key schema + value schema 两个参数。

**改动详情**：
- `ipcSchemas.ts:102`：`z.record(z.unknown())` → `z.record(z.string(), z.unknown())`

**验证结果**：TypeScript 编译通过

---

### #30 ipcApi.ts — setAlwaysOnTop 返回类型不匹配

**问题描述**：`ipcApi.ts` 中 `setAlwaysOnTop` 函数签名要求返回 `Promise<boolean>`，但两处实现返回了 `void`/`undefined`：
- 第 167 行：`raw` 为 null 时 `return;` 返回 `undefined`
- 第 393 行：mock 实现返回 `Promise<void>`

**改动详情**：
- `ipcApi.ts:169`：`if (!raw) return;` → `if (!raw) return false;`
- `ipcApi.ts:393`：`setAlwaysOnTop: async () => {}` → `setAlwaysOnTop: async () => false`

**验证结果**：TypeScript 编译通过

---

### 构建验证
- 前端 `tsc --noEmit`: ✓ 0 错误
- Electron `tsc -p tsconfig.electron.json --noEmit`: ✓ 0 错误

---

### #31 AddProviderDialog — 切回"自定义"时残留模板配置

**问题描述**：用户先选择 DeepSeek 等模板供应商，再切回"自定义"选项时，表单中仍残留上一个模板的 name、baseUrl、model、apiFormat、contextWindowSize、reasoningMode、modelConfigs，导致提交时带着错误配置。

**根因分析**：`handleSelectTemplate` 的 `!template` 分支（用户选择空选项即自定义时）只清除了 `selectedTemplateId`，没有重置其他表单字段。而选择模板时会用 `setName(template.name)` 等填充所有字段，但切回自定义时没有对应的清空逻辑。

**改动详情**：
- `AddProviderDialog.tsx:83-85`：`!template` 分支新增 7 个字段重置：
  - `setName("")`、`setApiFormat("openai")`、`setBaseUrl("")`、`setModel("")`
  - `setContextWindowSize(undefined)`、`setReasoningMode("off")`、`setModelConfigs([])`

**验证结果**：TypeScript 编译通过

---

### #32 ProviderCard — 编辑聚合供应商时删除当前模型不同步清空 provider.model

**问题描述**：编辑已有聚合供应商时，删除当前正在使用的模型只会更新 `modelConfigs`，不会清空 `provider.model`。导致当前模型指向一个已不存在的模型名，下拉框中仍显示该幽灵选项。

**根因分析**：`AddProviderDialog` 中删除模型时有 `if (model === modelConfigs[idx].name) setModel("")` 的同步逻辑，但 `ProviderCard` 的删除按钮（第 442 行）只做了 `onUpdate({ modelConfigs: newConfigs })`，缺少对 `provider.model` 的一致性检查。

**改动详情**：
- `ProviderCard.tsx:442-444`：删除模型时检查被删模型是否为当前 `provider.model`，若是则同时将 `model` 字段置空：
  ```ts
  if (removed && provider.model === removed.name) {
    patch.model = "";
  }
  ```

**验证结果**：TypeScript 编译通过

---

## Agent 命令执行沙箱化 + 系统提示词润色 (2026-06-26)

### 背景

参考 codex-reference 中 Codex 源码的沙箱实现思路（`sandboxing/` 跨平台抽象 + `linux-sandbox/` bubblewrap + `windows-sandbox-rs/` 受限令牌 + `execpolicy/` 命令策略 + `process-hardening/` 进程加固），在本 Excel AI 插件中落地一套不依赖原生模块、不破坏 Excel COM 主链路的命令执行沙箱；同时修复系统提示词在通用场景与公式场景的两处问题。

### 沙箱实现（阶段 0+1+4+5）

Codex 的关键约束对应到本项目：Excel 工具靠 `GetObject("Excel.Application")` 连回用户当前 Excel，受限令牌+私有桌面会让 COM 回连失败 → 沙箱只施加在 `shell.execute` 等不依赖 COM 回连的工具。

| 阶段 | 落地内容 | 对应 Codex 模块 |
|------|----------|-----------------|
| 0 接口重构 | `shell.execute` 改走 `-EncodedCommand` 避免明文拼接；命令切 token | `execpolicy` 命令语义分析前置 |
| 1 策略层 | `execpolicy` 前缀规则 `forbidden/prompt/allow` + cwd 白名单重定向 | `execpolicy` + `FileSystemSandboxPolicy` |
| 4 进程加固 | env allowlist 清洗（USERNAME/USERPROFILE/PATH…）| `process-hardening` env 剥离 + `WINDOWS_SANDBOX_WRAPPER_SETUP_ENV_ALLOWLIST` |
| 5 审计 | JSONL 落 `sandbox-logs/YYYY/MM/DD/audit-*.jsonl` | `windows-sandbox-rs/src/logging.rs` |

**新增 `electron/agent/sandbox/` 模块**：

| 模块 | 行数 | 职责 |
|------|------|------|
| `parseCommand.ts` | ~140 | 切子命令+token（管道/换行/引号/反引号转义/注释剥离） |
| `execPolicy.ts` | ~250 | `ExecPolicy` 类按首 token 索引、多规则取最严、`checkWorkdir` 白名单重定向 |
| `defaultRules.ts` | ~80 | 默认规则挡 `rm -rf /` / `Remove-Item -Recurse -Force` / `format` / `Stop-Computer` / `reg delete` / `iex` / `diskpart`；`curl`/`Invoke-WebRequest`/`powershell -c` 走 prompt |
| `audit.ts` | ~70 | 异步 JSONL 追加，失败不阻塞业务 |
| `index.ts` | ~230 | `evaluateCommand` + `runShellSpawn`（env 白名单 + `-EncodedCommand`）+ `killProcessTree`（taskkill /T 强杀） |
| `sandbox.test.ts` | ~115 | 14 个单测覆盖破坏性命令/prompt/cwd 重定向 |

### 接线点

**`executors.ts` `executeShellCommand`** — 重构成接收 `CommandEvaluation` 后再 spawn；forbidden 兜底拦截。
**`toolExecutor.ts` `processToolCalls`** — 在审批前对 `shell.execute` 评估策略：
- forbidden → 拒绝不进 spawn，记审计
- prompt → 覆盖 `permissionMode` 与 `alwaysAllowedTools`，把 `sandboxJustification` 透传审批对话框
- allow → 维持原审批流程

**IPC + 设置**：`sandbox:getConfig/setUserRules/setWritableRoots` + `applySandboxConfig()` 启动期热更到沙箱单例；`electron-store` 持久化 `sandboxUserRules` / `sandboxExtraWritableRoots`

**UI**：SettingsPage 新增「安全策略」tab（`ExecPolicySettings.tsx`）；`ToolConfirmDialog` 渲染 `sandboxJustification`

### 系统提示词润色

**问题 1：通用场景无脑 inspect** — `roleAndWorkflow()` 把 "用 `workbook.inspect` 获取结构 / `selection.get` 获取选区" 设为每轮固定第 1 步

**改法**：行动优先原则加第 6 条「探测要按需」；工作流程拆为 A 纯提问 / B 读/写表格 / C 通用系统操作 / D 附件触发 四类，仅在 B 类才做工作簿检查，且"用户已明确范围时可跳过"

**问题 2：公式场景与助手字段脱钩 + 与通用守则冲突**

通过 Explore 调研确认公式助手前端把 5 字段（任务说明 / 数据源选区 / 答案参考样例 / 写入锚点 / 是否动数组）拼成中文纯文本走 `agent:startTurn`，后端 system prompt 全量下发 `scenarioFormula()` 但对这 5 字段 0 覆盖。

**改法**：`scenarioFormula()` 全重写，逐字段消费：
- 样例分完整/部分两处理（完整要 value 完全一致，部分不纠结"是否完整"）
- #SPILL! 主动清理（探测→清空→告知用户清理了什么）
- 表头按需拼接（检测锚点是否含表头）
- 嵌套约束松绑（取消"3 层上限"死规矩，动态数组优先减少辅助列）
- 动数组"否"用"写入形态是否依赖溢出"作判定，不再按函数名禁用 XLOOKUP

**全局审计**：§可维护性 的"嵌套不超过3层"条款同步松绑；§安全底线 的"永远不覆盖公式单元格"对公式助手场景放宽（指定写入锚点含旧公式视为用户要重写）；神读校验单点真理（场景段不重列数值合理性，引用 §结果验证）

### 验证结果

- `npm run typecheck` ✅（前端 tsc + electron tsc 全过）
- `npx vitest run` ✅（58 个测试全部通过，含新增 14 个沙箱单测）
- Excel COM 主链路无回归：所有走 `excelBridge.executePowerShell`（`-EncodedCommand`）的工具保持原有 spawn 行为与完整 PSModulePath 环境

### 未落地（留待后续）

- 阶段 2：`utilityProcess.fork` 把 `shell.execute` spawn 搬出主进程 + Job Object 进程组强杀
- 阶段 3：复用 Codex `codex-windows-sandbox.exe` 拿 WRITE_RESTRICTED + WFP（仅对 A 类工具）
- execpolicy 升级到完整 Starlark DSL / `host_executable` / `network_rule`

---

## 2026-06-27 — 会话中断竞争修复 + 渲染性能优化 + Office/WPS 多宿主选择

### 分支
`feature/new-feature`

### 改动概览

| 类别 | 数目 | 说明 |
|------|------|------|
| Bug 修复 | 2 | 会话中断竞争条件、消息列表卡顿 |
| 功能新增 | 1 | Office/WPS 多宿主选择弹窗 |

---

### #1 会话中断竞争条件修复

**问题描述**：对话中点击"新建会话"或切换会话后，立即输入新消息会报错"Agent 正在运行中"。这是因为 `interrupt()` 仅调用 `abortController.abort()` 立即返回，不等旧 Turn 的 catch/finally 清理完毕。

**根因分析**：
- `agentLoop.ts` 的 `interrupt()` 是同步的，仅设置 abort，不等待清理
- `resetThread()` / `resumeThread()` 不检查 `isRunning`
- `threadActions.ts` 中 `ipcApi.agent.interrupt()` 未 `await`
- 旧 Turn 的 `finally` 块异步设置 `isRunning = false`，此时新 `runTurn()` 已因 `isRunning` 守卫抛出异常

**改动详情**：
- `agentLoop.ts`：新增 `_turnCompletionPromise` / `_resolveTurnCompletion`
  - `runTurn()` 启动时创建 Promise，`finally` 块中 resolve
  - `interrupt()` 改为 `async`，await `_turnCompletionPromise`
  - `resetThread()` / `resumeThread()` 增加 `isRunning` 守卫，自动 `await this.interrupt()`
- `ipcHandlers.ts`：`agent:interrupt` handler `await agent.interrupt()`
- `threadActions.ts`：`switchThread()` / `createNewThread()` 中改为 `await ipcApi.agent.interrupt()`

**验证结果**：TypeScript 编译通过，14 个 agentLoop 测试全部通过

---

### #2 消息列表渲染性能优化

**问题描述**：单轮对话中随轮次增加（100+ 条消息）桌面端明显卡顿。AI 流式输出时每秒 10-20 次全量重渲染。

**根因分析**：
- `ChatMessageList` 使用无选择器的 `useChatStore()`，任何状态变化（含 stream delta）触发全树重渲染
- 无虚拟化，全部消息渲染为 DOM
- 无 `React.memo`，子组件每次都执行 `cleanReasoningText` 和 `ReactMarkdown` 解析
- `sortItemsByRound` / `getItemDurationSeconds` 每次重渲染都执行
- `scrollIntoView` 在每次 stream delta 触发，迫使浏览器 layout 重算

**改动详情**：
- `ChatMessageList.tsx`：原子化 Zustand 选择器；`scrollIntoView` 仅在 `messages.length` 变化时触发；分组缓存依赖键稳定化
- `MessageBubble.tsx`：包裹 `React.memo`
- `ReasoningBubble.tsx`：包裹 `React.memo`
- `ToolCallBubble.tsx`：包裹 `React.memo`
- `AssistantGroupBlock.tsx`：`React.memo` + 自定义比较器；`sortItemsByRound` / `getItemDurationSeconds` 用 `useMemo` 缓存

**验证结果**：TypeScript 编译通过，91 个测试全部通过

---

### #3 Office / WPS 多宿主选择弹窗

**问题描述**：当 Office Excel 和 WPS 表格同时运行时，`detectExcelProcess` 的 `if/elseif` 始终选择 Excel，且 COM 的 `GetActiveObject('Excel.Application')` 可能因 WPS 覆盖注册而返回错误宿主，导致操作写错目标程序。

**改动详情**：
- `excelComBridge.ts`：`detectExcelProcess()` 同时检测两个进程，返回 `availableHosts[]`；`detectStatus()` / `connect()` 多宿主时返回 `availableHosts` 让前端弹窗；新增 `selectHost()` 保存 `_selectedHost` 并尝试 COM 连接；`getProgId()` 优先使用 `_selectedHost`
- IPC：新增 `excel:selectHost` 通道（preload / ipcApi / electronApi.d.ts）
- `HostSelectionDialog.tsx`：新建 dialog.css 风格弹窗，显示 Excel/WPS 两个选项按钮
- `useExcelConnection.ts`：新增 `pendingHosts` / `handleSelectHost` 状态管理
- `Sidebar.tsx`：重构为 `if/else` 统一管理 `sidebarContent`，底部渲染弹窗
- `i18n.ts` + `dialog.css`：新增中英文文案和宿主选择按钮样式

**验证结果**：TypeScript 编译通过，91 个测试全部通过

### 构建验证
- `npx tsc --noEmit` ✅（前端）
- `npx tsc -p tsconfig.electron.json --noEmit` ✅（Electron）
- `npx vitest run` ✅（91 个测试全部通过）

---

## 2026-06-28 — 功能模块输入链路增强 + 实时宿主环境修正 + 嵌入式 Python 补入

### 分支
`feature/new-feature`

### 改动概览

| 类别 | 数目 | 说明 |
|------|------|------|
| 功能增强 | 4 | 公式助手、代码生成、报告生成、功能模块发送链路 |
| 稳定性修复 | 3 | 会话切换隔离、工作簿宿主元数据、快速选区 |
| 运行时补强 | 1 | 嵌入式 Python 安装脚本与依赖配置 |

---

### #1 功能模块统一为“填入输入框并发送”

**目标**：所有功能模块的提交行为保持一致，先把结构化需求转成用户可见文本，再进入统一的对话发送链路，避免模块面板和普通输入框存在两套行为。

**改动详情**：
- 新增 `taskComposerPayloads.ts` 统一构造公式、代码、报告任务文本
- 公式助手与代码生成支持“部分样例 / 完整样例”
- 公式助手新增当前环境选项（WPS / Microsoft Excel），并自动同步连接状态
- 代码生成根据当前连接宿主自动选择运行环境
- 报告生成新增输出类型：Excel / Word / PPT
- Word / PPT 报告支持选择存储路径，默认桌面；Excel 报告仍在当前表格环境内处理

### #2 快速选区与环境识别优化

**问题描述**：原先功能模块点击“选区”会走完整读取链路，速度慢且在部分场景有延迟；工作簿读取结果中 `name/version` 可能显示固定旧值，容易误导模型判断当前宿主。

**改动详情**：
- 新增 `getSelectionAddress()`，仅读取当前选区地址和工作表名，用于功能模块快速填充选区字段
- 公式助手、代码生成、OCR 识别的选区按钮改为使用快速选区
- `workbook.inspect` 结果增加宿主归一化元数据：
  - `host`: `wps` / `microsoft_excel`
  - `name`: `WPS 表格` / `Microsoft Excel`
  - `version`: 优先使用实时连接版本，无法读取时才回退为 `unknown`

### #3 公式助手提示约束增强

**目标**：公式助手必须按公式场景交付，优先写入 Excel/WPS 函数公式，不应因前文普通对话或其它模块上下文偏移到 VBA、Python、脚本或手工步骤。

**改动详情**：
- 公式模块 payload 明确声明“本轮必须按公式助手模式处理”
- 交付物优先级改为可写入工作表的 Excel/WPS 函数公式
- 增加公式长度提示：未超过单元格公式长度限制时，若返回文本结果，视为公式使用或写入方式错误
- 保留“功能模块：生成公式”标识，隐藏内部系统提示词式描述，减少用户消息污染

### #4 嵌入式 Python 运行时补入

**问题描述**：项目中 Python 能力不可直接依赖用户本机全局 Python，导致 Python 相关链路在不同机器上不可用或不稳定。

**改动详情**：
- `.gitignore` 改为忽略 `desktop/python/*` 运行时产物，仅保留 README 与安装脚本
- `setup-python-embed.ps1` 重写为可重复执行的嵌入式 Python 安装脚本
- 安装脚本下载 Python Embedded Distribution，并启用 `pip` / `site-packages`
- 默认安装并验证 `xlwings`

### #5 会话切换与运行态隔离

**问题描述**：会话还在执行时切换到其它会话，可能因为全局运行态或事件转发混用导致旧会话停止、状态错乱或 UI 卡顿。

**改动详情**：
- Agent 运行态绑定 threadId，事件转发时携带并校验会话归属
- 中断与恢复流程补充等待逻辑，降低跨会话竞争
- 侧边栏运行/完成/失败提示作为“提醒查看”状态，用户查看后即清除，运行中保持显示

### 验证结果
- `npm --prefix desktop test` ✅（16 个测试文件，135 个测试全部通过）
- `npm --prefix desktop run typecheck` ✅
- `npm --prefix desktop run build` ✅（仅 Vite 既有 chunk / CJS API 警告）
- `git diff --check` ✅

---

## 2026-06-30 — Office action 删除 PPT 页、安装包 0.1.2 与文档同步

### 分支
`feature/new-feature`

### 改动概览

| 类别 | 数目 | 说明 |
|------|------|------|
| Office 能力 | 1 | 新增 PPT 删除页专用 operation |
| 稳定性 | 2 | Open XML 优先、COM 倒序删除兜底 |
| 文档 | 8+ | README、CHANGELOG、Agent/Tools/Memory README、开发规范同步 |
| 发布 | 1 | 生成 `Office AI 助手 Setup 0.1.2.exe` |

### #1 PPT 删除页专用能力

**问题描述**：用户要求“删除第 2~6 页，仅保留目录页”时，模型没有可选的专用工具，只能临时拼脚本处理。PowerShell COM 容易超时，Python 文件级处理又受文件占用、依赖和编码影响，失败率高。

**改动详情**：
- `office.action.apply` 新增 `presentation/deleteSlides` operation。
- Open XML 实现同步删除：
  - `ppt/presentation.xml` 中的 `<p:sldId>` 引用。
  - `ppt/_rels/presentation.xml.rels` 中对应 `Relationship`。
  - `ppt/slides/slideN.xml` slide 部件。
  - `[Content_Types].xml` 中对应 slide override。
- COM 兜底实现倒序删除 `$pres.Slides.Item($idx).Delete()`，避免索引变化导致删错页。
- 工具注册表和 Office 工具选择提示词明确要求删除 PPT 页优先使用统一 `office.action.apply`，禁止优先临场拼脚本。

### #2 安装包版本推进

**改动详情**：
- 桌面端版本从 `0.1.1` 升级到 `0.1.2`。
- 重新构建安装包：`desktop/release/Office AI 助手 Setup 0.1.2.exe`。
- 抽取 `app.asar` 验证包内版本为 `0.1.2`，并包含 `deleteSlides`、`删除 PPT 页`、`已删除幻灯片` 关键词。

### #3 项目文档同步

**改动详情**：
- 根 README 改为 Office 三件套总览，更新项目架构树、Office 操作路线、SQLite StateRuntime、测试基线和开发规范。
- CHANGELOG 增加 2026-06-30 变更摘要。
- Agent README 更新记忆层、Open XML/COM 统一 action 路线和 SQLite 主存储说明。
- Tools contracts/executors/registry/implementations README 补充统一 Office action、`deleteSlides` 和工具选择边界。
- Memory README 补充 SQLite 优先、JSONL 审计副本和长期记忆边界。
- Development standards 移除“必须 barrel/index”旧约束，改为按职责直接引用具体模块，避免重导出影子层。

### 验证结果
- `npm run typecheck` ✅
- `npm test` ✅（74 个测试文件，420 个测试全部通过）
- `npm run electron:build` ✅（生成 0.1.2 安装包）
