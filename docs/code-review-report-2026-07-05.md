# 代码审查报告 — Excel AI 插件项目

> **审查日期**：2026-07-05
> **审查依据**：`docs/code-review-standards.md`
> **审查范围**：desktop/ 全部核心源码（electron/ + src/）
> **审查维度**：安全性 · 正确性 · 可维护性 · 性能 · 测试 · 项目特定规范
> **测试基线**：105 文件 / 609 测试全绿 · typecheck 双通过

---

## 一、审查总结

### 整体评价

项目工程质量**中上**。架构分层清晰、沙箱设计专业、流式增量缓冲和增量持久化等核心模式成熟。但存在两类系统性问题：

1. **IPC 安防线有漏洞** — 多个处理外部输入的 IPC 通道缺少 Zod 校验和路径限制，是最高优先级风险
2. **文件行数大规模超标** — 29 个 TS/TSX 文件 + 4 个 CSS 文件超过规范上限，最严重的 `agentLoop.ts` 达 1276 行

### 问题统计

| 级别 | 数量 | 说明 |
|------|------|------|
| 🔴 P0 | **22** | 必须修复后才能合并 |
| 🟡 P1 | **30** | 应该修复 |
| 💭 P2 | **5** | 酌情处理 |
| ✅ 亮点 | **14** | 值得学习的好实践 |

---

## 修复进度（持续更新）

> 记录规则：每完成一个审查项，回写修复范围、验证证据和关联提交，避免只在代码里静默修复。

### 2026-07-05 — P0 安全：IPC 校验与路径授权

**状态**：✅ 已修复

**关联提交**：本节所在提交 `fix: harden ipc validation and file paths`

**覆盖项**：
- S1：为 `file:writeTempFile`、`ocr:recognize`、`sandbox:setUserRules`、`sandbox:setWritableRoots`、`app:openPath`、`app:openExternal`、`excel:selectHost`、`file:readAsBase64`、`agent:interrupt` 增加 Zod schema 校验。
- S2：接入已存在但未使用的 schema，覆盖 `window:setAlwaysOnTop`、`settings:get/set`、`excel:readRange/writeRange`、`thread:load/delete/resume/new`、`file:trashFile/openFile/copyPath/revealInExplorer`、`folder:listFiles`、`tool:confirm/cancel`、`stats:getSummary`。
- S5：新增主进程路径授权器，文件读取/打开/回收站/显示/文件夹枚举/数据目录打开均先校验授权范围。

**业务链路保护**：
- 授权来源包括：当前数据目录、系统临时目录、用户通过文件/图片/文件夹对话框选择的路径、设置中的 `pinnedFolders`、沙箱额外可写根。
- 拖拽/粘贴文件仍走 `webUtils.getPathForFile(file)`，preload 在返回路径前同步登记授权，避免附件预览、OCR、文件上传出现异步竞态。
- OCR 付费 MinerU → 免费 MinerU → 本地 fallback 的降级顺序未改。

**验证证据**：
- `npm exec vitest run electron/main-modules/ipcPathSecurity.test.ts electron/shared/ipcSchemas.test.ts electron/main-modules/ipcHandlers.ocr.test.ts src/hooks/useComposer.test.ts`
- `npm exec vitest run electron/agent/interaction/ipcAgentHandlers.test.ts electron/agent/interaction/eventForwarder.test.ts`
- `npm run typecheck`

### 2026-07-05 — P0 可维护性：`agentLoop.ts` 阶段性拆分（M1 进行中）

**状态**：🚧 进行中（阶段性拆分已提交，M1 尚未关闭）

**关联提交**：本节所在提交 `refactor: split agent loop helper modules`

**覆盖范围**：
- 从 `agentLoop.ts` 抽出上下文历史收集与用量事件：`electron/agent/core/agentLoop/contextUsage.ts`。
- 从 `agentLoop.ts` 抽出流结束后的 `TurnItem` 补发与落库顺序：`electron/agent/core/agentLoop/streamResultItems.ts`。
- 从 `agentLoop.ts` 抽出压缩进度事件、压缩参数 rollout 记录和冷 rollout 归档触发：`electron/agent/core/agentLoop/compactionProgress.ts`。
- 从 `agentLoop.ts` 抽出线程回调绑定、线程快照/运行态持久化、rollout 事件 sink 绑定和长期记忆抽取调度：`electron/agent/core/agentLoop/threadRuntime.ts`。
- 从 `agentLoop.ts` 抽出 pre-turn / mid-turn 压缩执行逻辑：`electron/agent/core/agentLoop/compactionRunner.ts`。
- 清理 `agentLoop.ts` 中已变成“转一手”的流式结果、压缩进度薄委托包装，减少重复跳转。
- 从 `agentLoop.ts` 抽出每轮模型请求参数装配：`electron/agent/core/agentLoop/roundStreamParams.ts`。
- 从 `agentLoop.ts` 抽出线程会话生命周期编排：`electron/agent/core/agentLoop/threadSession.ts`，覆盖重置、新建、恢复和空闲卸载。
- 移除 `AgentLoop` 内只转发 `turnState` 字段的私有 getter/setter，直接访问 `turnState`，减少主文件噪音。
- 从 `agentLoop.ts` 抽出配置热更新逻辑：`electron/agent/core/agentLoop/configUpdates.ts`，覆盖 AI 客户端重建、压缩 provider 重建、线程 metadata 更新和待压缩原因合并。
- 从 `agentLoop.ts` 抽出单轮 Turn 生命周期编排：`electron/agent/core/agentLoop/turnExecution.ts`，覆盖运行开始、活跃线程准备、用户消息落库、成功完成、失败记录和最终运行态收尾。
- 从 `agentLoop.ts` 抽出运行中补充输入/中断/队列续跑逻辑：`electron/agent/core/agentLoop/queuedTurns.ts`。
- 从 `agentLoop.ts` 抽出单轮模型采样、流式错误 item 落库和 usage 合并：`electron/agent/core/agentLoop/streamRound.ts`。
- 从 `agentLoop.ts` 抽出单轮工具调用处理和 mid-turn 压缩触发判断：`electron/agent/core/agentLoop/toolRound.ts`。
- 从 `agentLoop.ts` 抽出压缩摘要生成与 compact 请求重试配置：`electron/agent/core/agentLoop/compactionSummary.ts`。
- 从 `agentLoop.ts` 抽出每轮 Agent 主循环编排：`electron/agent/core/agentLoop/agentLoopRunner.ts`，覆盖模型请求、流式结果补发、工具轮继续和上下文用量事件。
- 从 `agentLoop.ts` 抽出 turn 前压缩计划判断：`electron/agent/core/agentLoop/preTurnCompaction.ts`，保持待处理压缩原因优先于自动阈值判断。
- 从 `agentLoop.ts` 抽出压缩 runner 依赖组装：`electron/agent/core/agentLoop/compactionRunnerDeps.ts`，集中管理进度事件、归档、历史写回和摘要生成回调。
- 从 `agentLoop.ts` 抽出空闲线程卸载 timer 管理：`electron/agent/core/agentLoop/idleThreadUnload.ts`，覆盖延迟计算、运行中跳过和失败重排。
- 从 `agentLoop.ts` 抽出单次 `runTurn` 生命周期编排：`electron/agent/core/agentLoop/turnFlow.ts`，串联线程准备、turn 前压缩、Agent 循环、成功/失败收尾和队列续跑。
- 从 `agentLoop.ts` 抽出 `AgentLoopConfig` 类型定义：`electron/agent/core/agentLoop/agentLoopConfig.ts`，并在原模块继续 re-export 保持外部导入兼容。
- 从 `agentLoop.ts` 抽出每轮 runner 依赖组装：`electron/agent/core/agentLoop/agentLoopRoundDeps.ts`，集中处理模型、工具、权限、上下文和落库回调接线。
- 精简 `agentLoop.ts` 中与 README 重复的长段注释、区块横线和简单 getter 注释，主循环设计说明改由 `electron/agent/core/agentLoop/README.md` 统一维护。
- 将输入队列 drain 调度和重排逻辑继续下沉到 `electron/agent/core/agentLoop/queuedTurns.ts`，`agentLoop.ts` 只保留状态接线。
- 从 `agentLoop.ts` 抽出基础状态与公共 API 基类：`electron/agent/core/agentLoop/agentLoopBase.ts`，覆盖构造、存储迁移、线程状态观察、队列入队、AI/权限热更新和 pending 压缩原因管理。
- `agentLoop.ts` 已从 1276 行降至 378 行，低于 400 行规范上限；M1 仍未关闭，后续继续处理其它超标文件。
- 从 `ipcHandlers.ts` 抽出 OCR IPC 注册、MinerU/本地 fallback、发票字段抽取和 OCR 结果归一化：`electron/main-modules/ipcOcrHandlers.ts`。`ipcHandlers.ts` 从 1115 行降至 723 行，OCR 新模块 397 行。
- 从 `ipcHandlers.ts` 抽出 AI 模型列表和连接测试 IPC：`electron/main-modules/ipcAiHandlers.ts`。`ipcHandlers.ts` 进一步降至 622 行，AI 新模块 110 行。
- 从 `ipcHandlers.ts` 抽出沙箱配置 IPC 和运行时规则刷新：`electron/main-modules/ipcSandboxHandlers.ts`。`ipcHandlers.ts` 进一步降至 520 行，沙箱新模块 103 行。
- 从 `ipcHandlers.ts` 抽出文件对话框、文件夹枚举、Base64 读取、临时文件写入和文件操作 IPC：`electron/main-modules/ipcFileHandlers.ts`。`ipcHandlers.ts` 进一步降至 347 行，低于 400 行规范上限。
- 将 `Sidebar.tsx` 的折叠态、展开态主体和底部连接/设置区按 UI 边界抽出到 `components/sidebar/SidebarCollapsed.tsx`、`SidebarExpanded.tsx`、`SidebarFooter.tsx`。父组件保留 store 状态、业务回调和数据编排，避免为了行数继续拆碎。
- 将 `settingsStore.ts` 的 Provider 模板、API 格式和推理选项静态配置抽出到 `store/settingsProviderTemplates.ts`。`settingsStore.ts` 从 749 行降至 479 行，仍由原入口 re-export `PROVIDER_TEMPLATES`、`API_FORMATS` 和相关类型，保持调用方兼容。
- 将 `services/ipcApi.ts` 的 `IIpcApi` 类型定义抽出到 `services/ipcApiTypes.ts`，测试 mock 工厂抽出到 `services/ipcApiMock.ts`。`ipcApi.ts` 从 745 行降至 430 行，并继续 re-export `IIpcApi` 与 `createMockIpcApi` 保持测试和调用方兼容。
- 将 `OCRTaskComposerPanel.tsx` 的 OCR 字段提取、预览行和写入矩阵构建纯函数抽出到 `components/task/ocrTaskResultHelpers.ts`。面板继续 re-export 原 helper 和类型，测试导入路径不变。
- 将 `sessionStore.ts` 的 rollout 使用统计解析抽出到 `electron/agent/memory/sessionUsageStats.ts`。`sessionStore.ts` 从 674 行降至 567 行，保留会话创建、恢复、删除、rollout 写入和路径缓存主链路在原类中。
- 将 `advancedExcel.ts` 的 OpenXML 公式单元格生成、动态数组函数识别和 `_xlfn` 前缀规范化抽出到 `officeOpenXml/excelFormulaXml.ts`。`advancedExcel.ts` 从 662 行降至 500 行，保留工作簿创建、写入范围、数据验证、条件格式和表格样式编排。
- 将 `toolExecutor.ts` 的工具执行日志记录类型、参数解析、摘要截断和安全写日志 helper 抽出到 `agentLoop/toolExecutionLog.ts`。`toolExecutor.ts` 从 566 行降至 513 行，并继续从原入口 re-export `ToolExecutionLogRecord`。
- 将 `webSearchExecutors.ts` 的 Bing/Baidu/360/Sogou/DuckDuckGo HTML 结果解析、URL 归一化和去重 helper 抽出到 `tools/executors/webSearchHtmlParsers.ts`。`webSearchExecutors.ts` 从 554 行降至 378 行，低于 400 行规范上限。
- 将 `documentParser.ts` 的 JSON 路径展开 helper 抽出到 `knowledge/jsonFlatten.ts`。`documentParser.ts` 从 544 行降至 510 行，保留 CSV/Excel/DOCX/PPTX/Markdown/Text 解析入口和分块 metadata 行为。
- 将 `chatStore.ts` 的 stream delta 合并、50ms 缓冲和 IPC listener 安装抽出到 `store/chatStreamBuffer.ts`。`chatStore.ts` 从 549 行降至 490 行，保留 Zustand 状态与用户动作主链路在原 store 中，避免为压线继续拆碎 action。
- 将 `sqliteStore.ts` 的 KnowledgeEntry/KnowledgeSource 行模型转换和余弦相似度计算抽出到 `knowledge/sqliteStoreRows.ts`。`sqliteStore.ts` 从 459 行降至 393 行，低于 400 行规范上限。
- 将 `documentParser.ts` 的 Excel OpenXML workbook/sheet/sharedStrings/worksheet 行解析抽出到 `knowledge/excelWorkbookParser.ts`。`documentParser.ts` 进一步降至 281 行，低于 400 行规范上限。
- 将 `advancedExcel.ts` 的工作簿基础部件、worksheet XML、sheetData 合并和单元格 XML 生成抽出到 `officeOpenXml/excelSheetXml.ts`。`advancedExcel.ts` 进一步降至 303 行，低于 400 行规范上限。
- 将 `settingsStore.ts` 的增量持久化 key 映射、compactionConfig 组合写入和 `savePartial` 抽出到 `store/settingsPersistence.ts`。`settingsStore.ts` 从 418 行降至 380 行，低于 400 行规范上限。
- 将 `ipcApi.ts` 的知识库 IPC wrapper 抽出到 `services/ipcKnowledgeApi.ts`。`ipcApi.ts` 从 408 行降至 378 行，低于 400 行规范上限。
- 将 `sessionStore.ts` 的 JSONL rollout → Thread 投影解析抽出到 `memory/sessionRolloutParser.ts`，原私有 `parseRolloutContent` 保留为委托入口兼容测试和调用路径。`sessionStore.ts` 从 469 行降至 353 行，低于 400 行规范上限。
- 新增对应单元测试，保护上下文顺序、流式结果事件顺序、压缩成功/失败事件和归档阈值行为。
- 同步更新 `electron/agent/core/agentLoop/README.md` 与 `electron/main-modules/README.md`，记录拆分后的模块职责。

**业务链路保护**：
- `AgentLoop` 对外 API 和主流程签名不变，保留原私有方法作为委托入口，降低调用链变更风险。
- reasoning、assistant_message、tool_call 的补发顺序保持原逻辑；工具执行、压缩触发、长期记忆写入和线程运行态写入链路未改。
- OCR 付费 MinerU → 免费 MinerU → 本地解析 fallback 顺序未改；`ocr:recognize` 仍先做 schema 校验和路径授权，再进入识别链路。
- `ai:listModels` 与 `ai:testConnection` 仅移动注册位置，保留 Anthropic/Responses/Chat Completions 的 endpoint、请求体、超时和错误返回行为。
- `sandbox:getConfig`、`sandbox:setUserRules`、`sandbox:setWritableRoots` 仅移动注册位置，保留 Zod 校验、规则规范化、electron-store 写入和 `applySandboxConfig()` 刷新行为。
- 文件/对话框 IPC 继续共用同一个 `pathAuthorizer` 实例，保留选择后授权、读取前授权、临时文件写入后授权和资源管理器/回收站路径校验。
- Sidebar 拆分只移动渲染层 JSX，`useChatStore`、`useSettingsStore`、Office 连接检测、文件夹加载、会话切换和文件右键操作仍由父组件统一编排。
- settingsStore 拆分只移动静态模板配置，`loadSettings` 迁移逻辑、增量保存、provider normalize、窗口透明度和知识库开关持久化链路未改。
- ipcApi 拆分只移动类型定义和测试 mock，运行时 `ipcApi` wrapper 的 fallback 行为、`readRange` 第三参 `expand` 透传、thread/runtimeStatus 与 threadGraph wrapper 保持不变。
- OCR 面板拆分只移动纯数据转换函数，拖拽/粘贴上传、静默识别、字段选择、选区获取和写入 Excel/WPS 链路未改。
- sessionStore 拆分只移动 `getUsageSummary` 的 JSONL 统计解析，`appendRolloutItems`、`loadThread`、`parseRolloutContent`、压缩归档搜索和数据库投影写入链路未改。
- advancedExcel 拆分只移动公式 XML 生成 helper，动态数组 `<f t="array" ref="...">` 输出、`_xlfn._xlws.FILTER` 前缀和 spill 占位清理行为由现有测试继续覆盖。
- toolExecutor 拆分只移动执行日志 helper，审批判断、沙箱 forbidden/prompt 覆盖、工具执行顺序和 `TurnItem` 事件补发顺序未改。
- webSearch 拆分只移动 HTML 解析和结果去重 helper，付费 Tavily/Bing/SerpAPI 优先、免费 HTML fallback 顺序、超时和错误汇总逻辑未改。
- documentParser 拆分只移动 JSON flatten 纯函数，知识库文件类型支持、OpenXML 文本提取、Excel 行数截断和 metadata 生成链路未改。
- chatStore 拆分只移动 stream delta 缓冲基础设施，`stream_delta` 合并规则、非流式事件前强制 flush、`onEvent`/`onStreamDelta` 双入口和原测试导入路径均保持不变。
- sqliteStore 拆分只移动纯行转换和余弦相似度 helper，建表/迁移、provider/model/dim 过滤、`a.length !== b.length` 返回 0 的维度不匹配语义、事务写入和来源汇总补全逻辑未改。
- Excel 文档解析拆分只移动 OpenXML workbook/sheet/sharedStrings 读取和 worksheet 行解析，`parseExcel` 的 25MB 限制、500 行截断、表头/行列数/tableRange metadata 和 RawChunk 结构未改。
- Excel OpenXML 写入拆分只移动 workbook/sheet XML 生成和 sheetData 合并 helper，`createWorkbook`/`writeRange` 输入归一化、动态数组判定、目标 sheet 解析、数据验证、条件格式、表格样式和 `excelDone` 返回结构未改。
- settingsStore 拆分只移动增量持久化适配层，`loadSettings` 迁移、`saveSettings` 全量保存、provider normalize、透明度、动态数组、知识库开关和 pinned folders action 均未改。
- ipcApi 拆分只移动知识库域 wrapper，`listSources/search/indexFile/indexFolder/deleteFile/reindexAll` 的无 IPC fallback 返回值和运行时调用路径未改。
- sessionStore 拆分只移动 JSONL rollout 解析纯逻辑，`appendRolloutItems`、`loadThread`、`loadThreadByPath`、压缩归档搜索、删除、metadata 追加和 usage 统计链路未改。
- 此阶段已关闭 `agentLoop.ts` 与 `ipcHandlers.ts` 单文件超标；Sidebar/settingsStore/ipcApi 已按可维护边界阶段性拆分，后续不再为追求行数继续拆碎。M1 仍未关闭，后续转向其它明确职责边界的超标文件。

**验证证据**：
- `npm exec vitest run electron/agent/core/agentLoop/streamResultItems.test.ts electron/agent/core/agentLoop/compactionProgress.test.ts electron/agent/core/agentLoop/contextUsage.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/threadRuntime.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/compactionRunner.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/roundStreamParams.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/threadSession.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/configUpdates.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/turnExecution.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/queuedTurns.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/streamRound.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/toolRound.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/compactionSummary.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/agentLoopRunner.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/preTurnCompaction.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/compactionRunner.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/idleThreadUnload.test.ts electron/agent/core/agentLoop/threadSession.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/turnExecution.test.ts electron/agent/core/agentLoop/preTurnCompaction.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/queuedTurns.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/agentLoop.test.ts electron/agent/core/agentLoop/configUpdates.test.ts electron/agent/core/agentLoop/threadSession.test.ts electron/agent/core/agentLoop/queuedTurns.test.ts`
- `npm exec vitest run electron/main-modules/ipcHandlers.ocr.test.ts electron/main-modules/mineruOcr.test.ts electron/main-modules/invoiceFieldExtraction.test.ts`
- `npm exec vitest run electron/shared/ipcSchemas.test.ts electron/main-modules/ipcPathSecurity.test.ts`
- `npm exec vitest run electron/shared/ipcSchemas.test.ts electron/main-modules/ipcPathSecurity.test.ts src/hooks/useComposer.test.ts`
- `npm exec vitest run src/store/settingsStore.test.ts`
- `npm exec vitest run src/services/ipcApi.test.ts`
- `npm exec vitest run src/components/task/OCRTaskComposerPanel.test.ts`
- `npm exec vitest run electron/agent/memory/sessionStore.test.ts`
- `npm exec vitest run electron/agent/tools/implementations/officeOpenXml/advancedExcel.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/toolExecutor.test.ts`
- `npm exec vitest run electron/agent/tools/executors/webSearchExecutors.test.ts`
- `npm exec vitest run electron/agent/knowledge/rag.test.ts`
- `npm exec vitest run src/store/chatStore.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P1 性能：`ChatPage` 文件夹加载取消保护

**状态**：✅ 已修复

**关联提交**：本节所在提交 `fix: guard chat folder file loading`

**覆盖范围**：
- 为 `src/components/ChatPage.tsx` 中按 `currentFolderId` 加载文件夹文件的 `useEffect` 增加 `cancelled` 标记。
- 当组件卸载或 `currentFolderId` 快速切换时，旧的 `ipcApi.folder.listFiles()` Promise 返回后不再调用 `setCurrentFolderFiles()`。

**业务链路保护**：
- 文件夹为空时仍立即清空 `currentFolderFiles`。
- 成功和失败 fallback 结果保持不变，仅阻止过期异步结果回写当前 UI。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P1 性能：`useExcelConnection` timeout 清理

**状态**：✅ 已修复

**关联提交**：本节所在提交 `fix: clean up excel connection timers`

**覆盖范围**：
- 为 `src/hooks/useExcelConnection.ts` 增加 timeout id registry。
- `pulseDot` 和 `connectFailed` 的动画复位统一通过 `scheduleTimeout()` 注册。
- hook 卸载时通过 `clearScheduledTimeouts()` 清理所有未触发的 timeout，避免卸载后继续 setState。

**业务链路保护**：
- 30 秒轮询 interval 的原 cleanup 保持不变。
- 手动连接、多宿主选择、连接失败动画和成功脉冲的延迟时长保持原值：`1500ms` / `600ms`。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P1 性能：`Sidebar` 排序派生数据 memo

**状态**：✅ 已修复

**关联提交**：本节所在提交 `perf: memoize sidebar sorted groups`

**覆盖范围**：
- 为 `src/components/Sidebar.tsx` 中的未分组会话排序结果增加 `useMemo`。
- 为项目文件夹分组、文件夹内会话排序和项目分组排序结果增加 `useMemo`。

**业务链路保护**：
- `sortSidebarItems()`、项目排序规则、会话排序规则、搜索过滤条件和传给 `SidebarExpanded` 的 props 结构均未改。
- memo 依赖显式覆盖 `threads`、`pinnedFolders`、`folderFiles`、排序模式、语言和搜索状态。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — PR4：IPC 类型声明同步

**状态**：✅ 已修复

**关联提交**：本节所在提交 `fix: sync preload ipc types`

**覆盖范围**：
- 补全 `desktop/src/electronApi.d.ts` 中 `excel.detectStatus()` 返回类型，与 `ipcApiTypes.ts` 保持一致，包含 `version`、`workbookName` 和 `availableHosts`。
- 补全 `desktop/electron/preload.ts` 中 `agent.onStreamDelta()` 回调 data 类型的 `clientId` 字段，与渲染层全局声明和 wrapper 类型保持一致。

**业务链路保护**：
- 仅同步 TypeScript 类型声明，没有改动 IPC channel、参数透传或运行时 fallback。
- `thread.runtimeStatus`、`threadGraph`、`excel.readRange(expand)` 等前一轮已补齐的 wrapper 能力保持不变。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P1 性能：`Sidebar` 搜索文件夹批量加载

**状态**：✅ 已修复

**关联提交**：本节所在提交 `perf: batch sidebar folder file listing`

**覆盖范围**：
- 新增 `folder:listFilesBatch` IPC，批量读取多个已授权文件夹内的 Office 文件。
- 将 `folder:listFiles` 的授权、扩展名过滤、stat 补全、排序和文件路径授权逻辑抽为 `listAuthorizedOfficeFiles()`，单条和批量接口共用，避免重复实现。
- `Sidebar` 搜索面板打开时只对未缓存的 pinned folder 发起一次批量请求，不再每个文件夹一次 IPC。
- `ipcApi`、`preload`、`electronApi.d.ts`、`ipcApiTypes.ts` 和 mock 工厂同步新增 `listFilesBatch()`。

**业务链路保护**：
- 添加文件夹、展开单个文件夹、删除/固定文件后的单文件夹刷新仍继续调用 `folder:listFiles`。
- wrapper 在旧 preload 未暴露 `listFilesBatch` 时回退为逐个 `listFiles`，避免运行环境未更新时直接失效。
- 批量接口对单个文件夹读取失败返回该文件夹空数组，不影响其它文件夹结果。

**验证证据**：
- `npm exec vitest run electron/main-modules/ipcFileHandlers.test.ts src/services/ipcApi.test.ts electron/shared/ipcSchemas.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — PR5：CHANGELOG 当前基线同步

**状态**：✅ 已修复

**关联提交**：本节所在提交 `docs: update changelog baseline`

**覆盖范围**：
- 在 `CHANGELOG.md` 的 Unreleased 顶部新增 2026-07-06 当前基线段，记录 `0.1.61` 版本线。
- 更新测试源统计为 124 个测试文件、654 个 `it/test` 用例，避免继续展示旧的 74/420 基线。
- 将 2026-06-30 的 `0.1.2` 描述改为历史阶段说明，并指向顶部当前版本基线。
- 汇总近期 IPC 安全、模块拆分、知识库/RAG、流式推理、Office/WPS、OCR 附件和侧边栏性能优化。

**业务链路保护**：
- 仅更新文档，不改动运行时代码、配置或打包脚本。

**验证证据**：
- `git diff --check`
- `npm run typecheck`
- `npm run build`

### 2026-07-06 — S7：依赖安全审计接入 CI

**状态**：✅ 已修复

**关联提交**：本节所在提交 `ci: audit npm dependencies`

**覆盖范围**：
- 在 `.github/workflows/ci.yml` 的 `npm ci` 之后新增 `npm audit --audit-level=high`。
- CI 会在发现 high/critical 级别漏洞时失败，避免新增依赖绕过安全检查。

**业务链路保护**：
- 不改动生产代码、打包脚本或依赖版本。
- 本地先执行 `npm audit --audit-level=high`，确认当前依赖树为 0 vulnerabilities 后再接入阻断式 CI。

**验证证据**：
- `npm audit --audit-level=high`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M3：`FolderSection` props 语义分组

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: group folder section props`

**覆盖范围**：
- 将 `FolderSection` 的文件夹操作合并为 `folderActions`。
- 将文件夹内线程切换和右键菜单合并为 `threadActions`。
- 将文件添加、文件右键菜单状态、关闭、打开、复制路径、资源管理器定位、删除和置顶合并为 `fileMenuApi`。
- `SidebarExpanded` 同步改为透传上述三个语义对象，`Sidebar.tsx` 使用 `useMemo` 组装，减少中间层 prop drilling。

**业务链路保护**：
- 仅调整 React props 形状，不改动菜单行为、文件操作 IPC、线程切换、文件夹展开和置顶逻辑。
- 未拆分组件或改变 DOM/CSS 结构，避免引出布局回归。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P2 性能：`Sidebar` 拖拽 resize rAF 节流

**状态**：✅ 已修复

**关联提交**：本节所在提交 `perf: throttle sidebar resize updates`

**覆盖范围**：
- 将 `Sidebar.tsx` 的 sidebar resize `mousemove` 处理改为 `requestAnimationFrame` 节流。
- 拖拽中只记录最新目标宽度，每帧最多 `setSidebarWidth()` 一次。
- `mouseup` 时取消未执行的 rAF 并提交最后一次宽度，避免松手时丢失最终尺寸。

**业务链路保护**：
- 保留原宽度边界 `180-400px`、`isResizing` 状态和 document 级 mousemove/mouseup 监听清理。
- 不改 DOM/CSS 和侧边栏折叠/展开行为。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P2 可维护性：shell/sandbox 执行限制命名化

**状态**：✅ 阶段性已修复

**关联提交**：本节所在提交 `refactor: name shell execution limits`

**覆盖范围**：
- 将 shell 执行默认超时、watchdog 宽限时间和毫秒换算提取到 `shellExecutionLimits.ts`，由执行器和工具 schema 共同引用。
- 将 sandbox 中 stdout/stderr 输出截断上限 `50000` / `10000` 提取为 `SHELL_STDOUT_MAX_CHARS` / `SHELL_STDERR_MAX_CHARS`。
- 补充测试，约束 `shell.execute` 注册描述中的默认超时与执行器默认值保持一致。

**业务链路保护**：
- 不改变 shell.execute 的审批、沙箱评估、cwd 校验、超时强杀、输出截断和错误返回行为。
- 仅治理审查报告点名的 shell/sandbox 执行边界数字；UI 尺寸、测试样例和业务阈值不做机械抽取，避免为减行数拆散上下文。

**验证证据**：
- `npm exec vitest run electron/agent/tools/executors/shellExecutor.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M8：electron 警告/错误日志统一 logger

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: route electron warnings through logger`

**覆盖范围**：
- 将 `desktop/electron` 下剩余 `console.warn` / `console.error` 替换为 `createLogger()` 结构化日志。
- 覆盖主进程启动失败、AgentLoop 队列失败、压缩归档失败、线程运行态持久化失败、长期记忆写入失败、Office bridge 清理失败、知识库初始化失败、沙箱审计写入失败和 settings 数据路径迁移/回退警告。
- 同步更新 sandbox audit 注释，不再描述为 `console.warn`。

**业务链路保护**：
- 仅替换日志出口，不改变错误吞吐、fallback、启动失败窗口展示、审计写入失败不阻塞业务等原有控制流。
- logger 仍保留控制台输出，并额外写入文件日志，便于用户环境排查。

**验证证据**：
- `Get-ChildItem -Path 'desktop/electron' -Recurse -Filter '*.ts' | Select-String -Pattern 'console\\.(warn|error)'`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — T4：`shellExecutor` 测试覆盖补强

**状态**：✅ 已修复

**关联提交**：本节所在提交 `test: cover shell executor failure paths`

**覆盖范围**：
- 将 `shellExecutor.test.ts` 从 1 个用例扩展到 5 个用例。
- 覆盖 sandbox evaluation 复用、forbidden 决策不 spawn、prompt 决策使用 effective workdir、缺失工作目录返回错误、watchdog 超时调用 `killProcessTree()`。
- 为 sandbox mock 增加 `@MOCK_INTERFACE` 注释，减少测试 mock 和真实接口漂移风险。

**业务链路保护**：
- 仅补测试，不修改 `shellExecutor.ts` 运行逻辑。
- 仍通过 mock 的 sandbox process primitives 避免测试启动真实 shell 进程。

**验证证据**：
- `npm exec vitest run electron/agent/tools/executors/shellExecutor.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — T3：`src/utils` 纯函数测试补强

**状态**：✅ 阶段性已修复

**关联提交**：本节所在提交 `test: cover reasoning utility contracts`

**覆盖范围**：
- 新增 `src/utils/reasoningSupport.test.ts`，覆盖 OpenAI/Responses、聚合平台推理模型、模板 fallback、默认值归一化和中英文标签格式化。
- 新增 `src/utils/textCleaner.test.ts`，覆盖 CJK token 空格清理、英文单词空格保留、Markdown 表格结构保留和连续空行归一化。
- 修复 `normalizeProviderReasoningConfig()` 中无效 per-model `reasoningMode` 未被剔除的问题，避免旧配置在模型切换后继续残留。

**业务链路保护**：
- 只补报告点名的 `src/utils/reasoningSupport.ts` / `src/utils/textCleaner.ts` 纯函数测试；electron/agent 的高风险测试缺口仍按后续项继续处理。
- 推理配置修复仅清理不在当前模型支持选项内的旧 `reasoningMode`，合法值保留，`reasoningOptions` 旧字段剔除行为不变。

**验证证据**：
- `npm exec vitest run src/utils/reasoningSupport.test.ts src/utils/textCleaner.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — T3：sandbox 策略安全边界测试补强

**状态**：✅ 阶段性已修复

**关联提交**：本节所在提交 `test: cover sandbox safety boundaries`

**覆盖范围**：
- 将 `electron/agent/security/sandbox/sandbox.test.ts` 从 15 个用例扩展到 19 个用例。
- 补充未闭合引号解析失败、解析失败命令进入 prompt、管道/分号后的危险子命令仍 forbidden、Windows 语义下默认规则大小写不敏感等回归测试。

**业务链路保护**：
- 仅补高风险沙箱模块测试，不修改 `parseCommand`、`ExecPolicy` 或默认规则运行逻辑。
- 覆盖的是 shell.execute 前置安全评估边界，确保危险命令不会因为组合命令、解析失败或大小写差异绕过策略。

**验证证据**：
- `npm exec vitest run electron/agent/security/sandbox/sandbox.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — T5/PR2：关键 mock 接口标记补齐

**状态**：✅ 阶段性已修复

**关联提交**：本节所在提交 `test: mark critical mock interfaces`

**覆盖范围**：
- 为 `src/store/chatStore.test.ts` 的 `ipcApi.agent/thread` mock 增加 `@MOCK_INTERFACE` 标记。
- 为 `src/services/ipcApi.test.ts` 的 `window.electronAPI` preload 形状 mock 增加 `@MOCK_INTERFACE` 标记。
- 为 `electron/main-modules/mineruOcr.test.ts` 的 MinerU 标准/Agent HTTP mock 增加 `@MOCK_INTERFACE` 标记。
- 为 `electron/agent/interaction/eventForwarder.test.ts` 的 Electron BrowserWindow/ipcMain mock 增加 `@MOCK_INTERFACE` 标记。
- 为 `electron/agent/providers/openaiResponsesClient.test.ts` 的 OpenAI Responses SSE fetch mock 增加 `@MOCK_INTERFACE` 标记；`shellExecutor.test.ts` 已有 sandbox process primitives 标记。

**业务链路保护**：
- 仅补测试注释，不修改 mock 行为、断言或生产代码。
- 标记范围收在报告点名的关键 mock，避免全仓库机械刷注释造成噪音。

**验证证据**：
- `npm exec vitest run src/store/chatStore.test.ts src/services/ipcApi.test.ts electron/main-modules/mineruOcr.test.ts electron/agent/interaction/eventForwarder.test.ts electron/agent/providers/openaiResponsesClient.test.ts electron/agent/tools/executors/shellExecutor.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P5 性能：`Sidebar` JSX 回调稳定化

**状态**：✅ 已修复

**关联提交**：本节所在提交 `perf: stabilize sidebar callbacks`

**覆盖范围**：
- 将 `Sidebar.tsx` 中传给 `SidebarCollapsed`、`SidebarExpanded`、`SidebarSearchPalette`、`HostSelectionDialog` 的搜索、展开、关闭、设置菜单、上下文菜单状态更新等内联箭头函数提取为 `useCallback`。
- `folderFileMenuApi.close` 改为复用稳定的 `closeFileContextMenu`，避免每次 render 生成新闭包。

**业务链路保护**：
- 仅稳定回调引用，不改动状态更新逻辑、菜单行为、搜索开关、设置入口或 Host 选择弹窗。
- `onPinThread` / `onRenameThread` 仍维持原先“关闭菜单”的占位行为。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

---

### 2026-07-06 — T3：StateRuntimeStore 跨库事务一致性补强

**状态**：✅ 已修复
**关联提交**：本节所在提交 `fix: rollback runtime transactions across databases`

**覆盖范围**：
- 为 `StateRuntimeStore.transaction()` 补充跨 `state/logs/goals/memories` 四个运行时库的回滚测试，覆盖 thread snapshot、rollout events、tool logs、goals、短期记忆、长期记忆和 pipeline cursor。
- 将外层事务从仅包裹 `state.db` 调整为同时 `BEGIN/COMMIT/ROLLBACK` 四个运行时库，避免日志、目标或记忆写入在事务失败后残留。
- `appendRolloutItems()` 与 backfill 继续保留 logs 局部事务；当已经处于外层运行时事务时，改为复用外层事务，避免 SQLite 嵌套 `BEGIN`。

**业务链路保护**：
- 非事务路径下 logs 写入仍由 `runSqliteTransaction()` 单库保护，rollout FTS 写入和 backfill 行为不变。
- 嵌套调用 `transaction()` 时复用外层事务语义，不再重复开启 SQLite 事务。
- 没有拆分 `stateRuntimeStore.ts` 文件结构，本轮只修正真实一致性边界，避免为了行数制造额外模块。

**验证证据**：
- `npm exec vitest run electron/agent/memory/stateRuntimeStore.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — T3：工具执行结果假值保真

**状态**：✅ 已修复
**关联提交**：本节所在提交 `fix: preserve falsy tool results`

**覆盖范围**：
- 为 `processToolCalls()` 增加成功工具返回 `false` 的回归测试，锁定 `tool_result.result` 不应被 `||` 误判为空。
- 将工具结果组装从 `result.data || result.error` 改为按 `result.success` 分支取值，保留 `false`、`0`、空字符串等合法结果。
- 修复 `summarizeForLog()` 对 `undefined` 的摘要处理，避免 `JSON.stringify(undefined)` 触发后续 `.length` 崩溃。

**业务链路保护**：
- 不改变工具审批、sandbox 判定、事件顺序和日志字段结构。
- 失败工具仍写入 `result.error`；成功工具按原始 `data` 写入前端事件与执行日志摘要。

**验证证据**：
- `npm exec vitest run electron/agent/core/agentLoop/toolExecutor.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P2：消息与工具详情样式令牌化

**状态**：✅ 已修复
**关联提交**：本节所在提交 `style: use theme tokens for tool details`

**覆盖范围**：
- 将 `message-bubble.css` 中工作详情摘要的硬编码灰色替换为 `var(--text-faint)` / `var(--text-muted)`。
- 将 `tool-call.css` 中工具调用头、状态图标、展开按钮、命令图标和结果头的硬编码灰色替换为 `var(--text-faint)`。
- 将 `settings.css` 暗色主题下 select option 与设置操作按钮背景从固定 `#111827` 改为 `var(--bg-primary)`。

**业务链路保护**：
- 不改 DOM、布局、间距、动画和交互状态，只把已有颜色接回主题变量系统。
- `tokens.css` 中保留颜色字面量作为设计令牌定义来源；普通样式文件不再直接使用本次清理的 `#9ca3af`、`#6b7280`、`#111827`。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — T3：OpenAI Responses SSE 末尾事件 flush

**状态**：✅ 已修复
**关联提交**：本节所在提交 `fix: flush trailing responses sse event`

**覆盖范围**：
- 为 `OpenAIResponsesClient` 增加无尾随空行的最后一个 SSE 事件回归测试，覆盖 `response.completed` 中的正文、usage 和 done 事件。
- 将 Responses SSE 单块解析抽出为 `processResponsesSSEChunk()`，正常 chunk 与结束时剩余 buffer 共用同一路径。
- 在 reader 完成后处理 `buffer.trim()` 仍有内容的最后事件，避免供应商流缺少结尾 `\n\n` 时丢正文或 done。

**业务链路保护**：
- 不改变现有 `response.output_text.delta`、`output_text.done`、`content_part.done`、工具调用和 usage 事件顺序。
- 保留畸形 chunk 忽略策略；仅补齐原本未处理的完整尾部 chunk。

**验证证据**：
- `npm exec vitest run electron/agent/providers/openaiResponsesClient.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`GeneralSettings` 表单字段复用

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: reuse general settings fields`

**覆盖范围**：
- 新增 `src/components/settings/SettingsFields.tsx`，集中 `SettingsSwitchField` 与 `SettingsSliderField` 两类通用设置字段。
- `GeneralSettings.tsx` 复用上述字段渲染关闭行为、Office 自动避让、窗口透明度、动态数组开关、自动压缩开关和压缩阈值滑块。
- `GeneralSettings.tsx` 从 441 行降至 398 行，低于 TSX 文件 400 行上限；抽出的字段组件为 89 行。

**业务链路保护**：
- 不移动文案、store action、IPC 数据路径迁移、MinerU token 保存、透明度/动态数组/压缩设置写入逻辑。
- 不改设置项顺序、CSS class、滑块填充变量、禁用态和百分比展示。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：OpenAI Responses 解析 helper 拆分

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: split responses parsing helpers`

**覆盖范围**：
- 新增 `electron/agent/providers/openaiResponsesParsing.ts`，集中 Responses content part 转换、工具调用状态、文本补发、usage 归一化和 reasoning effort 映射。
- `openaiResponsesClient.ts` 保留请求体组装、HTTP 请求、SSE 读取和事件分发主链路。
- `openaiResponsesClient.ts` 从 428 行降至 270 行，低于 400 行上限；新解析 helper 模块为 193 行。

**业务链路保护**：
- 不改变 Responses API 请求 endpoint、tools 名称清洗、reasoning 参数、流式正文补发、工具调用事件、usage 或 done 判定。
- 继续使用既有 `openaiResponsesClient.test.ts` 覆盖图片/PDF 输入、reasoning、正文 fallback、尾包 flush 和工具调用解析。

**验证证据**：
- `npm exec vitest run electron/agent/providers/openaiResponsesClient.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`OpenSourceSettings` 样式拆分

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: split open source settings styles`

**覆盖范围**：
- 新增 `src/styles/open-source-settings.css`，集中 `OpenSourceSettings.tsx` 使用的 `.open-source-*` 样式。
- `global.css` 在 `settings.css` 后继续导入该样式文件，保持设置页基础样式先加载、专属样式后加载。
- `settings.css` 从 1551 行降至 1444 行；新样式文件为 96 行。

**业务链路保护**：
- 不改 `OpenSourceSettings.tsx` 的 DOM 结构、className、开源项目数据或设置页 section 路由。
- 不改表格宽度、悬停、链接按钮和许可证标签的原有样式声明；暗色表头/悬停色改用既有 `--bg-tint` / `--bg-blue-tint` 主题变量，避免硬编码颜色残留。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：浮动任务入口样式拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: split floating task panel styles`

**覆盖范围**：
- 新增 `src/styles/floating-task-panel.css`，集中 `FeatureFloatingDock` 与 `FloatingTaskPanel` 使用的 `feature-floating-*` / `task-floating-*` 样式和对应 keyframes。
- `global.css` 在 `chat.css` 后导入浮动任务样式，保持聊天页面基础布局先加载、浮动入口组件样式后加载。
- `chat.css` 从 841 行降至 554 行；新样式文件为 286 行。

**业务链路保护**：
- 不改 `FeatureFloatingDock.tsx`、`FloatingTaskPanel.tsx` 和 `ChatPage.tsx` 的 DOM、className、拖拽/折叠/关闭逻辑。
- `.icon-btn`、聊天消息列表和 Office 预览样式仍留在 `chat.css`，避免通用按钮和聊天页面主体样式被误拆。
- `chat.css` 仍超 CSS 500 行上限 54 行，后续只在找到清晰职责边界时继续拆分。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`OfficePreviewPanel` 样式拆分

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: split office preview panel styles`

**覆盖范围**：
- 新增 `src/styles/office-preview-panel.css`，集中 ChatPage 顶部预览开关与 `OfficePreviewPanel.tsx` 使用的 `office-preview-*` 样式。
- `global.css` 在 `chat.css` 与 `floating-task-panel.css` 后导入 Office 预览样式，保持聊天页面基础布局先加载。
- `chat.css` 从 554 行降至 326 行，低于 CSS 500 行上限；新样式文件为 227 行。

**业务链路保护**：
- 不改 `ChatPage.tsx` 与 `OfficePreviewPanel.tsx` 的 DOM、className、侧栏开关、紧凑模式和移动端隐藏规则。
- 顶部预览按钮的 `flex: 0 0 auto`、`margin-left: 0` 等对齐修复随样式整体迁移，不回滚前序 UI 修复。
- 事件成功色改为直接使用已有 `--success` 主题令牌，避免在新样式文件保留硬编码 fallback。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：Composer 样式职责拆分

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: split composer control styles`

**覆盖范围**：
- 新增 `src/styles/composer-attachments.css`，集中 `ComposerArea.tsx` 附件 chip 列表样式。
- 新增 `src/styles/composer-controls.css`，集中附件/权限/思考按钮、连接状态、popover、发送/停止按钮样式。
- `global.css` 先导入 `composer-controls.css` 再导入 `composer.css`，让 `composer.css` 中既有紧凑模式和窄屏媒体查询继续覆盖控件基础样式。
- `composer.css` 从 673 行降至 344 行，低于 CSS 500 行上限；新附件样式为 61 行，控件样式为 274 行。

**业务链路保护**：
- 不改 `ComposerArea.tsx` 的附件上传/移除、权限切换、思考模式、上下文显示、发送和中止逻辑。
- 响应式规则仍保留在 `composer.css`，并通过 import 顺序保证窄屏隐藏文字、按钮尺寸和紧凑模式行为不被基础控件样式覆盖。
- 权限模式颜色改用已有 `--success` / `--warning` 令牌；发送按钮状态色先保持原视觉，避免引入 UI 状态回归。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

## 二、🔴 P0 问题清单（必须修复）

### 安全性（8 项）

---

#### 🔴 S1 — IPC 通道缺少 Zod Schema 校验

**位置**：`electron/main-modules/ipcHandlers.ts` 多处

**状态**：✅ 已修复（2026-07-05，见“P0 安全：IPC 校验与路径授权”）

**问题**：以下处理外部输入的 IPC 通道未使用 Zod schema 校验，仅依赖手写检查或直接消费参数：

| 通道 | 行号 | 风险 |
|------|------|------|
| `file:writeTempFile` | 422 | 接收 `{prefix,suffix,data}`，仅手写正则清洗 |
| `ocr:recognize` | 693 | 接收 `mode/filePaths`，仅 `normalizeOcrFilePaths` 手写检查 |
| `sandbox:setUserRules` | 490 | 仅 `Array.isArray` 检查 |
| `sandbox:setWritableRoots` | 503 | 仅 `Array.isArray` 检查 |
| `app:openPath` | 150 | 无任何校验 |
| `app:openExternal` | 154 | 有 URL 协议检查，但无 Zod |
| `excel:selectHost` | 272 | 无校验 |
| `file:readAsBase64` | 394 | 无校验 |
| `agent:interrupt` | ipcAgentHandlers.ts:198 | 无校验 |

**原因**：项目规范 PR3 明确要求"新 IPC 通道有 Zod schema"为 P0。`ipcSchemas.ts` 已定义了 `validateInput` 工具和多个 schema，但上述通道未接入。

**建议**：为上述通道在 `ipcSchemas.ts` 补充 Zod schema，在 handler 首行调用 `validateInput`。`file:writeTempFile` 和 `ocr:recognize` 处理文件路径/外部内容，优先级最高。

---

#### 🔴 S2 — IPC Schema 定义未接入（形同虚设）

**位置**：`electron/main-modules/ipcHandlers.ts` + `electron/agent/interaction/ipcAgentHandlers.ts`

**状态**：✅ 已修复（2026-07-05，见“P0 安全：IPC 校验与路径授权”）

**问题**：`ipcSchemas.ts` 中已定义的 schema 未被对应 handler 使用：

| Schema | 对应通道 | 状态 |
|--------|----------|------|
| `SetAlwaysOnTopInput` | `window:setAlwaysOnTop` | ❌ 未调用 |
| `SettingsGetInput/SettingsSetInput` | `settings:get/set` | ❌ 未调用 |
| `ExcelReadRangeInput/ExcelWriteRangeInput` | `excel:readRange/writeRange` | ❌ 未调用 |
| `ThreadIdInput` | `thread:load/delete/resume` | ❌ 未调用 |
| `ThreadNewInput` | `thread:new` | ❌ 未调用 |
| `FilePathInput` | `file:trashFile/openFile/copyPath/revealInExplorer` | ❌ 未调用 |
| `FolderPathInput` | `folder:listFiles` | ❌ 未调用 |
| `ToolConfirmInput/ToolCancelInput` | `tool:confirm/cancel` | ❌ 未调用 |
| `StatsGetSummaryInput` | `stats:getSummary` | ❌ 未调用 |

**原因**：schema 定义了却没调用，给人一种"已校验"的错觉，比"没定义 schema"更危险。

**建议**：在所有已定义 schema 的 handler 首行接入 `validateInput`，或删除未使用的 schema。

---

#### 🔴 S5 — 文件操作缺少路径穿越防护

**位置**：`electron/main-modules/ipcHandlers.ts`

**状态**：✅ 已修复（2026-07-05，见“P0 安全：IPC 校验与路径授权”）

**问题**：以下 IPC 通道接受任意文件路径，无路径范围限制：

```
file:readAsBase64 (394行)    — fs.promises.readFile(filePath)
file:trashFile    (442行)    — shell.trashItem(filePath)
file:openFile     (451行)    — shell.openPath(filePath)
file:revealInExplorer (470行)— shell.showItemInFolder(filePath)
folder:listFiles  (366行)    — fs.promises.readdir(folderPath)
app:openPath      (150行)    — shell.openPath(targetPath)
```

**原因**：虽然当前渲染进程受 `contextIsolation` 保护，但如果渲染进程被 XSS 攻击（如 Markdown 渲染漏洞），攻击者可通过这些 IPC 通道读取/删除/打开系统任意文件。

**建议**：
- 对用户可见的文件操作（trashFile/openFile），限制在用户通过 dialog 选择的路径范围内
- 对 `file:readAsBase64`，限制只允许读取临时目录或用户选择目录下的文件
- 对 `folder:listFiles`，校验 folderPath 在已授权的 pinnedFolders 范围内
- 至少对路径做 `path.resolve()` 后检查是否在允许的根目录下

---

#### 🔴 S3 — Electron 安全配置

**位置**：`electron/main.ts:128`、`electron/main-modules/windowManager.ts:73-74`

**状态**：✅ **通过** — 两个 BrowserWindow 均配置 `contextIsolation: true, nodeIntegration: false`

> 列出此项是为了确认检查通过，无需修复。

---

### 可维护性 — 文件行数超标（14 项）

> 规范上限：TS/TSX ≤ 400 行，React 组件 ≤ 300 行，Store ≤ 400 行，CSS ≤ 500 行

#### 🔴 M1 — 超标文件 TOP 15（完整列表见附录）

| 文件 | 行数 | 上限 | 超出 | 拆分建议 |
|------|------|------|------|----------|
| `electron/agent/core/agentLoop/agentLoop.ts` | **1276** | 400 | +876 | 按职责拆为 runner/streamHandler/snapshot |
| `electron/main-modules/ipcHandlers.ts` | **1057** | 400 | +657 | 按业务域拆为 threadIpc/settingsIpc/excelIpc 等 |
| `src/components/Sidebar.tsx` | **775** | 300 | +475 | 拆出 SidebarCollapsed/Expanded/SortMenu |
| `src/store/settingsStore.ts` | **749** | 400 | +349 | PROVIDER_TEMPLATES 常量抽到独立文件 |
| `src/services/ipcApi.ts` | **745** | 400 | +345 | 按域拆分（agentApi/threadApi/settingsApi） |
| `electron/agent/memory/sessionStore.ts` | **674** | 400 | +274 | — |
| `electron/agent/memory/stateRuntimeStore.ts` | **671** | 400 | +271 | — |
| `electron/agent/tools/implementations/officeOpenXml/advancedExcel.ts` | **662** | 400 | +262 | — |
| `src/store/chatStore.ts` | **623** | 400 | +223 | action 抽到 chatActions.ts |
| `src/components/task/OCRTaskComposerPanel.tsx` | **600** | 300 | +300 | 拆分子表单组件 |
| `electron/agent/core/agentLoop/toolExecutor.ts` | **566** | 400 | +166 | — |
| `src/styles/settings.css` | **1444** | 500 | +944 | 按子组件区块拆分 |
| `src/styles/sidebar.css` | **1101** | 500 | +601 | — |

**完整超标清单**：共 29 个 TS/TSX 文件 + 2 个 CSS 文件（详见附录 A）

---

## 三、🟡 P1 问题清单（应该修复）

### 性能（6 项）

---

#### 🟡 P1-perf — Sidebar 搜索触发 N+1 查询

**位置**：`src/components/Sidebar.tsx:210-220`

**状态**：✅ 已修复（2026-07-06，见“P1 性能：`Sidebar` 搜索文件夹批量加载”）

```typescript
// 当前：每个 pinnedFolder 一次 IPC
pinnedFolders.forEach((folder) => {
  if (folderFiles[folder.path]) return;
  ipcApi.folder.listFiles(folder.path).then(...);  // ← N+1
});
```

**原因**：用户固定的文件夹越多，IPC 调用次数线性增长。

**建议**：主进程新增 `folder:listFilesBatch(paths: string[])` 批量接口。

---

#### 🟡 P4-perf — ChatPage useEffect Promise 未取消

**位置**：`src/components/ChatPage.tsx:87-95`

**状态**：✅ 已修复（2026-07-06，见“P1 性能：`ChatPage` 文件夹加载取消保护”）

```typescript
useEffect(() => {
  if (currentFolderId) {
    ipcApi.folder.listFiles(currentFolderId).then((files) => {
      setCurrentFolderFiles(files);  // ← 组件卸载后仍可能被调用
    });
  }
}, [currentFolderId]);
```

**建议**：使用 cancelled flag 模式：
```typescript
useEffect(() => {
  let cancelled = false;
  if (currentFolderId) {
    ipcApi.folder.listFiles(currentFolderId).then((files) => {
      if (!cancelled) setCurrentFolderFiles(files);
    });
  }
  return () => { cancelled = true; };
}, [currentFolderId]);
```

---

#### 🟡 P4-perf — useExcelConnection 6 处 setTimeout 未清理

**位置**：`src/hooks/useExcelConnection.ts:63, 69, 72, 77, 93, 97, 102`

**状态**：✅ 已修复（2026-07-06，见“P1 性能：`useExcelConnection` timeout 清理”）

**问题**：6 处 `setTimeout` 用于 pulseDot/connectFailed 状态重置，均未在卸载时 `clearTimeout`。

**建议**：用 `useRef` 保存 timer id，在 `useEffect` cleanup 中统一清理。

---

#### 🟡 P3-perf — Sidebar 渲染内未 memo 的排序数组

**位置**：`src/components/Sidebar.tsx:429-461`

**状态**：✅ 已修复（2026-07-06，见“P1 性能：`Sidebar` 排序派生数据 memo”）

**问题**：每次 render 都重新创建 `compareText`、`sortThreads`、`ungroupedThreads`、`groupedByFolder`，当 threads > 50 时排序开销显著。

**建议**：用 `useMemo` 包裹排序结果，依赖 `[threads, pinnedFolders, conversationSortMode, ...]`。

---

#### 🟡 P5-perf — Sidebar 多处内联箭头函数

**位置**：`src/components/Sidebar.tsx:543, 547, 645-647, 765-770`

**状态**：✅ 已修复（2026-07-06，见“P5 性能：`Sidebar` JSX 回调稳定化”）

**问题**：内联箭头函数导致子组件 memo 失效。

**建议**：用 `useCallback` 提取回调。

---

#### 🟡 P2-perf — settingsStore 已正确使用增量持久化 ✅

> 此项为**亮点确认**，非问题。`savePartial` 实现完整，所有 setX action 均使用增量写入。

---

### 可维护性（15 项）

---

#### 🟡 M3 — FolderSection 组件 21 个 Props

**位置**：`src/components/sidebar/FolderSection.tsx:35-61`

**状态**：✅ 已修复（2026-07-06，见“M3：`FolderSection` props 语义分组”）

**问题**：Props 数量远超 10 个上限，是典型 prop drilling。

**建议**：将文件右键菜单相关 6 个 props 合并为 `fileContextMenuApi` 对象；线程回调合并为 `threadActions` 对象。

---

#### 🟡 M8 — electron/ 目录 13 处 console.warn/error 未用 logger

**状态**：✅ 已修复（2026-07-06，见“M8：electron 警告/错误日志统一 logger”）

| 文件 | 行号 | 当前 | 建议 |
|------|------|------|------|
| `electron/main.ts` | 121 | `console.error("Fatal startup error:", err)` | `logger.error` |
| `electron/main-modules/settingsManager.ts` | 97, 177 | `console.warn` | `logger.warn` |
| `electron/agent/security/sandbox/audit.ts` | 91 | `console.warn` | `logger.warn` |
| `electron/agent/runtime/knowledgeRuntime.ts` | 76 | `console.warn` | `logger.warn` |
| `electron/agent/runtime/bridgeRegistry.ts` | 143 | `console.warn` | `logger.warn` |
| `electron/agent/core/agentLoop/agentLoop.ts` | 539, 555, 565, 579, 1128 | `console.warn` × 5 | `logger.warn` |

> **注意**：`desktop/src/` 目录 ✅ 完全无 console.log 使用。

---

### 测试（5 项）

---

#### 🟡 T3 — 74 个 electron/agent 源文件无测试

**状态**：🚧 阶段性补强中（2026-07-06，见“T3：sandbox 策略安全边界测试补强”）

**高风险无测试文件**（节选）：

| 文件 | 风险说明 |
|------|----------|
| `security/sandbox/parseCommand.ts` | ✅ 已有 sandbox.test 覆盖，并补充解析失败边界 |
| `security/sandbox/execPolicy.ts` | ✅ 已有 sandbox.test 覆盖，并补充组合命令/大小写边界 |
| `security/sandbox/defaultRules.ts` | 默认安全规则 |
| `core/agentLoop/turnRunner.ts` | 轮次调度 |
| `core/agentLoop/threadStateManager.ts` | 线程状态机 |
| `knowledge/textChunker.ts` | 文本分块 |
| `knowledge/retriever.ts` | 知识检索 |
| `tools/executors/pythonExecutor.ts` | Python 执行器 |

**建议**：优先为 `security/sandbox/*`、`core/agentLoop` 状态机、`knowledge` 分块/检索补测试。

---

#### 🟡 T3 — src/ 下 45 个文件无测试

**状态**：✅ 阶段性已修复（2026-07-06，见“T3：`src/utils` 纯函数测试补强”）

**优先补测试的纯函数**：
- `src/utils/reasoningSupport.ts`：✅ 已补测试，并修复 invalid per-model reasoningMode 残留
- `src/utils/textCleaner.ts`：✅ 已补测试

---

#### 🟡 T5/PR2 — 全项目 0 个 @MOCK_INTERFACE 标记

**状态**：✅ 阶段性已修复（2026-07-06，见“T5/PR2：关键 mock 接口标记补齐”）

**问题**：规范要求 Mock 数据必须有 `@MOCK_INTERFACE` 注释，但全仓库 grep 返回 0 匹配。

**涉及文件**：
- `src/store/chatStore.test.ts:3-31`
- `src/services/ipcApi.test.ts:14-19`
- `electron/main-modules/mineruOcr.test.ts:51`
- `electron/agent/interaction/eventForwarder.test.ts:14`
- `electron/agent/providers/openaiResponsesClient.test.ts:13`
- `electron/agent/tools/executors/shellExecutor.test.ts:7`

**建议**：关键 mock 已补标记；后续新增或触碰测试 mock 时继续按接口边界补 `// @MOCK_INTERFACE` 注释，避免为历史所有普通 `vi.fn` 机械刷注释。

---

#### 🟡 T4 — shellExecutor 测试覆盖不足

**位置**：`electron/agent/tools/executors/shellExecutor.test.ts`

**状态**：✅ 已修复（2026-07-06，见“T4：`shellExecutor` 测试覆盖补强”）

**问题**：仅 1 个测试用例，只覆盖"复用 sandbox 评估"正常路径。缺少：拒绝路径、prompt 触发、cwd 重定向、超时强杀。

---

#### 🟡 PR4 — electronApi.d.ts 与 ipcApi.ts 类型不同步

**位置**：`src/electronApi.d.ts:335-337` vs `src/services/ipcApi.ts:59-62`

**状态**：✅ 已修复（2026-07-06，见“PR4：IPC 类型声明同步”）

**问题**：Excel 状态类型不一致：
- `electronApi.d.ts`: `detectStatus() => Promise<{ connected; host }>`（缺 version/workbookName/availableHosts）
- `ipcApi.ts`: `detectStatus() => Promise<{ connected; host; version?; workbookName?; availableHosts? }>`

**建议**：以 `ipcApi.ts` 为准补全 `electronApi.d.ts` 的类型声明。

---

### 项目规范（4 项）

---

#### 🟡 PR5 — CHANGELOG 严重滞后

**状态**：✅ 已修复（2026-07-06，见“PR5：CHANGELOG 当前基线同步”）

**问题**：
- 最新条目为 "2026-06-30"，但 git log 显示其后有多个未记录提交
- CHANGELOG 称安装包版本 0.1.2，实际 package.json 已是 0.1.61
- CHANGELOG 称"74 个测试文件、420 个测试"，实际已 105 文件 / 609 测试

**建议**：补 Unreleased 段，更新版本号和测试基线。

---

#### 🟡 S7 — 依赖安全

**状态**：✅ 已修复（2026-07-06，见“S7：依赖安全审计接入 CI”）

**问题**：项目无 `npm audit` / dependabot 配置，新增依赖无安全检查。

**建议**：CI 中增加 `npm audit --audit-level=high` 步骤。

---

## 四、💭 P2 问题清单（酌情处理）

| # | 位置 | 问题 |
|---|------|------|
| 1 | `Sidebar.tsx:190-206` | ✅ 已修复：拖拽 resize 改为 rAF 节流 |
| 2 | `useComposer.ts:326-331` | textarea onChange 每键触发 setState + DOM 写入，可接受 |
| 3 | `preload.ts:89` | ✅ 已修复：onStreamDelta 回调 data 类型补齐 clientId |
| 4 | `Sidebar.tsx` 多处 | ✅ 已修复：排序逻辑集中到 sortSidebarItems 并对派生数组 useMemo |
| 5 | 全项目 | ✅ 阶段性已修复：shell/sandbox 输出截断、默认超时和 watchdog 边界已命名化；其余业务数字按职责边界后续评估 |

---

## 五、✅ 亮点清单（值得学习）

### 架构设计

1. **✅ 沙箱命令执行设计专业** — `security/sandbox/index.ts` 使用 `execFile`（非 `exec`）、Windows `-EncodedCommand` 防注入、ENV_ALLOWLIST 清洗环境变量、进程树强杀。参考 Codex execpolicy 设计，审计日志完整。

2. **✅ 分层架构 + 自动守护测试** — `electron/agent/` 严格分层（interaction → runtime → core），`architecture.test.ts` 自动校验文件存在性和分层依赖。

3. **✅ IPC 抽象层完整** — `desktop/src/` 中 0 处直接 `window.electronAPI` 调用，全部通过 `ipcApi` 抽象层。

4. **✅ IPC schema 单一事实来源** — `ipcSchemas.ts` 集中定义 + `validateInput` 工具 + `z.infer` 推导类型（架构方向正确，只是接入不完整）。

### 性能优化

5. **✅ savePartial 增量持久化** — `settingsStore.ts:411-467`，所有 setX action 只写入变更字段，KEY_MAP + COMPACTION_FIELDS 双层映射，Promise.all 并行写入。

6. **✅ 流式增量缓冲合并** — `chatStore.ts:44-61, 187-206`，`mergeBufferedStreamDeltas` 合并相邻同源 delta，50ms 节流，大幅降低 React 渲染压力。

7. **✅ stats:getSummary 替代 N+1** — 注释明确"单次 IPC 调用，替代 N+1 的 thread:load"。

### 代码质量

8. **✅ ComposerArea composer prop 模式** — 将 useComposer hook 的 25+ 返回值作为单一 `composer` prop 传递，避免 28 props 的 prop drilling。

9. **✅ agentEventHandler 纯函数化** — 从 chatStore 提取为纯函数 `handleAgentEvent(event, current, patches) -> patches`，不依赖 Zustand，便于测试。

10. **✅ chatStore 事件监听器单例管理** — 模块级 unsubscribe 句柄，setupListeners 在 store 初始化时调用一次。

11. **✅ App.tsx 完整的 useEffect 清理** — resize 监听、displayMode 订阅、blur 监听全部有 return cleanup。

### 测试质量

12. **✅ streamCollector.test.ts 边界覆盖优秀** — 覆盖错误重试、JSON 解析失败保留 raw、流中断 emit 已收集内容。

13. **✅ chatStore.test.ts 竞态测试设计优秀** — 覆盖 streaming 中入队、thread id 未就绪拒绝、stale in_progress 元数据不复活。

14. **✅ IPC 通道命名 100% 合规** — 所有 ipcMain.handle 通道均符合 `domain:action` 格式。

---

## 六、修复优先级建议

### 第一周（P0 安全 + 最严重行数超标）

| 优先级 | 任务 | 预估影响 |
|--------|------|----------|
| 1 | 补全 IPC Zod schema 校验（S1 + S2） | 安全防线加固 |
| 2 | 文件操作路径穿越防护（S5） | 安全防线加固 |
| 3 | `agentLoop.ts` 1276 行拆分 | 可维护性 |
| 4 | `ipcHandlers.ts` 1057 行拆分 | 可维护性 |

### 第二周（P1 性能 + 测试）

| 优先级 | 任务 |
|--------|------|
| 5 | Sidebar N+1 查询改批量 |
| 6 | useExcelConnection setTimeout 清理 |
| 7 | ChatPage useEffect Promise 取消 |
| 8 | 沙箱核心模块补测试（parseCommand/execPolicy） |
| 9 | electron/ 13 处 console.warn → logger |

### 第三周（P1 规范 + 剩余行数）

| 优先级 | 任务 |
|--------|------|
| 10 | @MOCK_INTERFACE 标记补全 |
| 11 | electronApi.d.ts 类型同步 |
| 12 | CHANGELOG 更新 |
| 13 | 剩余超标文件拆分（Sidebar/settingsStore 等） |
| 14 | 引入 ESLint + Prettier 工具链 |

---

## 附录 A：超过行数限制的文件完整列表

### TS/TSX 文件（29 个）

| 行数 | 文件 |
|------|------|
| 1276 | `electron/agent/core/agentLoop/agentLoop.ts` |
| 1057 | `electron/main-modules/ipcHandlers.ts` |
| 775 | `src/components/Sidebar.tsx` |
| 749 | `src/store/settingsStore.ts` |
| 745 | `src/services/ipcApi.ts` |
| 674 | `electron/agent/memory/sessionStore.ts` |
| 671 | `electron/agent/memory/stateRuntimeStore.ts` |
| 662 | `electron/agent/tools/implementations/officeOpenXml/advancedExcel.ts` |
| 623 | `src/store/chatStore.ts` |
| 600 | `src/components/task/OCRTaskComposerPanel.tsx` |
| 566 | `electron/agent/core/agentLoop/toolExecutor.ts` |
| 554 | `electron/agent/tools/executors/webSearchExecutors.ts` |
| 544 | `electron/agent/knowledge/documentParser.ts` |
| 537 | `electron/agent/knowledge/sqliteStore.ts` |
| 509 | `electron/agent/tools/implementations/office/officeComActionBridge.ts` |
| 494 | `src/electronApi.d.ts` |
| 481 | `electron/agent/tools/implementations/officeOpenXml/advancedPresentation.ts` |
| 481 | `electron/agent/providers/openaiResponsesClient.ts` |
| 474 | `src/components/settings/GeneralSettings.tsx` |
| 473 | `electron/agent/tools/implementations/excel/excelComBridge.ts` |
| 470 | `electron/agent/shared/types.ts` |
| 440 | `src/components/settings/EditProviderDialog.tsx` |
| 431 | `src/hooks/useComposer.ts` |
| 427 | `electron/agent/tools/implementations/office/presentationComBridge.ts` |
| 426 | `electron/main-modules/settingsManager.ts` |
| 423 | `electron/agent/providers/openaiCompatibleClient.ts` |
| 419 | `src/components/settings/AddProviderDialog.tsx` |
| 416 | `src/components/settings/UsageStats.tsx` |
| 404 | `electron/main-modules/mineruOcr.ts` |

### CSS 文件（2 个）

| 行数 | 文件 |
|------|------|
| 1444 | `src/styles/settings.css` |
| 1101 | `src/styles/sidebar.css` |

---

## 附录 B：未覆盖测试的文件列表

### electron/agent 下无 .test.ts 的源文件（74 个）

<details>
<summary>点击展开完整列表</summary>

```
electron/agent/attachments/imageAttachmentResolver.ts
electron/agent/automation/jscript.ts
electron/agent/automation/json.ts
electron/agent/automation/processLimits.ts
electron/agent/automation/python.ts
electron/agent/core/agentLoop/sessionCompactionConfig.ts
electron/agent/core/agentLoop/summaryGenerator.ts
electron/agent/core/agentLoop/threadLifecycle.ts
electron/agent/core/agentLoop/threadStateManager.ts
electron/agent/core/agentLoop/turnRunner.ts
electron/agent/core/agentLoop/turnState.ts
electron/agent/knowledge/documentParser.ts
electron/agent/knowledge/embeddingService.ts
electron/agent/knowledge/knowledgeIndexer.ts
electron/agent/knowledge/retriever.ts
electron/agent/knowledge/sqliteStore.ts
electron/agent/knowledge/textChunker.ts
electron/agent/knowledge/workbookNotesStore.ts
electron/agent/memory/rolloutSearchContent.ts
electron/agent/memory/stateRuntimeMappers.ts
electron/agent/memory/stateRuntimePaths.ts
electron/agent/memory/stateRuntimeRecovery.ts
electron/agent/memory/stateRuntimeSchema.ts
electron/agent/memory/stateRuntimeTypes.ts
electron/agent/prompts/sections/folderContextPrompt.ts
electron/agent/prompts/sections/formulaAssistantPrompt.ts
electron/agent/prompts/sections/modelPrompt.ts
electron/agent/prompts/sections/officeToolsPrompt.ts
electron/agent/prompts/sections/permissionPrompt.ts
electron/agent/prompts/sections/qualityPrompt.ts
electron/agent/prompts/sections/scenarioPrompt.ts
electron/agent/prompts/sections/scriptPrompt.ts
electron/agent/providers/aiClient.ts
electron/agent/providers/aiClientFactory.ts
electron/agent/providers/aiClientTypes.ts
electron/agent/providers/anthropicClient.ts
electron/agent/providers/modelContextWindows.ts
electron/agent/runtime/agentGlobalSettings.ts
electron/agent/runtime/agentRuntime.ts
electron/agent/security/sandbox/audit.ts
electron/agent/security/sandbox/defaultRules.ts
electron/agent/security/sandbox/execPolicy.ts
electron/agent/security/sandbox/parseCommand.ts
electron/agent/storage/nodeSqlite.ts
electron/agent/tools/contracts/excel.ts
electron/agent/tools/contracts/office.ts
electron/agent/tools/contracts/scriptEnvironment.ts
electron/agent/tools/data/excelFunctionCatalog.ts
electron/agent/tools/executors/createToolExecutors.ts
electron/agent/tools/executors/fileExecutors.ts
electron/agent/tools/executors/localDocumentParser.ts
electron/agent/tools/executors/pythonExecutor.ts
electron/agent/tools/executors/validation.ts
electron/agent/tools/implementations/excel/connectionMetadata.ts
electron/agent/tools/implementations/excel/excelScriptBridgeCom.ts
electron/agent/tools/implementations/excel/excelUiComBridge.ts
electron/agent/tools/implementations/excel/excelVbaComBridge.ts
electron/agent/tools/implementations/excel/sheetOperations.ts
electron/agent/tools/implementations/excel/workbookOperations.ts
electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts
electron/agent/tools/implementations/officeOpenXml/presentationTemplate.ts
electron/agent/tools/registry/file.ts
electron/agent/tools/registry/formula.ts
electron/agent/tools/registry/knowledge.ts
electron/agent/tools/registry/memory.ts
electron/agent/tools/registry/ocr.ts
electron/agent/tools/registry/office.ts
electron/agent/tools/registry/python.ts
electron/agent/tools/registry/range.ts
electron/agent/tools/registry/script.ts
electron/agent/tools/registry/sheet.ts
electron/agent/tools/registry/shell.ts
electron/agent/tools/registry/toolDefinitions.ts
electron/agent/tools/registry/ui.ts
electron/agent/tools/registry/web.ts
electron/agent/tools/registry/workbook.ts
```

</details>

### src/ 下无 .test.ts 的源文件（45 个）

<details>
<summary>点击展开完整列表</summary>

```
src/App.tsx
src/components/chat/AssistantGroupBlock.tsx
src/components/chat/AttachmentImagePreview.tsx
src/components/chat/CompactionNotice.tsx
src/components/chat/ComposerArea.tsx
src/components/chat/MessageBubble.tsx
src/components/chat/ModelQuickSwitch.tsx
src/components/chat/ReasoningBubble.tsx
src/components/chat/ResumeHint.tsx
src/components/chat/ToolCallBubble.tsx
src/components/chat/ToolConfirmDialog.tsx
src/components/ChatPage.tsx
src/components/common/ErrorBoundary.tsx
src/components/common/FloatingTaskPanel.tsx
src/components/common/IconMap.tsx
src/components/common/LoadingSpinner.tsx
src/components/excel/HostSelectionDialog.tsx
src/components/office/OfficePreviewPanel.tsx
src/components/settings/AddProviderDialog.tsx
src/components/settings/DeleteConfirmDialog.tsx
src/components/settings/EditProviderDialog.tsx
src/components/settings/ExecPolicySettings.tsx
src/components/settings/GeneralSettings.tsx
src/components/settings/KnowledgeSettings.tsx
src/components/settings/ModelConfigList.tsx
src/components/settings/ModelSettings.tsx
src/components/settings/modelSettingsI18n.tsx
src/components/settings/OpenSourceSettings.tsx
src/components/settings/ReasoningModeSelect.tsx
src/components/settings/UsageStats.tsx
src/components/settings/useTestConnection.ts
src/components/SettingsPage.tsx
src/components/sidebar/FileContextMenu.tsx
src/components/sidebar/FolderSection.tsx
src/components/sidebar/SidebarSearchPalette.tsx
src/components/sidebar/ThreadContextMenu.tsx
src/components/Sidebar.tsx
src/components/task/CodeTaskComposerPanel.tsx
src/components/task/FormulaTaskComposerPanel.tsx
src/components/task/ReportTaskComposerPanel.tsx
src/hooks/useExcelConnection.ts
src/hooks/useOfficeConnection.ts
src/i18n.ts
src/main.tsx
src/utils/reasoningSupport.ts
src/utils/textCleaner.ts
```

</details>

---

> **审查完毕。** 建议按第六章的优先级顺序修复，第一周聚焦安全防线加固和最严重的文件拆分。
