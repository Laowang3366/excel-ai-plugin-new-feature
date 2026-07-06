# 代码审查报告 — Excel AI 插件项目

> **审查日期**：2026-07-05
> **审查依据**：`docs/code-review-standards.md`
> **审查范围**：desktop/ 全部核心源码（electron/ + src/）
> **审查维度**：安全性 · 正确性 · 可维护性 · 性能 · 测试 · 项目特定规范
> **测试基线**：105 文件 / 609 测试全绿 · typecheck 双通过

---

## 一、审查总结

### 整体评价

项目工程质量**中上**。架构分层清晰、沙箱设计专业、流式增量缓冲和增量持久化等核心模式成熟。本报告最初发现的两类系统性问题已经按修复进度持续收敛，当前应以本节下方的“修复进度”和第六章状态表为准：

1. **IPC 安防线历史风险** — 高风险 IPC 已接入 Zod 校验与路径授权，后续新增通道继续保持 schema、handler、preload/type wrapper 同步。
2. **文件行数历史超标** — 主入口和高风险大文件已阶段性收敛；当前仅剩少量集中协议/桥接边界略超，按自然职责边界继续优化，不为压线硬拆。

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

### 2026-07-06 — S1/S2：IPC Schema 接入状态校正

**状态**：✅ 已修复

**关联提交**：本节所在提交 `docs: correct ipc schema review status`

**覆盖范围**：
- 复核 `desktop/electron/main-modules/ipcHandlers.ts`，确认 `window:setAlwaysOnTop`、`settings:get/set`、`excel:readRange/writeRange/selectHost`、`app:openPath/openExternal/migrateDataPath` 已调用 `validateInput`。
- 复核 `desktop/electron/main-modules/ipcFileHandlers.ts`，确认 `file:writeTempFile/readAsBase64/trashFile/openFile/copyPath/revealInExplorer`、`folder:listFiles/listFilesBatch` 已调用 `validateInput` 并保留路径授权。
- 复核 `desktop/electron/agent/interaction/ipcAgentHandlers.ts`，确认 `agent:startTurn/continueTurn/enqueueTurn/interrupt`、`thread:load/delete/resume/new/updateMetadata`、`threadGraph:*`、`stats:getSummary`、`knowledge:*` 已调用 `validateInput`。
- 复核 OCR/沙箱/AI 子 handler，确认 `ocr:recognize`、`sandbox:setUserRules/setWritableRoots`、`ai:listModels/testConnection` 已接入 schema 校验。
- 将报告 S2 过时的未接入状态表格改为当前已接入状态，避免后续维护误判。

**业务链路保护**：
- 仅校正文档，不改运行时代码、IPC schema、路径授权或 handler 控制流。

**验证证据**：
- `Select-String` 复核上述 handler 中 `validateInput` 与对应 schema 的调用点。
- `git diff --check`

### 2026-07-06 — M1：Office COM 进程探测共享化

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: share office process detection`

**覆盖范围**：
- 新增 `detectOfficeProcess()` 到 `desktop/electron/agent/tools/implementations/office/officeComPowerShell.ts`，统一 Word/PPT 的 Office/WPS 进程探测脚本生成、token 映射和异常 fallback。
- `wordComBridge.ts` 与 `presentationComBridge.ts` 改为复用该 helper，保留原有 host 优先级：Word 优先 `WINWORD`，WPS 文字检测 `wps`；PPT 优先 `POWERPNT`，WPS 演示检测 `wpp/wps`。
- 新增 `presentationComScripts.ts`，集中 PPT 形状文本读取与 slide layout resolver PowerShell 片段；`presentationComBridge.ts` 只保留 COM 连接、文档操作和调用编排。
- `presentationComBridge.ts` 从 423 行降至 386 行，低于 400 行上限；当前全仓剩余超限文件收敛为 `electronApi.d.ts`、`shared/types.ts`、`excelComBridge.ts` 三个集中协议/桥接边界。
- 更新测试源静态基线为 154 个测试文件、768 个 `it/test` 用例。

**业务链路保护**：
- 不改 `detectStatus()`、`open/create/save`、PPT 读写、Word 读写或 COM ProgID 顺序。
- 进程探测仍通过 `executePowerShell`，失败时继续返回 `{ running: false, availableHosts: [] }`。
- PPT/WPS 共享进程名 `wps` 检测由现有测试继续覆盖。

**验证证据**：
- `npm exec vitest run electron/agent/tools/implementations/office/officeComPowerShell.test.ts electron/agent/tools/implementations/office/presentationComBridge.test.ts electron/agent/tools/implementations/office/wordComBridge.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — T3：知识库分块/检索测试补强

**状态**：✅ 阶段性已修复

**关联提交**：本节所在提交 `test: cover knowledge chunking and retrieval`

**覆盖范围**：
- 新增 `desktop/electron/agent/knowledge/textChunker.test.ts`，独立覆盖小文本元数据保真、CSV/表格每 100 行分块且重复表头、Markdown 二到四级标题切分、超长纯文本段落按 token budget 截断，以及普通段落合并。
- 新增 `desktop/electron/agent/knowledge/retriever.test.ts`，独立覆盖向量检索传递当前 embedding profile/source/path filter、低分过滤后按 topK 返回、embedding 失败或向量低分时关键词降级、prompt 注入去重来源，以及工具结果空状态文案。
- 更新测试源静态基线为 151 个测试文件、753 个 `it/test` 用例。

**业务链路保护**：
- 仅新增测试和文档，不改 `TextChunker`、`Retriever`、`SqliteStore`、`EmbeddingService` 的生产逻辑。
- 检索测试使用最小 mock 锁定调用契约，避免引入 SQLite/网络依赖；已有 `rag.test.ts` 集成覆盖继续保留。

**验证证据**：
- `npm exec vitest run electron/agent/knowledge/textChunker.test.ts electron/agent/knowledge/retriever.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — T3：Python 执行器测试补强

**状态**：✅ 阶段性已修复

**关联提交**：本节所在提交 `test: cover python executor contract`

**覆盖范围**：
- 新增 `desktop/electron/agent/tools/executors/pythonExecutor.test.ts`，覆盖 `python.execute` 与兼容别名 `python_execute` 注册到同一 executor。
- 覆盖缺少 `code` 时在 sandbox 评估和 Python 执行前直接返回参数错误。
- 覆盖 sandbox `effectiveWorkdir`、`redirected`、`decision`、请求工作目录、默认 `90000ms` timeout 和 `os.homedir()` fallback 的返回结构。
- 覆盖脚本失败时透出 stderr，避免调用方只能看到泛化失败。
- 更新测试源静态基线为 152 个测试文件、758 个 `it/test` 用例。

**业务链路保护**：
- 仅新增测试和文档，不改 `pythonExecutor.ts`、`automation/python.ts` 或 sandbox 执行逻辑。
- 测试 mock `executePlainPythonScript` 与 `evaluateCommand`，不依赖本机 Python、临时文件执行或真实安全策略状态。

**验证证据**：
- `npm exec vitest run electron/agent/tools/executors/pythonExecutor.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — T3：AgentLoop 状态机测试补强

**状态**：✅ 阶段性已修复

**关联提交**：本节所在提交 `test: cover agent loop state helpers`

**覆盖范围**：
- 新增 `desktop/electron/agent/core/agentLoop/threadStateManager.test.ts`，覆盖初始 `not_loaded` 快照、active idle unload 阈值、running 状态不卸载、markIdle 重置活跃时间、markUnloaded/clear 状态转换，以及 0/负数阈值禁用 idle unload。
- 新增 `desktop/electron/agent/core/agentLoop/turnRunner.test.ts`，覆盖 `createTurn`、`createUserMessageItem`、`completeTurn` 的结构契约、时间戳、附件与 clientId 保留。
- 更新测试源静态基线为 154 个测试文件、766 个 `it/test` 用例。

**业务链路保护**：
- 仅新增测试和文档，不改 AgentLoop 生产逻辑。
- 测试只控制 `Date.now()`，不接入真实模型、工具执行、rollout 写入或线程持久化。

**验证证据**：
- `npm exec vitest run electron/agent/core/agentLoop/threadStateManager.test.ts electron/agent/core/agentLoop/turnRunner.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

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

### 2026-07-05 — P0 可维护性：`agentLoop.ts` 阶段性拆分（M1 已关闭）

**状态**：✅ 已关闭（入口文件超标已收敛，后续仅按自然职责边界继续优化）

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
- `agentLoop.ts` 已从 1276 行降至当前约 346 行，低于 400 行规范上限；`ipcHandlers.ts` 当前约 307 行，Sidebar/settingsStore/ipcApi/chatStore 等前端主入口也已按职责边界收敛到上限内。
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
- 此阶段已关闭 `agentLoop.ts`、`ipcHandlers.ts`、Sidebar、settingsStore、ipcApi、chatStore 等主入口单文件超标；后续不再为追求行数继续拆碎，转向其它明确职责边界的高价值问题。

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

### 2026-07-06 — M1 可维护性：PPT OpenXML 内容页 helper 拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract presentation slide content helpers`

**覆盖范围**：
- 新增 `officeOpenXml/presentationSlideContent.ts`，收拢新增 PPT 内容页的参数归一化、bullet 列表合并、内容页 XML 生成和空 slide rels 生成。
- `advancedPresentation.ts` 改为导入上述 helper，保留 zip 包读写、presentation 关系维护、slide id/rel id 编号、Content Types 更新和删除页链路在原模块。
- 新增 `presentationSlideContent.test.ts`，覆盖数组 slides、单页参数、字段别名、bullet 合并、XML 转义和 blank 空页不生成占位文本框。
- 同步 `CHANGELOG.md` 当前测试源基线为 136 个测试文件、706 个 `it/test` 用例。

**业务链路保护**：
- `addSlides` 的 OpenXML 写包顺序、目标文件选择、关系插入、Content Types 覆盖写入和 `presentationDone()` 返回结构均未改动。
- 仅移动纯函数边界，没有把 PPT 包关系维护主链路拆散；`advancedPresentation.ts` 当前约 406 行，略超 400 但保留高内聚流程，后续只在出现明确边界时继续拆。

**验证证据**：
- `npm exec vitest run electron/agent/tools/implementations/officeOpenXml/presentationSlideContent.test.ts electron/agent/tools/implementations/officeOpenXml/advancedPresentation.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P1 可维护性：Composer 附件解析 helper 拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract composer attachment file helpers`

**覆盖范围**：
- 新增 `hooks/composerAttachmentFiles.ts`，收拢附件扩展名/MIME 判断、本地路径获取、路径缺失时写入临时文件，以及 `resolveDroppedFiles()`。
- `useComposer.ts` 改为导入并 re-export `resolveDroppedFiles()`，保留旧测试和外部导入路径兼容。
- 粘贴附件与拖拽附件统一调用 `resolveDroppedFiles()`，移除 `handlePaste` 中重复的本地路径/临时文件写入循环。
- 新增 `composerAttachmentFiles.test.ts`，覆盖无 MIME 的图片扩展名识别、本地文件路径优先和路径缺失图片写入临时文件。
- 同步 `CHANGELOG.md` 当前测试源基线为 137 个测试文件、709 个 `it/test` 用例。

**业务链路保护**：
- `ipcApi.file.getPathForFile()` 优先、`file:writeTempFile` fallback、图片前缀 `image`、普通附件前缀 `attachment`、去重合并附件列表等行为保持不变。
- 对话输入框状态、草稿切换、发送/恢复中断、文件/图片/文件夹选择、弹层关闭和拖拽 hover 逻辑未改动。

**验证证据**：
- `npm exec vitest run src/hooks/composerAttachmentFiles.test.ts src/hooks/useComposer.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P1 可维护性：工具审批策略 helper 拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract tool approval policy`

**覆盖范围**：
- 新增 `agentLoop/toolApproval.ts`，集中权限模式判断、always-allowed 工具集合、审批请求默认放行和审批回调委托。
- `toolExecutor.ts` 改为导入审批策略，并继续 re-export `ToolApprovalConfig`、`shouldRequireApproval()`、`requestToolApproval()`、`markToolAlwaysAllowed()`、`getAlwaysAllowedTools()`、`clearAlwaysAllowedTools()`，保持旧入口兼容。
- 新增 `toolApproval.test.ts`，覆盖权限模式、always-allowed 覆盖、无回调默认批准和带回调委托审批。
- 同步 `CHANGELOG.md` 当前测试源基线为 138 个测试文件、712 个 `it/test` 用例。

**业务链路保护**：
- `processToolCalls()` 的 sandbox forbidden/prompt 覆盖、审批取消/异常处理、工具执行、`TurnItem` 事件顺序、执行日志落库和结果 item 生成均未改动。
- 旧测试仍从 `toolExecutor` 导入审批 API，证明外部调用路径保持兼容。

**验证证据**：
- `npm exec vitest run electron/agent/core/agentLoop/toolApproval.test.ts electron/agent/core/agentLoop/toolExecutor.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：`UsageStats` 数据 helper 拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract usage stats data helpers`

**覆盖范围**：
- 新增 `components/settings/usageStatsData.ts`，集中使用统计文案、模型颜色、数字/日期格式化、`stats:getSummary` 行转换和时间范围聚合。
- `UsageStats.tsx` 改为通过 `buildUsageStatsData()` 获取视图数据，组件保留 IPC 加载、状态和 JSX 渲染职责。
- 新增 `usageStatsData.test.ts`，覆盖零 token 过滤、行排序、按范围聚合、模型排序、连续天数、估算标记和格式化输出。
- 同步 `CHANGELOG.md` 阶段性测试源基线为 139 个测试文件、715 个 `it/test` 用例。

**业务链路保护**：
- `ipcApi.stats.getSummary()` 调用、loading/error 状态、范围切换、刷新按钮、图表 DOM/className 和样式入口均未改动。
- 本地日界线聚合规则保持与原组件一致，仅移动纯计算逻辑并补测试。

**验证证据**：
- `npm exec vitest run src/components/settings/usageStatsData.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：`AddProviderDialog` 草稿 helper 拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract add provider draft helpers`

**覆盖范围**：
- 新增 `components/settings/addProviderDraft.ts`，集中新增供应商弹窗的空白草稿、模板草稿和 `AiProviderConfig` 构造逻辑。
- `AddProviderDialog.tsx` 保留表单状态、模板选择事件、测试连接、聚合模型列表和 JSX 渲染职责，仅将字段批量赋值与最终配置拼装委托给 helper。
- 新增 `addProviderDraft.test.ts`，覆盖空白/模板草稿、模板供应商配置构造、自定义供应商默认值和空可选字段处理。
- 同步 `CHANGELOG.md` 阶段性测试源基线为 140 个测试文件、718 个 `it/test` 用例。

**业务链路保护**：
- 不改 `useTestConnection`、`ipcApi.ai.testConnection()`、`ReasoningModeSelect`、`ModelConfigList`、模板列表和按钮启用条件。
- 继续沿用 `effectiveReasoningMode` 写入 `enableReasoning/reasoningMode`，聚合平台 `modelConfigs` 为空时仍不写入可选字段。

**验证证据**：
- `npm exec vitest run src/components/settings/addProviderDraft.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：Markdown 表格解析共享化

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: share markdown table extraction`

**覆盖范围**：
- 新增 `electron/shared/markdownTables.ts`，集中 GitHub 风格 Markdown 表格行提取逻辑。
- `mineruOcr.ts` 改为消费共享解析函数，文件从 404 行降至 367 行，低于 400 行上限。
- `localDocumentParser.ts` 不再从 `main-modules/mineruOcr` 导入通用 helper，改为依赖 `electron/shared/markdownTables.ts`，解除 agent executor 对主进程 OCR 模块的反向依赖。
- 新增 `markdownTables.test.ts`，并保留 `mineruOcr.test.ts` 对 MinerU 付费/免费流程和表格结果结构的覆盖。
- 同步 `CHANGELOG.md` 阶段性测试源基线为 141 个测试文件、719 个 `it/test` 用例。

**业务链路保护**：
- 不改 MinerU 付费批量流程、Agent 免费流程、签名上传、轮询、ZIP/Markdown 下载、错误格式化和本地文档解析策略。
- Markdown 表格解析算法原样迁移，OCR 结果中的 `rows` 结构保持不变。

**验证证据**：
- `npm exec vitest run electron/shared/markdownTables.test.ts electron/main-modules/mineruOcr.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：线程 IPC wrapper 域拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract thread ipc wrapper`

**覆盖范围**：
- 新增 `src/services/ipcThreadApi.ts`，集中 `thread` 与 `threadGraph` wrapper 的转发、fallback 和错误兜底行为。
- `ipcApi.ts` 改为组合 `createThreadIpcApi(getRaw)`，保留原 `ipcApi.thread.*` 与 `ipcApi.threadGraph.*` 调用面不变，文件从 414 行降至 354 行。
- 新增 `ipcThreadApi.test.ts`，覆盖 runtimeStatus / upsertSpawnEdge 转发，以及 raw IPC 不可用时的安全 fallback。
- 保留 `ipcApi.test.ts` 对整体 wrapper 的 thread runtime、threadGraph 和 fallback 行为覆盖。
- 同步 `CHANGELOG.md` 阶段性测试源基线为 142 个测试文件、721 个 `it/test` 用例。

**业务链路保护**：
- 不改 preload 暴露、`IIpcApi` 类型、线程加载/删除/恢复/新建/metadata 更新、线程图边创建/关闭/后代列表的参数与返回值。
- 仍由 `getRaw()` 统一读取 `window.electronAPI`，缺失 IPC 时继续返回原有安全空值或抛出原有错误。

**验证证据**：
- `npm exec vitest run src/services/ipcThreadApi.test.ts src/services/ipcApi.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：`GeneralSettings` 文案模块拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract general settings text`

**覆盖范围**：
- 新增 `src/components/settings/generalSettingsText.ts`，集中常规设置页中英文文案和窗口透明度标签/提示文案。
- `GeneralSettings.tsx` 保留设置状态、IPC 读取/迁移、MinerU token 保存、store action 和 JSX 渲染职责，文件从 427 行降至 318 行。
- 新增 `generalSettingsText.test.ts`，覆盖中英文文案可用性、动态数组设置文案和窗口透明度提示。
- 同步 `CHANGELOG.md` 阶段性测试源基线为 143 个测试文件、723 个 `it/test` 用例。

**业务链路保护**：
- 不改 `ipcApi.app.getDataPath/selectDataPath/migrateDataPath/openPath`、`ipcApi.settings.get/set`、`useSettingsStore` action、设置卡片 DOM/className 和现有开关/滑块交互。
- 窗口透明度、动态数组函数、自动压缩、OCR token 和数据目录迁移继续走原组件事件处理。

**验证证据**：
- `npm exec vitest run src/components/settings/generalSettingsText.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：OpenAI 兼容工具名解析拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract openai tool names`

**覆盖范围**：
- 新增 `electron/agent/providers/openaiToolNames.ts`，集中 Chat Completions 工具名点号转下划线、以及 provider 返回工具名还原逻辑。
- `openaiCompatibleClient.ts` 改为导入并 re-export `sanitizeToolName` / `desanitizeToolName`，保留旧导入路径兼容，文件从 423 行降至 385 行。
- 新增 `openaiToolNames.test.ts`，覆盖 Office 复合前缀、常规命名空间、未知工具名保留；保留 `openaiCompatibleClient.test.ts` 对兼容导出与 SSE/tool call 行为覆盖。
- 同步 `CHANGELOG.md` 当前测试源基线为 144 个测试文件、725 个 `it/test` 用例。

**业务链路保护**：
- 不改 Chat Completions 请求体、SSE 解析、reasoning 配置、工具调用 begin/delta/end 事件顺序和错误格式化。
- 旧的 `import { sanitizeToolName } from "./openaiCompatibleClient"` 继续可用。

**验证证据**：
- `npm exec vitest run electron/agent/providers/openaiToolNames.test.ts electron/agent/providers/openaiCompatibleClient.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：供应商模型选择器复用

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: share provider model selector`

**覆盖范围**：
- 新增 `src/components/settings/ProviderModelSelector.tsx`，集中聚合平台模型下拉、预设模型下拉、自定义模型输入三种 UI 分支。
- `AddProviderDialog.tsx` 与 `EditProviderDialog.tsx` 改为复用共享选择器，分别从 409/429 行降至 379/393 行。
- 新增 `ProviderModelSelector.test.ts`，覆盖选择器类型判定、聚合平台空模型选项和编辑时保留不在预设列表内的当前模型。
- 同步 `CHANGELOG.md` 当前测试源基线为 145 个测试文件、728 个 `it/test` 用例。

**业务链路保护**：
- 不改新增供应商的模板选择、草稿构造、连接测试和添加逻辑。
- 不改编辑供应商的 `applyModelConfig()` per-model 上下文窗口/推理模式同步、模型刷新按钮、保存 patch 和聚合模型列表管理。
- 两个弹窗仍使用原有 CSS className，避免改变设置页布局样式。

**验证证据**：
- `npm exec vitest run src/components/settings/ProviderModelSelector.test.ts src/components/settings/addProviderDraft.test.ts src/components/settings/editProviderPatch.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：`KnowledgeSettings` 文案与格式化拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract knowledge settings text`

**覆盖范围**：
- 新增 `src/components/settings/knowledgeSettingsText.ts`，集中知识库设置页双语文案、知识来源类型标签、索引时间、来源统计和文件夹索引结果汇总。
- `KnowledgeSettings.tsx` 保留知识库开关、来源列表加载、添加文件/文件夹、删除来源、重建索引和 JSX 渲染职责，文件从 378 行降至 255 行。
- 新增 `knowledgeSettingsText.test.ts`，覆盖来源统计、文件夹索引结果汇总、类型标签映射和时间格式化。
- 同步 `CHANGELOG.md` 当前测试源基线为 146 个测试文件、732 个 `it/test` 用例。

**业务链路保护**：
- 不改 `ipcApi.knowledge.listSources/indexFile/indexFolder/deleteFile/reindexAll` 调用顺序和返回处理。
- 不改 `useSettingsStore` 的 `knowledgeEnabled/setKnowledgeEnabled` 设置链路。
- 不改知识库设置页 DOM className、来源列表结构和操作按钮布局。

**验证证据**：
- `npm exec vitest run src/components/settings/knowledgeSettingsText.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：`ChatPage` 简单任务面板拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract simple task composer`

**覆盖范围**：
- 新增 `src/components/task/SimpleTaskComposerPanel.tsx`，承接清洗/图表任务的选区输入、需求输入、选区拾取和提交按钮 UI。
- `ChatPage.tsx` 只保留 simple task 草稿更新、选区拾取、任务提交和各任务面板编排职责，文件从 341 行降至 309 行。
- 新增 `SimpleTaskComposerPanel.test.ts`，覆盖 simple task payload 对 range/task 的 trim 和空字段省略规则。
- 同步 `CHANGELOG.md` 当前测试源基线为 147 个测试文件、734 个 `it/test` 用例。

**业务链路保护**：
- 不改公式、代码、OCR、报告任务面板的渲染和提交链路。
- 清洗/图表任务仍通过原 `handleTaskSubmit()` 写入输入框、调用 `sendMessage()`、清空输入并关闭意图。
- 保留原 `task-composer-panel`、`task-field`、`range-input-row`、`btn-pick-range`、`task-submit-btn` className。
- `ChatPage.tsx` 仍略超 React 300 行上限，剩余内容主要是页面编排；后续只在有自然边界时继续拆。

**验证证据**：
- `npm exec vitest run src/components/task/SimpleTaskComposerPanel.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：`ComposerArea` 思考模式按钮拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract composer thinking mode`

**覆盖范围**：
- 新增 `src/components/chat/ComposerThinkingModeButton.tsx`，集中 composer 底部思考模式按钮的 provider/model 推理选项解析、当前模式标签和弹层选项渲染。
- `ComposerArea.tsx` 保留输入框、附件、权限、上下文用量、模型切换、发送/停止等编排职责，文件从 324 行降至 267 行。
- 新增 `ComposerThinkingModeButton.test.ts`，覆盖无 active provider 时隐藏、per-model reasoningMode 优先于 provider 级配置、无配置时按 provider/template 默认值适配。
- 同步 `CHANGELOG.md` 当前测试源基线为 148 个测试文件、737 个 `it/test` 用例。

**业务链路保护**：
- 不改 `useComposer` 返回结构、附件上传/移除、权限切换、上下文用量显示、模型快速切换、发送和中止逻辑。
- 思考模式仍通过 `updateProvider(activeProviderId, { reasoningMode })` 写回设置，保留原 `thinking-mode-*`、`composer-popover`、`popover-item` className。
- 打开思考模式弹层时仍关闭附件与权限弹层，避免多个 popover 同时展开。

**验证证据**：
- `npm exec vitest run src/components/chat/ComposerThinkingModeButton.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：`FeatureFloatingDock` 几何逻辑拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract feature dock geometry`

**覆盖范围**：
- 新增 `src/components/common/featureFloatingDockGeometry.ts`，集中浮动功能栏的 pointer action 判定、外部点击折叠判断、初始定位、resize 定位和展开态边界约束。
- `FeatureFloatingDock.tsx` 继续从原入口 re-export 已有 helper，保留 `FeatureFloatingDock.test.ts` 的导入路径不变，组件从 324 行降至 255 行。
- 复用现有 14 个 FeatureFloatingDock 单测覆盖拖拽/点击行为、默认定位、窗口 resize 和展开态边界约束。

**业务链路保护**：
- 不改 `feature-floating-*` DOM className、展开/折叠渲染结构、intent 点击回调和拖拽事件绑定。
- 不改默认定位常量、折叠球尺寸、展开卡片宽度和边缘留白语义，仅移动纯函数位置。
- `FeatureFloatingDock.tsx` 的公共 helper 导出保持兼容。

**验证证据**：
- `npm exec vitest run src/components/common/FeatureFloatingDock.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：`FolderSection` 会话项复用

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: share sidebar thread item`

**覆盖范围**：
- 新增 `src/components/sidebar/SidebarThreadItem.tsx`，集中侧边栏会话项的 active class、时间、预览、运行/完成/失败状态标记和文件夹内图标渲染。
- `FolderSection.tsx` 的文件夹内会话列表与 `UngroupedThreadList` 复用同一组件，文件从 305 行降至 259 行。
- 新增 `SidebarThreadItem.test.ts`，覆盖 running runtime state、active thread 视为已读隐藏非运行状态、已读 completed metadata 不再显示状态标记。
- 同步 `CHANGELOG.md` 当前测试源基线为 149 个测试文件、740 个 `it/test` 用例。

**业务链路保护**：
- 不改文件夹展开、文件列表切换、文件加入会话、文件右键菜单、文件夹新建会话和移除文件夹逻辑。
- 保留 `sidebar-thread-item`、`sidebar-thread-in-folder`、`thread-item-*`、`thread-status-indicator` className。
- 右键菜单仍按原语义传入 `inFolder=true`；未分组会话仍不传第三参。

**验证证据**：
- `npm exec vitest run src/components/sidebar/SidebarThreadItem.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1 可维护性：`GeneralSettings` 数据存储卡片拆分

**状态**：✅ 已修复

**关联提交**：本节所在提交 `refactor: extract general storage card`

**覆盖范围**：
- 新增 `src/components/settings/GeneralSettingsStorageCard.tsx`，集中常规设置页的数据目录展示、打开、复制和迁移按钮 UI。
- `GeneralSettings.tsx` 保留 `ipcApi.app.getDataPath/openPath/selectDataPath/migrateDataPath`、剪贴板复制、迁移状态和 `loadSettings()` 回调职责，文件从 318 行降至 282 行。
- `CHANGELOG.md` 当前测试源基线保持为 149 个测试文件、740 个 `it/test` 用例。

**业务链路保护**：
- 不改数据目录读取、打开、复制、迁移、迁移后重载设置和错误展示逻辑。
- 保留 `settings-card`、`storage-path-row`、`storage-path-input`、`settings-action-btn` className。
- MinerU token、透明度、动态数组、自动压缩和当前模型上下文显示链路未改。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

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

### P1 可维护性：Excel 连接探测职责拆分

**状态**：已修复

**关联提交**：`refactor: extract excel connection probe`

**覆盖范围**：
- 新增 `electron/agent/tools/implementations/excel/excelConnectionProbe.ts`，集中维护 Excel/WPS 进程探测和 COM 活跃对象验证。
- `excelComBridge.ts` 只替换原私有方法调用点，继续保留连接状态机、宿主选择、重试、工作簿操作、range 读写和公式上下文入口。
- 文件拆分以职责边界为准，没有为了压行数拆散连接流程；`excelComBridge.ts` 从约 403 行收敛到约 368 行。

**业务链路保护**：
- `detectExcelProcess()` 的返回结构仍包含 `running`、`host`、`availableHosts`，继续兼容 Excel/WPS 双宿主选择。
- `verifyExcelComAvailable()` 仍按 `Excel.Application` / `Ket.Application` ProgID 获取活跃 COM 对象，并保留版本号和活动工作簿名称回传。
- `ensureConnectedInternal`、`checkStatus`、`selectHost` 的状态写入、重试和 fallback 语义保持不变。

**验证证据**：
- `npm exec vitest run electron/agent/tools/implementations/excel/excelComBridge.test.ts electron/agent/tools/implementations/excel/rangeOperations.test.ts electron/agent/tools/implementations/excel/formulaOperations.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### P1 可维护性：设置数据路径职责拆分

**状态**：已修复

**关联提交**：`refactor: extract settings data path helpers`

**覆盖范围**：
- 新增 `electron/main-modules/settingsDataPath.ts`，集中维护安装目录/用户目录选择、bootstrap 数据路径读取写入、旧默认目录迁移、路径边界判断和同步/异步目录复制工具。
- `settingsManager.ts` 从约 427 行收敛到约 282 行，继续负责 settings store 实例、SessionStore/AgentGraphStore/StateRuntimeStore 生命周期、数据目录迁移编排、AI 配置读取和窗口主题/透明度应用。
- 保留 `settingsManager.ts` 对外导出 `getActiveDataPath`，避免影响 `main.ts`、`ipcHandlers.ts` 等既有导入入口。

**业务链路保护**：
- `migrateDataPath` 的互斥锁、目标路径校验、会话/知识库/日志复制、settings store 切换、日志目录刷新、AgentLoop store 刷新和知识库 runtime 重载顺序保持不变。
- 旧默认数据目录迁移仍在 settings store 初始化前执行，继续复制 settings、sessions、knowledge、logs 四类数据。
- bootstrap `dataPath` 写入封装为 `setConfiguredDataPath`，底层仍使用原 `excel-ai-bootstrap` store。

**验证证据**：
- `npm run typecheck`
- `npm exec vitest run electron/main-modules/ipcPathSecurity.test.ts electron/main-modules/ipcFileHandlers.test.ts`
- `npm run build`
- `git diff --check`

### P1 可维护性：SessionStore 文件扫描职责拆分

**状态**：已修复

**关联提交**：`refactor: extract session store file helpers`

**覆盖范围**：
- 新增 `electron/agent/memory/sessionStoreFiles.ts`，集中维护默认 sessions 根目录、按日期生成 rollout JSONL 路径、递归收集 rollout 文件和扫描线程元数据。
- `sessionStore.ts` 从约 443 行收敛到约 374 行，继续负责线程创建、rollout 写入、数据库投影、压缩归档搜索、加载/删除/元数据更新和使用统计入口。
- 保留 `getDefaultSessionsRoot` 从 `sessionStore.ts` re-export，兼容现有导入路径。

**业务链路保护**：
- `appendRolloutItems` 仍先写数据库投影，再把 JSONL 行交给 `AsyncRolloutWriter`，写入顺序和审计副本语义不变。
- `listThreads` 仍在 flush 后扫描 `rollout-*.jsonl`，跳过损坏文件，并按 `updatedAt` 降序返回。
- `findRolloutPath` 仍维护 `threadId -> filePath` 缓存；`getUsageSummary` 与路径查找共用同一递归扫描 helper，避免重复实现再次分叉。

**验证证据**：
- `npm exec vitest run electron/agent/memory/sessionStore.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### P1 可维护性：IPC Office wrapper 拆分

**状态**：已修复

**关联提交**：`refactor: extract office ipc wrapper`

**覆盖范围**：
- 新增 `src/services/ipcOfficeApi.ts`，集中维护 `excel` 与 `office` 两组前端 IPC wrapper。
- `ipcApi.ts` 改为组合 `createOfficeIpcApi`、`createThreadIpcApi` 和 `createKnowledgeIpcApi`，主入口从约 355 行收敛到约 299 行。
- 保留 `ipcApi.ts` 对外导出和 `IIpcApi` 类型不变，调用方无需迁移。

**业务链路保护**：
- `excel.detectStatus/connect/selectHost/getSelection/getSelectionAddress/readRange/inspectWorkbook/writeRange` 的 fallback 返回值保持不变。
- `excel.readRange(sheetName, range, expand)` 仍透传第三参 `expand`，避免前端 wrapper 再次丢失 spill/currentRegion 能力。
- Word/PPT 状态检测仍在 IPC 不可用时返回 `{ connected: false, host: "unknown" }`。

**验证证据**：
- `npm exec vitest run src/services/ipcApi.test.ts src/services/ipcThreadApi.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### P1 可维护性：SQLite 知识库 schema 与来源摘要职责拆分

**状态**：已修复

**关联提交**：`refactor: extract sqlite knowledge schema helpers`

**覆盖范围**：
- 新增 `electron/agent/knowledge/sqliteStoreSchema.ts`，集中维护 `knowledge_entries` / `knowledge_sources` 建表、索引和 `embedding_provider/model/dimensions` 迁移。
- 新增 `electron/agent/knowledge/sqliteSourceSummaries.ts`，集中维护从 `knowledge_entries` 回填 `knowledge_sources` 的来源摘要逻辑。
- `sqliteStore.ts` 从约 468 行收敛到约 357 行，继续保留条目写入、批量事务、删除、向量检索、关键词检索、来源 API、统计和维护入口。

**业务链路保护**：
- `init()` 仍先创建目录、打开 SQLite、设置 WAL，再执行建表与 embedding profile 迁移。
- `listSources()` 仍在查询前执行来源摘要回填，保留旧索引缺少 `knowledge_sources` 记录时的兼容补全行为。
- 向量检索的 provider/model/dim 过滤、维度不匹配返回 0、关键词去重和 `topK` 截断语义保持不变。

**验证证据**：
- `npm run typecheck`
- `npm exec vitest run electron/agent/knowledge/rag.test.ts`
- `npm exec vitest run electron/agent/knowledge/knowledgeRegistry.test.ts electron/agent/knowledge/knowledgeWriter.test.ts electron/agent/runtime/knowledgeRuntime.test.ts electron/agent/tools/executors/knowledgeExecutors.test.ts`
- `npm run build`
- `git diff --check`

### P1 可维护性：StateRuntime 线程状态表操作收敛

**状态**：已修复

**关联提交**：`refactor: extract state runtime thread tables`

**覆盖范围**：
- 新增 `electron/agent/memory/stateRuntimeThreads.ts`，集中维护 `thread_snapshots` 与 `thread_runtime` 的 upsert/list/get SQL 和行映射调用。
- `StateRuntimeStore` 从约 428 行收敛到约 361 行，继续负责四库生命周期、恢复报告、迁移/WAL、跨库事务、thread name 联动和 rollout/tool/goal/memory helper 编排。
- `metadata.name` 写入 `thread_names` 的联动仍保留在主 store 中，避免把跨表业务编排藏进单表 helper。

**业务链路保护**：
- `upsertThreadSnapshot` 的字段、JSON 序列化、`ON CONFLICT` 更新列和 `metadata.name` 追加逻辑保持不变。
- `listThreadSnapshots` 仍按 `updated_at DESC` 排序，`getThreadSnapshot` 仍复用 `mapThreadSnapshot`。
- `updateThreadRuntime` 仍用当前时间写 `updated_at`，`getThreadRuntime` 的 undefined 归一化保持不变。

**验证证据**：
- `npm exec vitest run electron/agent/memory/stateRuntimeStore.test.ts electron/agent/memory/stateRuntimeMappers.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### P1 可维护性：PPT OpenXML 包结构 helper 拆分

**状态**：已修复

**关联提交**：`refactor: extract presentation package helpers`

**覆盖范围**：
- 新增 `electron/agent/tools/implementations/officeOpenXml/presentationPackageParts.ts`，集中维护 slide part 编号、presentation slide id、relationship id、content type override、slide entry 收集和 XML 属性读取。
- `advancedPresentation.ts` 从约 409 行收敛到约 323 行，继续保留 `createPresentation`、`applyTheme`、`deleteSlides`、`addSlides` 和 COM fallback 判定等业务流程。
- `[Content_Types].xml` 部件名改为复用同一 `CONTENT_TYPES_PART` 常量，避免新增 helper 后继续双写。

**业务链路保护**：
- `addSlides` 的 slide XML、rels XML、content type 更新顺序和返回 changed parts 保持不变。
- `deleteSlides` 的 slide id 删除、relationship 删除、content type override 移除和“至少保留一张”校验保持不变。
- `applyTheme` 的 run 颜色替换逻辑仍留在主业务文件中，避免把主题操作与包结构 helper 混在一起。

**验证证据**：
- `npm exec vitest run electron/agent/tools/implementations/officeOpenXml/advancedPresentation.test.ts electron/agent/tools/implementations/officeOpenXml/presentationSlideContent.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### P1 可维护性：settingsStore 加载归一化拆分

**状态**：已修复

**关联提交**：`refactor: extract settings loaded state`

**覆盖范围**：
- 新增 `src/store/settingsLoadedState.ts`，集中维护 `settings.getAll()` 原始对象到 Zustand state patch 的归一化，并检测 provider `enableReasoning -> reasoningMode` 迁移。
- 新增 `src/store/settingsValues.ts`，集中维护窗口透明度上下限与 `normalizeWindowOpacity()`，`settingsStore.ts` 保留 re-export 兼容原导入路径。
- `settingsStore.ts` 从约 347 行收敛到约 317 行，`loadSettings` 改为一次构建 patch、一次 `set()`，继续负责 IPC 读取、迁移写回、`isConfigured` 计算和所有 setter 持久化。

**业务链路保护**：
- 动态数组函数支持仍在缺省时保持开启；保存/切换仍写入 `dynamicArrayFunctionsEnabled`。
- provider 迁移仍会规范化 reasoning 配置并写回 `aiProviders`，`isConfigured` 仍按 active provider 计算。
- 透明度归一化上下限仍为 `0.55` 到 `1`，原 `settingsStore` 导出的常量和函数路径保持可用。

**验证证据**：
- `npm exec vitest run src/store/settingsStore.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### P1 可维护性：chatStore 初始状态 helper 拆分

**状态**：已修复

**关联提交**：`refactor: extract chat initial state`

**覆盖范围**：
- 新增 `src/store/chatInitialState.ts`，集中维护 `createInitialChatState()` 和 `createClearedMessagesPatch()`。
- `chatStore.ts` 从约 324 行收敛到约 290 行，继续保留事件监听、stream delta 处理、turn action、thread action、工具审批和输入框文件桥接编排。
- 清空消息仍只清理消息、流式内容、当前 turn、用量、上下文和错误，不清理会话列表、运行中线程、pending tool 或 composer 文件。

**业务链路保护**：
- 初始状态字段和值与原内联对象保持一致，数组/对象由函数每次新建，避免共享可变引用。
- `clearMessages` 的 patch 字段与原行为一致，不影响 thread 列表和后台运行态。
- Agent 事件处理、stream delta 跨轮清理、会话切换/新建/删除/移动和工具审批 IPC 调用未改。

**验证证据**：
- `npm exec vitest run src/store/chatStore.test.ts src/store/chatThreadRuntimeState.test.ts src/store/chatTurnState.test.ts src/store/threadActions.test.ts src/store/agentEventHandler.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### P1 可维护性：Office COM action 脚本模板拆分

**状态**：已修复

**关联提交**：`refactor: extract office com action scripts`

**覆盖范围**：
- 新增 `electron/agent/tools/implementations/office/officeComActionScripts.ts`，集中维护 Excel、Word、PowerPoint 三类 COM action PowerShell 模板。
- 新增 `officeComActionScriptHelpers.ts`，集中维护参数解析、默认输出路径、颜色转换、PPT 删除页码解析和变更描述。
- `officeComActionBridge.ts` 从约 469 行降至约 67 行，保留支持判断、执行、JSON 解析和结果归一化。

**业务链路保护**：
- `executePowerShell(script, 120000)`、`safeJsonParse`、`doneResult/failedResult`、unsupported 分支和 changes 透传保持不变。
- Excel 图表、Word 目录、PPT 快照、PPT 删除页的脚本关键片段由原测试继续覆盖。

**验证证据**：
- `npm exec vitest run electron/agent/tools/implementations/office/officeComActionBridge.test.ts electron/agent/tools/implementations/office/officeComPowerShell.test.ts`
- `npm run typecheck`

### P1 可维护性：StateRuntime goals/memories 表级 CRUD 收敛

**状态**：已修复

**关联提交**：`refactor: extract state runtime memory tables`

**覆盖范围**：
- 新增 `electron/agent/memory/stateRuntimeGoals.ts`，集中维护 goals upsert/get。
- 新增 `electron/agent/memory/stateRuntimeMemories.ts`，集中维护短期记忆、长期记忆、归档、筛选分页和 memory pipeline cursor。
- `StateRuntimeStore` 从约 501 行降至约 378 行，保留初始化、关闭、迁移、事务、线程快照/运行时和派生索引回填编排。

**业务链路保护**：
- goals payload JSON、long-term memory metadata/citations JSON、状态归档、limit/offset 归一化和 pipeline cursor 默认值保持不变。
- 跨库事务仍由 `StateRuntimeStore.transaction()` 统一管理，helper 只执行传入数据库连接上的同步 SQL。

**验证证据**：
- `npm exec vitest run electron/agent/memory/stateRuntimeStore.test.ts`
- `npm run typecheck`

### P1 可维护性：StateRuntime 工具执行日志链路收敛

**状态**：已修复

**关联提交**：`refactor: extract state runtime tool logs`

**覆盖范围**：
- 新增 `electron/agent/memory/stateRuntimeToolLogs.ts`，集中维护 `tool_execution_logs` 写入、limit 归一化和行映射。
- `StateRuntimeStore` 保留 `appendToolExecutionLog` / `listToolExecutionLogs` 公开 API、连接生命周期和跨库事务编排。
- `stateRuntimeStore.ts` 继续收敛到约 501 行，logs.db 的 rollout 与 tool log SQL 细节已下沉到同层 helper。

**业务链路保护**：
- `durationMs` 非负取整、`metadata_json` 序列化、按 `id ASC` 查询和默认/最大 limit 保持不变。
- 跨库事务失败时，工具日志仍随 logs.db 一起回滚。

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

### 2026-07-06 — M1：侧边栏搜索样式拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: split sidebar search styles`

**覆盖范围**：
- 新增 `src/styles/sidebar-search.css`，集中侧边栏搜索框、搜索弹层、结果列表、空状态和搜索弹层动画。
- `global.css` 在 `sidebar.css` 后导入搜索样式，保持侧边栏基础布局先加载。
- `sidebar.css` 从 1100 行降至 891 行；新搜索样式文件为 210 行。

**业务链路保护**：
- 不改 `SidebarSearchPalette.tsx` 的 DOM、className、搜索 tab、结果点击、空状态和弹层关闭逻辑。
- 不改搜索框尺寸、弹层 z-index、玻璃背景、暗色覆盖和动画表现，只按搜索交互职责移动样式。
- `sidebar.css` 仍超 CSS 500 行上限，后续继续按文件夹分组、意图入口、footer 等自然边界拆分。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：侧边栏文件夹样式拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: split sidebar folder styles`

**覆盖范围**：
- 新增 `src/styles/sidebar-folder.css`，集中 `FolderSection.tsx` 使用的文件夹分组、文件项、文件夹内会话缩进和文件夹操作按钮样式。
- `global.css` 在 `sidebar.css` / `sidebar-search.css` 后导入文件夹样式，保持侧边栏基础布局先加载。
- `sidebar.css` 从 891 行降至 739 行；新文件夹样式文件为 155 行。

**业务链路保护**：
- 不改 `FolderSection.tsx` 的 DOM、className、文件夹展开、文件列表切换、文件加入会话、右键菜单和新建文件夹会话逻辑。
- 右键菜单分隔线、disabled 和箭头仍保留在 `sidebar.css`，因为它们同时服务 `ThreadContextMenu` 与 `FileContextMenu`，不是文件夹专属样式。
- `sidebar.css` 仍超 CSS 500 行上限，后续继续按 footer、意图入口、线程列表等自然边界拆分。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：侧边栏底部样式拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: split sidebar footer styles`

**覆盖范围**：
- 新增 `src/styles/sidebar-footer.css`，集中 `SidebarFooter.tsx` 使用的 Office 连接状态、重连按钮、设置按钮、设置菜单和暗色覆盖样式。
- `global.css` 在侧边栏基础/搜索/文件夹样式后导入 footer 样式，保持基础布局先加载。
- `sidebar.css` 从 739 行降至 563 行；新 footer 样式文件为 203 行。

**业务链路保护**：
- 不改 `SidebarFooter.tsx` 的 DOM、className、连接状态、重连按钮、设置菜单打开和设置 section 跳转逻辑。
- 保留 `.sidebar-nav-btn span` 的 `translateY(1px)` 对齐修复，避免设置按钮图标和文字再次错位。
- `sidebar.css` 仍超 CSS 500 行上限，后续再按意图入口或线程列表等自然边界拆分收口。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：侧边栏会话与右键菜单样式拆分

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: split sidebar thread styles`

**覆盖范围**：
- 新增 `src/styles/sidebar-thread.css`，集中普通会话条目、文件夹内会话共享基础样式、状态点、线程右键菜单和文件右键菜单样式。
- `global.css` 在侧边栏基础、搜索、文件夹和 footer 样式后导入线程样式，保持共享基础样式可被 `sidebar-folder.css` 的文件夹增量规则配合使用。
- `sidebar.css` 从 563 行降至 391 行，低于 CSS 500 行上限；新线程样式文件为 174 行。

**业务链路保护**：
- 不改 `FolderSection.tsx`、`ThreadContextMenu.tsx`、`FileContextMenu.tsx` 的 DOM、className、菜单打开/确认/移动到文件夹和删除逻辑。
- 文件夹内会话的增量缩进仍保留在 `sidebar-folder.css`，共享的会话基础样式进入 `sidebar-thread.css`。
- 至此 `sidebar.css` 已从 CSS 超标清单关闭。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`UsageStats` 样式拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: split usage stats styles`

**覆盖范围**：
- 新增 `src/styles/usage-stats.css`，集中 `UsageStats.tsx` 使用的 `usage-*` 页面、指标卡、热力图、柱状图、tooltip、模型分布和响应式样式。
- `global.css` 在 `settings.css` 与 `open-source-settings.css` 后导入使用统计样式，保持设置页基础布局先加载。
- `settings.css` 从 1444 行降至 949 行；新使用统计样式文件为 494 行，低于 CSS 500 行上限。

**业务链路保护**：
- 不改 `UsageStats.tsx` 的 DOM、className、数据加载、范围切换、刷新、tooltip 和图表计算逻辑。
- 仅迁移 `usage-*` 及对应 980px 响应式规则，设置页通用表单和其他设置模块样式仍留在 `settings.css`。
- `settings.css` 仍超 CSS 500 行上限，后续继续按 `ExecPolicySettings`、`KnowledgeSettings` 等自然边界拆分。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`KnowledgeSettings` 样式拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: split knowledge settings styles`

**覆盖范围**：
- 新增 `src/styles/knowledge-settings.css`，集中 `KnowledgeSettings.tsx` 使用的启用开关行、来源列表、空状态、成功/错误提示、type badge、操作按钮和删除按钮样式。
- `global.css` 在设置页基础样式和使用统计样式后导入知识库样式。
- `settings.css` 从 949 行降至 686 行；新知识库样式文件为 260 行。

**业务链路保护**：
- 不改 `KnowledgeSettings.tsx` 的 DOM、className、上传/索引/删除/重建索引和开关逻辑。
- `settings-empty-state`、`settings-success-banner`、`settings-error-banner`、`type-badge` 当前仅由知识库设置使用，因此随知识库样式迁移。
- 该块中的 danger/success fallback 字面量改为直接使用既有主题语义变量，避免新文件继续携带硬编码颜色。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：设置页布局样式拆分

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: split settings layout styles`

**覆盖范围**：
- 新增 `src/styles/settings-layout.css`，集中设置页外壳、侧栏、主视图、账号 profile、旧版 `.settings-page` 兼容布局、双列网格、profile 卡片和编辑器空态样式。
- `global.css` 在 `settings.css` 前导入布局样式，让设置页框架先加载，表单/模块样式后加载。
- `settings.css` 从 686 行降至 333 行，低于 CSS 500 行上限；新布局样式文件为 369 行。

**业务链路保护**：
- 不改 `SettingsPage.tsx`、`ProviderCard.tsx` 或各设置页组件的 DOM、className、导航、折叠侧栏和 profile 展示逻辑。
- 旧版 `.settings-page` / `.settings-grid` / `.profile-card` 兼容样式随布局文件迁移，避免遗留页面入口丢样式。
- 至此审查报告中的 4 个 CSS 超标文件已全部关闭。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — P1：文档级弹层关闭逻辑复用

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: reuse document dismiss hook`

**覆盖范围**：
- 扩展 `src/hooks/useDocumentDismiss.ts`，支持 `boundaryRefs`，允许点击指定 ref 内部时不触发关闭。
- `SidebarSearchPalette.tsx` 复用该 hook 替换手写 `mousedown` + `keydown[Escape]` document listener。
- `ModelQuickSwitch.tsx` 复用该 hook 替换手写外部点击监听，保留原有“不按 Escape 关闭”的行为。
- `useDocumentDismiss.test.ts` 增加 ref 边界回归测试，覆盖弹层内部点击不关闭。

**业务链路保护**：
- 不改搜索面板、模型切换下拉的 DOM、样式、选中模型逻辑、搜索结果逻辑和定位逻辑。
- 搜索面板仍在打开时清空 query、聚焦输入框，并支持 Escape 关闭。
- 模型快速切换仍只在外部 mousedown 时关闭，不额外引入 Escape 行为。

**验证证据**：
- `npx vitest run src/hooks/useDocumentDismiss.test.ts`
- `npm run typecheck`

### 2026-07-06 — P1：文件大小格式化入口收敛

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: use shared file size formatter`

**覆盖范围**：
- `ChatPage.tsx` 直接从 `src/utils/fileSize.ts` 导入 `formatFileSize`。
- 移除 `chatHelpers.tsx` 中仅转发共享 formatter 的 `formatFileSize()` 导出，保留 `chatHelpers` 的聊天分组、任务元数据和小组件职责。
- 附件预览仍通过 `formatAttachmentSize()` 调用同一个共享 formatter，以保留空值隐藏和紧凑格式差异。

**业务链路保护**：
- 文件夹弹层中的大小展示仍使用默认一位小数格式，例如 `120.0 KB`。
- 图片附件预览的紧凑大小格式不变，例如 `120 KB`。
- 不改 ChatPage 文件夹弹层 DOM、CSS class、文件打开/添加逻辑。

**验证证据**：
- `npx vitest run src/utils/fileSize.test.ts src/utils/attachmentPreview.test.ts`
- `npm run typecheck`

### 2026-07-06 — M1：`OCRTaskComposerPanel` 文件流程 helper 拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract ocr task file helpers`

**覆盖范围**：
- 新增 `src/components/task/ocrTaskFileHelpers.ts`，集中 OCR 文件类型判断、发票文件名识别、临时文件落盘、Sheet/range 解析和写入目标解析。
- `OCRTaskComposerPanel.tsx` 复用 helper，保留面板状态、识别动作、字段选择、写入动作和 JSX 渲染。
- 新增 `ocrTaskFileHelpers.test.ts` 覆盖 Sheet 引用解析、OCR 文件格式白名单和发票文件名识别。

**业务链路保护**：
- 不改 OCR 面板 DOM、CSS class、拖拽/粘贴/选择文件、发票模式自动切换、识别接口和 Excel 写入接口。
- `ipcApi.ocr.recognize` 与 `ipcApi.excel.writeRange` 仍留在面板流程中，避免把用户动作隐藏到工具函数里。
- 本次不为追求行数继续拆 JSX 子组件；组件从 517 行降至约 455 行，仍略超 React 组件规范，但职责边界更清晰。

**验证证据**：
- `npx vitest run src/components/task/ocrTaskFileHelpers.test.ts src/components/task/OCRTaskComposerPanel.test.ts`
- `npm run typecheck`

### 2026-07-06 — M1：`toolExecutor` sandbox 策略预评估拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract tool sandbox policy`

**覆盖范围**：
- 新增 `electron/agent/core/agentLoop/toolSandboxPolicy.ts`，集中 shell 工具的 sandbox 预评估、prompt 理由汇总、forbidden/prompt 标志和异常兜底审批。
- `toolExecutor.ts` 只消费 `evaluateToolSandboxPolicy()` 的结果，主流程继续负责创建/更新 `TurnItem`、审批、执行工具和写执行日志。
- 新增 `toolSandboxPolicy.test.ts` 覆盖非 shell 跳过评估、prompt 理由汇总和评估异常强制审批。

**业务链路保护**：
- forbidden 仍在审批和 spawn 前直接拒绝，并写入 blocked 执行日志。
- prompt 仍覆盖 `permissionMode` 与 always-allowed 工具，强制进入用户审批。
- shell 工具执行时仍把 `sandboxEvaluation` 透传到 executor context；非 shell 工具仍不产生 sandbox context。
- `toolExecutor.ts` 从 488 行降至约 458 行；不拆审批/执行主状态机，避免改变工具事件顺序。

**验证证据**：
- `npx vitest run electron/agent/core/agentLoop/toolSandboxPolicy.test.ts electron/agent/core/agentLoop/toolExecutor.test.ts`
- `npm run typecheck`

### 2026-07-06 — M1：`chatStore` turn 启动状态构造收敛

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract chat turn start state`

**覆盖范围**：
- 新增 `src/store/chatTurnState.ts`，集中构造 turn 开始时的流式状态重置、`activeClientId` 绑定、`turnStatus` 标记和 stopped thread 清理。
- 新增 `chatTurnState.test.ts`，覆盖有 active thread 时移除 stopped 标记、无 active thread 时保留原 stopped map 引用，以及调用方额外清理字段。
- `sendMessage` 仍额外清理 `compactionNotice`，`resumeFromInterruption` 仍额外清理 `lastInterruptContext`，差异通过 `extraPatch` 明确传入。

**业务链路保护**：
- 不改 `ensureAgentThread()`、`ipcApi.agent.startTurn()`、`ipcApi.agent.continueTurn()`、`enqueueTurn()` 和 `loadThreads()` 的调用顺序。
- 不改变 streaming 中入队、未绑定 threadId 时拒绝入队、发送失败/恢复失败回滚状态和 active thread 绑定逻辑。
- 仅抽取纯状态 patch，避免把发送/恢复 action 拆成更碎的异步流程模块。

**验证证据**：
- `npx vitest run src/store/chatTurnState.test.ts src/store/chatStore.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`EditProviderDialog` 保存 patch 计算抽取

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract edit provider patch builder`

**覆盖范围**：
- 新增 `src/components/settings/editProviderPatch.ts`，集中计算编辑供应商弹窗保存时的 `Partial<AiProviderConfig>`。
- 新增 `editProviderPatch.test.ts`，覆盖草稿无变化时不提交 patch、字段变化时只提交差异，以及旧 `ModelConfig.reasoningOptions` 不再写回。
- `EditProviderDialog.tsx` 仅改为把当前草稿传给 helper，组件继续负责表单状态、模型切换、连接测试和弹窗渲染。

**业务链路保护**：
- 不改 `useTestConnection`、`ipcApi.ai.listModels()`、`ModelConfigList`、`ReasoningModeSelect` 和聚合平台模型选择链路。
- 不改变 `onSave(patch)` 的调用时机；仍在用户点击保存时一次性提交。
- `EditProviderDialog.tsx` 从约 411 行降至约 406 行；本项目标是固定保存语义，不为行数继续拆碎 JSX。

**验证证据**：
- `npx vitest run src/components/settings/editProviderPatch.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`Sidebar` 排序与派生列表 helper 抽取

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract sidebar derived lists`

**覆盖范围**：
- 将 `compareSidebarText()`、`sortSidebarItems()`、文件夹分组、未分组会话和空状态判断移动到 `src/utils/sidebarHelpers.ts`。
- `SidebarExpanded` 继续暴露原 `SidebarSortMode` / `SidebarGroupedFolder` 类型名，但来源改为 utils，保持父组件 props 形状不变。
- `sidebarHelpers.test.ts` 新增排序模式、文件夹会话分组、文件列表透传和空状态派生测试。

**业务链路保护**：
- 不改文件夹加载、批量 `folder:listFilesBatch`、会话切换、右键菜单、固定文件和 Office 连接状态逻辑。
- `Sidebar.tsx` 仍统一编排 store 状态、IPC 动作和 UI 回调；本次只移动纯排序/分组数据派生。
- `Sidebar.tsx` 从约 520 行降至约 493 行；仍高于规范上限，后续继续按职责拆，不为行数硬拆状态机。

**验证证据**：
- `npx vitest run src/utils/sidebarHelpers.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`chatStore` 线程运行态重算 helper 抽取

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract chat thread runtime reconciliation`

**覆盖范围**：
- 新增 `src/store/chatThreadRuntimeState.ts`，集中 `loadThreads` 后根据线程元数据重算 `runningThreadIds` 的逻辑。
- 新增 `chatThreadRuntimeState.test.ts`，覆盖 activeTurnId / in_progress 元数据恢复运行态、已完成线程移除运行态，以及用户停止线程不被旧元数据复活。
- `chatStore.ts` 的 `loadThreads` 只负责调用 action、写入 threads 和接入运行态 helper。

**业务链路保护**：
- 不改 `thread:list` / `loadThreadsAction()` 调用，不改变 `stoppedThreadIds` 对 stale in-progress 的屏蔽语义。
- 保留旧行为：本次线程列表未返回的既有 running id 不在 helper 中主动删除，避免误伤后台运行态。
- `chatStore.ts` 从约 467 行降至约 463 行；本次重点是固定运行态合并规则，而不是继续拆 action。

**验证证据**：
- `npx vitest run src/store/chatThreadRuntimeState.test.ts src/store/chatStore.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`toolExecutor` 工具名解析 helper 抽取

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract tool name resolution`

**覆盖范围**：
- 新增 `electron/agent/core/agentLoop/toolNameResolution.ts`，集中点号/下划线工具名、OpenAI compatible sanitize/desanitize 候选解析。
- 新增 `toolNameResolution.test.ts`，覆盖原 executor 名优先、`ocr_parseDocument` 解析到 `ocr.parseDocument`、无匹配返回 `null`。
- `toolExecutor.ts` 改为复用 `resolveExecutableToolName()`，保留 `desanitizeToolName()` 用于缺 executor 时的 fallback 展示。

**业务链路保护**：
- 不改 `processToolCalls` 的 sandbox forbidden/prompt、审批、执行、日志和 `TurnItem` 事件顺序。
- 不改 `executeTool()` 的未知工具错误和 executor context 透传行为；既有 `toolExecutor.test.ts` 覆盖 shell/OCR alias、falsy result 和日志链路。
- `toolExecutor.ts` 从约 426 行降至约 409 行；仍略超 400 行，后续只在明确职责边界处继续拆。

**验证证据**：
- `npx vitest run electron/agent/core/agentLoop/toolNameResolution.test.ts electron/agent/core/agentLoop/toolExecutor.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`toolExecutor` 工具结果 item 构造收敛

**状态**：✅ 已修复
**关联提交**：本节所在提交 `refactor: extract tool result item builder`

**覆盖范围**：
- 新增 `electron/agent/core/agentLoop/toolResultItems.ts`，统一生成 `tool_result` 的 id、timestamp、result、isError 和工具名字段。
- 新增 `toolResultItems.test.ts`，覆盖成功结果和错误结果的稳定字段输出。
- `toolExecutor.ts` 四处重复 `tool_result` 对象字面量改为调用 `createToolResultItem()`。

**业务链路保护**：
- 不改 sandbox、审批、执行、日志和 callbacks 事件顺序；仅替换结果 item 的结构创建入口。
- `toolExecutor.test.ts` 继续覆盖审批取消、fallback tool_call、falsy 成功结果、日志、shell alias 和 OCR alias。
- `toolExecutor.ts` 从约 409 行降至约 398 行，低于 400 行规范上限。

**验证证据**：
- `npx vitest run electron/agent/core/agentLoop/toolResultItems.test.ts electron/agent/core/agentLoop/toolExecutor.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：供应商 reasoning 自动适配提示复用

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: reuse provider reasoning hint`

**覆盖范围**：
- 新增 `src/components/settings/providerReasoningHint.ts`，集中新增/编辑供应商弹窗的 reasoning 自动适配提示文案。
- 新增 `providerReasoningHint.test.ts`，覆盖中英文提示和既有 ` / ` 选项分隔格式。
- `AddProviderDialog.tsx` 与 `EditProviderDialog.tsx` 改为复用 `buildReasoningAutoHint()`。

**业务链路保护**：
- 不改 `resolveReasoningOptionValues()`、`defaultReasoningModeForOptions()`、`coerceReasoningMode()` 和 `ReasoningModeSelect` 的适配逻辑。
- 不改新增/编辑供应商保存、连接测试、模型列表和聚合平台配置链路。
- `EditProviderDialog.tsx` 从约 406 行降至约 404 行；本次重点是移除重复文案，不为行数硬拆 UI。

**验证证据**：
- `npx vitest run src/components/settings/providerReasoningHint.test.ts src/utils/reasoningSupport.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`StateRuntimeStore` row mapper 与查询 helper 下沉

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extend state runtime mappers`

**覆盖范围**：
- 将 `mapThreadSnapshot()`、`mapToolExecutionLog()`、`buildRolloutFtsQuery()` 和 `clampMemoryListOffset()` 移入 `stateRuntimeMappers.ts`。
- 新增 `stateRuntimeMappers.test.ts`，覆盖线程快照 JSON 字段解析、工具日志 metadata JSON、FTS term quote 和 offset clamp。
- `stateRuntimeStore.ts` 改为复用 mapper/helper，保留 SQLite statement、事务和写入编排。

**业务链路保护**：
- 不改四库初始化、迁移、WAL 配置、恢复备份、事务 begin/commit/rollback 和 `runLogsWrite()` 行为。
- 不改 thread snapshot、rollout event、tool log、goal、memory、long-term memory 的 SQL 语句和 API 签名。
- `stateRuntimeStore.ts` 从约 626 行降至约 584 行；剩余体量主要来自仓储 API 和事务边界，后续继续按职责拆。

**验证证据**：
- `npx vitest run electron/agent/memory/stateRuntimeMappers.test.ts electron/agent/memory/stateRuntimeStore.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`ChatPage` 文件夹徽标组件拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract chat folder badge`

**覆盖范围**：
- 新增 `src/components/chat/ChatFolderBadge.tsx`，承接聊天页顶部当前文件夹 badge、文件数量展示和文件弹层渲染。
- `ChatPage.tsx` 改为通过语义化回调传入打开/关闭、附件注入和隐藏 badge 行为，主页面继续保留会话、任务面板和文件夹文件加载编排。
- `ChatPage.tsx` 从 309 行降至 276 行；新增组件约 80 行，边界为 UI 展示与用户点击，不拆散业务状态。

**业务链路保护**：
- 不改 `ipcApi.folder.listFiles(currentFolderId)` 的加载时机、取消保护和失败兜底。
- 不改 `addFilesToComposer([{ filePath, fileName, fileType: "document", size }])` 的附件注入结构。
- 不改点击文件后的关闭弹层、隐藏 badge 行为，也不改现有 CSS className，避免影响视觉和交互。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`SidebarExpanded` 分组标题与排序弹层拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract sidebar expanded sections`

**覆盖范围**：
- 新增 `src/components/sidebar/SidebarSectionHeader.tsx`，复用项目区和会话区的展开、排序、右侧动作按钮结构。
- 新增 `src/components/sidebar/SidebarSortMenu.tsx`，承接排序弹层渲染、当前模式高亮和排序选项列表。
- `SidebarExpanded.tsx` 从 319 行降至 295 行，继续保留展开态侧边栏的整体布局、上下文菜单和 footer 编排。

**业务链路保护**：
- 不改 `Sidebar.tsx` 中的展开状态、排序状态、新建会话、新建文件夹和排序菜单定位逻辑。
- 不改现有 CSS className：`sidebar-section-header`、`sidebar-section-toggle`、`sidebar-section-add`、`sidebar-sort-menu`、`sidebar-sort-menu-item`。
- 排序菜单仍按 `projects/conversations` 分区选择对应 active mode，回调签名保持 `onSelectSortMode(section, mode)` 不变。

**验证证据**：
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`OCRTaskComposerPanel` 展示块拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract ocr task sections`

**覆盖范围**：
- 新增 `src/components/task/OCRModeSelector.tsx`，承接通用 OCR / 发票识别模式选择 UI。
- 新增 `src/components/task/OCRFileUploadSection.tsx`，承接拖拽/点击上传区和已选文件列表。
- 新增 `src/components/task/OCRRecognizeButton.tsx`，承接识别按钮禁用态与文案展示。
- 新增 `src/components/task/OCRResultSection.tsx`，承接错误提示、字段勾选、目标单元格、预览表格和写入按钮展示。
- `OCRTaskComposerPanel.tsx` 从 455 行降至 297 行；剩余体量主要是 OCR 状态、识别调用、字段选择和 Excel 写入编排，不拆散识别/写入状态机。

**业务链路保护**：
- 不改 `ipcApi.ocr.recognize(effectiveOcrMode, await resolveOcrFilePaths(files))` 调用和模式判断。
- 不改 `resolveOcrFilePaths()` 临时文件落盘、文件类型过滤、发票文件推断和最多文件数量规则。
- 不改 `buildOcrWriteValues()`、`resolveWriteTarget()`、`ipcApi.excel.writeRange()` 写入链路，也保留 `OCRTaskComposerPanel` 对 helper 的 re-export，兼容现有测试和调用方。

**验证证据**：
- `npm exec vitest run src/components/task/OCRTaskComposerPanel.test.ts src/components/task/ocrTaskFileHelpers.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：新增/编辑供应商弹窗表单块复用

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: share provider dialog fields`

**覆盖范围**：
- 新增 `src/components/settings/ProviderDialogFields.tsx`，复用供应商名称、API 格式、API 地址、API Key、上下文窗口、模型字段、测试按钮和测试结果展示。
- 新增 `src/components/settings/ProviderDialogFrame.tsx` 与 `ProviderDialogActions.tsx`，复用弹窗外壳和新增/编辑底部操作区。
- 新增 `src/components/settings/AddProviderTemplateSelect.tsx`，承接新增供应商模板分组下拉。
- `AddProviderDialog.tsx` 从 379 行降至 295 行，`EditProviderDialog.tsx` 从 393 行降至 299 行。

**业务链路保护**：
- 不改 `providerDraftFromTemplate()`、`createEmptyProviderDraft()`、`buildProviderConfigFromDraft()` 和 `buildEditProviderPatch()` 的配置构造链路。
- 不改 `useTestConnection()`、`ipcApi.ai.testConnection()`、`ipcApi.ai.listModels()` 的调用条件和回调入口。
- 保留新增弹窗“添加 / 测试并添加”、编辑弹窗“测试连接 / 保存”、聚合模型列表和 reasoning 自动适配语义。

**验证证据**：
- `npm exec vitest run src/components/settings/addProviderDraft.test.ts src/components/settings/editProviderPatch.test.ts src/components/settings/ProviderModelSelector.test.ts src/components/settings/providerReasoningHint.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`Sidebar` 状态 hook 拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract sidebar state hooks`

**覆盖范围**：
- 新增 `src/hooks/useSidebarFolderFiles.ts`，承接固定文件夹添加、展开加载、搜索批量文件加载、文件右键菜单、回收站/打开/复制/显示/置顶操作。
- 新增 `useSidebarResize.ts`、`useSidebarViewedThreads.ts`、`useSidebarSortMenu.ts`、`useSidebarThreadContextMenu.ts`、`useSidebarThreadCreation.ts`、`useSidebarSettingsNavigation.ts`、`useSidebarSectionToggles.ts`，分别拆出侧边栏宽度拖拽、会话已读状态、排序菜单、线程右键菜单、新建会话、设置导航和 section 展开状态。
- 移除 `Sidebar` 已不消费的 `activeIntent/onIntentClick` props，`App.tsx` 只继续把功能意图传给 `ChatPage`。
- `Sidebar.tsx` 从 545 行降至 300 行；前端 TSX 超 300 行组件阶段性清零。

**业务链路保护**：
- 不改 `FolderSection` / `SidebarExpanded` / `SidebarCollapsed` / `SidebarSearchPalette` 的 props 语义和 CSS className。
- 文件操作仍走原 IPC：`dialog.openFolder`、`folder.listFiles/listFilesBatch`、`file.trashFile/openFile/copyPath/revealInExplorer`，附件注入仍走 `addFilesToComposer`。
- 会话动作仍走 `createNewThread`、`switchThread`、`deleteThread`、`moveThreadToFolder`；设置入口仍优先调用 `onOpenSettingsSection`，否则回退到 `onNavigate("settings")`。

**验证证据**：
- `npm exec vitest run src/components/sidebar/SidebarThreadItem.test.ts src/utils/sidebarHelpers.test.ts`
- `npm run typecheck`
- `npm run build`
- `git diff --check`

### 2026-07-06 — M1：`chatStore` turn 动作拆分

**状态**：✅ 阶段性已修复
**关联提交**：本节所在提交 `refactor: extract chat turn actions`

**覆盖范围**：
- 新增 `src/store/chatTurnActions.ts`，承接 `sendMessage`、`resumeFromInterruption`、`interruptTurn` 的 IPC 调用、turn start patch、resume 检查和中断状态更新。
- `chatStore.ts` 保留 Zustand state shape、事件入口、stream delta、线程管理、工具审批和 composer 文件桥接；主文件从 516 行降至 367 行。

**业务链路保护**：
- 对外 store action 名称、参数和返回 Promise 行为不变。
- 不改 `buildTurnStartPatch()`、`ipcApi.agent.startTurn/enqueueTurn/continueTurn/interrupt`、`thread.resume` 和 `loadThreads()` 调用语义。
- 不改运行中追加队列、恢复中断、停止线程状态、`stoppedThreadIds`、`runningThreadIds` 和中断提示文案。

**验证证据**：
- `npm exec vitest run src/store/chatStore.test.ts src/store/chatTurnState.test.ts src/store/threadActions.test.ts src/store/chatThreadRuntimeState.test.ts src/store/agentEventHandler.test.ts`
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

**历史问题**：`ipcSchemas.ts` 中已定义的 schema 曾未被对应 handler 使用；当前已复核为接入状态：

| Schema | 对应通道 | 状态 |
|--------|----------|------|
| `SetAlwaysOnTopInput` | `window:setAlwaysOnTop` | ✅ 已调用 `validateInput` |
| `SettingsGetInput/SettingsSetInput` | `settings:get/set` | ✅ 已调用 `validateInput` |
| `ExcelReadRangeInput/ExcelWriteRangeInput` | `excel:readRange/writeRange` | ✅ 已调用 `validateInput` |
| `ThreadIdInput` | `thread:load/delete/resume` | ✅ 已调用 `validateInput` |
| `ThreadNewInput` | `thread:new` | ✅ 已调用 `validateInput` |
| `FilePathInput` | `file:trashFile/openFile/copyPath/revealInExplorer` | ✅ 已调用 `validateInput` |
| `FolderPathInput` | `folder:listFiles` | ✅ 已调用 `validateInput` |
| `ToolConfirmInput/ToolCancelInput` | `tool:confirm/cancel` | ✅ 已调用 `validateInput` |
| `StatsGetSummaryInput` | `stats:getSummary` | ✅ 已调用 `validateInput` |

**当前状态**：schema 定义与 handler 调用已对齐，避免继续给维护者造成"定义了但未生效"的误判。

**建议**：本节点已完成；后续新增 IPC 时继续保持 schema 定义、handler `validateInput`、preload/type wrapper 三处同步。

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

**状态**：✅ 主入口与业务模块已阶段性收敛；当前仅剩 3 个集中协议/桥接边界略超，后续只在有自然职责边界时继续拆分。

| 文件 | 行数 | 上限 | 超出 | 拆分建议 |
|------|------|------|------|----------|
| `src/electronApi.d.ts` | **483** | 400 | +83 | 集中 preload/API 类型声明，暂不为行数硬拆 |
| `electron/agent/shared/types.ts` | **455** | 400 | +55 | Agent 协议集中类型，暂不为行数硬拆 |
| `electron/agent/tools/implementations/excel/excelComBridge.ts` | **409** | 400 | +9 | Excel COM 连接状态机，后续有自然边界时再收敛 |

**完整超标清单**：当前共 3 个 TS/TSX 文件 + 0 个 CSS 文件（详见附录 A）

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

**状态**：✅ 阶段性已关闭（2026-07-06，节选高风险文件已由 sandbox、AgentLoop、知识库和 Python 执行器测试补强覆盖）

**高风险无测试文件**（节选）：

| 文件 | 风险说明 |
|------|----------|
| `security/sandbox/parseCommand.ts` | ✅ 已有 sandbox.test 覆盖，并补充解析失败边界 |
| `security/sandbox/execPolicy.ts` | ✅ 已有 sandbox.test 覆盖，并补充组合命令/大小写边界 |
| `security/sandbox/defaultRules.ts` | ✅ 已有 sandbox.test 覆盖默认安全规则 |
| `core/agentLoop/turnRunner.ts` | ✅ 已补独立测试，覆盖轮次调度结构 |
| `core/agentLoop/threadStateManager.ts` | ✅ 已补独立测试，覆盖线程状态机 |
| `knowledge/textChunker.ts` | ✅ 已补独立测试，覆盖文本分块 |
| `knowledge/retriever.ts` | ✅ 已补独立测试，覆盖知识检索 |
| `tools/executors/pythonExecutor.ts` | ✅ 已补独立测试，覆盖 Python 执行器 |

**建议**：本节点名的高风险节选已完成阶段性覆盖；后续不为历史数量做机械补测，改为在触碰模块、修复缺陷或新增行为时同步补契约测试。

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

## 六、修复优先级完成情况

> 原始优先级表保留为闭环账本。后续继续修复时，优先选择“风险真实、职责边界自然、能补验证”的项目；集中类型声明、协议聚合和桥接状态机允许少量超出行数上限。

### 第一周（P0 安全 + 最严重行数超标）

| 优先级 | 任务 | 当前状态 | 备注 |
|--------|------|----------|------|
| 1 | 补全 IPC Zod schema 校验（S1 + S2） | ✅ 已完成 | 高风险 IPC 已接入 `validateInput`，后续新增通道按同一规则维护 |
| 2 | 文件操作路径穿越防护（S5） | ✅ 已完成 | 文件读取/打开/回收站/枚举已接入路径授权 |
| 3 | `agentLoop.ts` 1276 行拆分 | ✅ 已关闭 | 主循环按生命周期、工具轮、压缩、线程运行态等职责拆分 |
| 4 | `ipcHandlers.ts` 1057 行拆分 | ✅ 已关闭 | OCR/AI/沙箱/文件 IPC 已拆到子 handler |

### 第二周（P1 性能 + 测试）

| 优先级 | 任务 | 当前状态 | 备注 |
|--------|------|----------|------|
| 5 | Sidebar N+1 查询改批量 | ✅ 已完成 | 新增批量文件夹列表读取，减少多文件夹搜索时 IPC 次数 |
| 6 | useExcelConnection setTimeout 清理 | ✅ 已完成 | timeout 统一登记并在卸载时清理 |
| 7 | ChatPage useEffect Promise 取消 | ✅ 已完成 | 文件夹加载加入 cancelled guard |
| 8 | 沙箱核心模块补测试（parseCommand/execPolicy） | ✅ 阶段性完成 | 高风险边界已有 sandbox 测试覆盖 |
| 9 | electron/ 13 处 console.warn → logger | ✅ 已完成 | 主进程警告/错误统一走结构化 logger |

### 第三周（P1 规范 + 剩余行数）

| 优先级 | 任务 | 当前状态 | 备注 |
|--------|------|----------|------|
| 10 | @MOCK_INTERFACE 标记补全 | ✅ 阶段性完成 | 关键 mock 已标记；后续触碰测试 mock 时继续补 |
| 11 | electronApi.d.ts 类型同步 | ✅ 已完成 | preload、类型声明与前端 IPC wrapper 已对齐 |
| 12 | CHANGELOG 更新 | ✅ 持续执行 | 每项修复同步追加记录 |
| 13 | 剩余超标文件拆分（Sidebar/settingsStore 等） | ✅ 阶段性关闭 | 主入口已收敛；剩余集中类型/桥接文件不为行数硬拆 |
| 14 | 引入 ESLint + Prettier 工具链 | ⏳ 待决策 | 当前已有 `typecheck`、`vitest`、`build`、`git diff --check` 门禁；引入格式化链路需单独评估规则和历史代码扰动 |

---

## 附录 A：超过行数限制的文件完整列表

### TS/TSX 文件（3 个）

| 行数 | 文件 |
|------|------|
| 483 | `src/electronApi.d.ts` |
| 455 | `electron/agent/shared/types.ts` |
| 409 | `electron/agent/tools/implementations/excel/excelComBridge.ts` |

### CSS 文件（0 个）

| 行数 | 文件 |
|------|------|

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

> **审查闭环说明。** 原始高风险优先级已基本处理；后续继续以真实风险和自然职责边界为准推进，避免为单纯行数把聚合良好的协议、类型或桥接状态机拆碎。

## 追加修复记录：2026-07-06

### P1 可维护性：settingsStore Provider 状态收敛

**状态**：已修复

**关联提交**：`refactor: extract settings provider state`

**覆盖范围**：
- 新增 `src/store/settingsProviderState.ts`，集中维护 provider template 匹配、reasoning 配置归一化、provider 增删改状态构造、`isConfigured` 计算和 provider id 生成。
- `settingsStore.ts` 保留加载、保存、通用设置、透明度、动态数组、知识库开关和 pinned folders 编排，provider action 改为调用纯 helper。
- 修复 provider 配置状态误判：新增非激活供应商不会把当前设置误标为已配置；更新激活供应商后立即按新配置重算。

**业务链路保护**：
- `loadSettings` 的迁移和 normalize 链路仍调用同一归一化逻辑。
- `savePartial` 持久化 key、`aiProviders` / `activeProvider` 写入时机和原有 action 名称保持不变。
- 新增 helper 仅处理无副作用状态推导，不引入新的 IPC、UI 或模型配置入口。

**验证证据**：
- `npm run typecheck`
- `npm exec vitest run src/store/settingsStore.test.ts`

### P1 可维护性：StateRuntime rollout 日志链路收敛

**状态**：已修复

**关联提交**：`refactor: extract state runtime rollout events`

**覆盖范围**：
- 新增 `electron/agent/memory/stateRuntimeRolloutEvents.ts`，集中维护 rollout events 写入、列表查询和 FTS 搜索。
- `StateRuntimeStore` 保留连接生命周期、四库事务、公开 API、工具日志、目标和记忆读写编排，不继续硬拆 class。
- `stateRuntimeStore.ts` 从 584 行收敛到约 520 行，rollout SQL 细节从主 store 下沉到日志领域 helper。

**业务链路保护**：
- `appendRolloutItems` 仍写入 `rollout_events` 与 `rollout_events_fts`，继续通过 `runLogsWrite` 复用现有事务语义。
- `listRolloutEvents`、`searchRolloutMatches` 的返回结构、排序和 FTS query 规则保持不变。
- `backfillDerivedIndexes` 暂不迁移，避免触碰旧损坏 JSON 回填容错逻辑。

**验证证据**：
- `npm exec vitest run electron/agent/memory/stateRuntimeStore.test.ts`
- `npm run typecheck`
