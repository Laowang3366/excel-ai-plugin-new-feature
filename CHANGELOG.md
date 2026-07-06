# Changelog

## Unreleased

### 2026-07-06 工程质量、RAG/知识库与桌面体验

- **版本基线更新到 `0.1.61`**：以 `desktop/package.json` 为准，后续安装包和验收记录均按该版本线继续递增。
- **当前测试源基线**：静态统计为 147 个测试文件、734 个 `it/test` 用例；具体通过情况以每项修复记录中的 `vitest`、`typecheck` 和 `build` 验证命令为准。
- **代码审查整改闭环**：新增 `docs/code-review-report-2026-07-05.md`，持续记录 IPC 安全、路径授权、大文件拆分、性能优化、类型同步和测试补强的修复状态、验证证据与关联提交。
- **CI 依赖安全审计**：现有 GitHub Actions 在 `npm ci` 后执行 `npm audit --audit-level=high`，阻断 high/critical 级别漏洞进入主线。
- **IPC 安全与路径授权加固**：补齐高风险 IPC 的 Zod 校验，统一文件/文件夹授权边界，保护拖拽、粘贴、OCR、文件预览、回收站、资源管理器打开等链路。
- **AgentLoop 与主进程模块继续拆分**：将 Agent loop 状态、队列、单轮执行、工具执行日志、IPC OCR/AI/沙箱/文件处理等职责拆到独立模块，保留原入口兼容。
- **侧边栏组件接口收束**：将 `FolderSection` 的文件夹动作、线程动作和文件菜单能力合并为语义化 API 对象，减少 prop drilling。
- **侧边栏拖拽性能优化**：侧边栏宽度拖拽改为 requestAnimationFrame 节流，每帧最多触发一次宽度状态更新。
- **侧边栏回调稳定化**：将搜索、展开、设置菜单和上下文菜单等 JSX 内联回调提取为 `useCallback`，减少子组件无效更新。
- **主进程日志统一**：electron 主进程警告/错误输出改走结构化 logger，保留控制台和文件日志双输出。
- **shell 执行器测试补强**：覆盖安全策略拒绝、prompt 决策、cwd 重定向、缺失工作目录和超时强杀路径。
- **shell/sandbox 执行限制命名化**：将 shell 默认超时、watchdog 宽限和 stdout/stderr 截断上限提取为命名常量，并约束工具 schema 文案与执行器默认值一致。
- **推理与文本清理工具测试补强**：为 reasoning 选项适配和 reasoning 文本清理补充纯函数测试，并修复不支持的 per-model `reasoningMode` 旧值残留问题。
- **sandbox 安全边界测试补强**：补充解析失败 prompt、组合命令危险子命令拦截和 Windows 大小写不敏感匹配等回归测试。
- **关键 mock 接口标记补齐**：为 chatStore/ipcApi/MinerU/eventForwarder/OpenAI Responses/shellExecutor 等关键测试 mock 增加 `@MOCK_INTERFACE` 说明，降低真实接口与测试替身漂移风险。
- **运行时状态库事务一致性补强**：`StateRuntimeStore.transaction()` 现在同时覆盖 `state/logs/goals/memories` 四个运行时库，并补充跨库回滚测试，避免事务失败后日志、目标或记忆写入残留。
- **知识库/RAG 稳定性增强**：SQLite 知识库行转换、Excel/OpenXML 解析、session rollout 解析等逻辑拆分为可测试 helper，并修复 provider/model/dim 变更后的 runtime 与索引一致性问题。
- **StateRuntime row 映射收敛**：将线程快照、工具执行日志、rollout FTS 查询和长期记忆分页 offset 处理下沉到 `stateRuntimeMappers.ts`，让运行时仓储主文件更聚焦数据库读写与事务。
- **PPT OpenXML 内容页 helper 拆分**：将 `advancedPresentation.ts` 中的新增幻灯片参数归一化、内容页 XML 和空关系部件生成抽到 `presentationSlideContent.ts`，补充 helper 单元测试，同时保留 PPT 包关系维护、页码编号和写包主流程在原模块。
- **Composer 附件解析 helper 拆分**：将 `useComposer.ts` 中的拖拽/粘贴附件类型判断、本地路径解析和临时文件落盘抽到 `composerAttachmentFiles.ts`，拖拽与粘贴共用同一解析链路，hook 主文件降至 336 行。
- **流式与推理显示优化**：调整流式增量缓冲、工具事件顺序、上下文用量估算、思考正文滚动和展开详情时间线，减少长 reasoning 输出导致的 UI 堵塞。
- **OpenAI Responses 流式尾包补齐**：Responses SSE 解析在流结束时会处理未带尾随空行的最后事件，避免正文、usage 或 done 事件被漏掉。
- **OpenAI 兼容工具名解析拆分**：将 Chat Completions 工具名点号/下划线清洗与还原逻辑抽到 `openaiToolNames.ts`，保留 `openaiCompatibleClient.ts` re-export 兼容入口。
- **工具执行结果保真修复**：工具成功返回 `false`、`0`、空字符串等假值时不再被误当成无结果，执行日志摘要也能安全处理 `undefined`。
- **消息与工具详情样式令牌化**：工作详情、工具调用和设置按钮的剩余硬编码灰色/暗色背景改为复用主题变量，减少暗色主题漂移风险。
- **常规设置职责收敛**：抽出通用开关/滑块字段与双语文案模块，`GeneralSettings.tsx` 降至 318 行，同时保留透明度、动态数组、压缩、OCR token 和数据目录设置链路不变。
- **OpenAI Responses 解析职责拆分**：将 Responses 文本补发、工具调用状态、usage 归一化和 reasoning effort 映射抽到独立 helper，客户端主文件降至 400 行以内。
- **开源项目信息样式拆分**：将 `OpenSourceSettings` 专属样式从 `settings.css` 移到 `open-source-settings.css`，保持全局样式入口和设置页面加载顺序不变，并复用暗色主题变量。
- **浮动任务入口样式拆分**：将功能悬浮入口与任务浮窗样式从 `chat.css` 移到 `floating-task-panel.css`，保留聊天页面、消息列表和 Office 预览样式在原文件中。
- **Office 预览侧栏样式拆分**：将 `office-preview-*` 样式从 `chat.css` 移到 `office-preview-panel.css`，`chat.css` 降至 CSS 上限以内。
- **Composer 样式职责拆分**：将输入框附件 chip 与操作控件样式拆到 `composer-attachments.css` / `composer-controls.css`，`composer.css` 降至 CSS 上限以内。
- **侧边栏搜索样式拆分**：将侧边栏搜索框与搜索弹层样式移到 `sidebar-search.css`，保留侧边栏主体布局在 `sidebar.css`。
- **侧边栏文件夹样式拆分**：将固定文件夹分组、文件项和文件夹内会话缩进样式移到 `sidebar-folder.css`，继续收束 `sidebar.css`。
- **侧边栏底部样式拆分**：将 Office 连接状态、设置按钮和设置菜单样式移到 `sidebar-footer.css`，保留设置按钮对齐修复。
- **侧边栏会话样式拆分**：将会话条目、状态点和线程/文件右键菜单样式移到 `sidebar-thread.css`，`sidebar.css` 降至 CSS 上限以内。
- **使用统计样式拆分**：将 `UsageStats` 页面样式移到 `usage-stats.css`，显著降低 `settings.css` 体量。
- **使用统计数据 helper 拆分**：将 `UsageStats.tsx` 中的文案、数字格式化、stats 行转换和时间范围聚合抽到 `usageStatsData.ts`，补充纯函数测试，页面组件降至 252 行。
- **新增供应商草稿 helper 拆分**：将 `AddProviderDialog.tsx` 中的模板草稿初始化、空白重置和 `AiProviderConfig` 构造抽到 `addProviderDraft.ts`，补充纯函数测试，保留表单状态、测试连接和渲染链路不变。
- **供应商模型选择器复用**：将新增/编辑供应商弹窗中的聚合模型下拉、预设模型下拉和自定义模型输入分支收敛到 `ProviderModelSelector.tsx`，两个弹窗继续保留各自保存和测试连接语义。
- **Markdown 表格解析共享化**：将 MinerU 和本地文档解析共用的 Markdown 表格提取逻辑移到 `electron/shared/markdownTables.ts`，解除 agent executor 对 `main-modules/mineruOcr` 的反向依赖。
- **线程 IPC wrapper 域拆分**：将 `ipcApi.ts` 中的 `thread` / `threadGraph` wrapper 抽到 `ipcThreadApi.ts`，补充转发与 fallback 测试，主 IPC wrapper 降至 354 行。
- **知识库设置样式拆分**：将 `KnowledgeSettings` 来源列表、提示 banner 和操作按钮样式移到 `knowledge-settings.css`，并复用主题语义色。
- **知识库设置文案与格式化拆分**：将 `KnowledgeSettings.tsx` 内的双语文案、来源统计、文件夹索引结果汇总、类型标签和索引时间格式化抽到 `knowledgeSettingsText.ts`，设置组件降至 255 行。
- **简单任务面板拆分**：将 `ChatPage.tsx` 内联的清洗/图表 simple task 表单抽到 `SimpleTaskComposerPanel.tsx`，并用测试锁定 range/task payload 拼装规则；主页面保留任务调度和会话发送职责。
- **设置页布局样式拆分**：将设置页外壳、侧栏、主视图、旧版兼容布局和 profile 卡片样式移到 `settings-layout.css`，`settings.css` 降至 CSS 上限以内。
- **文档级弹层关闭逻辑复用**：扩展 `useDocumentDismiss` 支持 ref 边界，侧边栏搜索面板和模型快速切换下拉统一复用该 hook，减少重复 document listener。
- **文件大小格式化入口收敛**：`ChatPage` 直接使用共享 `utils/fileSize`，移除 `chatHelpers` 中的薄包装导出，避免格式规则分散。
- **OCR 面板流程 helper 拆分**：将 OCR 文件类型判断、临时文件落盘和 Excel 写入目标解析移到 `ocrTaskFileHelpers.ts`，面板组件保留状态、识别、写入和渲染职责。
- **工具执行 sandbox 策略 helper 拆分**：将 `toolExecutor` 中的 shell 命令安全策略预评估抽到 `toolSandboxPolicy.ts`，保留 forbidden 永拒、prompt 强制审批和执行上下文透传语义。
- **工具审批策略 helper 拆分**：将 `toolExecutor` 中的权限模式判断、always-allowed 工具集合和审批回调兜底抽到 `toolApproval.ts`，旧入口继续 re-export，主执行器降至 355 行。
- **工具名解析 helper 拆分**：将 `toolExecutor` 中点号/下划线工具别名解析抽到 `toolNameResolution.ts`，补充 canonical executor 优先级和 OCR alias 回归测试。
- **工具结果 item 构造收敛**：将 `toolExecutor` 中重复的 `tool_result` 结构生成抽到 `toolResultItems.ts`，统一 id/timestamp/result/isError 字段并让主执行器降至 400 行以内。
- **chatStore turn 启动状态收敛**：将发送消息与恢复中断共用的流式状态重置、`activeClientId` 绑定和 stopped thread 清理抽到 `chatTurnState.ts` 纯 helper，保留两个 action 各自的 compaction/interrupt 清理差异。
- **chatStore 线程运行态重算收敛**：将 `loadThreads` 中根据线程元数据重算 `runningThreadIds` 的逻辑抽到 `chatThreadRuntimeState.ts`，并用测试锁定用户停止线程不被旧 in-progress 元数据复活的行为。
- **编辑供应商保存 patch 可测试化**：将 `EditProviderDialog` 中的保存差异计算和旧 `reasoningOptions` 清理抽到 `editProviderPatch.ts`，避免弹窗组件内混杂配置迁移细节。
- **供应商 reasoning 提示文案复用**：将新增/编辑供应商弹窗中的自动适配提示抽到 `providerReasoningHint.ts`，统一中英文提示和选项标签格式。
- **侧边栏排序与分组派生收敛**：将会话排序、文件夹分组和空状态判断移到 `sidebarHelpers.ts`，`Sidebar.tsx` 保留状态编排与事件回调，排序规则补充单元测试。
- **Office/WPS 操作体验优化**：增加紧凑模式、透明度、动态数组环境支持设置，修复窗口恢复、按钮对齐、公式写入策略和 WPS 动态数组提示约束。
- **OCR 与附件体验完善**：接入 MinerU 付费 → 免费 → 本地降级链路，补充图片附件预览、拖拽上传、发票字段静默提取和 OCR 工具化能力。
- **侧边栏与文件夹性能优化**：搜索打开时批量加载 pinned folder 文件列表，避免每个文件夹一次 IPC；新增旧 preload 环境的 wrapper fallback。

### 2026-06-30 文档、Office 操作与安装包

- **PPT 删除页专用能力**：新增 `office.action.apply` 的 `presentation/deleteSlides` 操作，支持 `params.slides`、`params.from/to` 和 `target: "slide:2-6"`。
  - Open XML 优先删除 `presentation.xml` 中的 slide 引用、`presentation.xml.rels` 关系和对应 `ppt/slides/slideN.xml` 部件。
  - COM 兜底按倒序删除幻灯片，避免索引变化导致删错页。
  - 工具注册表和系统提示词已要求模型优先使用统一 Office action，避免回退到临场 PowerShell/Python 脚本。
- **版本线已继续推进**：该阶段安装包曾推进到 `0.1.2`；当前仓库版本以顶部 `0.1.61` 基线为准。
- **文档同步**：更新根 README、Agent 架构文档、工具层文档、记忆层文档和开发规范，统一描述 Office 三件套、Open XML 优先、SQLite StateRuntime、当前测试基线和打包注意事项。
- **验证基线**：该阶段 `npm run typecheck` 通过；测试规模后续已增长，当前统计见顶部 2026-07-06 基线。

### Added

- **Office / WPS 多宿主选择弹窗**：当 Microsoft Excel 和 WPS 表格同时运行时，弹出 `HostSelectionDialog` 让用户选择目标程序
  - `excelComBridge.detectExcelProcess()` 同时检测 `EXCEL` + `et` 进程，返回 `availableHosts[]`
  - 新增 `selectHost()` 方法，用户选择后保存 `_selectedHost` 并尝试 COM 连接
  - 新增 IPC 通道 `excel:selectHost`，preload + ipcApi 桥接
  - 新增 `HostSelectionDialog` 组件（dialog.css 风格），`useExcelConnection` 新增 host 选择状态
  - 连接后验证 `$app.Name`（需确认实际返回 vs 预期），避免 ProgID 注册冲突
- **Agent 命令执行沙箱化（阶段 0+1+4+5）**：参考 Codex `execpolicy` + `process-hardening`，对 `shell.execute` 工具加入命令语义级过滤、cwd 白名单、env 清洗与审计
  - 新增 `electron/agent/sandbox/`：`parseCommand.ts`（命令切 token）、`execPolicy.ts`（前缀规则策略引擎，最严 `forbidden > prompt > allow`）、`defaultRules.ts`（默认规则挡 `rm -rf /` / `Remove-Item -Recurse -Force` / `format` / `Stop-Computer` / `reg delete` / `iex` 等）、`audit.ts`（JSONL 审计 `sandbox-logs/YYYY/MM/DD/audit-*.jsonl`）、`index.ts`（一站式 `evaluateCommand` + `runShellSpawn` + `killProcessTree`）
  - `shell.execute` 改走 `-EncodedCommand`（Base64 UTF-16LE），消除明文 `-Command` 拼接注入；子进程环境只保留 USERNAME / USERPROFILE / PATH 等白名单变量
  - 任意 `permissionMode`（含 `confirm_all` + 始终允许）下，`forbidden` 仍直接拒绝；`prompt` 覆盖 `alwaysAllowedTools` 强制审批并展示理由
  - 越界 cwd 自动重定向到临时目录并记审计；超时 `taskkill /T /F /PID` 强杀整棵进程树
  - 新增 IPC：`sandbox:getConfig` / `sandbox:setUserRules` / `sandbox:setWritableRoots`，可在安全策略设置页增删用户规则与可写根
  - 新增设置页 `ExecPolicySettings.tsx`（SettingsPage 新增「安全策略」tab）
  - `ToolConfirmDialog` 命中 prompt 规则时展示安全策略理由（`sandboxJustification`）
  - 14 个单元测试覆盖破坏性命令拒绝、prompt 命令、cwd 重定向
- **系统提示词按场景分支化**：修正"任何场景都无脑触发 `workbook.inspect` / `selection.get`"问题
  - 「行动优先原则」追加第 6 条「探测要按需」
  - 工作流程拆分为 A 纯提问 / B 读/写表格 / C 通用系统操作 / D 附件触发 四类，仅在"用户要读/写表格"时才做工作簿检查
- **公式生成场景提示词重写**（`scenarioFormula()`）：对齐公式助手 5 个结构化字段
  - 样例分完整/部分两种处理：完整样例 value 必须完全一致；部分样例不纠结"是否完整"
  - #SPILL! 主动清理：探测溢出方向非空单元格 → 清空 → 告知用户清理内容
  - 表头按需拼接：检测锚点是否含表头，含则不拼、不含则在首行拼语义标题
  - 嵌套约束松绑：取消"嵌套不超过3层"死规矩，动态数组优先减少辅助列
  - 动数组"否"态用"写入形态是否依赖溢出"作为判定标准，不再按函数名禁用 XLOOKUP
  - 与 §可维护性 / §安全底线 / §结果验证的冲突经全局审计调和：覆盖公式单元格安全底线对公式助手场景放宽

### Changed

- `electron/main-modules/settingsManager.ts` — `DEFAULT_SETTINGS` 增加 `sandboxUserRules` / `sandboxExtraWritableRoots`
- `electron/main.ts` — `app.whenReady` 中调 `applySandboxConfig()` 热更沙箱单例
- `electron/preload.ts` — 暴露 `electronAPI.sandbox.*`
- `src/electronApi.d.ts` + `src/services/ipcApi.ts` — 新增 `SandboxPrefixRule` / `SandboxConfig` 与桥接
- `src/store/agentEventHandler.ts` + `src/store/chatStore.ts` — 透传 `sandboxJustification`
- `src/i18n.ts` — 中英文补 `assistant.sandboxJustification` 文案
- `src/styles/tool-confirm.css` + `src/styles/settings.css` — 沙箱 notice 与 ExecPolicy 样式

### Refactored

- **模块化拆分：大型文件高内聚低耦合重构**
  - `ChatPage.tsx`：从 1,378 行降至 300 行，拆分为 `chatHelpers.tsx`、`FloatingTaskPanel.tsx`、`AssistantGroupBlock.tsx`、`ChatMessageList.tsx`、`ComposerArea.tsx`、`useComposer.ts`、`useTaskDrafts.ts` 共 7 个模块
  - `Sidebar.tsx`：从 786 行降至 436 行，拆分为 `sidebarHelpers.ts`、`useExcelConnection.ts`、`FolderSection.tsx`、`ThreadContextMenu.tsx` 共 4 个模块
  - `chatStore.ts`：从 645 行降至 418 行，拆分为 `agentEventHandler.ts`、`threadActions.ts` 共 2 个模块
  - `main.ts`：从 947 行降至 124 行，拆分为 `settingsManager.ts`、`windowManager.ts`、`ipcHandlers.ts`、`eventForwarder.ts` 共 4 个模块
  - `agentLoop.ts`：从 1,101 行降至 1,054 行，拆分为 `imageAttachmentResolver.ts`、`toolExecution.ts`、`compactionManager.ts` 共 3 个模块
  - `agentLoop.ts` 深度拆分（Week 4）：进一步拆为 `agentLoop/` 目录 6 个子模块（agentLoop、streamCollector、toolExecutor、compactionManager、buildStreamParams、index）
  - `toolRegistry.ts`：从 1,228 行拆为 `toolRegistry/` 目录 4 个子模块（interfaces、definitions、excelFunctions、executors）
  - `excelBridge.ts`：拆为 `excelBridge/` 目录 5 个子模块（comBridge、vbaBridge、scriptBridge、uiBridge、index）
  - `ProviderCard` / `AddProviderDialog` 重复逻辑提取：新增 `useTestConnection` hook、`ReasoningModeSelect` 组件、`ModelConfigList` 组件（含 useRef 替代 document.querySelector），消除 4 块重复代码及行为分叉风险
  - 所有模块遵循 ≤ 400 行/文件、单文件单职责、React 组件 ≤ 300 行的编码规范
- **ComposerArea Props 精简（Week 4）**：28 个 props → 7 个
  - 将 useComposer hook 返回值整体作为 `composer` prop 传入
- **settingsStore 增量持久化（Week 4）**：9 个 settings IPC → 仅写变更 key
  - 新增 `savePartial(keys, get)` 函数，13 个 setter action 全部改为按 key 增量写入
- **IPC 依赖注入抽象层（Week 4）**：新建 `src/services/ipcApi.ts`
  - 封装 71 处 `window.electronAPI` 直接调用，已迁移 5 个核心模块
  - `createMockIpcApi()` 辅助函数，测试时一键注入 mock

### Added

- **项目开发规范文档（Week 4）**：`docs/development-standards.md`
  - 基于 4 周 18 项代码审查修复经验，按六大审查方向编写
  - 涵盖：大文件模块拆分、Props 与接口精简、持久化优化、测试基础设施、IPC 依赖注入、Bug 修复与防御
  - 包含 11 项提交前自查清单
- **单元测试基础设施（Week 4）**：vitest + @vitest/coverage-v8
  - 早期基线为 35 tests；当前基线已扩展到 74 个测试文件、420 个测试
- **IPC Zod Schema 输入验证（Week 2）**：所有 IPC 通道新增 zod schema 运行时校验
- **结构化日志系统（Week 2）**：`electron/agent/logger.ts`，支持 JSON 格式 + 日志级别 + 文件输出
- **OCR Mock 标记（Week 4）**：`OCRTaskComposerPanel.tsx` 新增 `@MOCK_INTERFACE` 标记
- **文件夹上下文感知**：对话界面显示文件夹信息 + AI 模型感知文件夹及文件列表
  - 对话区标题新增文件夹标签（badge），显示当前会话所属文件夹名称
  - 输入框下方新增文件夹上下文行，显示文件夹名和 Excel 文件数量
  - `systemPrompt.ts` 新增 `FolderFileItem` 接口和 `appendFolderContext()` 函数
  - `agentLoop.ts` 在 `runAgentLoop` 中动态注入文件夹路径和文件列表到 `systemPrompt`
  - 新增 i18n 键 `folderFileCount`（中英双语）
  - 新增 `.chat-folder-badge` / `.composer-folder-context` 等 CSS 样式
- **文件夹线程组织**：会话可归属到固定文件夹，侧边栏采用合并布局统一展示
  - `ThreadMetadata` / `SessionMeta` 新增 `folderId` 字段，关联 `pinnedFolders[].path`
  - 文件夹分组可折叠，内含该文件夹的会话和文件列表
  - 文件夹头部 [+] 按钮直接在该文件夹中新建会话
  - 右键会话菜单新增"移动到文件夹"子菜单（支持移入/移出文件夹）
  - 修复文件夹创建会话后 `activeThreadId` 未设置的问题
  - 新增 `thread:updateMetadata` IPC 通道
  - `AgentLoop` 新增 `pendingFolderId` 机制
  - 侧边栏合并布局：`sidebar-content` 统一滚动区域
  - 新增 i18n 键：`newThreadInFolder` / `moveToFolder` / `noFolder` / `addFolderFirst`（中英双语）
- **新建会话过渡动画**：点击 [+] 新建会话时显示旋转图标 + 弹性缩放动画
- **即时文件夹信息显示**：从文件夹创建新会话后，导航栏和输入框立即显示文件夹信息
- **侧边栏 UI 优化**：顶部操作区 + 排序 + 文件夹交互增强
- **CSS 模块化拆分**：`global.css` 从 4,827 行单文件拆分为 23 个独立模块文件

### Changed

- Removed the Web Add-in, backend API, admin console, Docker/Nginx deployment stack, server-side tests, and backend-oriented documentation.
- Repositioned the repository as a desktop-only Electron application.

### Fixed

- **修复 7 项 TypeScript 编译错误**
  - `agentLoop.ts`：`TurnItem` 联合类型上访问 `.message` 属性需先窄化为 `ErrorItem`
  - `compactionManager.ts`：`TurnItem[] | null` 不可赋值给 `TurnItem[]` 返回类型
  - `modelSettingsI18n.ts`：`.ts` 文件包含 JSX 语法，重命名为 `.tsx`
  - `agentLoop.test.ts`：缺少 `beforeEach` 导入（从 vitest 补充）
  - `settingsManager.ts`：`ElectronStore<具体类型>` 不可赋值给 `Store<Record<string, unknown>>`
  - `ipcSchemas.ts`：Zod v4 `z.record()` 需要 key + value 两个参数
  - `ipcApi.ts`：`setAlwaysOnTop` 返回 `void/undefined` 而非签名要求的 `boolean`
- **修复添加供应商弹窗切回"自定义"时残留模板配置**：`AddProviderDialog.tsx` 中 `handleSelectTemplate` 的空模板分支现在重置所有表单字段（name、apiFormat、baseUrl、model、contextWindowSize、reasoningMode、modelConfigs）
- **修复编辑聚合供应商时删除当前模型不同步清空 provider.model**：`ProviderCard.tsx` 中删除模型时检查是否为当前选中模型，若是则同步清空 `model` 字段
- **修复对话信息显示顺序（Week 1）**：改为按 API 真实事件顺序渲染，而非按类型排序
  - `agentLoop.ts`：流式阶段不再立即发出 `tool_call` 的 `item_started` 事件；流结束后按 `reasoning → assistant_message → tool_call` 顺序依次发出
  - `chatStore.ts`：`item_completed` handler 简化为按事件到达顺序 `push`
  - `ChatPage.tsx`：流式渲染顺序调整；`sortItemsByRound` 组内排序调整为 JSONL 历史恢复兜底
  - 折叠"本轮工作时长"时，`phase=commentary` 的阶段性正文也一起折叠
- **修复 API 400 错误（Week 1）**：assistant 消息含 `tool_calls` 但缺少对应 `tool` result 消息导致请求被拒
  - `turnItemsToChatMessages()` 重写为三遍处理：识别孤立 `tool_call`、跳过孤立项、清理空壳 assistant 消息
  - `buildRequestMessages()` 新增双重安全校验
  - 孤立 `tool_call` 采用跳过策略而非合成假 result
- **修复 Electron 类型检查错误（Week 1）**：Agent loop 和 Excel bridge 的 TypeScript 类型问题
- **修复 AgentLoopConfig 缺少 reasoningMode（Week 1）**：从 `aiConfig.reasoningMode` 读取，支持顶层覆盖
- **修复 UsageStats N+1 查询（Week 2）**：单次查询替代循环，减少 IPC 调用次数
- **修复 vitest 测试 bug（Week 4）**：`toolExecutor.ts` import 路径修正 + 工具名修正（`read_range` → `range.read`）
