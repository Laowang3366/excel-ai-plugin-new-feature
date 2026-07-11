# 项目文件架构与调用链路图

更新时间：2026-07-07
范围：`desktop/` Electron 桌面端、Renderer 前端、Agent 运行时、Office/WPS 桥接、知识库、OCR、记忆与设置链路。

## 1. 总体分层

```mermaid
flowchart TB
  subgraph Renderer["Renderer 前端: desktop/src"]
    App["App.tsx"]
    Chat["components/ChatPage.tsx + chat/* + task/*"]
    Sidebar["components/Sidebar.tsx + sidebar/*"]
    Settings["components/SettingsPage.tsx + settings/*"]
    Store["store/chatStore.ts + settingsStore.ts"]
    IpcApi["services/ipcApi.ts + ipcOfficeApi.ts + ipcThreadApi.ts + ipcKnowledgeApi.ts"]
  end

  subgraph Bridge["隔离桥: desktop/electron/preload.ts"]
    ElectronAPI["window.electronAPI"]
  end

  subgraph Main["Electron 主进程: desktop/electron"]
    MainEntry["main.ts"]
    WindowMgr["main-modules/windowManager.ts"]
    SettingsMgr["main-modules/settingsManager.ts"]
    IpcHandlers["main-modules/ipcHandlers.ts + ipc*Handlers.ts"]
    AgentIpc["agent/interaction/ipcAgentHandlers.ts"]
    EventForwarder["agent/interaction/eventForwarder.ts"]
  end

  subgraph Agent["Agent 运行时: desktop/electron/agent"]
    Runtime["runtime/agentRuntime.ts"]
    Loop["core/agentLoop/*"]
    Providers["providers/*"]
    Prompts["prompts/*"]
    Tools["tools/registry + tools/executors + tools/implementations"]
    Knowledge["knowledge/*"]
    Memory["memory/*"]
    Security["security/sandbox/*"]
  end

  subgraph External["外部能力与本地资源"]
    Office["Excel / WPS / Word / PowerPoint COM"]
    OpenXml["Open XML 文件包: xlsx/docx/pptx"]
    Data["用户数据目录: sessions / SQLite / logs"]
    AI["AI Provider: OpenAI 兼容 / Responses / Anthropic / 聚合平台"]
    MinerU["MinerU OCR / 免费降级 / 本地兜底"]
    FS["本地文件系统"]
  end

  App --> Chat
  App --> Sidebar
  App --> Settings
  Chat --> Store
  Sidebar --> Store
  Settings --> Store
  Store --> IpcApi
  IpcApi --> ElectronAPI
  ElectronAPI --> IpcHandlers
  ElectronAPI --> AgentIpc
  MainEntry --> Runtime
  MainEntry --> IpcHandlers
  MainEntry --> WindowMgr
  IpcHandlers --> SettingsMgr
  IpcHandlers --> AgentIpc
  AgentIpc --> Loop
  Runtime --> Loop
  Runtime --> Tools
  Runtime --> Knowledge
  Runtime --> Memory
  Loop --> Providers
  Loop --> Prompts
  Loop --> Tools
  Tools --> Security
  Tools --> Office
  Tools --> OpenXml
  Tools --> FS
  Knowledge --> Data
  Memory --> Data
  Providers --> AI
  IpcHandlers --> MinerU
  Tools --> MinerU
  EventForwarder --> ElectronAPI
  ElectronAPI --> Store
```

核心边界：

- `desktop/src` 只通过 `services/ipcApi.ts` 访问主进程，避免组件直接散落调用 `window.electronAPI`。
- `desktop/electron/preload.ts` 是唯一暴露到 Renderer 的隔离桥。
- `main-modules/ipcHandlers.ts` 负责通用 IPC、Office 当前窗口 IPC、设置、文件、OCR 子 handler 汇总。
- `agent/interaction/ipcAgentHandlers.ts` 负责 Agent、Thread、Knowledge、Stats、工具审批相关 IPC。
- `agent/runtime/agentRuntime.ts` 负责把 AI 配置、Office bridge、知识库、记忆、工具执行器、AgentLoop 装配到一起。
- `agent/core/agentLoop/*` 只做对话轮次、模型流、工具调用、压缩、中断/恢复编排，不直接依赖 COM 具体实现。
- `agent/tools/registry` 是模型可见工具定义；`agent/tools/executors` 是工具路由；`agent/tools/implementations` 才触达 COM、OpenXML、PowerShell、Python 等具体实现。

## 2. 启动与运行时装配链路

```mermaid
sequenceDiagram
  participant Electron as electron/main.ts
  participant Settings as main-modules/settingsManager.ts
  participant Runtime as agent/runtime/agentRuntime.ts
  participant Bridges as runtime/bridgeRegistry.ts
  participant Knowledge as runtime/knowledgeRuntime.ts
  participant Tools as tools/executors/createToolExecutors.ts
  participant Loop as core/agentLoop/AgentLoop
  participant IPC as main-modules/ipcHandlers.ts
  participant Window as main-modules/windowManager.ts

  Electron->>Settings: getSessionStoreInstance()
  Electron->>Runtime: getOrCreateAgentRuntime(deps)
  Runtime->>Bridges: getOrCreateOfficeBridges()
  Runtime->>Knowledge: initializeKnowledgeRuntime(aiConfig, dataPath)
  Runtime->>Tools: createToolExecutors(bridges, retriever, memoryStore, deps)
  Runtime->>Loop: new AgentLoop(config)
  Runtime-->>Electron: AgentRuntime
  Electron->>IPC: registerIpcHandlers()
  IPC->>IPC: registerAgentIpcHandlers()
  IPC->>IPC: registerToolApprovalHandlers()
  Electron->>Window: createWindow()
```

启动关键文件：

| 文件 | 入口/职责 | 主要调用 | 主要产物 |
| --- | --- | --- | --- |
| `desktop/electron/main.ts` | Electron 生命周期入口 | `getOrCreateAgentRuntime`、`registerIpcHandlers`、`createWindow`、退出时 flush/close | 主窗口、Agent runtime、IPC 注册 |
| `desktop/electron/main-modules/settingsManager.ts` | electron-store、数据目录、Session/StateRuntime 生命周期 | `getActiveAIConfig`、`getSessionStoreInstance`、`getStateRuntimeStoreInstance` | 设置、数据路径、SQLite/JSONL 存储实例 |
| `desktop/electron/main-modules/windowManager.ts` | BrowserWindow、托盘、普通/紧凑模式、透明度/主题 | Electron `BrowserWindow`、设置值 | 主窗口状态、托盘行为 |
| `desktop/electron/main-modules/ipcHandlers.ts` | 通用 IPC 注册入口 | `registerAgentIpcHandlers`、Office bridge refs、settings/file/ocr/sandbox 子 handler | 所有主进程 IPC handler |
| `desktop/electron/agent/runtime/agentRuntime.ts` | Agent 依赖装配 | `bridgeRegistry`、`knowledgeRuntime`、`createToolExecutors`、`AgentLoop` | `AgentRuntime`、`AgentLoopManager` |
| `desktop/electron/agent/runtime/bridgeRegistry.ts` | Office/WPS bridge 单例 | Excel/Word/PPT/OpenXML bridge 构造 | `OfficeBridgeRegistry` |
| `desktop/electron/agent/runtime/knowledgeRuntime.ts` | RAG runtime 初始化/刷新 | `SqliteStore`、`EmbeddingService`、`KnowledgeIndexer`、`Retriever` | 知识库 store/indexer/retriever |

## 3. Renderer 到主进程 IPC

```mermaid
flowchart LR
  UI["React 组件"]
  Hooks["hooks/*"]
  Stores["Zustand store"]
  Wrappers["services/ipcApi.ts"]
  Preload["electron/preload.ts"]
  IpcMain["ipcMain.handle(...)"]
  Modules["main-modules/* 或 agent/interaction/*"]

  UI --> Hooks
  UI --> Stores
  Hooks --> Wrappers
  Stores --> Wrappers
  Wrappers --> Preload
  Preload --> IpcMain
  IpcMain --> Modules
```

前端连接表：

| 功能模块 | 前端入口 | 状态/服务 | preload 暴露 | 主进程 handler |
| --- | --- | --- | --- | --- |
| 聊天对话 | `components/ChatPage.tsx`、`components/chat/*`、`components/task/*` | `store/chatStore.ts`、`store/chatTurnActions.ts`、`store/chatStreamBuffer.ts` | `electronAPI.agent.*` | `agent/interaction/ipcAgentHandlers.ts` |
| 会话/文件夹 | `components/Sidebar.tsx`、`components/sidebar/*` | `store/threadActions.ts`、`hooks/useSidebar*` | `electronAPI.thread.*`、`threadGraph.*`、`folder.*` | `ipcAgentHandlers.ts`、`ipcFileHandlers.ts` |
| 设置 | `components/SettingsPage.tsx`、`components/settings/*` | `store/settingsStore.ts`、`settingsPersistence.ts`、`settingsProviderState.ts` | `electronAPI.settings.*`、`sandbox.*` | `main-modules/ipcHandlers.ts`、`ipcSandboxHandlers.ts` |
| Excel 当前窗口 | `components/excel/HostSelectionDialog.tsx`、任务面板 | `hooks/useExcelConnection.ts`、`services/ipcOfficeApi.ts` | `electronAPI.excel.*` | `main-modules/ipcHandlers.ts` |
| Word/PPT 当前窗口 | `components/office/OfficePreviewPanel.tsx` | `hooks/useOfficeConnection.ts` | `electronAPI.office.*` | `main-modules/ipcHandlers.ts` |
| OCR 面板 | `components/task/OCRTaskComposerPanel.tsx` + OCR 子组件 | `utils/fileBase64.ts`、`ocrTaskFileHelpers.ts` | `electronAPI.ocr.recognize` | `main-modules/ipcOcrHandlers.ts`、`mineruOcr.ts` |
| 知识库设置页 | `components/settings/KnowledgeSettings.tsx` | `services/ipcKnowledgeApi.ts` | `electronAPI.knowledge.*` | `agent/interaction/ipcAgentHandlers.ts` |
| 使用统计 | `components/settings/UsageStats.tsx` | `usageStatsData.ts` | `electronAPI.stats.getSummary` | `agent/interaction/ipcAgentHandlers.ts` |

## 4. 用户发消息与流式回显链路

```mermaid
sequenceDiagram
  participant User as 用户
  participant Composer as ComposerArea/useComposer
  participant ChatStore as store/chatStore.ts
  participant IpcApi as services/ipcApi.ts
  participant Preload as electron/preload.ts
  participant AgentIPC as agent/interaction/ipcAgentHandlers.ts
  participant Loop as core/agentLoop/*
  participant Provider as providers/*
  participant Forwarder as eventForwarder.ts
  participant UIStore as chatStreamBuffer + agentEventHandler
  participant View as ChatMessageList/AssistantGroupBlock

  User->>Composer: 输入文本/拖拽附件/选择功能模块
  Composer->>ChatStore: sendMessage / enqueueTurn
  ChatStore->>IpcApi: agent.startTurn(input)
  IpcApi->>Preload: window.electronAPI.agent.startTurn
  Preload->>AgentIPC: ipcRenderer.invoke("agent:startTurn")
  AgentIPC->>Loop: runAgentLoop / enqueue / continue
  Loop->>Provider: stream(params)
  Provider-->>Loop: reasoning/content/tool_call 增量
  Loop-->>Forwarder: onStreamDelta / onEvent
  Forwarder-->>Preload: webContents.send("agent:event")
  Preload-->>UIStore: agent.onEvent / onStreamDelta listener
  UIStore->>ChatStore: merge patches
  ChatStore->>View: 消息、思考、工具详情、最终回答渲染
```

关键文件连接：

| 阶段 | 文件 | 被谁调用 | 调用谁 |
| --- | --- | --- | --- |
| 输入组织 | `src/hooks/useComposer.ts`、`src/hooks/composerAttachmentFiles.ts` | `ComposerArea.tsx` | `ipcApi.file.*`、附件 base64/临时文件工具 |
| 功能模块提示词组装 | `src/utils/taskComposerPayloads.ts`、`components/task/*` | `ChatPage.tsx`、`FloatingTaskPanel.tsx` | `chatStore.sendMessage` |
| Turn 启动 | `src/store/chatTurnActions.ts` | `chatStore.ts` | `ipcApi.agent.startTurn` / `continueTurn` / `enqueueTurn` |
| IPC 请求 | `src/services/ipcApi.ts` | Store/Hooks/Components | `window.electronAPI.agent.*` |
| IPC 注册 | `electron/agent/interaction/ipcAgentHandlers.ts` | `main-modules/ipcHandlers.ts` | `AgentLoopManager.getLoopForThread`、SessionStore、Knowledge runtime |
| Agent 编排 | `electron/agent/core/agentLoop/turnRunner.ts`、`turnExecution.ts`、`streamRound.ts`、`toolRound.ts` | `AgentLoop` | Provider、ToolExecutor、Memory/Session |
| 模型请求 | `electron/agent/core/agentLoop/buildStreamParams.ts`、`roundStreamParams.ts` | AgentLoop round | system prompt、history、tool definitions、knowledge/date/memory context |
| 事件转发 | `electron/agent/interaction/eventForwarder.ts` | AgentLoop callbacks | `BrowserWindow.webContents.send("agent:event", ...)` |
| 前端投影 | `src/store/chatStreamBuffer.ts`、`agentEventHandler.ts` | `chatStore.ts` | Zustand state patches |
| UI 渲染 | `components/chat/ChatMessageList.tsx`、`AssistantGroupBlock.tsx`、`ReasoningBubble.tsx`、`ToolCallBubble.tsx` | `ChatPage.tsx` | Markdown/StreamingOutput/ToolConfirmDialog |

## 5. Agent 核心与模型 Provider

```mermaid
flowchart TB
  Loop["core/agentLoop/agentLoop.ts"]
  Runner["turnRunner.ts / turnExecution.ts / turnFlow.ts"]
  Stream["streamRound.ts / streamCollector.ts / streamResultItems.ts"]
  ToolRound["toolRound.ts / toolExecutor.ts"]
  Params["buildStreamParams.ts / roundStreamParams.ts"]
  Compaction["compactionRunner.ts / preTurnCompaction.ts / contextUsage.ts"]
  Prompts["prompts/systemPrompt.ts + sections/*"]
  Definitions["tools/registry/toolDefinitions.ts"]
  Providers["providers/aiClientFactory.ts"]
  OpenAICompat["providers/openaiCompatibleClient.ts"]
  Responses["providers/openaiResponsesClient.ts + openaiResponsesParsing.ts"]
  Anthropic["providers/anthropicClient.ts"]
  Vendor["providers/providerClients.ts"]

  Loop --> Runner
  Runner --> Params
  Runner --> Stream
  Stream --> ToolRound
  ToolRound --> Definitions
  Params --> Prompts
  Params --> Definitions
  Runner --> Compaction
  Params --> Providers
  Providers --> OpenAICompat
  Providers --> Responses
  Providers --> Anthropic
  Providers --> Vendor
```

Agent 核心文件：

| 文件/目录 | 职责 | 上游 | 下游 |
| --- | --- | --- | --- |
| `core/agentLoop/agentLoop.ts` | AgentLoop 门面，承接 start/continue/enqueue/interrupt | `ipcAgentHandlers.ts`、`AgentLoopManager` | `turnRunner.ts`、thread/session helpers |
| `core/agentLoop/turnRunner.ts` | 单轮 turn 编排 | `AgentLoop` | `turnExecution.ts`、`streamRound.ts`、`toolRound.ts` |
| `core/agentLoop/buildStreamParams.ts` | 组装系统提示词、上下文、工具定义、history | `roundStreamParams.ts` | `prompts/*`、`messageBuilder.ts` |
| `core/agentLoop/streamRound.ts` | 调模型流，收集增量 | `turnRunner.ts` | `providers/*`、`streamCollector.ts` |
| `core/agentLoop/toolRound.ts` | 处理模型工具调用 | `streamRound.ts` | `toolExecutor.ts` |
| `core/agentLoop/toolExecutor.ts` | 权限、沙箱策略、工具执行日志、结果封装 | `toolRound.ts` | `tools/executors/*`、`security/sandbox/*` |
| `core/agentLoop/compaction*.ts` | 上下文压缩、token 估算、历史裁剪 | `turnRunner.ts` | `memory/compaction.ts`、Provider |
| `providers/aiClientFactory.ts` | 按配置创建模型客户端 | `buildStreamParams.ts` / runtime | Responses/OpenAI Compatible/Anthropic/厂商子类 |
| `prompts/systemPrompt.ts` | 基础系统提示词 + 动态场景片段入口 | `buildStreamParams.ts` | `prompts/templates/*`、`promptComposer.ts` |

## 6. 工具注册、路由与执行链路

```mermaid
flowchart LR
  Model["模型 tool_call"]
  ToolRound["core/agentLoop/toolRound.ts"]
  ToolExecutor["core/agentLoop/toolExecutor.ts"]
  Registry["tools/registry/*.ts"]
  Executors["tools/executors/*.ts"]
  Contracts["tools/contracts/*.ts"]
  Implementations["tools/implementations/*"]
  Result["tool_result + AgentEvent"]

  Model --> ToolRound
  ToolRound --> ToolExecutor
  ToolExecutor --> Registry
  ToolExecutor --> Executors
  Executors --> Contracts
  Executors --> Implementations
  Implementations --> Result
```

工具层连接表：

| 工具域 | 模型可见定义 | 执行器 | 实现/依赖 | 典型能力 |
| --- | --- | --- | --- | --- |
| Workbook/Range/Formula/Sheet/UI | `tools/registry/workbook.ts`、`range.ts`、`formula.ts`、`sheet.ts`、`ui.ts` | `tools/executors/excelExecutors.ts` | `implementations/excel/*`、`contracts/excel.ts` | 检查工作簿、读写选区、公式验证、工作表操作、宿主选择 |
| Script/VBA | `tools/registry/script.ts` | `excelExecutors.ts` | `excelScriptBridgeCom.ts`、`excelVbaComBridge.ts`、`automation/*` | Excel/WPS 脚本执行、VBA |
| File | `tools/registry/file.ts` | `fileExecutors.ts` | 本地 FS、路径授权 | 读写项目/附件文件 |
| Shell | `tools/registry/shell.ts` | `shellExecutor.ts` | `security/sandbox/*`、`automation/processLimits.ts` | 安全 shell 执行、审批、审计 |
| Python | `tools/registry/python.ts` | `pythonExecutor.ts` | `automation/python.ts` | Python 脚本执行 |
| Knowledge | `tools/registry/knowledge.ts` | `knowledgeExecutors.ts` | `knowledge/retriever.ts`、`knowledgeWriter.ts` | 检索、列出、写入、修改、追加、删除知识库内容 |
| Web | `tools/registry/web.ts` | `webSearchExecutors.ts` | HTTP fetch、HTML parser | 模型上网搜索 |
| OCR | `tools/registry/ocr.ts` | `ocrExecutors.ts` | MinerU token、免费降级、本地解析 | 图片/PDF OCR、发票字段提取辅助 |
| Memory | `tools/registry/memory.ts` | `memoryExecutors.ts` | `memory/longTerm/*` | 长期记忆列出、写入、删除 |
| Office 文件级 | `tools/registry/office.ts` | `officeExecutors.ts` | `officeCore/*`、`officeOpenXml/*`、`office/*` | Word/PPT/Excel 文件创建、编辑、美化、快照 |

## 7. Office/WPS 当前窗口与 OpenXML 文件级编辑

```mermaid
flowchart TB
  subgraph DirectIPC["前端直接 Office IPC"]
    ExcelUi["excel.readRange/writeRange/getSelection"]
    WordStatus["office.detectWordStatus"]
    PptStatus["office.detectPptStatus"]
  end

  subgraph AgentTools["模型 Office 工具"]
    ExcelTools["workbook/range/formula/script/sheet/ui"]
    OfficeTools["office.action.*"]
  end

  subgraph Bridges["桥接实现"]
    ExcelBridge["implementations/excel/excelComBridge.ts"]
    RangeOps["implementations/excel/rangeOperations.ts"]
    FormulaOps["implementations/excel/formulaOperations.ts"]
    WordBridge["implementations/office/wordComBridge.ts"]
    PptBridge["implementations/office/presentationComBridge.ts"]
    ComAction["implementations/office/officeComActionBridge.ts"]
    OpenXmlEngine["implementations/officeOpenXml/officeOpenXmlEngine.ts"]
  end

  DirectIPC --> ExcelBridge
  DirectIPC --> WordBridge
  DirectIPC --> PptBridge
  AgentTools --> ExcelBridge
  AgentTools --> RangeOps
  AgentTools --> FormulaOps
  AgentTools --> ComAction
  ComAction --> OpenXmlEngine
  ComAction --> WordBridge
  ComAction --> PptBridge
```

Office 连接详情：

| 场景 | 入口文件 | 调用链 | 输出 |
| --- | --- | --- | --- |
| 读取/写入当前 Excel/WPS | `preload.ts` 的 `excel.*`、`ipcOfficeApi.ts`、`ipcHandlers.ts` | `excelBridgeRef()` -> `excelComBridge.ts` -> `rangeOperations.ts`/`workbookOperations.ts` | 单元格值、选区、工作簿结构、写入结果 |
| 公式生成/验证 | `tools/registry/formula.ts`、`prompts/templates/scenarios/formula.zh-CN.md` | `excelExecutors.ts` -> `formulaOperations.ts` / `rangeOperations.ts` | 公式写入、回读、动态数组 spill 校验 |
| Excel 脚本 | `tools/registry/script.ts` | `excelExecutors.ts` -> `excelScriptBridgeCom.ts` / `automation/scriptEngine.ts` | Python/JScript/PowerShell 执行结果 |
| Word/PPT 当前窗口状态 | `preload.ts` 的 `office.*` | `ipcHandlers.ts` -> `wordComBridge.ts` / `presentationComBridge.ts` | 当前宿主连接状态 |
| 文件级 Word/PPT/Excel 编辑 | `tools/registry/office.ts` | `officeExecutors.ts` -> `officeCore/officeActionAdapter.ts` -> `officeOpenXml/*`，必要时 COM fallback | 修改后的 Office 文件、视觉快照、变更摘要 |
| OpenXML Excel | `officeOpenXml/advancedExcel.ts`、`excelSheetXml.ts`、`excelFormulaXml.ts` | `officeOpenXmlEngine.ts` 包读写 | sheet/cell/formula/table/style XML |
| OpenXML Word | `officeOpenXml/advancedWord.ts` | `officeOpenXmlEngine.ts` | docx 段落、表格、样式 |
| OpenXML PPT | `officeOpenXml/advancedPresentation.ts`、`presentationPackageParts.ts`、`presentationSlideContent.ts` | `officeOpenXmlEngine.ts` | pptx slide、rels、theme、layout |

## 8. 知识库 RAG 与模型可修改内容链路

```mermaid
flowchart TB
  subgraph UI["前端知识库"]
    KnowledgeSettings["components/settings/KnowledgeSettings.tsx"]
    IpcKnowledge["services/ipcKnowledgeApi.ts"]
  end

  subgraph IPC["主进程知识库 IPC"]
    KHandlers["agent/interaction/ipcAgentHandlers.ts knowledge:*"]
  end

  subgraph Runtime["知识库运行时"]
    KRuntime["runtime/knowledgeRuntime.ts"]
    Indexer["knowledge/knowledgeIndexer.ts"]
    Writer["knowledge/knowledgeWriter.ts"]
    Retriever["knowledge/retriever.ts"]
    Parser["knowledge/documentParser.ts + excelWorkbookParser.ts + jsonFlatten.ts"]
    Chunker["knowledge/textChunker.ts"]
    Store["knowledge/sqliteStore.ts + sqliteStoreSchema.ts"]
    Embedding["knowledge/embeddingService.ts"]
  end

  subgraph Agent["模型工具入口"]
    KTools["tools/registry/knowledge.ts"]
    KExecutors["tools/executors/knowledgeExecutors.ts"]
  end

  KnowledgeSettings --> IpcKnowledge --> KHandlers
  KHandlers --> KRuntime
  KRuntime --> Indexer
  KRuntime --> Writer
  KRuntime --> Retriever
  Indexer --> Parser --> Chunker --> Embedding --> Store
  Retriever --> Embedding
  Retriever --> Store
  KTools --> KExecutors --> Retriever
  KExecutors --> Writer
```

知识库文件连接：

| 文件 | 职责 | 上游 | 下游/数据 |
| --- | --- | --- | --- |
| `components/settings/KnowledgeSettings.tsx` | 展示已索引来源、索引文件/文件夹、删除/重建 | 设置页 | `ipcKnowledgeApi.ts` |
| `services/ipcKnowledgeApi.ts` | 前端知识库 IPC wrapper | KnowledgeSettings、其他前端 | `window.electronAPI.knowledge.*` |
| `agent/interaction/ipcAgentHandlers.ts` | `knowledge:listSources/search/indexFile/indexFolder/deleteFile/reindexAll` | preload | `knowledgeRuntime`、path authorizer |
| `runtime/knowledgeRuntime.ts` | 创建/刷新 RAG runtime；provider/model 变化时重载 | Agent runtime、settings:set | `SqliteStore`、`EmbeddingService`、`KnowledgeIndexer`、`Retriever` |
| `knowledge/documentParser.ts` | 解析 txt/md/csv/json/docx/pptx/pdf 等文档文本 | `KnowledgeIndexer` | 纯文本段落 |
| `knowledge/excelWorkbookParser.ts` | Excel 工作簿解析为文本/表格语义 | `documentParser.ts` | sheet/table 文本 |
| `knowledge/textChunker.ts` | 按标题、表格、长度切块 | `KnowledgeIndexer` | chunk records |
| `knowledge/embeddingService.ts` | 根据当前 AI provider 生成 embedding | indexer/retriever | 向量与 profile |
| `knowledge/sqliteStore.ts` | 知识库 SQLite 持久化、搜索、来源摘要 | indexer/retriever/writer | `knowledge.sqlite` |
| `knowledge/knowledgeWriter.ts` | 模型写入、替换、追加、删除知识库内文本来源 | `knowledgeExecutors.ts` | 可写文本来源和索引重建 |
| `tools/registry/knowledge.ts` | 模型可见知识库工具 schema | `toolDefinitions.ts` | `knowledge.search/listSources/write/updateSource/deleteSource` |
| `tools/executors/knowledgeExecutors.ts` | 工具参数校验、调用 retriever/writer | `toolExecutor.ts` | tool_result |

## 9. OCR、附件与视觉解析链路

```mermaid
flowchart TB
  Drag["拖拽/粘贴/选择附件"]
  ComposerFiles["hooks/composerAttachmentFiles.ts + utils/fileBase64.ts"]
  ChatInput["聊天附件"]
  OcrPanel["components/task/OCRTaskComposerPanel.tsx"]
  PreloadOcr["preload.ts ocr.recognize"]
  IpcOcr["main-modules/ipcOcrHandlers.ts"]
  Mineru["main-modules/mineruOcr.ts"]
  Invoice["main-modules/invoiceFieldExtraction.ts"]
  OcrTool["tools/registry/ocr.ts"]
  OcrExec["tools/executors/ocrExecutors.ts"]
  LocalParser["tools/executors/localDocumentParser.ts"]

  Drag --> ComposerFiles
  ComposerFiles --> ChatInput
  ComposerFiles --> OcrPanel
  OcrPanel --> PreloadOcr --> IpcOcr
  IpcOcr --> Mineru
  IpcOcr --> Invoice
  OcrTool --> OcrExec
  OcrExec --> Mineru
  OcrExec --> LocalParser
  Mineru --> LocalParser
```

OCR/附件连接：

| 场景 | 文件 | 调用关系 |
| --- | --- | --- |
| 聊天附件上传 | `hooks/useComposer.ts`、`composerAttachmentFiles.ts`、`utils/fileBase64.ts` | 解析拖拽/粘贴文件，生成附件对象，随 `AgentTurnInput.attachments` 进入 Agent |
| 图片预览 | `components/chat/AttachmentImagePreview.tsx`、`utils/attachmentPreview.ts` | 消息列表读取附件路径/base64 展示 |
| OCR 功能面板 | `components/task/OCRTaskComposerPanel.tsx`、`OCRFileUploadSection.tsx`、`OCRResultSection.tsx` | 前端静默识别并展示字段和预览，可写回 Excel |
| 主进程 OCR IPC | `main-modules/ipcOcrHandlers.ts`、`ipcHandlers.ocr.test.ts` | `ocr:recognize` 入口，走 MinerU/降级/字段提取 |
| MinerU 接口 | `main-modules/mineruOcr.ts` | 付费 token 优先、免费限制后降级、本地兜底 |
| 发票字段提取 | `main-modules/invoiceFieldExtraction.ts` | 将 OCR 文本整理为字段结构 |
| 模型 OCR 工具 | `tools/registry/ocr.ts`、`tools/executors/ocrExecutors.ts` | 让无多模态模型通过工具识别图片/PDF |
| 本地文档解析兜底 | `tools/executors/localDocumentParser.ts` | 解析 docx/pptx/pdf/txt/json 等本地内容 |

## 10. 会话、记忆、压缩与线程拓扑

```mermaid
flowchart TB
  ThreadIPC["thread:* IPC"]
  Loop["AgentLoop"]
  Session["memory/sessionStore.ts"]
  RuntimeState["memory/stateRuntimeStore.ts"]
  Logs["stateRuntimeToolLogs.ts + stateRuntimeRolloutEvents.ts"]
  Threads["stateRuntimeThreads.ts"]
  Goals["stateRuntimeGoals.ts"]
  Memories["stateRuntimeMemories.ts + longTerm/*"]
  Graph["memory/agentGraphStore.ts"]
  Compaction["memory/compaction.ts + core/agentLoop/compaction*.ts"]
  Jsonl["JSONL rollout 兼容副本"]
  Sqlite["state/logs/goals/memories SQLite"]

  ThreadIPC --> Session
  ThreadIPC --> Graph
  Loop --> Session
  Loop --> RuntimeState
  Loop --> Compaction
  RuntimeState --> Logs
  RuntimeState --> Threads
  RuntimeState --> Goals
  RuntimeState --> Memories
  Session --> Jsonl
  RuntimeState --> Sqlite
  Memories --> Sqlite
  Graph --> Sqlite
```

记忆/会话文件连接：

| 文件 | 职责 | 连接 |
| --- | --- | --- |
| `memory/sessionStore.ts` | 会话、线程、turn、rollout 兼容写入 | AgentLoop 写入；thread IPC 读取；`sessionStoreFiles.ts` 扫描文件 |
| `memory/stateRuntimeStore.ts` | state/logs/goals/memories 四库门面 | AgentLoop、Stats、LongTermMemoryStore |
| `memory/stateRuntimeThreads.ts` | 线程快照与运行态 SQL | `StateRuntimeStore` |
| `memory/stateRuntimeToolLogs.ts` | 工具执行日志 | `toolExecutionLog.ts`、UsageStats |
| `memory/stateRuntimeRolloutEvents.ts` | rollout 事件和 FTS 搜索 | `StateRuntimeStore` |
| `memory/stateRuntimeGoals.ts` | 任务目标持久化 | `StateRuntimeStore` |
| `memory/stateRuntimeMemories.ts` | 长期记忆表级 CRUD | `LongTermMemoryStore` |
| `memory/longTerm/*` | 长期记忆抽取、合并、裁剪、成功画像 | `memoryExecutors.ts`、AgentLoop 启动任务 |
| `memory/agentGraphStore.ts` | 线程派生关系图 | `threadGraph:* IPC` |
| `memory/compaction.ts` | 历史压缩和 token 估算基础能力 | `core/agentLoop/compaction*.ts` |

## 11. 设置、Provider、沙箱与窗口体验

```mermaid
flowchart TB
  SettingsUI["SettingsPage + settings/*"]
  SettingsStore["store/settingsStore.ts"]
  IpcSettings["ipcApi.settings"]
  MainSettings["main-modules/settingsManager.ts"]
  IpcHandlers["main-modules/ipcHandlers.ts settings:*"]
  RuntimeRefresh["agent/runtime/agentRuntime.ts refreshKnowledgeRuntime"]
  SandboxUI["ExecPolicySettings.tsx"]
  SandboxMain["main-modules/ipcSandboxHandlers.ts"]
  SandboxCore["agent/security/sandbox/*"]
  WindowMgr["main-modules/windowManager.ts"]
  Providers["agent/providers/*"]

  SettingsUI --> SettingsStore --> IpcSettings --> IpcHandlers --> MainSettings
  IpcHandlers --> RuntimeRefresh
  SettingsStore --> SandboxUI --> SandboxMain --> SandboxCore
  MainSettings --> WindowMgr
  MainSettings --> Providers
```

设置链路：

| 模块 | 文件 | 说明 |
| --- | --- | --- |
| 设置页面 | `components/SettingsPage.tsx`、`settings/*` | 模型、常规、知识库、安全策略、开源项目、统计页 |
| 设置状态 | `store/settingsStore.ts`、`settingsLoadedState.ts`、`settingsProviderState.ts`、`settingsValues.ts` | 从 electron-store 加载，局部持久化更新 |
| Provider UI | `AddProviderDialog.tsx`、`EditProviderDialog.tsx`、`ProviderCard.tsx`、`ReasoningModeSelect.tsx` | 模型配置、推理模式、聚合平台适配 |
| Provider 运行时 | `providers/aiClientFactory.ts`、`openaiCompatibleClient.ts`、`openaiResponsesClient.ts`、`providerClients.ts` | 生成统一 AI client，适配不同协议 |
| 常规体验 | `GeneralSettings.tsx`、`windowManager.ts` | 紧凑模式、透明度、动态数组函数环境支持 |
| 沙箱策略 | `ExecPolicySettings.tsx`、`ipcSandboxHandlers.ts`、`security/sandbox/*` | shell 命令前置评估、审批、审计、可写根 |
| 设置变更副作用 | `ipcHandlers.ts` 的 `settings:set` | AI 配置变化刷新 Agent/Knowledge runtime；动态数组开关注册到全局 |

## 12. 文件级模块索引

| 顶层目录/文件 | 模块定位 | 主要上游 | 主要下游 |
| --- | --- | --- | --- |
| `desktop/src/main.tsx` | React 入口 | Vite/Electron renderer | `App.tsx` |
| `desktop/src/App.tsx` | 应用壳、页面切换、窗口按钮 | React root | ChatPage、Sidebar、SettingsPage、settingsStore |
| `desktop/src/components/chat/*` | 消息、思考、工具、Markdown、流式渲染 | ChatPage、chatStore | UI 展示、ToolConfirmDialog |
| `desktop/src/components/task/*` | 公式/代码/OCR/报告/简单任务面板 | FloatingTaskPanel | task payload、OCR IPC、chatStore |
| `desktop/src/components/settings/*` | 设置子页面和弹窗 | SettingsPage | settingsStore、ipcApi |
| `desktop/src/components/sidebar/*` | 会话/文件夹/搜索/排序/底部状态 | Sidebar | threadActions、folder/file IPC |
| `desktop/src/hooks/*` | 输入框、Office 连接、侧边栏交互 hooks | 组件 | ipcApi、utils |
| `desktop/src/services/*` | 前端 IPC wrapper 和 mock | stores/components/tests | preload 暴露的 electronAPI |
| `desktop/src/store/*` | Zustand 状态、事件投影、线程操作 | UI、preload event | ipcApi、agentEventHandler、chatStreamBuffer |
| `desktop/src/utils/*` | 纯函数和格式化工具 | UI/store/hooks | 无运行时副作用 |
| `desktop/electron/preload.ts` | 安全隔离桥 | Renderer | ipcRenderer.invoke/on |
| `desktop/electron/main.ts` | 主进程生命周期入口 | Electron app | runtime、IPC、window、shutdown |
| `desktop/electron/main-modules/*` | 主进程设置、窗口、文件、AI、OCR、沙箱 IPC | main.ts/preload | settingsManager、Agent runtime、MinerU、本地 FS |
| `desktop/electron/agent/interaction/*` | Agent IPC 与事件转发 | main-modules/ipcHandlers.ts | AgentLoop、Session/Knowledge/Stats、BrowserWindow |
| `desktop/electron/agent/runtime/*` | Agent 装配层 | main.ts、settings:set | bridges、knowledge、memory、toolExecutors、AgentLoop |
| `desktop/electron/agent/core/agentLoop/*` | 对话轮次和工具编排核心 | ipcAgentHandlers、AgentLoopManager | providers、prompts、tools、memory |
| `desktop/electron/agent/providers/*` | AI 协议客户端 | AgentLoop | OpenAI/Responses/Anthropic/厂商 API |
| `desktop/electron/agent/prompts/*` | 系统提示词和场景片段 | buildStreamParams | 模型请求 |
| `desktop/electron/agent/tools/registry/*` | 工具 schema | buildStreamParams/toolExecutor | 模型可见工具列表 |
| `desktop/electron/agent/tools/executors/*` | 工具执行器 | toolExecutor | contracts、implementations、knowledge、OCR、sandbox |
| `desktop/electron/agent/tools/implementations/excel/*` | Excel/WPS COM 具体能力 | excelExecutors、ipcHandlers | COM、automation |
| `desktop/electron/agent/tools/implementations/office/*` | Word/PPT COM 与 Office action fallback | officeExecutors、officeCore | COM、PowerShell |
| `desktop/electron/agent/tools/implementations/officeOpenXml/*` | xlsx/docx/pptx 文件级 OpenXML 引擎 | officeCore/officeExecutors | ZIP/XML 文件包 |
| `desktop/electron/agent/tools/officeCore/*` | Office action 统一定位、能力、结果适配 | officeExecutors | OpenXML 优先、COM fallback |
| `desktop/electron/agent/knowledge/*` | RAG 文档解析、切块、embedding、检索、写入维护 | runtime、knowledge IPC、knowledge tools | SQLite、AI embedding |
| `desktop/electron/agent/memory/*` | 会话、运行态、长期记忆、压缩、线程图 | AgentLoop、thread IPC、memory tools | SQLite、JSONL |
| `desktop/electron/agent/security/sandbox/*` | shell 安全策略和审计 | shellExecutor、ipcSandboxHandlers | spawn 包装、audit logs |
| `desktop/electron/agent/automation/*` | PowerShell/JScript/Python 进程执行基础设施 | Office/Excel implementations、shell/python executors | 本地脚本进程 |
| `desktop/electron/agent/shared/*` | Agent 共享类型、消息转换、数值限制 | core/providers/tools | 轻量公共能力 |

## 13. 维护注意事项

- 新增前端能力时，优先走 `services/ipcApi.ts` 子 wrapper；如果新增 preload API，需要同步 `src/electronApi.d.ts`、`services/ipcApiTypes.ts` 和对应 wrapper。
- 新增模型工具时，必须同时补齐 `tools/registry/*`、`tools/executors/*`、必要的 `contracts/*` 或 `implementations/*`，并确认 `toolDefinitions.ts` 暴露顺序。
- 新增 Office 文件级能力时，优先接 `office.action.*`，OpenXML 能完成的能力先放在 `officeOpenXml/*`，COM 只做当前窗口、动态对象、快照或兜底。
- 新增知识库解析格式时，入口应在 `knowledge/documentParser.ts` 或独立 parser，再接 `KnowledgeIndexer -> textChunker -> embeddingService -> sqliteStore`；模型可修改能力走 `KnowledgeWriter`。
- 修改 AI provider、reasoning、context 逻辑时，需要同步 `providers/*`、`settingsProviderState.ts`、`ReasoningModeSelect.tsx` 和 `buildStreamParams.ts`。
- 修改流式渲染时，要同时考虑主进程 `eventForwarder.ts` 的 32ms 合并和前端 `chatStreamBuffer.ts` 的 50ms 合并，避免工具事件、思考正文和最终回答时间线错位。
