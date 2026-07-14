# Agent 架构树

`desktop/electron/agent` 按架构层分目录。各层 README 记录当前职责、主要模块和依赖方向。

```text
agent/
├─ interaction/                 # 交互层：承接 Electron/UI 事件、IPC、审批回调、流式事件转发
│  └─ README.md                 # 交互层职责说明
│
├─ runtime/                     # 运行时装配层：创建并组装模型、工具、记忆、知识库、提示词和核心循环
│  └─ README.md                 # 运行时装配层职责说明
│
├─ core/                        # 核心层：Agent 对话轮次、模型流、工具调用编排、中断与继续
│  └─ agentLoop/                # Agent 核心循环与 turn/stream/tool-call 编排
│
├─ memory/                      # 记忆层：SQLite 运行态、JSONL 审计副本、上下文压缩、长期记忆
│  ├─ stateRuntimeStore.ts      # StateRuntime 四库门面：state/logs/goals/memories
│  ├─ sessionStore.ts           # 会话/线程/轮次写入，兼容 JSONL 并投影到 logs.db
│  ├─ rolloutWriter.ts          # 有界批量写入队列，降低高频写入阻塞
│  ├─ rolloutArchive.ts         # 冷 JSONL gzip 归档快照
│  └─ compaction.ts             # 历史压缩、token 估算、恢复上下文
│
├─ knowledge/                   # 知识层：本地文档解析、切块、embedding、索引、检索、文本来源维护和知识库注册表
│  ├─ documentParser.ts         # 文件解析
│  ├─ knowledgeIndexer.ts       # 索引构建
│  ├─ knowledgeWriter.ts        # 模型写入笔记、修改/追加可写文本来源、删除来源索引内容
│  ├─ retriever.ts              # 检索与结果格式化
│  └─ sqliteStore.ts            # 知识库持久化
│
├─ tools/                       # 工具层：模型可调用工具的定义、契约和路由
│  ├─ registry/                 # 工具注册表：工具 Schema、风险等级、工具目录
│  ├─ contracts/                # 工具契约：Excel/Word/PPT/Knowledge 等实现无关接口
│  ├─ executors/                # 工具执行器：参数校验、工具路由、结果封装
│
├─ officeWorker/                # .NET 8 COM 与 DocumentFormat.OpenXml Worker 薄桥
│
├─ providers/                   # 模型供应商层：OpenAI 兼容、Anthropic、厂商适配、模型上下文窗口
│  ├─ aiClient.ts               # 模型客户端统一出口
│  ├─ openaiCompatibleClient.ts # OpenAI 兼容协议实现
│  ├─ anthropicClient.ts        # Anthropic 协议适配
│  └─ providerClients.ts        # 各供应商子类
│
├─ prompts/                     # 提示词层：系统提示词、工具选择指南、文件夹上下文注入
│  └─ systemPrompt.ts           # 提示词总装配入口
│
├─ attachments/                 # 附件层：图片/文件等输入进入核心层前的预处理
│  └─ imageAttachmentResolver.ts
│
├─ shared/                      # 共享层：跨层类型、常量、消息转换等轻量公共能力
│  ├─ types.ts                  # Agent 共享类型
│  └─ messageBuilder.ts         # TurnItem 到模型消息的转换
│
├─ architecture.test.ts         # 架构归位测试：防止关键文件漂回旧目录
└─ README.md                    # 本架构说明
```

## 依赖方向

```text
interaction -> runtime -> core
core -> providers / prompts / memory / knowledge / tools / shared
tools/executors -> tools/contracts + tools/registry + officeCore + officeWorker
officeWorker -> .NET Worker JSON-RPC
memory -> shared
knowledge -> 本层存储与 embedding
providers -> shared
prompts -> shared
attachments -> providers 消息类型
```

## 约束

- `core` 只做编排，不直接依赖 Excel/Word/PPT 具体 COM 实现。
- `tools/contracts` 只放实现无关接口，不反向依赖 `officeWorker`。
- `tools/executors` 负责工具路由和结果封装，不依赖 Electron UI。
- `memory`、`knowledge` 不依赖 `core`。
- `providers` 不依赖 `core`、`tools`、`memory`、`knowledge`。
- `shared` 保持轻量，不引入其他 Agent 层的运行时依赖。

## 当前归位状态

- `tools/contracts` 保存实现无关接口，`tools/executors` 保存工具执行路由，`tools/registry` 只保存模型可见工具定义。
- Excel、Word、PowerPoint 和 WPS 的 COM/Open XML 能力统一由 `officeWorker` 类型化薄桥调用 `desktop/dotnet/Wengge.OfficeWorker`；不保留任意脚本执行层。
- Office 三件套文件级编辑统一走 `office.action.*`，Open XML 优先，COM 只作为动态对象、快照、当前窗口交互或 Open XML 不适用时的兜底。
- Agent IPC 和事件转发归入 `interaction`，运行期装配归入 `runtime`。
- `core/agentLoop` 已拆出 turn、stream、tool-call、压缩配置和上下文窗口相关辅助模块。
- 系统提示词由 `prompts/systemPrompt.ts` 装配，静态和场景片段存放在 `prompts/templates`，运行时文件夹上下文由 `prompts/sections/folderContextPrompt.ts` 渲染。
- 运行态主存储已迁移到 `memory/stateRuntimeStore.ts` 四个 SQLite 库；JSONL 仅作为兼容审计副本保留。
