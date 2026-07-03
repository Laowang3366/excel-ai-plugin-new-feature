# 长期记忆系统设计

## 目标

为 Office 助手增加轻量、可持续演进的长期记忆系统。当前客户端主要服务 Excel / Word / PPT 等办公软件任务，大多数任务是一次性的文件创建、编辑、排版和修复，因此首版不追求记住每个任务的完整内容，而是重点记住跨任务稳定复用的用户偏好、规则约束、操作方式偏好、用户纠正、用户过往编辑文件的轻量印象，以及系统内部的工具成功率经验。

系统以现有 `StateRuntimeStore` 的 `memories.db` 为主存储，继续保留 JSONL/rollout 作为审计来源，不引入 `raw_memories.md` 作为主存储。

## 当前状态

当前项目已经有四库 SQLite 运行态：`state.db`、`logs.db`、`goals.db`、`memories.db`。其中 `memories.db` 只有基础 `memories` 表，以及 `StateRuntimeStore.upsertMemory()` / `listMemories()` 这类人工写入能力。

缺失的核心能力包括：

- 从 rollout 自动提取高价值候选记忆，重点识别用户偏好、长期约束和纠正反馈。
- 用 AI 对候选记忆做清洗、去重、合并，避免把一次性任务细节写成长期记忆。
- 记忆引用溯源和来源审计。
- 记忆过期裁剪。
- 用户纠正沉淀，例如“失败在哪里、哪个环节不对、下次应如何规避”。
- 过往编辑文件印象，例如文件类型、文件名、最近操作、用户关注点和高层摘要。
- 系统内部工具调用经验，例如某类 Office 操作哪条执行路线成功率更高、失败集中在哪个环节。
- 启动时后台记忆管道。
- 模型可自主调用的 `memory.*` 工具。

## 设计原则

1. 以数据库为主，不把 markdown 文件作为主存储。
2. 先做稳定的本地管道，再扩展远程或插件式能力。
3. 记忆写入必须结构化，不能让模型直接写任意大段上下文。
4. 默认不记一次性文件正文、表格明细、临时数据和中间生成内容。
5. 所有 AI 生成的记忆必须保留 citation，能追溯到 thread / rollout event。
6. 用户纠正优先级高于模型自行总结，但仍要结构化、去重、可撤回。
7. 文件印象只记录低敏高层信息，不替代文件本身，也不作为恢复文件内容的来源。
8. 用户可见记忆和系统内部策略记忆必须隔离，内部工具策略不能伪装成用户偏好。
9. 工具成功率经验只能作为路由优先级建议，不能把某条执行路线永久写死。
10. 自动清理采用“标记 stale + 后续 prune”策略，避免误删。
11. 模块职责单一，但不为了拆分而拆碎。

## 架构

新增长期记忆首版核心目录：

```text
desktop/electron/agent/memory/longTerm/
  memoryTypes.ts
  memoryStore.ts
  memoryExtraction.ts
  memoryConsolidation.ts
  memoryPruning.ts
  memoryStartupTask.ts
```

模块职责：

- `memoryTypes.ts`: 定义长期记忆 schema、AI 提取输出、citation、状态枚举和可选 scope。
- `memoryStore.ts`: 封装 `StateRuntimeStore` 的长期记忆读写、搜索、过期、引用查询。
- `memoryExtraction.ts`: 从 `logs.db rollout_events` 读取最近会话事件，提取偏好、约束、纠正、文件印象和工具调用经验类候选记忆。
- `memoryConsolidation.ts`: 对候选记忆做去重、合并、更新已有记忆，过滤一次性任务内容。
- `memoryPruning.ts`: 按过期时间、stale 状态、namespace 策略裁剪。
- `memoryStartupTask.ts`: Agent 启动后后台运行轻量记忆同步任务。

后续扩展目录：

- `memoryWorkspace.ts`: 仅在存在工作区绑定记忆时计算 fingerprint 并标记 stale。
- `memoryExtensions.ts`: 定义扩展接口，允许后续按项目、工具、业务域注入额外记忆指令。

## 数据模型

扩展 `memories.db` schema，而不是新建外部文件。建议迁移到：

```sql
CREATE TABLE IF NOT EXISTS memories (
  memory_id TEXT PRIMARY KEY,
  namespace TEXT NOT NULL,
  kind TEXT NOT NULL,
  content TEXT NOT NULL,
  summary TEXT,
  confidence REAL,
  status TEXT NOT NULL,
  source_thread_id TEXT,
  source_event_id INTEGER,
  workspace_fingerprint TEXT,
  expires_at INTEGER,
  metadata_json TEXT,
  citations_json TEXT,
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_memories_namespace_status_updated_at
  ON memories(namespace, status, updated_at);

CREATE INDEX IF NOT EXISTS idx_memories_source_thread
  ON memories(source_thread_id);
```

用户可见 `kind` 初始支持：

- `preference`: 用户偏好，例如“回复尽量直接，先给结论再补充说明”。
- `constraint`: 用户长期约束，例如“每个阶段完成后先确认，通过后再进入下一步”。
- `correction`: 用户纠正和失败复盘，例如“创建 PPT 失败后，不要反复尝试打开应用，应改用更稳定的文件级创建方式”。
- `style_preference`: 办公文档风格偏好，例如“标题字号减少 2px，顶部导航更紧凑”。
- `operation_preference`: 用户可理解的操作方式偏好，例如“优先使用稳定的文件级编辑，只有需要预览或用户明确要求时再打开办公软件”。
- `file_impression`: 过往编辑文件印象，例如“健康饮食宣传学生版 PPT，约 10 页，目标受众是学生，曾尝试创建/优化排版”。

系统内部策略 `kind` 初始支持：

- `tool_success_profile`: 工具调用成功率经验，例如“PPT 创建/编辑在当前环境下文件级编辑成功率高于应用自动化打开”。该类记忆只给工具路由器和执行层使用，不注入普通对话提示词，也不展示为用户偏好。

后续可选支持：

- `project_fact`: 项目事实，例如“StateRuntime 使用四个 SQLite 库”。首版只在用户明确要求记住项目规则时写入。
- `workflow`: 长期工作流。首版只记录跨多次任务稳定出现的流程偏好。

`status` 初始支持：

- `active`: 可注入上下文和工具检索。
- `stale`: 工作区变化或低置信度后暂停使用。
- `archived`: 已保留审计但默认不返回。

## AI 提取管道

### Stage One：候选记忆提取

从 `logs.db rollout_events` 读取最近未处理的用户消息、助手总结、压缩摘要和工具结果，拼成受控输入。提取器只保留跨任务稳定复用的信息：

- 用户明确偏好：输出格式、文档风格、操作方式、权限策略。
- 用户长期约束：阶段 review、提交 git、测试文件清理等。
- 用户纠正：指出哪里错了、哪个环节不对、下次应如何执行。
- 反复出现的失败模式：例如某类编辑任务多次失败在“打开办公软件”环节，后续同类任务应优先使用文件级编辑。
- 文件印象：用户编辑过的 Office 文件、文件类型、文件名、任务目标、最近操作和非敏感高层摘要。
- 系统内部工具调用经验：同类任务中工具族、应用、操作和执行结果的稳定倾向。

提取器必须忽略一次性内容：

- 当前文件的大段正文、表格明细、完整 PPT 文案。
- 临时路径、临时脚本、一次性 shell 输出。
- 单次偶然成功或失败的工具结果。
- 某次任务的中间推理和未被用户确认的猜测。

AI 输出结构化 JSON：

```ts
interface StageOneOutput {
  memories: Array<{
    kind: "preference" | "constraint" | "correction" | "style_preference" | "operation_preference" | "file_impression" | "tool_success_profile";
    namespace: string;
    content: string;
    summary?: string;
    confidence: number;
    citations: MemoryCitation[];
    errorStage?: "planning" | "tool_selection" | "tool_execution" | "file_editing" | "ui" | "review";
    nextTimeRule?: string;
    file?: {
      displayName: string;
      fileType: "xlsx" | "docx" | "pptx" | "csv" | "other";
      pathHash?: string;
      lastAction?: string;
    };
    toolProfile?: {
      toolFamily: "openxml" | "com" | "script" | "shell" | "python" | "office_action" | "other";
      app?: "excel" | "word" | "powerpoint" | "office";
      operation?: string;
      successCount: number;
      failureCount: number;
      preferredRoute?: string;
      avoidRoute?: string;
      reason?: string;
    };
    expiresInDays?: number;
  }>;
}
```

解析失败时不写入记忆，只记录 diagnostics，避免污染长期存储。

### Stage Two：清洗合并

`memoryConsolidation.ts` 读取同 namespace 下的相似 active 记忆，构建 consolidation prompt，让 AI 判断：

- 新增。
- 合并到已有记忆。
- 更新已有记忆。
- 忽略低价值或重复记忆。

合并结果必须保留所有 citation，不能覆盖来源。

纠正类记忆的合并规则：

- 同一错误模式多次出现时合并为一条更强规则，增加 citation 和置信度。
- 新纠正与旧偏好冲突时，以最新用户纠正为准，并将旧记忆标记为 `archived` 或更新为新规则。
- 只记录可执行规则，不记录情绪化或一次性抱怨。

文件印象的合并规则：

- 同一 `pathHash` 或同名同类型文件反复出现时，更新最近操作和高层摘要，不重复新增。
- 文件印象可以记录“用户想做什么”和“我们做过什么”，不能记录完整正文、表格明细或隐私字段。
- 如果文件路径变化但文件名和任务目标高度相似，保留为同一印象并追加 citation。

工具成功率经验的合并规则：

- 只有同类任务出现多次可观察结果时才升级为 `tool_success_profile`，避免单次结果污染策略。
- 同一 app + operation + toolFamily 下累计成功/失败计数，并保留最近 citation。
- 当环境变化或近期结果反转时降低 confidence，不直接删除旧经验。
- `preferredRoute` 只影响工具路由器的候选排序；如果用户明确指定操作方式或当前任务能力不匹配，应服从当前任务。
- `tool_success_profile` 不进入普通对话提示词，不作为“用户说过”的偏好展示给模型。

## Prompt 模板

新增：

```text
desktop/electron/agent/prompts/templates/memory/stage_one_system.zh-CN.md
desktop/electron/agent/prompts/templates/memory/consolidation.zh-CN.md
desktop/electron/agent/prompts/templates/memory/instructions.zh-CN.md
```

对应 loader：

```text
desktop/electron/agent/prompts/memoryPrompt.ts
```

模板用中文默认版本，后续可追加英文或用户覆盖。

## 上下文注入边界

长期记忆分为两类消费路径：

- 用户可见记忆：`preference`、`constraint`、`correction`、`style_preference`、`operation_preference`、`file_impression`。这类记忆可以按相关性注入普通对话上下文，但内容必须使用用户可理解的表述。
- 系统内部策略记忆：`tool_success_profile`。这类记忆默认不进入普通对话上下文，不作为“用户偏好”告诉模型，也不展示给用户；只由工具路由器、执行器或内部策略模块读取，用于候选工具排序、兜底顺序和失败规避。

`buildStreamParams` 只能读取用户可见记忆。内部策略记忆如果必须参与决策，应在工具执行前由路由层读取，不通过用户提示词传递。

## 工具设计

新增工具定义：

- `memory.write`: 写入一条结构化长期记忆。
- `memory.read`: 按 `memoryId` 读取完整记忆和引用。
- `memory.search`: 按 query / namespace / kind 检索 active 记忆，默认只返回用户可见记忆。
- `memory.list`: 列出某 namespace 下的记忆摘要，默认只返回用户可见记忆。

工具默认 `riskLevel: safe`，但 `memory.write` 必须做参数校验：

- `content` 长度限制。
- `kind` 枚举限制。
- `namespace` 默认 `global` 或当前项目。
- citation 可选；模型主动写入时 metadata 标记 `source: "tool"`.
- `tool_success_profile` 只能由内部调用写入或读取；普通模型检索默认不可见。

工具使用边界：

- 模型可以主动写入用户明确说出的偏好、约束和纠正。
- 模型可以写入低敏文件印象，例如文件名、类型、最近操作、高层摘要。
- 稳定工具调用经验主要由工具执行器或内部采集器写入；如果由模型触发写入，必须标记为 `source: "telemetry"`，并包含样本数量或多条 citation。
- 模型不能把用户文件正文、表格明细、隐私数据、临时路径、一次性任务结果写入长期记忆。
- 当用户说“记住/以后都/下次不要/这里错了/不是这样”时，应优先考虑写入 `memory.write`。

## 启动任务

`memoryStartupTask.ts` 在 Agent runtime 初始化后启动，但不阻塞客户端启动。

启动任务步骤：

1. 检查 `memories.db` migration。
2. 扫描 `logs.db` 中上次处理位置之后的 rollout events。
3. 如果记忆功能启用，运行 Stage One，只提取高置信度偏好、约束、纠正、文件印象和工具调用经验。
4. 对候选记忆运行 consolidation。
5. 执行 prune。
6. 将处理位置写入 `memory_pipeline_state` 表。

失败策略：

- 单次失败只记录错误，不影响聊天。
- 连续失败时退避，避免启动时反复请求模型。

## 工作区 Diff 联动

工作区 diff 不进入首版核心。当前用户场景以一次性 Office 任务为主，长期价值主要来自用户偏好和纠正，不来自项目文件事实。

后续如果需要支持长期项目型任务，再做轻量 fingerprint：

- 记录当前 workspace root、关键文件路径、mtime、size。
- 当 fingerprint 变化时，标记 `metadata.workspaceBound === true` 的相关记忆为 `stale`。
- 不直接删除，后续 prune 决定是否 archive。

后续可以接入更细粒度 diff，将 citation 与文件路径绑定。

## 过期裁剪

默认策略：

- `constraint` / `preference` / `style_preference` / `operation_preference`: 默认不过期，除非用户删除或 AI 合并替换。
- `correction`: 默认 30 天后重新验证；如果相同纠正反复出现，则升级为长期规则。
- `file_impression`: 默认 90 天后归档；如果用户再次编辑同一文件，则刷新有效期。
- `tool_success_profile`: 默认 30 天后重新验证；如果最近 5 次同类工具结果反转，则降低置信度或归档。
- `project_fact` / `workflow`: 非首版核心，默认 7 天后重新验证。
- `stale` 超过 7 天归档。

裁剪只改 `status`，不物理删除；物理 vacuum/清理放到后续维护任务。

## Rollout 摘要同步

`sync_rollout_summaries_from_memories()` 的本项目版本应做两件事：

1. 将 `compacted` rollout 摘要作为候选输入送入 Stage One，但只提取偏好、约束、纠正、文件印象和工具调用经验。
2. 将高置信度长期记忆同步成可检索上下文，在构建模型请求时按需注入，不直接写回历史 turn。

这样保留 rollout 审计完整性，同时避免历史消息被长期记忆反复污染。

## 运行时接入点

- `runtime/agentRuntime.ts`: 初始化 `MemoryRuntime`，传给工具执行器和 AgentLoop。
- `tools/registry`: 注册 `memory.*` 工具定义。
- `tools/executors`: 注册 `memory.*` 执行器。
- `core/agentLoop/buildStreamParams.ts`: 只注入用户可见的高相关长期记忆摘要，不注入 `tool_success_profile`。
- 工具路由/执行层：读取 `tool_success_profile`，用于同类任务的候选工具排序和兜底策略选择。
- `memory/stateRuntimeStore.ts`: 扩展 schema 和底层查询。

## 测试策略

核心测试：

- `stateRuntimeStore.test.ts`: migration、upsert、search、stale、prune。
- `memoryExtraction.test.ts`: rollout event 到 StageOne 输入、AI JSON 解析失败不写入。
- `memoryConsolidation.test.ts`: 新增、合并、忽略重复；工具成功率经验需要多次同类结果才升级为 active 记忆。
- `memoryTools.test.ts`: `memory.write/read/search/list` 参数校验和返回格式。
- `memoryStartupTask.test.ts`: 启动任务不阻塞、记录处理位置、失败退避。
- `buildStreamParams.test.ts`: 只注入 active 且相关的用户可见长期记忆，排除 `tool_success_profile`。
- 工具路由测试：验证 `tool_success_profile` 只影响内部候选排序，不进入普通对话提示词。

每个阶段必须先写失败测试，再实现最小代码，通过 review 后提交。

## 分阶段落地

### 阶段一：存储和 Schema

扩展 `memories.db` schema，增加偏好、约束、纠正、文件印象和工具成功率经验类长期记忆字段、pipeline state、基础 store API。此阶段不调用 AI，不改模型上下文。

### 阶段二：AI 提取和清洗

增加 memory prompt 模板、Stage One 提取和 consolidation。提取范围只包含用户偏好、规则约束、纠正反馈、用户可理解的操作方式偏好、低敏文件印象和内部工具调用经验。先提供手动触发函数，避免启动时自动请求模型造成不可控成本。

### 阶段三：模型工具

注册 `memory.write/read/search/list`，让模型能自主读写查长期记忆。所有写入必须通过 schema 校验，并禁止写入一次性文件内容。

### 阶段四：启动任务和上下文注入

接入 `memoryStartupTask`、过期裁剪和构建请求时的记忆注入。上下文注入只放入少量高相关偏好、约束、纠正规则和当前任务相关文件印象；工具成功率经验只接入工具路由/执行层。

### 阶段五：引用和同步增强

完善 citation 展示、rollout 摘要同步和扩展系统。工作区 diff 和项目事实记忆放到这一阶段之后评估。

## 非目标

- 不引入独立向量数据库。
- 不把 markdown 文件作为长期记忆主存储。
- 不在第一版实现复杂 AST 级 workspace diff。
- 不在第一版记住用户每次 Office 文件的具体内容；只记录低敏文件印象。
- 不把临时路径、临时脚本、一次性任务数据当成长期记忆。
- 不因单次工具成功或失败永久改变工具优先级。
- 不把内部工具名、执行路线或成功率统计作为用户偏好提示词注入模型。
- 不让模型绕过 schema 直接写任意自由文本到长期记忆。

## 验收标准

- 长期记忆以 `memories.db` 为主存储，并具备 migration。
- 模型可通过 `memory.*` 工具读写查记忆。
- 自动提取和清洗都输出结构化 JSON，且只保留偏好、约束、纠正、操作方式偏好、低敏文件印象和内部工具调用经验。
- 用户纠正可沉淀为包含错误环节和下次规避规则的 `correction` 记忆。
- 用户过往编辑文件可沉淀为 `file_impression`，能帮助模型识别用户提到的历史文件和之前做过的操作。
- AI 过往工具调用结果可沉淀为 `tool_success_profile`，在同类任务中影响内部工具选择优先级，但不进入普通对话提示词。
- 每条 AI 自动写入的记忆都有 citation。
- 过期和 workspace 变化不会物理删除记忆，只标记 stale/archive。
- 启动任务失败不影响正常聊天。
