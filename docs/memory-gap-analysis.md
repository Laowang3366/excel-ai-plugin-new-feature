# 记忆模块差距分析

> 对比对象：`desktop/electron/agent/` (当前项目) vs `codex-reference/codex-rs/` (Codex Reference)
>
> 评估原则：每一项都先判断"在当前 Excel AI 插件上下文中是否有价值"，再做对比。
>
> 状态说明：
> - 🔴 **值得做**：有明确价值，建议纳入路线图
> - 🟡 **可做可不做**：有一定价值但优先级低，或场景不明确
> - ⚪ **不值得做**：Codex 的特性但本项目不需要

---

## 一、长期记忆体系（Memories）

Codex 有一套完整的 AI 记忆系统：启动时扫描历史 rollout → Phase 1 提取原始记忆 → Phase 2 清洗/去重/合并 → 持久化到 `raw_memories.md` → AI 通过 tools 读写。当前项目只有 RAG 文档检索（`knowledge/`），没有 AI 自主记忆能力。

### 1.1 AI 自动记忆提取 + 清洗（Phase 1 → Phase 2）

**Codex 做法**：
- Phase 1：`stage_one_system.md` 提示词，逐条扫描已完成的 rollout，用模型提取原始记忆
- Phase 2：`build_consolidation_prompt()` 对 Phase 1 的结果做清洗、去重、合并
- 最终持久化为 `raw_memories.md` + 每个线程的 `rollout_summary.md`
- 启动时触发，后台异步执行，有限流和守卫

**价值判断**：🟡 **可做可不做**

| 维度 | 评估 |
|------|------|
| 用户场景 | Excel AI 助手的使用模式偏"一次性任务"（"帮我做一份销售报表"），而非 Codex 的"长期编码项目"。用户对 AI 记住偏好（"我上次用的蓝色主题"）有需求，但当前可以通过 Excel 文件本身的状态来满足 |
| 工程成本 | 高：需要 Phase 1 + Phase 2 两条独立的 AI 调用管线，需要守卫限流，需要存储层 |
| 替代方案 | Excel 文件本身是"状态"——用户保存文件即保存了工作成果。AI 记忆在此场景下价值不如 Codex（编码项目需要记住上下文跨会话） |
| 建议 | 当前不做完整管线。可以考虑轻量方案：仅存储用户显式偏好（如"always use bold headers"），通过 `memories.db` 的 upsert 接口支持 |

### 1.2 AI 可写记忆工具（memory.write / search / read / list）

**Codex 做法**：通过 extension 暴露 4 个 tools 给 AI：
- `memory.write` — AI 主动写入一段记忆
- `memory.search` — AI 搜索相关记忆
- `memory.read` — AI 读取特定记忆
- `memory.list` — AI 列出所有记忆

**价值判断**：🟡 **可做可不做**

| 维度 | 评估 |
|------|------|
| 用户场景 | Excel AI 助手"记住用户偏好"的场景确实存在。用户可能说"表头用深蓝色"，AI 应能在后续会话中记住 |
| 方案替代 | 当前可以通过 system prompt 注入用户偏好 + Excel 模板文件来达到类似效果，不一定需要 AI 自主记忆 |
| 风险 | AI 自主写入记忆存在误写、遗忘、冲突等问题，需要额外的审查机制 |
| 建议 | 可以做一个简化版：只暴露 `memory.write` 和 `memory.search` 两个工具，不实现完整 CRUD。但优先级不高 |

---

## 二、Session / Turn 循环

### 2.1 comp_hash 变更检测

**Codex 做法**：每个模型的 `comp_hash` 指示其压缩兼容性。当两个模型的 `comp_hash` 不同时，切换模型前自动执行压缩，确保上下文格式一致。

**价值判断**：🔴 **值得做**

| 维度 | 评估 |
|------|------|
| 用户场景 | 当前项目支持 11 个模型供应商（`providers/`），用户可能会在对话中切换模型。不同模型的上下文格式、tokenizer、压缩策略可能不同 |
| 影响面 | 不检测直接切换可能导致上下文理解错误、token 计数偏差 |
| 工程成本 | 低：在每个模型信息中加一个 `comp_hash` 字段，切换时判断是否需要重新压缩 |
| 建议 | 纳入路线图，短期可做 |

### 2.2 InputQueue 中断输入处理

**Codex 做法**：模型运行期间用户可以通过 `InputQueue` 插入新消息，当前轮结束后自动处理插入的消息，不用等到下一轮。

**价值判断**：🔴 **值得做**

| 维度 | 评估 |
|------|------|
| 用户场景 | 用户等待 AI 回复时可能想补充信息（"等一下，用 Sheet2 的数据"），当前只能等 AI 回复完再发新消息 |
| 影响面 | 用户体验改善明显 |
| 工程成本 | 中：需要修改 `agentLoop.ts` 的主循环，在 sampling 和 tool call 之间检查 pending input |
| 建议 | `PendingInterruptQueue` 已存在基础结构，扩展为支持消息插入即可 |

### 2.3 TurnContext 快照

**Codex 做法**：`TurnContext` 是一个包含环境、模型信息、配置、状态的一次性快照，整个 turn 生命周期内不可变，保证状态一致性。

**价值判断**：🟡 **可做可不做**

| 维度 | 评估 |
|------|------|
| 影响面 | 更多是架构层面的规范性和可测试性，对功能影响有限 |
| 工程成本 | 中：需要提取 `TurnContext` 类型，重构 agentLoop 中分散的状态访问。改动面大 |
| 建议 | 短期不做。可以在 `turnState.ts` 中逐步收敛状态访问点 |

### 2.4 重试策略（max_retries）

**Codex 做法**：模型调用失败时根据 `max_retries` + backoff 自动重试，支持 `ResponsesStreamRequest` 不同阶段（sampling / compact）的独立重试策略。

**价值判断**：🔴 **值得做**

| 维度 | 评估 |
|------|------|
| 用户场景 | AI API 调用可能因网络抖动、限流、500 错误而失败。当前直接报错，用户体验差 |
| 影响面 | 可靠性的核心保障 |
| 工程成本 | 低：在 `aiClient.ts` 中实现重试包装器即可 |
| 建议 | 高优先级，可直接复用 `responses_retry.rs` 的设计模式 |

### 2.5 Token budget 提醒

**Codex 做法**：`token_budget::maybe_record()` 在上下文接近上限时向用户推送预算提醒。

**价值判断**：🟡 **可做可不做**

| 维度 | 评估 |
|------|------|
| 用户场景 | 用户知道剩余空间后可以主动要求压缩或简化回复 |
| 影响面 | 用户体验细节改善 |
| 建议 | 在 context_usage event 中已经携带了 estimatedTokens / threshold / percentage 数据，前端可以利用。后端不需要额外改动 |

---

## 三、持久化层

### 3.1 分页查询

**Codex 做法**：`ListThreadsParams` 使用 cursor 分页，避免全量扫描。支持 `page_size`、`sort_key`、`cwd_filters`、`search_term` 等多种过滤条件。

**价值判断**：🔴 **值得做**

| 维度 | 评估 |
|------|------|
| 用户场景 | `listThreads()` 当前全量扫描 JSONL + 逐个 `readFile` 解析首行。用户有几十个会话时还可接受，几百个时性能明显下降 |
| 影响面 | 会话列表页面的加载性能 |
| 工程成本 | 中：`logs.db` 已经是结构化存储，把列表查询从全量扫描切换为 `SELECT ... LIMIT ? OFFSET ?` 即可。`state.db` 的 `thread_snapshots` 表已具备分页条件 |
| 建议 | 改为从 `state.db` 的 `thread_snapshots` 做分页查询，JSONL 扫描作为 fallback。短期可做 |

### 3.2 ThreadMetadataPatch merge 语义

**Codex 做法**：`ThreadMetadataPatch` 使用 merge-presence 语义，每个字段可以独立更新／清除，不会因为并发写入而覆盖其他字段。

**价值判断**：🟡 **可做可不做**

| 维度 | 评估 |
|------|------|
| 用户场景 | 当前项目是单用户桌面应用，不存在多进程并发写同一线程元数据的场景 |
| 影响面 | 只影响 `updateThreadMetadata()` 的接口设计 |
| 建议 | 当前覆盖式写入在单进程场景下足够。如需支持 IPC 并发，再改为 merge 语义 |

### 3.3 ThreadStore trait 抽象

**Codex 做法**：`trait ThreadStore` 定义 12 个抽象方法，`LocalThreadStore` 和 `InMemoryThreadStore` 分别实现。

**价值判断**：⚪ **不值得做**

| 维度 | 评估 |
|------|------|
| 理由 | 当前项目是 Electron 桌面插件，运行在单一用户本地环境，只有文件系统一个后端。没有多后端切换需求 |
| 成本 | 引入 trait 增加抽象层复杂度，对开发效率和可读性反而是负担 |
| 建议 | 当前 `SessionStore` 单类实现足够了。如果后续需要远程同步（如 OneDrive 同步），再考虑抽象 |

### 3.4 Git 元数据 / Permission 快照

**Codex 做法**：`StoredThread` 包含 `git_info`（commit SHA / branch / origin_url）和 `permission_profile`（运行时权限快照）。

**价值判断**：⚪ **不值得做**

| 维度 | 评估 |
|------|------|
| 理由 | Excel AI 插件的工作场景不涉及 Git 版本控制，也没有 Codex 那样的 permission profile 系统 |
| 建议 | 永不需要 |

### 3.5 ExtraConfig 分离

**Codex 做法**：`ExtraConfig` 独立类型，不混入 `StoredThread` 元数据。

**价值判断**：⚪ **不值得做**

| 维度 | 评估 |
|------|------|
| 理由 | 元数据字段数量可控（15 字段），不需要额外分离。等字段超过 25 个时再考虑 |
| 建议 | 当前不做 |

---

## 四、压缩（Compaction）

### 4.1 远程压缩

**Codex 做法**：`compact_remote_v2.rs` 把压缩任务交给服务端执行，减少本地 CPU 和 token 消耗。

**价值判断**：⚪ **不值得做**

| 维度 | 评估 |
|------|------|
| 理由 | 当前项目是纯本地桌面应用，没有配套的服务端 |
| 建议 | 永不需要 |

### 4.2 内存守卫（guard limit）

**Codex 做法**：`guard.rs` 限流记忆生成，防止无限循环。`CODEX_LIMIT_ID` 硬上限。

**价值判断**：🟡 **可做可不做**

| 维度 | 评估 |
|------|------|
| 理由 | 如果不上完整记忆管线，就不需要守卫。如果上了，则必须做 |
| 建议 | 跟随记忆管线一起实现 |

---

## 五、汇总

### 🔴 值得做（按优先级排序）

| 优先级 | 特性 | 成本 | 预估工时 |
|-------|------|------|---------|
| P1 | 重试策略（max_retries） | 低 | 0.5d |
| P1 | 分页查询替代全量扫描 | 中 | 1d |
| P2 | comp_hash 变更检测 | 低 | 0.5d |
| P2 | InputQueue 中断输入处理 | 中 | 1.5d |

### 🟡 可做可不做（按推荐排序）

| 推荐 | 特性 | 理由 |
|------|------|------|
| 2 | AI 可写记忆工具（简化版） | 有用户场景但可通过别的方式实现 |
| 3 | AI 自动记忆提取（完整管线） | 成本高，场景不如 Codex 强烈 |
| 4 | TurnContext 快照 | 架构规范但当前无痛点 |
| 5 | ThreadMetadataPatch merge 语义 | 单进程场景不需要 |
| 6 | 内存守卫 | 跟随记忆管线 |
| 7 | Token budget 提醒 | 前端已有数据 |

### ⚪ 不值得做

| 特性 | 原因 |
|------|------|
| ThreadStore trait 抽象 | 没有多后端需求 |
| Git 元数据 | Excel 场景不需要 |
| Permission 快照 | Excel 场景不需要 |
| ExtraConfig 分离 | 当前字段数可控 |
| 远程压缩 | 没有配套服务端 |
| 跨进程文件锁 | 单进程应用 |
| 工作区 diff → 清理记忆 | 当前无记忆管线 |

---

### 总结

**当前项目最值得投入的 4 个改进**：

1. **重试策略** — 0.5 天，可靠性收益大
2. **分页查询** — 1 天，大数据量场景必备
3. **comp_hash 变更检测** — 0.5 天，多模型场景的兼容性保障
4. **InputQueue 中断输入** — 1.5 天，用户体验显著提升

**完整记忆管线（Phase 1 → Phase 2 + AI 工具）目前不建议投入**，因为 Excel AI 助手的会话使用模式与 Codex（长期编码项目）差异较大，投入产出比不高。可以等用户需求明确后再考虑。
