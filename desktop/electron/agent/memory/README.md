# 记忆层

职责：管理会话持久化、上下文压缩、线程恢复和 token 估算，保证 Agent 可以在多轮对话中恢复状态。

模块说明：

- `sessionStore.ts`: 会话、轮次、rollout item 和元数据的兼容 JSONL 审计副本，同时把事件投影写入 `logs.db`。
- `agentGraphStore.ts`: 线程拓扑图存储，记录父线程/子线程关系和边状态。
- `stateRuntimeStore.ts`: StateRuntime 四库门面，使用 `better-sqlite3` 管理 `state.db`、`logs.db`、`goals.db`、`memories.db`。
- `stateRuntimeSchema.ts`: 四库 schema、迁移记录和 WAL 配置。
- `stateRuntimePaths.ts`: 四库路径解析，以及旧 `state-runtime.db` 到 `state-runtime/state.db` 的兼容迁移。
- `stateRuntimeTypes.ts`: StateRuntime 共享类型。
- `stateRuntimeMappers.ts`: SQLite 行到领域对象的映射。
- `rolloutWriter.ts`: rollout 写入队列，按文件批量追加，并在读取/退出前 flush。
- `rolloutArchive.ts`: 为过大的 rollout JSONL 生成 gzip 归档快照，不删除活跃 JSONL。
- `compaction.ts`: 上下文压缩纯函数、摘要上下文构建、token 粗估和中断恢复上下文。

Rollout 中的 `compact_params` 用于审计压缩触发原因、压缩前后 token 和失败信息；`compacted` 用于恢复压缩后的上下文摘要。

压缩摘要生成不在 memory 层执行，由 `../core/agentLoop/compactionProvider.ts` 选择本地模型或远程压缩服务。

SQLite 是当前运行态主存储：

- `state.db`: 线程元数据快照、活跃状态、卸载状态。
- `logs.db`: rollout 事件流，供后续会话搜索、回放和审计查询使用。
- `goals.db`: 目标状态、预算和完成状态。
- `memories.db`: 长期记忆条目和命名空间索引。

JSONL 仍保留为兼容审计副本，避免旧会话恢复链路一次性硬切；会话列表、工具日志和全文搜索优先使用 SQLite 派生索引。gzip 归档是字节级快照，用于降低历史文件留存成本，不参与当前会话恢复主路径。

长期记忆边界：

- 用户可见记忆只记录办公场景有价值的信息：偏好、规则约束、纠正反馈、过往编辑文件印象。
- 内部工具画像记录在受控命名空间，用于提高工具选择成功率，不直接暴露为用户偏好提示。
- 记忆写入必须带来源、置信度和可见性，避免把一次性任务内容污染成长期偏好。

关联模块：

- `../core/agentLoop`: 运行时读取历史、写入 turn item，并在阈值触发时调用压缩能力。
- `../shared/types.ts`: 定义 Thread、Turn、TurnItem 和压缩配置类型。
