# agentLoop 核心循环模块

本目录负责模型会话的运行主循环，只处理 Agent 内部状态、上下文构建、流式收集、工具调用编排和线程运行态管理。

## 模块职责

- `agentLoop.ts`：Agent 主循环入口，协调线程、Turn、模型流式响应、工具调用和自动压缩。
- `threadLifecycle.ts`：线程创建与恢复，只负责调用 `memory/sessionStore` 读写持久化会话。
- `threadStateManager.ts`：线程运行态观察与空闲卸载判定，只记录内存态，不直接读写 rollout。
- `turnState.ts`：当前运行 Turn 的可变状态容器。
- `turnRunner.ts`：创建和完成 Turn、用户消息等基础对象。
- `streamCollector.ts`：收集模型流式事件并转换为 Agent 内部事件。
- `streamResultItems.ts`：在模型流结束后，按 reasoning、assistant message、tool call 的真实顺序补齐 `TurnItem` 落库与事件发送。
- `toolExecutor.ts`：执行模型请求的工具调用，并处理审批与工具结果回写。
- `buildStreamParams.ts`：构建模型请求参数、系统提示词和推理配置。
- `contextUsage.ts`：收集用于提示词的历史条目，并生成上下文使用量事件。
- `compactionProgress.ts`：封装上下文压缩进度事件、压缩参数 rollout 记录和冷 rollout 归档触发。
- `summaryGenerator.ts`：会话标题/摘要生成。
- `sessionCompactionConfig.ts`：根据会话上下文窗口生成压缩配置。
- `maxTokens.ts`：输出 token 预算计算。

## 压缩可观察性

自动压缩开始时，`agentLoop.ts` 会发出 `compact_progress` 的 `item_started`；压缩完成或失败时，发出同一条目的 `item_completed`。压缩参数同时通过 `compact_params` 写入 rollout，最终摘要继续通过 `compacted` 写入，用于恢复上下文。

压缩策略来自 `runtime/compactionRuntime.ts` 装配的 `CompactionConfig`：支持阈值百分比、mid-turn 触发比例、最近消息保留数量、摘要失败重试次数和 rollout gzip 归档阈值。归档只生成 `.jsonl.gz` 快照，活跃 JSONL 仍由 `SessionStore` 负责恢复。

## 线程卸载约定

空闲卸载只释放 `AgentLoop` 内存中的 `activeThread`，不会删除 `SessionStore` 中的 rollout 文件。前端继续向当前会话发送消息前，会通过 `thread:resume` 重新加载线程，因此用户视角仍是同一个会话。

线程创建、恢复、运行、完成和卸载时，`agentLoop.ts` 会把线程快照与运行态写入 `memory/stateRuntimeStore.ts`。当前运行态查询以 SQLite 为主，rollout JSONL 保留为兼容审计副本和归档来源。
