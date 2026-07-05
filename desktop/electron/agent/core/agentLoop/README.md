# agentLoop 核心循环模块

本目录负责模型会话的运行主循环，只处理 Agent 内部状态、上下文构建、流式收集、工具调用编排和线程运行态管理。

## 模块职责

- `agentLoop.ts`：Agent 主循环入口，协调线程、Turn、模型流式响应、工具调用和自动压缩。
- `threadLifecycle.ts`：线程创建与恢复，只负责调用 `memory/sessionStore` 读写持久化会话。
- `threadSession.ts`：编排重置、新建、恢复和空闲卸载线程，保持 `AgentLoop` 公共 API 不变。
- `threadStateManager.ts`：线程运行态观察与空闲卸载判定，只记录内存态，不直接读写 rollout。
- `threadRuntime.ts`：线程回调绑定、线程快照/运行态持久化、rollout 事件 sink 绑定和 Turn 完成后的长期记忆抽取调度。
- `turnExecution.ts`：编排单个 Turn 的运行开始、活跃线程准备、用户消息落库、成功完成、失败记录和最终运行态收尾。
- `queuedTurns.ts`：处理运行中补充输入入队、用户中断等待和队列自动续跑。
- `configUpdates.ts`：处理 AI/压缩配置热更新后的客户端重建、线程 metadata 更新和待压缩原因合并。
- `turnState.ts`：当前运行 Turn 的可变状态容器。
- `turnRunner.ts`：创建和完成 Turn、用户消息等基础对象。
- `streamCollector.ts`：收集模型流式事件并转换为 Agent 内部事件。
- `streamRound.ts`：封装单轮模型采样请求、流式错误 item 落库和最终 token usage 合并。
- `streamResultItems.ts`：在模型流结束后，按 reasoning、assistant message、tool call 的真实顺序补齐 `TurnItem` 落库与事件发送。
- `toolExecutor.ts`：执行模型请求的工具调用，并处理审批与工具结果回写。
- `toolRound.ts`：封装单轮工具调用后的执行、工具日志写入和 mid-turn 压缩触发判断。
- `buildStreamParams.ts`：构建模型请求参数、系统提示词和推理配置。
- `roundStreamParams.ts`：为每轮模型调用装配消息历史、图片附件、恢复上下文、动态系统提示词、工具定义和流式请求参数。
- `contextUsage.ts`：收集用于提示词的历史条目，并生成上下文使用量事件。
- `compactionProgress.ts`：封装上下文压缩进度事件、压缩参数 rollout 记录和冷 rollout 归档触发。
- `compactionRunner.ts`：执行 pre-turn / mid-turn 上下文压缩，负责摘要写入、历史替换、当前用户消息保留和压缩完成事件。
- `compactionSummary.ts`：生成压缩摘要并集中处理 compact 请求的重试配置。
- `summaryGenerator.ts`：会话标题/摘要生成。
- `sessionCompactionConfig.ts`：根据会话上下文窗口生成压缩配置。
- `maxTokens.ts`：输出 token 预算计算。

## 压缩可观察性

自动压缩开始时，`agentLoop.ts` 会发出 `compact_progress` 的 `item_started`；压缩完成或失败时，发出同一条目的 `item_completed`。压缩参数同时通过 `compact_params` 写入 rollout，最终摘要继续通过 `compacted` 写入，用于恢复上下文。

压缩策略来自 `runtime/compactionRuntime.ts` 装配的 `CompactionConfig`：支持阈值百分比、mid-turn 触发比例、最近消息保留数量、摘要失败重试次数和 rollout gzip 归档阈值。归档只生成 `.jsonl.gz` 快照，活跃 JSONL 仍由 `SessionStore` 负责恢复。

## 线程卸载约定

空闲卸载只释放 `AgentLoop` 内存中的 `activeThread`，不会删除 `SessionStore` 中的 rollout 文件。前端继续向当前会话发送消息前，会通过 `thread:resume` 重新加载线程，因此用户视角仍是同一个会话。

线程创建、恢复、运行、完成和卸载时，`agentLoop.ts` 会把线程快照与运行态写入 `memory/stateRuntimeStore.ts`。当前运行态查询以 SQLite 为主，rollout JSONL 保留为兼容审计副本和归档来源。
