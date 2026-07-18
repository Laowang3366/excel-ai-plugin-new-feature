# Desktop Vitest 测试内容清单（中文版，2026-07-16）

- 基线提交：`5cc5053c`
- 测试文件：204
- 实际执行测试：1101
- 通过：1101
- 失败：0

说明：本清单按测试文件列出中文测试内容概述。文件路径、函数名、协议名等技术标识保留原名；说明文字均使用中文。测试数量来自实际 Vitest JSON 报告。

## 目录汇总

| 目录 | 文件数 | 测试数 |
| --- | ---: | ---: |
| `electron/agent` | 1 | 11 |
| `electron/agent/core/agentLoop` | 33 | 159 |
| `electron/agent/interaction` | 3 | 15 |
| `electron/agent/knowledge` | 7 | 75 |
| `electron/agent/memory` | 10 | 73 |
| `electron/agent/memory/longTerm` | 4 | 24 |
| `electron/agent/officeWorker` | 3 | 10 |
| `electron/agent/prompts` | 4 | 33 |
| `electron/agent/providers` | 5 | 23 |
| `electron/agent/runtime` | 4 | 14 |
| `electron/agent/shared` | 3 | 11 |
| `electron/agent/tools/executors` | 10 | 82 |
| `electron/agent/tools/officeCore` | 10 | 69 |
| `electron/agent/tools/registry` | 15 | 139 |
| `electron/main-modules` | 28 | 86 |
| `electron/main-modules/localDataProtection` | 8 | 29 |
| `electron/shared` | 14 | 54 |
| `src` | 1 | 2 |
| `src/components` | 1 | 2 |
| `src/components/chat` | 5 | 23 |
| `src/components/common` | 1 | 4 |
| `src/components/office` | 1 | 1 |
| `src/components/settings` | 6 | 16 |
| `src/components/sidebar` | 1 | 3 |
| `src/components/task` | 3 | 9 |
| `src/hooks` | 4 | 17 |
| `src/services` | 2 | 12 |
| `src/store` | 6 | 60 |
| `src/utils` | 11 | 45 |

## 文件清单

### 1. `electron/agent/architecture.test.ts`

- 测试数量：11
- 测试内容：验证 Agent 运行时的模块分层、依赖方向、公共入口和 Office 工具边界，防止架构职责重新混杂。

### 2. `electron/agent/core/agentLoop/agentLoop.test.ts`

- 测试数量：27
- 测试内容：覆盖 Agent 主循环的线程执行、流式回复、工具调用、审批、中断、排队、错误处理和回合状态变化。

### 3. `electron/agent/core/agentLoop/agentLoopRunner.test.ts`

- 测试数量：1
- 测试内容：验证主循环执行器能够按约定启动并完成一次 Agent 回合。

### 4. `electron/agent/core/agentLoop/aiRequestRetry.test.ts`

- 测试数量：6
- 测试内容：验证 AI 请求在瞬时错误、重试上限、已产生可见输出和不可重试错误下的处理。

### 5. `electron/agent/core/agentLoop/buildStreamParams.test.ts`

- 测试数量：14
- 测试内容：验证模型流式请求参数的构建，包括消息、工具、系统提示、推理配置和令牌限制。

### 6. `electron/agent/core/agentLoop/compactionProgress.test.ts`

- 测试数量：2
- 测试内容：验证上下文压缩进度事件的生成与状态表达。

### 7. `electron/agent/core/agentLoop/compactionProvider.test.ts`

- 测试数量：6
- 测试内容：验证压缩请求选择正确的模型供应商、配置和降级行为。

### 8. `electron/agent/core/agentLoop/compactionRunner.test.ts`

- 测试数量：4
- 测试内容：验证上下文压缩任务的执行、成功结果和失败处理。

### 9. `electron/agent/core/agentLoop/compactionSummary.test.ts`

- 测试数量：3
- 测试内容：验证压缩摘要的结构、内容提取和空结果处理。

### 10. `electron/agent/core/agentLoop/configUpdates.test.ts`

- 测试数量：2
- 测试内容：验证运行中配置更新能够正确合并，并只影响允许动态调整的字段。

### 11. `electron/agent/core/agentLoop/contextUsage.test.ts`

- 测试数量：2
- 测试内容：验证上下文令牌使用量和占用比例的计算。

### 12. `electron/agent/core/agentLoop/idleThreadUnload.test.ts`

- 测试数量：3
- 测试内容：验证空闲线程的自动卸载、活跃线程保护和卸载时机。

### 13. `electron/agent/core/agentLoop/inputQueue.test.ts`

- 测试数量：3
- 测试内容：验证当前回合运行期间追加输入的入队、读取和清理。

### 14. `electron/agent/core/agentLoop/maxTokens.test.ts`

- 测试数量：3
- 测试内容：验证不同模型与配置下最大输出令牌数的限制和归一化。

### 15. `electron/agent/core/agentLoop/modelCompHash.test.ts`

- 测试数量：4
- 测试内容：验证模型压缩配置哈希的稳定性，以及配置变化时哈希能够变化。

### 16. `electron/agent/core/agentLoop/pendingInterruptQueue.test.ts`

- 测试数量：1
- 测试内容：验证等待执行的中断请求能够按顺序保存并取出。

### 17. `electron/agent/core/agentLoop/preTurnCompaction.test.ts`

- 测试数量：3
- 测试内容：验证回合开始前达到上下文阈值时触发压缩，并正确处理无需压缩的情况。

### 18. `electron/agent/core/agentLoop/queuedTurns.test.ts`

- 测试数量：5
- 测试内容：验证同一线程追加回合的排队顺序、串行执行、失败隔离和队列清理。

### 19. `electron/agent/core/agentLoop/roundStreamParams.test.ts`

- 测试数量：3
- 测试内容：验证每轮模型请求使用的流式参数会结合当前消息、工具结果和配置重新构建。

### 20. `electron/agent/core/agentLoop/streamCollector.test.ts`

- 测试数量：4
- 测试内容：验证流式事件收集器对正文、推理、工具调用和结束状态的聚合。

### 21. `electron/agent/core/agentLoop/streamResultItems.test.ts`

- 测试数量：2
- 测试内容：验证流式结果能够转换为会话中的助手消息和工具调用条目。

### 22. `electron/agent/core/agentLoop/streamRound.test.ts`

- 测试数量：6
- 测试内容：验证单轮流式交互中的事件转发、工具调用识别、结束原因和异常处理。

### 23. `electron/agent/core/agentLoop/summaryGenerator.test.ts`

- 测试数量：2
- 测试内容：验证会话摘要生成请求和摘要结果的处理。

### 24. `electron/agent/core/agentLoop/threadRuntime.test.ts`

- 测试数量：3
- 测试内容：验证线程运行时对象的创建、活动状态、资源持有和释放。

### 25. `electron/agent/core/agentLoop/threadSession.test.ts`

- 测试数量：3
- 测试内容：验证线程会话的加载、消息追加、持久化和状态同步。

### 26. `electron/agent/core/agentLoop/threadStateManager.test.ts`

- 测试数量：5
- 测试内容：验证线程状态管理器对回合、消息、错误、目标和持久化状态的更新。

### 27. `electron/agent/core/agentLoop/threadWatchManager.test.ts`

- 测试数量：2
- 测试内容：验证线程监听器的注册、事件通知和取消监听。

### 28. `electron/agent/core/agentLoop/toolApproval.test.ts`

- 测试数量：7
- 测试内容：验证工具审批规则，包括危险工具、未知工具、授权范围、过期授权和默认拒绝。

### 29. `electron/agent/core/agentLoop/toolExecutionLog.test.ts`

- 测试数量：2
- 测试内容：验证工具执行日志记录的字段、成功失败状态和敏感信息边界。

### 30. `electron/agent/core/agentLoop/toolExecutor.test.ts`

- 测试数量：15
- 测试内容：覆盖工具执行器的参数校验、名称解析、审批、路由、结果封装、错误和取消处理。

### 31. `electron/agent/core/agentLoop/toolNameResolution.test.ts`

- 测试数量：3
- 测试内容：验证模型工具名称与内部工具名称之间的解析、别名兼容和未知名称拒绝。

### 32. `electron/agent/core/agentLoop/toolRound.test.ts`

- 测试数量：3
- 测试内容：验证包含工具调用的一轮交互能够执行工具、回送结构化结果并继续模型请求。

### 33. `electron/agent/core/agentLoop/turnExecution.test.ts`

- 测试数量：3
- 测试内容：验证单个回合的执行入口、运行状态和异常传播。

### 34. `electron/agent/core/agentLoop/turnLifecycle.test.ts`

- 测试数量：7
- 测试内容：验证回合创建、开始、完成、失败、取消以及对应事件和持久化生命周期。

### 35. `electron/agent/interaction/eventForwarder.test.ts`

- 测试数量：3
- 测试内容：验证 Agent 流式事件向渲染进程转发时的顺序、缓冲和清理。

### 36. `electron/agent/interaction/ipcAgentHandlers.test.ts`

- 测试数量：10
- 测试内容：验证 Agent IPC 的输入校验、线程操作、回合启动、中断、排队和错误返回。

### 37. `electron/agent/interaction/ipcKnowledgeHandlers.test.ts`

- 测试数量：2
- 测试内容：验证知识库 IPC 查询与管理请求的校验和路由。

### 38. `electron/agent/knowledge/builtinKnowledge.test.ts`

- 测试数量：3
- 测试内容：验证内置知识资源的注册、读取和缺失资源处理。

### 39. `electron/agent/knowledge/knowledgeIndexer.test.ts`

- 测试数量：2
- 测试内容：验证知识文件索引流程及其写入结果。

### 40. `electron/agent/knowledge/knowledgeWriter.test.ts`

- 测试数量：5
- 测试内容：验证知识条目的新增、更新、来源关联、删除和写入失败处理。

### 41. `electron/agent/knowledge/rag.test.ts`

- 测试数量：51
- 测试内容：全面覆盖 RAG 的文档解析、切块、索引、检索、排序、过滤、引用、去重和异常边界。

### 42. `electron/agent/knowledge/retriever.test.ts`

- 测试数量：7
- 测试内容：验证知识检索器的查询构建、结果排序、数量限制和无结果处理。

### 43. `electron/agent/knowledge/sqliteStoreSearch.test.ts`

- 测试数量：2
- 测试内容：验证 SQLite 知识存储的全文搜索和结果映射。

### 44. `electron/agent/knowledge/textChunker.test.ts`

- 测试数量：5
- 测试内容：验证文本切块的长度限制、重叠、段落边界、空文本和中文内容处理。

### 45. `electron/agent/memory/agentGraphStore.test.ts`

- 测试数量：2
- 测试内容：验证 Agent 图状态的保存、读取和线程隔离。

### 46. `electron/agent/memory/compaction.test.ts`

- 测试数量：26
- 测试内容：覆盖会话压缩的触发条件、保留内容、摘要插入、工具条目处理、令牌预算和失败恢复。

### 47. `electron/agent/memory/longTerm/memoryAutoExtraction.test.ts`

- 测试数量：3
- 测试内容：验证从对话中自动识别可长期保存记忆的条件和过滤规则。

### 48. `electron/agent/memory/longTerm/memoryExtraction.test.ts`

- 测试数量：6
- 测试内容：验证长期记忆候选内容的提取、引用来源和无效内容拒绝。

### 49. `electron/agent/memory/longTerm/memoryStore.test.ts`

- 测试数量：6
- 测试内容：验证长期记忆的写入、查询、更新、删除、线程归属和持久化。

### 50. `electron/agent/memory/longTerm/memoryTypes.test.ts`

- 测试数量：9
- 测试内容：验证长期记忆类型、字段校验、状态转换和序列化边界。

### 51. `electron/agent/memory/rolloutArchive.test.ts`

- 测试数量：4
- 测试内容：验证会话 rollout 归档路径、归档内容、重复归档和读取。

### 52. `electron/agent/memory/rolloutWriter.test.ts`

- 测试数量：5
- 测试内容：验证 rollout 事件以追加方式写入、目录创建和写入失败处理。

### 53. `electron/agent/memory/sessionStore.cipher.test.ts`

- 测试数量：2
- 测试内容：验证会话存储加密后的写入与读取，以及密文不暴露明文。

### 54. `electron/agent/memory/sessionStore.test.ts`

- 测试数量：8
- 测试内容：验证会话存储的创建、读取、追加、恢复、兼容回退和损坏数据处理。

### 55. `electron/agent/memory/stateRuntimeMappers.test.ts`

- 测试数量：4
- 测试内容：验证 SQLite 运行时状态与 Agent 领域对象之间的双向映射。

### 56. `electron/agent/memory/stateRuntimeRolloutEvents.canary.test.ts`

- 测试数量：1
- 测试内容：以金丝雀测试确认关键 rollout 事件仍会写入运行时状态库。

### 57. `electron/agent/memory/stateRuntimeStore.test.ts`

- 测试数量：18
- 测试内容：全面验证运行时 SQLite 存储的线程、回合、消息、目标、日志、事务和恢复行为。

### 58. `electron/agent/memory/threadRepository.test.ts`

- 测试数量：3
- 测试内容：验证线程仓库优先读取 SQLite、必要时回退 JSONL，并保持线程数据一致。

### 59. `electron/agent/officeWorker/dotNetOfficeActionBridge.test.ts`

- 测试数量：4
- 测试内容：验证 Electron 到 .NET Office 动作桥的请求映射、响应校验、错误和协议字段。

### 60. `electron/agent/officeWorker/dotNetOfficeDocumentBridge.test.ts`

- 测试数量：2
- 测试内容：验证 Electron 到 .NET Office 文档桥的文档请求和结果转换。

### 61. `electron/agent/officeWorker/officeWorkerClient.test.ts`

- 测试数量：4
- 测试内容：验证 Office Worker 子进程的启动、协议握手、并发请求、响应匹配、超时和退出处理。

### 62. `electron/agent/prompts/officeToolBoundaryPrompts.test.ts`

- 测试数量：6
- 测试内容：验证系统提示明确限制 Office 工具边界、进程所有权、COM 与 Open XML 使用规则。

### 63. `electron/agent/prompts/promptComposer.test.ts`

- 测试数量：6
- 测试内容：验证系统提示、场景提示、记忆、知识和工具说明的组合顺序与内容边界。

### 64. `electron/agent/prompts/promptRouting.test.ts`

- 测试数量：7
- 测试内容：验证不同任务和应用场景能够选择正确的提示词模块。

### 65. `electron/agent/prompts/systemPrompt.test.ts`

- 测试数量：14
- 测试内容：验证完整系统提示包含安全、工具、数据、记忆、Office 和交互规则，并防止关键约束丢失。

### 66. `electron/agent/providers/openaiCompatibleClient.test.ts`

- 测试数量：7
- 测试内容：验证 OpenAI 兼容接口的请求格式、流式解析、工具调用、错误映射和自定义端点。

### 67. `electron/agent/providers/openaiResponsesClient.test.ts`

- 测试数量：10
- 测试内容：验证 OpenAI Responses 接口的请求、事件流、推理内容、工具调用和异常处理。

### 68. `electron/agent/providers/openaiToolNames.test.ts`

- 测试数量：2
- 测试内容：验证 OpenAI 协议工具名称的编码、长度限制和反向还原。

### 69. `electron/agent/providers/providerClients.test.ts`

- 测试数量：2
- 测试内容：验证不同模型供应商配置会创建正确的客户端实现。

### 70. `electron/agent/providers/providerErrors.test.ts`

- 测试数量：2
- 测试内容：验证供应商错误被统一分类为认证、限流、网络、模型和服务异常。

### 71. `electron/agent/runtime/agentRuntime.test.ts`

- 测试数量：6
- 测试内容：验证 Agent Runtime 的初始化、线程访问、回合调度、服务装配和释放。

### 72. `electron/agent/runtime/bridgeRegistry.test.ts`

- 测试数量：4
- 测试内容：验证各运行时桥接服务的注册、读取、覆盖和清理。

### 73. `electron/agent/runtime/compactionRuntime.test.ts`

- 测试数量：1
- 测试内容：验证运行时能够调用会话压缩能力并返回结果。

### 74. `electron/agent/runtime/knowledgeRuntime.test.ts`

- 测试数量：3
- 测试内容：验证知识运行时的初始化、索引、查询和服务缺失处理。

### 75. `electron/agent/shared/messageBuilder.test.ts`

- 测试数量：5
- 测试内容：验证发送给模型的消息结构构建，包括系统消息、用户消息、附件和工具结果。

### 76. `electron/agent/shared/numberLimits.test.ts`

- 测试数量：2
- 测试内容：验证数字配置的有限值、上下限和默认值归一化。

### 77. `electron/agent/shared/xmlEntities.test.ts`

- 测试数量：4
- 测试内容：验证 XML 实体的转义、反转义和非法实体处理。

### 78. `electron/agent/tools/executors/excelExecutors.test.ts`

- 测试数量：19
- 测试内容：验证 Excel 工具执行器的参数校验、操作路由、Worker 调用、结果转换和错误处理。

### 79. `electron/agent/tools/executors/excelMacroExecutors.test.ts`

- 测试数量：2
- 测试内容：验证 Excel 宏写入与运行工具的审批信息、参数和执行路由。

### 80. `electron/agent/tools/executors/excelUiExecutors.test.ts`

- 测试数量：2
- 测试内容：验证面向当前 Excel 窗口的界面操作执行器及其参数传递。

### 81. `electron/agent/tools/executors/knowledgeExecutors.test.ts`

- 测试数量：5
- 测试内容：验证知识库工具的索引、搜索、写入和错误结果封装。

### 82. `electron/agent/tools/executors/memoryExecutors.test.ts`

- 测试数量：14
- 测试内容：验证记忆读取、写入、确认、引用校验、权限边界和错误处理。

### 83. `electron/agent/tools/executors/ocrExecutorResult.test.ts`

- 测试数量：1
- 测试内容：验证 OCR 执行结果能够转换为模型可用的结构化内容。

### 84. `electron/agent/tools/executors/ocrExecutors.test.ts`

- 测试数量：6
- 测试内容：验证 OCR 工具的文件授权、模式选择、执行路由、结果和失败处理。

### 85. `electron/agent/tools/executors/officeExecutors.test.ts`

- 测试数量：23
- 测试内容：全面验证 Office 工具执行器的操作选择、参数校验、审批、事务、Worker 路由和结构化结果。

### 86. `electron/agent/tools/executors/webSearchExecutors.test.ts`

- 测试数量：7
- 测试内容：验证 Web 搜索工具的查询校验、供应商调用、结果整理、引用和错误处理。

### 87. `electron/agent/tools/executors/webSearchProviders.test.ts`

- 测试数量：3
- 测试内容：验证 Web 搜索供应商的选择、配置缺失和不支持供应商处理。

### 88. `electron/agent/tools/officeCore/capabilities.test.ts`

- 测试数量：5
- 测试内容：验证 Office 应用、动作和操作能力表的查询与支持范围。

### 89. `electron/agent/tools/officeCore/locator.test.ts`

- 测试数量：1
- 测试内容：验证 Office 对象定位信息的标准化。

### 90. `electron/agent/tools/officeCore/officeActionAdapter.test.ts`

- 测试数量：19
- 测试内容：验证模型可见 Office 动作到内部原语的转换、参数映射、能力判断和结果处理。

### 91. `electron/agent/tools/officeCore/officeActionTransactionAdapter.test.ts`

- 测试数量：2
- 测试内容：验证 Office 动作事务适配器能够在事务上下文中执行和记录操作。

### 92. `electron/agent/tools/officeCore/operationPolicy.test.ts`

- 测试数量：13
- 测试内容：验证 Office 操作的风险等级、审批要求、备份、事务和宿主限制策略。

### 93. `electron/agent/tools/officeCore/transactionJournal.test.ts`

- 测试数量：6
- 测试内容：验证事务日志的创建、步骤记录、提交、回滚和持久化。

### 94. `electron/agent/tools/officeCore/transactions.test.ts`

- 测试数量：6
- 测试内容：验证 Office 事务的开始、执行、提交、回滚、冲突和恢复。

### 95. `electron/agent/tools/officeCore/workflow.test.ts`

- 测试数量：9
- 测试内容：验证 Office 工作流的创建、变量解析、步骤执行、暂停恢复和失败处理。

### 96. `electron/agent/tools/officeCore/workflowStepExecution.test.ts`

- 测试数量：6
- 测试内容：验证单个工作流步骤的变量取值、工具执行、结果写回和错误传播。

### 97. `electron/agent/tools/officeCore/workflowTemplates.test.ts`

- 测试数量：2
- 测试内容：验证内置工作流模板的读取和实例化。

### 98. `electron/agent/tools/registry/officeChartParamSchemas.test.ts`

- 测试数量：3
- 测试内容：验证 Office 图表参数 Schema 对合法配置的接受和非法字段的拒绝。

### 99. `electron/agent/tools/registry/officeCrossOfficeParamSchemas.test.ts`

- 测试数量：5
- 测试内容：验证跨 Word、Excel、PowerPoint 操作参数的严格 Schema。

### 100. `electron/agent/tools/registry/officeExcelFormulaParamSchemas.test.ts`

- 测试数量：4
- 测试内容：验证 Excel 公式相关参数、动态数组设置和非法输入拒绝。

### 101. `electron/agent/tools/registry/officeExcelObjectParamSchemas.test.ts`

- 测试数量：4
- 测试内容：验证 Excel 工作表、区域、表格和对象操作参数 Schema。

### 102. `electron/agent/tools/registry/officeExcelPrintParamSchemas.test.ts`

- 测试数量：3
- 测试内容：验证 Excel 页面设置和打印参数 Schema。

### 103. `electron/agent/tools/registry/officeExcelTemplateParamSchemas.test.ts`

- 测试数量：3
- 测试内容：验证 Excel 模板操作参数和变量结构 Schema。

### 104. `electron/agent/tools/registry/officeExportParamSchemas.test.ts`

- 测试数量：3
- 测试内容：验证 Office 导出格式、目标路径和选项参数 Schema。

### 105. `electron/agent/tools/registry/officePresentationBrandingParamSchemas.test.ts`

- 测试数量：6
- 测试内容：验证 PowerPoint 品牌规范、主题和版式参数 Schema。

### 106. `electron/agent/tools/registry/officePresentationPlaybackParamSchemas.test.ts`

- 测试数量：6
- 测试内容：验证 PowerPoint 播放、切换和演示设置参数 Schema。

### 107. `electron/agent/tools/registry/officeTools.test.ts`

- 测试数量：8
- 测试内容：验证 Office 工具定义、动作枚举、风险元数据和注册完整性。

### 108. `electron/agent/tools/registry/officeToolVisibility.test.ts`

- 测试数量：7
- 测试内容：验证允许模型调用的 Office 工具可见性，并防止内部工具意外暴露。

### 109. `electron/agent/tools/registry/officeWordFormattingParamSchemas.test.ts`

- 测试数量：5
- 测试内容：验证 Word 字体、段落、样式、表格和页面格式参数 Schema。

### 110. `electron/agent/tools/registry/officeWordReviewParamSchemas.test.ts`

- 测试数量：6
- 测试内容：验证 Word 批注、修订、比较和审阅操作参数 Schema。

### 111. `electron/agent/tools/registry/officeWordTemplateParamSchemas.test.ts`

- 测试数量：7
- 测试内容：验证 Word 模板、字段、目录和内容控件参数 Schema。

### 112. `electron/agent/tools/registry/toolSchema.test.ts`

- 测试数量：69
- 测试内容：全面枚举模型可见工具，验证 JSON Schema 严格性、必填字段、附加属性拒绝和审批元数据一致性。

### 113. `electron/main-modules/appShutdown.test.ts`

- 测试数量：3
- 测试内容：验证应用退出时按顺序停止 Agent、Worker、后台资源，并避免重复关闭。

### 114. `electron/main-modules/asyncResource.test.ts`

- 测试数量：3
- 测试内容：验证异步资源的创建、复用、释放和异常清理。

### 115. `electron/main-modules/dataMaintenance.test.ts`

- 测试数量：3
- 测试内容：验证本地数据维护任务的调度、执行和失败记录。

### 116. `electron/main-modules/dataPathMigration.test.ts`

- 测试数量：4
- 测试内容：验证数据目录迁移的复制范围、目标校验、失败回滚和状态更新。

### 117. `electron/main-modules/excelIpcOperations.test.ts`

- 测试数量：2
- 测试内容：验证 Excel IPC 操作的参数校验和主进程路由。

### 118. `electron/main-modules/hotPatchArchive.test.ts`

- 测试数量：2
- 测试内容：验证热补丁压缩包的路径、文件类型和目录穿越防护。

### 119. `electron/main-modules/hotPatchManager.test.ts`

- 测试数量：9
- 测试内容：验证热补丁清单校验、签名、哈希、允许路径、安装和失败恢复。

### 120. `electron/main-modules/invoiceFieldExtraction.test.ts`

- 测试数量：1
- 测试内容：验证发票 OCR 文本到结构化字段的提取。

### 121. `electron/main-modules/ipcFileHandlers.test.ts`

- 测试数量：3
- 测试内容：验证文件 IPC 的路径授权、读取写入路由和非法路径拒绝。

### 122. `electron/main-modules/ipcHandlers.ocr.test.ts`

- 测试数量：2
- 测试内容：验证 OCR IPC 注册入口和请求转发。

### 123. `electron/main-modules/ipcOcrHandlers.test.ts`

- 测试数量：2
- 测试内容：验证 OCR IPC 的输入校验、文件授权、任务执行和错误返回。

### 124. `electron/main-modules/ipcOfficeAutomationHandlers.test.ts`

- 测试数量：3
- 测试内容：验证 Office 自动化 IPC 的工作流、事务、模板和文档操作路由。

### 125. `electron/main-modules/ipcPathSecurity.test.ts`

- 测试数量：3
- 测试内容：验证文件路径授权、防目录穿越、防链接逃逸和授权根边界。

### 126. `electron/main-modules/localDataMaintenance.test.ts`

- 测试数量：3
- 测试内容：验证本地数据定期清理的保留期、目标范围和失败处理。

### 127. `electron/main-modules/localDataProtection/aesGcm.test.ts`

- 测试数量：2
- 测试内容：验证 AES-GCM 加密解密、随机数、认证标签和篡改检测。

### 128. `electron/main-modules/localDataProtection/archiveValidation.test.ts`

- 测试数量：2
- 测试内容：验证本地数据归档的文件清单、路径边界和损坏归档拒绝。

### 129. `electron/main-modules/localDataProtection/dataKeystore.test.ts`

- 测试数量：5
- 测试内容：验证本地数据密钥的创建、加密保存、读取、轮换和不可用处理。

### 130. `electron/main-modules/localDataProtection/lifecycleRegistry.test.ts`

- 测试数量：3
- 测试内容：验证受保护数据生命周期登记、查询、更新和清理。

### 131. `electron/main-modules/localDataProtection/localDataEraseAll.test.ts`

- 测试数量：5
- 测试内容：验证彻底删除本地数据时覆盖所有登记位置，并正确报告部分失败。

### 132. `electron/main-modules/localDataProtection/localDataMigrator.test.ts`

- 测试数量：1
- 测试内容：验证旧版明文或旧密钥数据迁移到当前保护格式。

### 133. `electron/main-modules/localDataProtection/recoveryJournal.test.ts`

- 测试数量：9
- 测试内容：验证数据恢复日志的步骤记录、断点恢复、提交和异常状态。

### 134. `electron/main-modules/localDataProtection/rotationFailure.test.ts`

- 测试数量：2
- 测试内容：验证密钥轮换失败时保留旧数据可读性并记录恢复信息。

### 135. `electron/main-modules/mineruOcr.test.ts`

- 测试数量：4
- 测试内容：验证 MinerU OCR API、Agent 端点、响应解析、回退和错误处理。

### 136. `electron/main-modules/ocrDocumentResultBuilder.test.ts`

- 测试数量：4
- 测试内容：验证 OCR 文档结果的文本、Markdown、页信息和元数据组装。

### 137. `electron/main-modules/officeProcessLauncher.test.ts`

- 测试数量：3
- 测试内容：验证 Office 进程启动时只管理任务拥有的进程，并正确处理启动失败。

### 138. `electron/main-modules/settingRuntimeEffects.test.ts`

- 测试数量：3
- 测试内容：验证设置变更对运行时服务、日志、自动清理和模型配置的即时影响。

### 139. `electron/main-modules/settingsDataPath.test.ts`

- 测试数量：4
- 测试内容：验证数据目录设置的路径校验、规范化和默认目录行为。

### 140. `electron/main-modules/settingsManager.migrateDataPath.test.ts`

- 测试数量：1
- 测试内容：验证设置管理器迁移数据目录时调用正确的迁移流程。

### 141. `electron/main-modules/settingsSecrets.test.ts`

- 测试数量：4
- 测试内容：验证供应商、OCR 和远程服务密钥的安全存储、掩码、迁移与删除。

### 142. `electron/main-modules/updateManifest.test.ts`

- 测试数量：3
- 测试内容：验证更新清单的签名、版本、文件大小、SHA-256 和资源字段。

### 143. `electron/main-modules/updateUrl.test.ts`

- 测试数量：2
- 测试内容：验证更新地址只接受允许的 HTTPS 来源和受控本地地址。

### 144. `electron/main-modules/userDataErase.test.ts`

- 测试数量：4
- 测试内容：验证用户数据删除覆盖规定的数据类型，并返回删除证明和失败详情。

### 145. `electron/main-modules/userDataEraseCoordinator.test.ts`

- 测试数量：4
- 测试内容：验证用户数据删除协调器的执行顺序、互斥、进度和部分失败处理。

### 146. `electron/main-modules/userDataExport.test.ts`

- 测试数量：2
- 测试内容：验证用户数据导出的文件范围、结构、脱敏和归档结果。

### 147. `electron/main-modules/userDataExportCoordinator.test.ts`

- 测试数量：3
- 测试内容：验证用户数据导出协调器的互斥、进度、结果和失败清理。

### 148. `electron/main-modules/windowManager.test.ts`

- 测试数量：2
- 测试内容：验证主窗口创建、状态恢复、导航限制和窗口生命周期。

### 149. `electron/shared/currentDocumentation.test.ts`

- 测试数量：5
- 测试内容：验证当前文档索引只引用现行文档，并识别历史或失效文档引用。

### 150. `electron/shared/egressPolicy.test.ts`

- 测试数量：4
- 测试内容：验证外发数据策略对目标、数据类型、用户设置和禁止场景的判断。

### 151. `electron/shared/ipcRateLimiter.test.ts`

- 测试数量：2
- 测试内容：验证 IPC 调用频率限制、时间窗口、不同通道隔离和恢复。

### 152. `electron/shared/ipcSchemas.test.ts`

- 测试数量：9
- 测试内容：验证各 IPC 通道的 Zod 输入 Schema 接受合法输入并拒绝未知或危险字段。

### 153. `electron/shared/jsonResourceBudget.test.ts`

- 测试数量：3
- 测试内容：验证 JSON 资源的深度、节点数、字符串长度和总大小预算。

### 154. `electron/shared/logger.test.ts`

- 测试数量：2
- 测试内容：验证主进程日志级别、结构化字段和敏感信息清理。

### 155. `electron/shared/markdownTables.test.ts`

- 测试数量：1
- 测试内容：验证 Markdown 表格识别与规范化。

### 156. `electron/shared/outboundUrlPolicy.test.ts`

- 测试数量：5
- 测试内容：验证出站 URL 必须使用允许协议和主机，并阻断私网、重定向和 DNS 重绑定。

### 157. `electron/shared/rendererBundleBudget.test.ts`

- 测试数量：4
- 测试内容：验证渲染进程构建产物的入口大小和分包预算。

### 158. `electron/shared/repositoryGovernance.test.ts`

- 测试数量：4
- 测试内容：验证仓库所需文档、工作流、所有者和治理文件保持一致。

### 159. `electron/shared/sensitiveData.test.ts`

- 测试数量：5
- 测试内容：验证日志和错误信息中的 API Key、令牌、密码及其他敏感数据脱敏。

### 160. `electron/shared/sourceGovernance.test.ts`

- 测试数量：5
- 测试内容：验证源文件行数、格式、旧文件和禁止模式等代码治理规则。

### 161. `electron/shared/trustedIpc.test.ts`

- 测试数量：3
- 测试内容：验证敏感 IPC 只接受可信窗口、主 frame 和允许来源。

### 162. `electron/shared/workflowProvenance.test.ts`

- 测试数量：2
- 测试内容：验证持久化工作流的来源、版本和完整性信息。

### 163. `src/components/AppTitlebar.test.ts`

- 测试数量：2
- 测试内容：验证标题栏透明度显示、范围计算和可访问性属性。

### 164. `src/components/chat/ChatMessageList.test.ts`

- 测试数量：6
- 测试内容：验证聊天消息列表的消息渲染、流式状态、工具内容和滚动行为。

### 165. `src/components/chat/ComposerArea.test.ts`

- 测试数量：7
- 测试内容：验证输入区的文本、附件、发送状态、快捷操作和禁用条件。

### 166. `src/components/chat/ComposerThinkingModeButton.test.ts`

- 测试数量：3
- 测试内容：验证思考模式按钮的选项、当前状态和交互。

### 167. `src/components/chat/MarkdownContent.test.ts`

- 测试数量：4
- 测试内容：验证 Markdown 内容的安全渲染、链接处理、代码块和表格。

### 168. `src/components/chat/StreamingOutput.test.ts`

- 测试数量：3
- 测试内容：验证流式正文和推理内容的增量显示与状态切换。

### 169. `src/components/common/FeatureSidebarPanel.test.ts`

- 测试数量：4
- 测试内容：验证功能侧栏的打开状态、当前功能、无障碍属性和非活动内容隔离。

### 170. `src/components/office/OfficeAutomationPanel.test.ts`

- 测试数量：1
- 测试内容：验证 Office 自动化视图模型的模板变量解析、应用名称和路径缩写。

### 171. `src/components/settings/addProviderDraft.test.ts`

- 测试数量：3
- 测试内容：验证新增模型供应商草稿的默认值、类型切换和字段清理。

### 172. `src/components/settings/editProviderPatch.test.ts`

- 测试数量：2
- 测试内容：验证编辑模型供应商时只生成实际变化的补丁字段。

### 173. `src/components/settings/GeneralSettingsStorageCard.test.ts`

- 测试数量：1
- 测试内容：验证通用设置存储卡的状态显示和主要操作。

### 174. `src/components/settings/knowledgeSettingsText.test.ts`

- 测试数量：4
- 测试内容：验证知识库设置中的数量、索引结果和来源类型中文本格式化。

### 175. `src/components/settings/ProviderModelSelector.test.ts`

- 测试数量：3
- 测试内容：验证供应商模型选择器的模型列表、当前值、加载和自定义模型处理。

### 176. `src/components/settings/usageStatsData.test.ts`

- 测试数量：3
- 测试内容：验证用量统计数据的聚合、时间范围和图表数据转换。

### 177. `src/components/sidebar/SidebarThreadItem.test.ts`

- 测试数量：3
- 测试内容：验证侧栏线程项的标题、选中状态、重命名和菜单操作。

### 178. `src/components/task/OCRTaskComposerPanel.test.ts`

- 测试数量：4
- 测试内容：验证 OCR 任务面板的文件选择、模式、提交参数和禁用状态。

### 179. `src/components/task/ocrTaskFileHelpers.test.ts`

- 测试数量：3
- 测试内容：验证 OCR 任务附件的文件类型判断、去重和限制。

### 180. `src/components/task/SimpleTaskComposerPanel.test.ts`

- 测试数量：2
- 测试内容：验证通用任务面板的输入、提交和嵌入模式行为。

### 181. `src/hooks/composerAttachmentFiles.test.ts`

- 测试数量：3
- 测试内容：验证输入区附件文件的添加、去重、删除和大小限制。

### 182. `src/hooks/useComposer.test.ts`

- 测试数量：3
- 测试内容：验证输入框 Hook 的草稿、附件、发送、清空和错误处理。

### 183. `src/hooks/useDocumentDismiss.test.ts`

- 测试数量：4
- 测试内容：验证点击外部区域或按键时关闭文档浮层，并正确清理监听器。

### 184. `src/hooks/useTaskDrafts.test.ts`

- 测试数量：7
- 测试内容：验证不同任务类型草稿的保存、恢复、更新、清除和线程隔离。

### 185. `src/i18n.test.ts`

- 测试数量：2
- 测试内容：验证中英文资源的约束选项一致，以及时间格式和未知语言回退。

### 186. `src/services/ipcApi.test.ts`

- 测试数量：10
- 测试内容：验证渲染进程 IPC API 的域包装、参数转发、错误和事件订阅。

### 187. `src/services/ipcThreadApi.test.ts`

- 测试数量：2
- 测试内容：验证线程相关 IPC API 的线程查询和操作封装。

### 188. `src/store/agentEventHandler.test.ts`

- 测试数量：15
- 测试内容：验证 Agent 事件到聊天状态的投影，包括回合、消息、流式内容、工具和错误。

### 189. `src/store/chatStore.test.ts`

- 测试数量：25
- 测试内容：全面验证聊天 Store 的线程加载、切换、消息、回合、流式状态、持久化和并发更新。

### 190. `src/store/chatThreadRuntimeState.test.ts`

- 测试数量：3
- 测试内容：验证每个线程独立保存运行中、排队、中断和流式状态。

### 191. `src/store/chatTurnState.test.ts`

- 测试数量：2
- 测试内容：验证聊天回合状态的创建、更新、完成和错误状态。

### 192. `src/store/settingsStore.test.ts`

- 测试数量：8
- 测试内容：验证设置 Store 的加载、更新、供应商配置、密钥掩码和持久化。

### 193. `src/store/threadActions.test.ts`

- 测试数量：7
- 测试内容：验证线程的新建、选择、重命名、删除、归档和异常处理。

### 194. `src/utils/attachmentPreview.test.ts`

- 测试数量：3
- 测试内容：验证附件预览类型、可预览条件和预览信息生成。

### 195. `src/utils/chatHelpers.test.ts`

- 测试数量：9
- 测试内容：验证聊天消息和工具结果的格式化、分组、状态判断和显示辅助逻辑。

### 196. `src/utils/featureSidebarState.test.ts`

- 测试数量：6
- 测试内容：验证功能侧栏状态的打开关闭、功能切换、宿主限制和持久化。

### 197. `src/utils/fileBase64.test.ts`

- 测试数量：2
- 测试内容：验证文件与 Base64 之间的转换。

### 198. `src/utils/fileSize.test.ts`

- 测试数量：2
- 测试内容：验证文件大小的人类可读格式化和边界值。

### 199. `src/utils/modelContextWindows.test.ts`

- 测试数量：2
- 测试内容：验证不同模型的上下文窗口识别和默认值。

### 200. `src/utils/reasoningSupport.test.ts`

- 测试数量：6
- 测试内容：验证模型推理能力、思考等级、自动适配提示和供应商差异。

### 201. `src/utils/sidebarHelpers.test.ts`

- 测试数量：5
- 测试内容：验证侧栏线程排序、标题生成、时间分组和显示辅助逻辑。

### 202. `src/utils/sidebarSearch.test.ts`

- 测试数量：2
- 测试内容：验证侧栏搜索对线程、文件和操作建议的匹配。

### 203. `src/utils/taskComposerPayloads.test.ts`

- 测试数量：4
- 测试内容：验证公式、代码、报告等任务提交载荷的构建和宿主名称归一化。

### 204. `src/utils/textCleaner.test.ts`

- 测试数量：4
- 测试内容：验证流式文本清理中的中英文空格、单元格引用、标点、Markdown 表格和空行处理。

## 核对结果

- 文件条目：204 / 204
- 测试数量：1101 / 1101 通过
- 中文概述：204 / 204

