# 知识层

职责：提供本地知识库的文档解析、切块、embedding、索引、检索和持久化能力。

模块说明：

- `documentParser.ts`: 解析文档内容并生成可索引文本。
- `textChunker.ts`: 将文档文本切分为检索块。
- `embeddingService.ts`: 调用 embedding 服务生成向量。
- `sqliteStore.ts`: 存储知识库、文档、切块和向量数据。
- `knowledgeIndexer.ts`: 编排解析、切块、embedding 和入库流程。
- `knowledgeWriter.ts`: 将模型明确沉淀的 note 写入知识库并生成可检索索引。
- `retriever.ts`: 执行检索并格式化返回结果。
- `knowledgeRegistry.ts`: 保存当前运行期知识库检索依赖。
- `workbookNotesStore.ts`: 存储与工作簿相关的知识备注。
- `types.ts`: 知识库内部类型定义。

关联模块：

- `../runtime/knowledgeRuntime.ts`: 初始化和注册知识库运行期实例。
- `../tools/executors/knowledgeExecutors.ts`: 将知识库能力暴露为模型可调用工具。
- `../interaction/ipcAgentHandlers.ts`: 注册知识库管理相关 IPC。
