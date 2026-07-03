# 运行时装配层

职责：创建并组装 Agent 运行所需依赖，包括模型客户端、工具执行器、记忆存储、知识库、提示词和核心循环。

模块说明：

- `agentRuntime.ts`: Agent 总装配入口，关联桥接、知识库、工具执行器和核心循环。
- `bridgeRegistry.ts`: Office/Excel 桥接实例注册表，供主进程和 IPC 共享。
- `knowledgeRuntime.ts`: RAG 知识库初始化，负责注册全局检索、索引、存储实例。
- `compactionRuntime.ts`: 上下文压缩配置构建，包含阈值、重试、归档和远程压缩服务配置，供 Agent 初始化和设置更新复用。

关联模块：

- `../core`: AgentLoop 核心循环。
- `../tools`: 工具定义和执行器。
- `../knowledge`: RAG 知识层。
- `../../main.ts`: Electron 生命周期入口，仅调用本层完成 Agent 初始化。
