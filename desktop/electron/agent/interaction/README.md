# 交互层

职责：承接 Electron/UI 到 Agent 的交互适配，包括 IPC、工具审批回调、流式事件转发和前端可见事件协议。

模块说明：

- `eventForwarder.ts`: Agent 事件、流式增量和工具审批请求转发到渲染进程。
- `ipcAgentHandlers.ts`: 注册 Agent 会话、线程、线程拓扑图、工具定义、统计和知识库相关 IPC。

关联模块：

- `../core`: AgentLoop 提供运行、恢复、中断和线程能力。
- `../memory`: SessionStore 提供线程列表、元数据和统计，AgentGraphStore 提供线程父子关系存储。
- `../knowledge`: 知识库检索与索引能力。
- `../../main-modules/ipcHandlers.ts`: Electron 通用 IPC 注册入口，调用本层注册 Agent IPC。
