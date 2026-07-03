# 模型供应商层

职责：封装不同模型供应商协议，向核心循环提供统一的流式和非流式模型客户端接口。

模块说明：

- `aiClient.ts`: 统一导出模型客户端入口和兼容类型。
- `aiClientTypes.ts`: 模型请求、响应、消息和流式事件类型。
- `aiClientFactory.ts`: 根据供应商配置创建具体客户端。
- `openaiCompatibleClient.ts`: OpenAI 兼容协议实现。
- `anthropicClient.ts`: Anthropic Messages 协议实现。
- `providerClients.ts`: 供应商适配和配置归一化。
- `modelContextWindows.ts`: 常见模型上下文窗口配置。

关联模块：

- `../core/agentLoop`: 通过统一客户端发起模型请求。
- `../runtime/agentRuntime.ts`: 根据用户配置创建 Agent 运行期。
