# AI Gateway（同源 AI 代理）

独立 Node.js 22 服务：把浏览器/同源前端的 AI 请求代理到**服务端显式配置**的上游，避免在客户端暴露 API Key，并阻断开放代理 / SSRF。

本目录是仓库内的独立项目（与 `desktop/`、`product-site/`、`excel-addin/` 平级），**不**修改加载项或桌面运行时代码。

## HTTP 合同

| 方法 | 路径 | 说明 |
|------|------|------|
| `GET` | `/healthz` | 健康检查 |
| `GET` | `/api/ai/v1/:upstreamId/models` | 列出模型（无 body） |
| `POST` | `/api/ai/v1/:upstreamId/chat/completions` | OpenAI 兼容 chat |
| `POST` | `/api/ai/v1/:upstreamId/responses` | OpenAI Responses |
| `POST` | `/api/ai/v1/:upstreamId/messages` | Anthropic Messages |

- `upstreamId` **只能**来自环境变量 `AI_GATEWAY_UPSTREAMS_JSON` 的键；客户端**不能**传任意 `baseUrl` / URL。
- 上游 URL 在启动时校验：绝对 `https`（测试可用 loopback `http`/`https`）、禁止 credentials / fragment / query、默认拒绝私网目标；`AI_GATEWAY_ALLOW_LOCAL_UPSTREAMS=1` **仅**放行 loopback，不放行 10/172.16/192.168/169.254 等 LAN。
- 总超时覆盖响应头后的完整 body 流；连接超时仅覆盖建连到响应头；客户端断开可中止上游 body 与 `drain` 等待。
- 固定拼接：`{baseUrl}/{models|chat/completions|responses|messages}`。

## 安全行为

- **Header 白名单（请求转发）**：`content-type`、`accept`、`anthropic-version`、`anthropic-beta`、`openai-organization`、`openai-project`。
- **认证**：服务端按 upstream 配置注入 `Authorization: Bearer …` 或 `x-api-key`；客户端的 auth / host / cookie / hop-by-hop 头会被丢弃。
- **POST**：仅 `application/json`（或 `+json`），body ≤ **4 MiB**，超限 **413** 且不访问上游。
- **透传**：上游状态码与响应体按字节流式转发；响应头仅安全白名单。
- **取消**：客户端断开会 abort 上游请求。
- **限制**：连接/总超时、并发上限、速率限制均为 fail-closed（超时 504、并发 503、速率 429）。
- **日志**：不含 API key、`Authorization`、请求/响应正文、完整 URL query；错误响应不返回堆栈或密钥。
- **配置**：缺失或非法配置时进程拒绝启动，无动态代理 fallback。

## 本地开发

```bash
cd ai-gateway
cp .env.example .env   # 填入占位密钥与 upstream 映射
# 导出环境变量后：
npm start
# 或
npm test
```

Node 内置测试使用本地假上游（loopback），不访问公网。

环境变量见 [`.env.example`](./.env.example)。生产**不要**设置 `AI_GATEWAY_ALLOW_LOCAL_UPSTREAMS=1`。

## 生产部署要点

- 进程只监听 **`127.0.0.1`**，由本机 Nginx 反代同域路径 `/api/ai/`。
- 流式响应需要：`proxy_buffering off`；网关对 `text/event-stream` 额外设置 `X-Accel-Buffering: no` 与 `Cache-Control: no-store`、足够大的 `proxy_read_timeout`。
- 示例：`deploy/ai-gateway.service`、`deploy/nginx-ai-gateway.conf`。

## Scripts

| 命令 | 说明 |
|------|------|
| `npm start` | 启动网关 |
| `npm run dev` | `--watch` 启动 |
| `npm test` | `node --test` 全量测试 |

无第三方运行时依赖；仅使用 Node 22 内置模块。
