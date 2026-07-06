# Electron main modules

本目录承接主进程侧 IPC 注册、窗口控制、文件路径授权、OCR 和 AI 连接测试等 Electron 主进程能力。

## 模块职责

- `ipcHandlers.ts`：主进程 IPC 聚合入口，负责创建共享依赖并注册各业务域 handler。
- `ipcOcrHandlers.ts`：OCR IPC 注册、MinerU 付费/免费降级、本地解析兜底、发票字段抽取和 OCR 结果归一化。
- `ipcAiHandlers.ts`：AI 模型列表读取和连接测试，保留 Anthropic、OpenAI Responses、Chat Completions 三类协议的原有请求逻辑。
- `ipcSandboxHandlers.ts`：沙箱默认规则展示、用户规则保存、额外可写根保存和运行时 sandbox 单例刷新。
- `ipcFileHandlers.ts`：文件/图片/文件夹选择、文件夹枚举、Base64 读取、临时文件写入、回收站、打开文件、复制路径和资源管理器定位。
- `ipcPathSecurity.ts`：主进程文件路径授权和路径范围校验。

## 拆分约定

`ipcHandlers.ts` 只保留共享依赖创建和小型聚合逻辑。新增或继续拆分 IPC 时，优先按业务域放入独立模块，并保持 schema 校验、路径授权、超时和错误返回行为与原链路一致。
