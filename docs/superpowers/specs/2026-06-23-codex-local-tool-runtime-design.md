# Codex 本地工具执行逻辑与 Excel 桥接设计

## 背景

本项目当前需要让桌面端 AI 稳定读取和写入 Excel/WPS。直接让模型生成 PowerShell 或脚本去操作表格，容易出现语法、转义、COM 状态、并发和错误恢复问题。

Codex 的开源实现提供了一个更稳定的参考范式：模型不直接操作本机文件或应用，而是只发起结构化工具调用；本地 Runtime 负责权限判断、沙箱、审批、执行和结果回传。

本文基于 `openai/codex` 源码 commit `49614a0391d83eec442ffeca1d4aa0fdeb119818` 整理实现逻辑，并映射到本项目后续 Excel/WPS 桥接层设计。

## Codex 的核心实现逻辑

Codex 的本机文件创建和修改不是由模型直接完成，而是由本地客户端通过工具 Runtime 完成。

整体链路如下：

```text
用户需求
  -> 模型推理
  -> 模型输出 tool call
  -> Tool Handler 解析参数
  -> Runtime 校验权限和沙箱策略
  -> 本地进程 / 文件系统 / MCP 服务执行
  -> 结构化结果返回模型
  -> 模型继续下一步
```

关键原则是：模型只决定“做什么”，本地 Runtime 决定“能不能做”和“怎么安全地做”。

## 工具暴露

Codex 会把可用能力注册成工具规格 `ToolSpec`，再暴露给模型。模型看到的是工具名、参数结构和调用格式，而不是直接文件系统权限。

典型工具包括：

- `shell_command`：执行本机 shell 命令。
- `apply_patch`：以受控补丁方式修改文件。
- MCP 工具：通过 MCP 调用外部服务或本地连接器。
- code mode 工具：执行受控代码单元，并允许嵌套调用其它工具。

源码参考：

- `codex-rs/core/src/tools/handlers/shell/shell_command.rs`
- `codex-rs/core/src/tools/handlers/apply_patch.rs`
- `codex-rs/core/src/tools/code_mode/execute_spec.rs`

## Shell 命令执行链路

当模型调用 `shell_command` 时，Codex 不会直接把字符串扔给系统执行，而是先进入 `ShellCommandHandler`。

处理流程：

```text
模型输出 shell_command 参数
  -> ShellCommandHandler::handle_call
  -> 解析 JSON 参数
  -> 解析 workdir
  -> 构造 ExecParams
  -> run_exec_like
  -> ShellRuntime
  -> 审批 / 沙箱 / 环境变量 / 超时 / 取消
  -> execute_env
  -> 返回 stdout、stderr、exit_code、duration
```

`ExecParams` 中包含：

- `command`
- `cwd`
- `expiration`
- `capture_policy`
- `env`
- `sandbox_permissions`
- `windows_sandbox_level`
- `justification`

这说明 Codex 把“执行命令”抽象成受控请求，而不是无约束执行。

## 文件修改链路

Codex 修改文件优先使用 `apply_patch`，而不是让模型生成 `echo`、`cat`、`sed` 之类命令。

`apply_patch` 的处理流程：

```text
模型输出 patch
  -> ApplyPatchHandler::handle_call
  -> parse_patch
  -> verify_apply_patch_args
  -> 计算变更文件路径
  -> 计算需要的写权限
  -> assess_patch_safety
  -> 必要时请求审批
  -> ApplyPatchRuntime::run
  -> codex_apply_patch::apply_patch
  -> 返回 delta / stdout / stderr / exit_code
```

这里有几个非常关键的设计点：

- patch 必须先被解析，不能直接落盘。
- patch 会基于目标环境文件系统做校验。
- 写路径会经过权限计算。
- 安全评估可能自动允许、请求用户确认或拒绝。
- 真正写入由 Runtime 完成，不由模型完成。

这套机制降低了误改文件、越权写入、格式错误和不可追踪变更的风险。

## MCP 工具调用链路

MCP 工具与 shell、patch 的思路一致：模型只提交工具名和参数，本地 Codex 负责解析、权限策略、审批和实际调用。

处理流程：

```text
模型输出 MCP tool call
  -> handle_mcp_tool_call
  -> 解析 arguments JSON
  -> 查询工具元数据
  -> 判断 app/plugin/custom approval policy
  -> 必要时请求审批
  -> 调用 MCP server
  -> 截断或整理结果
  -> 返回模型可读结果
```

这说明 Codex 对“外部能力”的接入也是统一模式：能力通过工具边界暴露，调用前后都有治理层。

## 可借鉴的架构原则

Codex 的实现可以总结为五层：

```text
1. Model Layer
   模型推理，只输出工具调用意图。

2. Tool Spec Layer
   定义工具名、参数格式、描述和可调用边界。

3. Handler Layer
   解析参数、校验格式、解析路径、转换成内部请求。

4. Runtime Layer
   执行权限、审批、沙箱、队列、超时、取消和错误归一化。

5. Native Capability Layer
   调用本机文件系统、shell、MCP、浏览器、桌面应用或其它本地服务。
```

对 Excel/WPS 场景，最重要的是把复杂性放在 Runtime 和 Native Capability Layer，而不是放在模型输出内容里。

## 映射到本项目的实现设计

本项目应该避免让模型直接输出 PowerShell COM 脚本操作 Excel/WPS。更稳定的做法是照 Codex 的工具 Runtime 模式设计本地桥接层。

推荐链路：

```text
模型
  -> workbook.inspect / range.read / range.write 等工具调用
  -> Tool Registry 分发
  -> Tool Handler 解析和校验参数
  -> Excel Runtime 串行队列、权限、目标锁定、超时、重试
  -> 本地桥接进程
  -> Excel/WPS COM 或文件 API
  -> 结构化结果返回模型
```

## 建议工具接口

基础工具建议保持少而稳定：

```ts
workbook.inspect(): WorkbookSnapshot

selection.get(): SelectionSnapshot

range.read({
  sheet?: string
  address: string
  valueMode?: "display" | "raw" | "formula"
}): RangeReadResult

range.write({
  sheet?: string
  address: string
  values: unknown[][]
  mode?: "value" | "formula"
  verify?: boolean
}): RangeWriteResult

range.clear({
  sheet?: string
  address: string
}): RangeClearResult
```

后续再按需要增加：

```ts
sheet.list()
sheet.activate()
formula.set()
table.detect()
workbook.save()
```

不要一开始暴露任意脚本执行能力。任意脚本能力应该作为高风险工具单独审批。

## Tool Handler 职责

Handler 只负责把模型参数转换为安全的内部请求：

- 校验 JSON schema。
- 规范化工作簿、工作表和区域地址。
- 限制最大读取/写入范围。
- 拒绝歧义地址。
- 为每次工具调用生成 `toolCallId`。
- 把请求放入 Excel Runtime 队列。

Handler 不应该直接拼接 PowerShell，也不应该直接触碰 COM。

## Excel Runtime 职责

Excel Runtime 是稳定性的核心，应负责：

- 串行执行所有 Excel/WPS 操作，避免 COM 并发冲突。
- 维护当前连接状态。
- 绑定目标应用、工作簿和工作表，避免写错窗口。
- 设置超时和取消。
- 对短暂失败进行有限重试。
- 把 COM/WPS 错误归一化为结构化错误码。
- 记录工具调用摘要、耗时、状态和错误。
- 对写操作做可选 read-back 验证。

示例错误码：

```ts
type ExcelToolErrorCode =
  | "APP_NOT_RUNNING"
  | "WORKBOOK_NOT_FOUND"
  | "SHEET_NOT_FOUND"
  | "INVALID_RANGE"
  | "RANGE_TOO_LARGE"
  | "APP_BUSY"
  | "EDIT_MODE"
  | "PERMISSION_DENIED"
  | "WRITE_VERIFY_FAILED"
  | "UNKNOWN_COM_ERROR"
```

## 本地桥接进程设计

如果要替代当前 PowerShell COM，建议用常驻本地桥接进程：

```text
Electron 主进程
  -> 启动 C#/.NET bridge
  -> stdio JSON-RPC 通信
  -> bridge 直接使用 COM 自动化 Excel/WPS
```

推荐优先使用 stdio，而不是本地 HTTP 端口：

- 不暴露本地端口，安全边界更小。
- 生命周期跟随 Electron 子进程。
- 更容易做请求/响应匹配。
- 不需要处理端口冲突。

请求格式示例：

```json
{
  "id": "tool-123",
  "method": "range.read",
  "params": {
    "sheet": "Sheet1",
    "address": "A1:C10",
    "valueMode": "display"
  }
}
```

响应格式示例：

```json
{
  "id": "tool-123",
  "ok": true,
  "result": {
    "sheet": "Sheet1",
    "address": "A1:C10",
    "values": [["姓名", "金额"]]
  },
  "meta": {
    "durationMs": 42
  }
}
```

失败格式示例：

```json
{
  "id": "tool-123",
  "ok": false,
  "error": {
    "code": "APP_BUSY",
    "message": "Excel 当前正忙或处于编辑状态",
    "retryable": true
  }
}
```

## 写入安全策略

写操作建议默认比读操作更严格。

`range.write` 应该执行：

```text
1. 校验目标 range 大小
2. 校验 values 行列数
3. 捕获写入前快照，必要时用于回滚或提示
4. 批量写入 Value2 或 Formula
5. read-back 验证
6. 返回写入单元格数量和验证结果
```

对于大范围写入、清空、覆盖公式、跨表写入等操作，应由权限模式决定是否需要用户确认。

## 与当前项目的落地方式

当前桌面端已经有类似边界：

- `desktop/electron/agent/toolRegistry.ts` 定义工具注册和执行。
- `desktop/electron/agent/excelBridge.ts` 负责 Excel/WPS 桥接。
- `desktop/electron/agent/agentLoop.ts` 负责模型循环和工具调用。

后续可以在不改变模型层和 UI 层的情况下替换桥接实现：

```text
现状：
toolRegistry -> ExcelComBridge -> PowerShell COM

目标：
toolRegistry -> ExcelRuntime -> DotNetExcelBridge -> Excel/WPS COM
```

这样可以保留现有工具名和前端交互，只替换不稳定的执行层。

## 迁移步骤建议

第一阶段：稳定现有接口

- 保持现有 `ExcelWorkbookBridge` 接口不变。
- 增加工具调用队列。
- 为读写操作增加结构化错误。
- 增加读写范围限制。
- 写入后增加验证。

第二阶段：引入常驻桥接进程

- 新增 `DotNetExcelBridge`。
- Electron 启动 C# bridge 子进程。
- 用 stdio JSON-RPC 实现 `inspect/read/write/selection`。
- 保留 PowerShell bridge 作为 fallback。

第三阶段：移除 PowerShell 复杂脚本

- 将读写迁移到 bulk COM 调用。
- 减少字符串拼接脚本。
- 统一错误码和日志。
- 覆盖 Excel/WPS 典型状态：未启动、忙碌、编辑中、无工作簿、受保护表。

## 成功标准

实现完成后，应满足：

- 模型不再输出复杂 PowerShell 操作表格。
- 读写工具参数稳定、可校验。
- Excel/WPS 操作串行化。
- 写入结果可验证。
- 错误能被模型理解并继续修正。
- UI 能展示工具调用状态、错误和耗时。
- PowerShell 脚本复杂度显著下降，最终可以移除。

## 结论

Codex 能修改本机文件，靠的不是模型直接控制电脑，而是本地客户端提供了受控工具 Runtime。模型只发起工具调用，本地 Runtime 做参数校验、权限审批、沙箱执行和结果回传。

本项目也应该使用同样思想：把 Excel/WPS 能力封装成稳定工具，而不是让模型临时生成操作脚本。这样可以显著降低 PowerShell 语法错误、COM 状态漂移、并发冲突和写入不可验证的问题。
