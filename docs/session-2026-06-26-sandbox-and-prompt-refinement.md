# 会话总结：Agent 命令执行沙箱化 + 系统提示词润色

> 日期：2026-06-26
> 关联：`docs/sandbox-implementation-plan.md`、`desktop/electron/agent/sandbox/`、`desktop/electron/agent/systemPrompt.ts`

## 目标

1. 参考 Codex 源码（`codex-reference/codex-rs/{sandboxing,linux-sandbox,windows-sandbox-rs,execpolicy,process-hardening}`）的沙箱实现思路，在本 Electron Excel AI 插件中做一套可落地的命令执行沙箱。
2. 修复系统提示词的两个具体问题：
   - 任何场景都会无脑触发 `workbook.inspect` / `selection.get` 的工作簿上下文检查
   - 写公式（公式助手）场景提示词与公式助手字段脱钩、与通用质量守则反复冲突

## Codex 沙箱实现剖析（关键点）

- **`sandboxing/`**：跨平台抽象层。`SandboxManager::transform()` 把命令+权限 profile 转成"沙箱化 argv"，按 `cfg!(target_os)` 分派到平台后端。
- **`linux-sandbox/`**：bubblewrap `--ro-bind / /` + landlock + seccomp + `--unshare-user/pid/net`。
- **`windows-sandbox-rs/`**：WRITE_RESTRICTED 受限令牌 + capability SID + ACL deny-ACE + 私有桌面 + WFP 网络封禁；复杂情况走 elevated backend。
- **`execpolicy/`**：Starlark DSL 命令策略，按前缀 token 匹配，多规则取最严 `forbidden > prompt > allow`，**独立于 OS 沙箱**，是命令语义级过滤。
- **`process-hardening/`**：`#[ctor::ctor]` 在 main 前禁 core dump、ptrace、危险 envvars。
- 与 **Excel 主链路的关键约束**：本插件 Excel 工具靠 `GetObject("Excel.Application")` 连回用户当前 Excel，受限令牌+私有桌面+`--unshare-net` 会让 COM 回连失败 → 沙箱只施加在 `shell.execute` 等不依赖 COM 回连的工具。

## 本方案落地范围（阶段 0+1+4+5）

详见 `docs/sandbox-implementation-plan.md`。本次只实现在不引入原生依赖、不破坏 Excel 主链路前提下的纯 TS 沙箱：

- 阶段 0：`shell.execute` 接口重构（命令切 token + Windows `-EncodedCommand` 注入）
- 阶段 1：`execpolicy` 命令策略层（前缀规则 + 决策聚合 + cwd 白名单重定向）
- 阶段 4：进程加固（env allowlist 清洗 + `-EncodedCommand` 标准化）
- 阶段 5：JSONL 审计
- **未落地**（留作后续可演进项）：阶段 2 (utilityProcess 隔离)、阶段 3 (OS 级受限令牌)

## 新增文件

| 路径 | 职责 |
|------|------|
| `desktop/electron/agent/sandbox/parseCommand.ts` | 命令切子命令+token（管道/换行/引号/转义/注释） |
| `desktop/electron/agent/sandbox/execPolicy.ts` | 策略引擎（前缀 token 规则 + 最严合并 + cwd 白名单重定向） |
| `desktop/electron/agent/sandbox/defaultRules.ts` | 默认规则：`Remove-Item -Recurse` / `rm -rf /` / `Format` / `Stop-Computer` / `reg delete` / `iex` / `diskpart` → forbidden；`curl` / `Invoke-WebRequest` / `powershell -c` → prompt |
| `desktop/electron/agent/sandbox/audit.ts` | 按 `sandbox-logs/YYYY/MM/DD/audit-*.jsonl` 落审计 |
| `desktop/electron/agent/sandbox/index.ts` | 一站式入口 `evaluateCommand` + `runShellSpawn`（env 白名单 + `-EncodedCommand`）+ `killProcessTree`（taskkill /T 强杀整棵进程树） |
| `desktop/electron/agent/sandbox/sandbox.test.ts` | 14 个单测覆盖破坏性命令拒绝、prompt 命令、cwd 重定向 |
| `desktop/src/components/settings/ExecPolicySettings.tsx` | 安全策略设置页（用户规则增删、可写根增删、保存） |

## 改动文件

### Electron 主进程

- `executors.ts` — `executeShellCommand` 改为先 `evaluateCommand`，forbidden 直接返失败；Windows 改走 `-EncodedCommand`（避免明文 `-Command` 拼接注入），env 白名单清洗；超时 `taskkill /T /F /PID` 强杀整树
- `toolExecutor.ts` — `processToolCalls` 在审批前对 `shell.execute` 评估策略：
  - forbidden → 拒绝不进 spawn，记审计并回失败结果
  - prompt → 覆盖 `permissionMode` 与 `alwaysAllowedTools` 强制审批，把 `sandboxJustification` 透传审批对话框
  - allow → 维持原审批流程
- `ipcHandlers.ts` — 新增 `sandbox:getConfig / setUserRules / setWritableRoots` IPC + `applySandboxConfig()` 热更到沙箱单例 + 文件尾辅助函数 `normalizeUserRules` / `getSandboxDefaultRulesForUI`
- `settingsManager.ts` — `DEFAULT_SETTINGS` 增加 `sandboxUserRules` / `sandboxExtraWritableRoots`
- `main.ts` — `app.whenReady` 中调 `applySandboxConfig()`
- `preload.ts` — 暴露 `window.electronAPI.sandbox.*`

### 前端

- `electronApi.d.ts` — `SandboxPrefixRule` / `SandboxConfig` 类型 + `sandbox` 命名空间接口
- `services/ipcApi.ts` — 实桥接 `sandbox.getConfig/setUserRules/setWritableRoots`
- `components/SettingsPage.tsx` — 新增「安全策略」tab (`safety` section + `ShieldAlert` 图标)
- `components/settings/ExecPolicySettings.tsx` + 配套 CSS — 安全策略设置页
- `components/chat/ToolConfirmDialog.tsx` — 渲染 `sandboxJustification`（命中 prompt 规则时显示理由）
- `store/agentEventHandler.ts` + `store/chatStore.ts` — 透传 `sandboxJustification`
- `i18n.ts` — 中英文补 `assistant.sandboxJustification`
- `styles/tool-confirm.css` — 沙箱 notice 样式
- `styles/settings.css` — ExecPolicySettings 样式

### 系统提示词

- `agent/systemPrompt.ts` — 两轮润色：
  1. **工作流程按场景分支**（A/B/C/D）+「探测要按需」原则：纯问答/解释/系统命令类请求不再无脑 `workbook.inspect`/`selection.get`
  2. **公式生成场景**（`scenarioFormula()`）大改：
     - 对齐公式助手 5 个结构化字段（任务说明 / 数据源选区 / 答案参考样例 / 写入锚点 / 是否动数组）
     - 样例分完整/部分两种处理：完整样例需 value 完全一致；部分样例不做"是否完整"的纠结
     - #SPILL! 主动清理：先 `range.read` 探测溢出方向非空单元格 → 清空 → 告知用户清理内容
     - 表头按需拼接：检测锚点是否含表头，含则不拼、不含则在首行拼语义标题
     - 嵌套约束松绑：取消"嵌套不超过3层"死规矩，动态数组优先减少辅助列
     - 覆盖公式单元格安全底线放宽：公式助手指定写入锚点含旧公式视为用户要重写
     - 动数组"否"态改用"写入形态是否依赖溢出"作为判定标准，不再按函数名禁用 XLOOKUP 等
     - 与 §可维护性 / §安全底线 / §结果验证的冲突经全局审计调和

## 重要安全语义

- 任意 `permissionMode`（含 `confirm_all` + 始终允许）下，`forbidden` 规则仍直接拒绝
- `prompt` 命令无视 `alwaysAllowedTools`，永远走审批并展示理由
- 越界 cwd 自动重定向到临时目录并记审计
- Windows shell.execute 明文 `-Command` 注入入口改为 Base64(UTF-16LE) `-EncodedCommand`
- 子进程环境只保留 USERNAME / USERPROFILE / PATH 等白名单变量
- 超时用 `taskkill /T /F /PID` 强杀整棵进程树

## 验证结果

- `npm run typecheck` ✅（前端 tsc + electron tsc 全过）
- `npx vitest run` ✅（58 个测试全部通过，含新增 14 个沙箱单测）
- Excel COM 主链路无回归：所有走 `excelBridge.executePowerShell`（`-EncodedCommand`）的工具保持原有 spawn 行为与完整 PSModulePath 环境，沙箱前置评估仅施加在 `shell.execute` 这一外部任意命令入口

## 后续可演进项

- 阶段 2：`utilityProcess.fork` 把 `shell.execute` spawn 搬出主进程，加 Job Object 进程组强杀
- 阶段 3：复用 Codex `codex-windows-sandbox.exe` 拿 WRITE_RESTRICTED + WFP（仅对 A 类工具，不影响 COM 路径）
- execpolicy 升级到完整 Starlark DSL / `host_executable` / `network_rule`