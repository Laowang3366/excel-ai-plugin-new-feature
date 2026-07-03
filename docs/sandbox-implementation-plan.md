# Agent 命令执行沙箱化实现方案

> 参考：`codex-reference/codex-rs/{sandboxing, linux-sandbox, windows-sandbox-rs, execpolicy, process-hardening}`
> 落地点：本仓库 `desktop/electron/agent/` 与 `desktop/electron/main-modules/`
> 目标平台：Windows（Excel/WPS COM 自动化场景）为主线，非 Windows 走现有 fallback

---

## 一、Codex 沙箱实现剖析

Codex 把"沙箱"拆成 **5 个 crate**，分层清晰、平台相关代码用 `#[cfg]` 严格隔离。整体架构大致是：

```
core/tools ─┐
            ├─ SandboxManager.select_initial(transform) ─→ SandboxExecRequest ─→ spawn
exec ───────┘                                              │
                                                           ▼
                       ┌──────── 平台后端 ────────┐
                       │ macos : seatbelt (sandbox-exec + .sbpl 策略)
                       │ linux  : bubblewrap --ro-bind / + landlock + seccomp
                       │ windows: WRITE_RESTRICTED token + capability SID + ACL deny-ACE + WFP + 私有桌面
                       └──────────────────────────┘
                       execpolicy(STarlark DSL)  ↘ prefix_rule(allow|prompt|forbidden)
                       process-hardening(ctor)   ↘ 禁 core dump / ptrace / 危险 envvars
```

### 1. `sandboxing/` — 跨平台抽象层（`sandboxing/src/manager.rs`）

- `SandboxType` 枚举：`None / MacosSeatbelt / LinuxSeccomp / WindowsRestrictedToken`（`manager.rs:34`）
- `SandboxablePreference`：`Auto / Require / Forbid`（`manager.rs:53`），决定是否进沙箱
- `get_platform_sandbox()`：按 `cfg!(target_os=...)` 选平台后端，Windows 默认**关闭**，需 `windows_sandbox_enabled=true` 才启用 `WindowsRestrictedToken`（`manager.rs:59`）
- `SandboxManager::transform(SandboxTransformRequest) -> SandboxExecRequest`（`manager.rs:318`）是核心：
  - 解析 `PathUri` → 绝对路径，套上 MITM CA 信任根（`with_managed_mitm_ca_readable_root`，`manager.rs:75`）
  - **平台分支决定 argv 形态**：
    - macOS / Linux：把原 argv 前置 `sandbox-exec` 或 `codex-linux-sandbox` 的参数（命令被打包进沙箱包装器）
    - Windows：**argv 不变**，仅准备 `PendingSandboxedExecRequest`（filesystem overrides / deny ACL）以供 spawn 时降权
- `transform_for_direct_spawn()`（`manager.rs:458`）：Windows 直跑时把 executable 替换成 codex-windows-sandbox helper，并以 `create_windows_sandbox_command_args_for_permission_profile(...)` 生成 wrapper argv、env allowlist 仅留 `USERNAME/USERPROFILE`

### 2. `linux-sandbox/`（见 `linux-sandbox/README.md`）

- 默认 `bwrap` POL：`--ro-bind / /` 全盘只读，`--bind <writable_root> <writable_root>` 给写根打补丁；`.git / .codex / gitdir:` 子路径再 `--ro-bind` 切回只读
- 命名空间隔离：`--unshare-user --unshare-pid`，需要时 `--unshare-net`
- 进程加固：`PR_SET_NO_NEW_PRIVS` + seccomp 网络过滤；blob glob 扫描不可读文件并 `/dev/null` 掩盖
- WSL1 不支持（无 user namespace）；优先用系统 `bwrap`，缺失时回退 bundled

### 3. `windows-sandbox-rs/`（`.../src/lib.rs` + `windows_impl` 模块）

- 受限令牌（`token.rs`：`create_readonly_token_with_caps_from`、`create_workspace_write_token_with_caps_from`）。`WRITE_RESTRICTED` token 让 restricting SID 只在**写入**检查中生效（这是不能执行 deny-read 限制的根因，见 `windows.rs:110-118`）
- **capability SID + ACL deny-ACE**（`acl.rs` / `cap.rs` / `workspace_acl.rs`）：给 workspace 写目录分配一个唯一的 workspace write-cap SID，对该 SID 添加 Allow-Write ACE；对 cap SID 拒绝其他写权限 → 子进程只能写白名单目录
- `apply_deny_read_acls` / `apply_deny_write_acls`：在敏感路径上铺 Deny ACE
- **私有桌面** `CreateDesktop`（`desktop.rs` / `LaunchDesktop`）：子进程独立桌面，看不到用户桌面的窗口
- **网络封禁**：WFP（Windows Filtering Platform，`wfp*.rs`）按账户 SID 装过滤器封掉出站
- **elevated backend**：deny-read、复杂 ACL 需要 Admin，通过命名管道与一个 elevated runner 进程通信（`elevated*.rs` + `ipc_framed.rs`）
- spawn 走 `CreateProcessAsUser`（`process.rs: create_process_as_user`），不依赖 Node 子进程

### 4. `execpolicy/`（见 `execpolicy/README.md`，`src/rule.rs`）

Starlark DSL 描述命令策略：

```starlark
prefix_rule(
  pattern = ["git", "reset", "--hard"],  # 第一个 token 固定，其余可 [xC1,C2] 表互斥
  decision = "allow" | "prompt" | "forbidden",   # 默认 allow
  justification = "...", match = [...], not_match = [...]
)
host_executable(name="git", paths=["/usr/bin/git", "/opt/homebrew/bin/git"])  # 约束 basename fallback
network_rule(host="example.com", protocol="https", decision="allow")
```

- 按前缀 token 匹配；多规则命中取最严 `forbidden > prompt > allow`
- 提供 `codex execpolicy check --rules ...` CLI 与 `Policy::matches_for_command_with_options`
- 这一层 **独立于 OS 沙箱**，是 Codex 在权限 profile 之外再加一道"命令语义级"过滤

### 5. `process-hardening/`（`README.md`）

`pre_main_hardening()` 用 `#[ctor::ctor]` 在 `main()` 前跑：禁 core dumps、禁 ptrace attach、清掉 `LD_PRELOAD / DYLD_*`。这是宿主进程加固，不在沙箱进程树内，进程级早期硬化。

### 6. 与 Agent 的集成（`core/src/tools/sandboxing.rs`）

- `ApprovalStore` / `with_cached_approval`：缓存"会话内始终允许"，避免重复弹窗
- `Sandboxable` trait + `ToolRuntime`：每个工具声明自己是否可沙箱化、需要的权限
- `exec.rs` / `orchestrator.rs` 调 `SandboxManager::transform` 产 `SandboxExecRequest`，再由 spawn 层落盘

---

## 二、本项目现状与差距

| 维度 | 本项目现状 | Codex 对应 |
|------|------------|-----------|
| 执行位置 | 全部 AgentLoop + 23 工具在 **Electron 主进程**内，副作用经 `child_process.execFile` spawn | 主进程编排 + spawn 沙箱化命令 |
| Shell 入口 | `executors.ts:80 executeShellCommand` 明文把 `command` 拼到 `powershell.exe -Command <command>`；`shell.execute` risk=dangerous 但用户可"始终允许"后无任何过滤 | argv 包装器 + 受限令牌 |
| Excel 工具 | `excelBridgeHelpers.executePowerShell/executePythonScript/executeJScript`，**靠 `GetActiveObject` 连接用户当前 Excel** —— 不能失去会话/桌面可见性 | — |
| 权限层 | 3 档 `permissionMode`（normal / auto_approve_safe / confirm_all）+ 单工具 `alwaysAllowedTools` Set | `permission_profile` + OS 沙箱 + execpolicy 三层 |
| 命令策略 | **无** | `execpolicy` 前缀规则 |
| 文件系统 | 仅校验 `workdir` 是目录 | `FileSystemSandboxPolicy` writable_roots / deny_read / deny_write |
| 网络 | 无 | WFP / `--unshare-net` |
| 进程加固 | 只有 `windowsHide / maxBuffer / timeout` | Job Object / seccomp / PR_SET_NO_NEW_PRIVS / 私有桌面 |

**核心缺口**：`shell.execute` 在 `confirm_all` 或被勾"始终允许"后可无确认地执行**任意 PowerShell**；明文拼接 `-Command`，无策略、无 cwd 白名单、无网络封禁、无审计；失败也无结构化日志。

**关键约束（与 Codex 最大的不同）**：Excel 工具依赖 **跨进程 COM 连回用户当前 Excel**，不能像 Codex 那样做受限令牌 + 私有桌面 + `--unshare-net`，否则 `GetObject("Excel.Application")` 拿不到主桌面 Excel 实例。因此沙箱只能施加在 **不依赖 COM 回连的工具** —— 主要是 `shell.execute` 与 `script.execute`（看少，多数 `script.execute` 仍会注入 xlwings 连 Excel，需要分类处理）。

---

## 三、设计目标与原则

1. **安全分级**：把 23 个工具按"是否依赖 COM 回连"和"危险面"分为三类，沙箱强度按级递增
2. **不破坏 Excel 主链路**：所有通过 `GetObject("Excel.Application/Ket.Application")` 连回当前 Excel 的工具保持原 spawn 行为，只在外层加审计与确认
3. **能力可降级**：用户机器若没有 AppContainer / 受限令牌基因，回退到策略层 + Job Object，不阻塞工作
4. **零原生依赖起步**：先做纯 TS 能落地的命令策略层（对应 Codex 的 `execpolicy`），再渐进引入 OS 沙箱

---

## 四、工具风险分级（沙箱适用范围）

| 类别 | 工具 | COM 回连 | 是否沙箱化 | 策略层 | OS 沙箱 |
|------|------|----------|------------|--------|---------|
| **A 类：外部任意命令** | `shell.execute` | 否 | ✅ 全部施加 | ✅ | ✅ Job Object + AppContainer |
| **B 类：用户代码注入脚本** | `script.execute`（纯 Python 处理数据分支） | 否 | ✅ | ✅（针对代码内 `import socket`、`subprocess` 检测） | ✅ Job Object |
| **B 类：脚本连 Excel** | `script.execute`（含 xlwings 分支）、`ui.createForm`、`vba.runMacro`、`vba.writeModule` | ✅ | ⚠️ | ✅ | ❌（保持 COM） |
| **C 类：Excel COM 读写** | `range.*`、`sheet.operation`、`workbook.*`、`selection.get`、`formula.*`、`ui.addControl/removeControl/addMenu/listControls` | ✅ | ❌ | ❌ | ❌ |
| **D 类：纯本地** | `formula.search`、`file.getPaths`、`script.detect` | 否 | — | — | — |

> 设计取向：**A/B(无COM)** 进沙箱，**B(COM)/C** 保持现状但统一加审计与确认，**D** 不动。

---

## 五、实现方案：分层落地

### 阶段 0 ── 重构 `shell.execute` 命令接口（先决条件）

Codex 的命令策略层只对**已经是 token 数组**的命令工作；本项目现在传 `command: string` 明文拼接 PS `-Command`。先做：

- 新增 `desktop/electron/agent/sandbox/parseCommand.ts`：用 `fast-glob` 风格的轻量 shlex（或直接基于 PowerShell AST 的 `System.Management.Automation.PSParser::Tokenize` 通过 PS 自身的解析能力 —— 用一段固定 `-EncodedCommand` 启动一段解析脚本输出 JSON Token）把命令切成 `string[]`
- `executeShellCommand(command, workdir, timeoutMs)` 保留字符串入参兼容，但内部先 `parseCommand` 后再交给策略层
- 仍保留 `-EncodedCommand` Base64(UTF-16LE) 模式注入，避免 PS 字符串转义/注入（Codex 沙箱没暴露这块，但本项目 `excelBridgeHelpers.executePowerShell` 已经是 `-EncodedCommand` 模式，照搬到 `shell.execute` 同时收紧）

### 阶段 1 ── `execpolicy` 命令策略层（最易落地，对应 Codex `execpolicy`）

**新增** `desktop/electron/agent/sandbox/execPolicy.ts`：

- 数据结构对齐 Codex `rule.rs`：`PrefixRule = { first: string, rest: PatternToken[], decision: "allow"|"prompt"|"forbidden", justification?: string }`，`PatternToken = Single | Alts`
- 规则文件格式选 **TypeScript / JSON**（不上 Starlark，我们不在主进程嵌解释器；DSL 留下可演进）
- 决策聚合：多规则命中取最严 `forbidden > prompt > allow`；未命中走 `permissionMode` 默认
- 默认策略文件 `desktop/electron/agent/sandbox/default.rules.json`（示例）：

```jsonc
[
  // 破坏性命令直接拒绝
  { "pattern": ["rm", "-rf", "/"],            "decision": "forbidden", "justification": "防止清盘" },
  { "pattern": ["Remove-Item", ["-Recurse", "-r"], ["-Force", "-f"]], "decision": "forbidden" },
  { "pattern": ["Format"],                   "decision": "forbidden" },
  { "pattern": ["reg", "delete"],            "decision": "forbidden" },
  { "pattern": ["Stop-Computer", "Restart-Computer", "shutdown"], "decision": "forbidden" },
  { "pattern": ["diskpart"],                 "decision": "forbidden" },
  // 远程下载/执行放 prompt
  { "pattern": ["Invoke-WebRequest", "iwr"], "decision": "prompt" },
  { "pattern": ["curl", "wget"],             "decision": "prompt" },
  // 远程脚本无脑执行拒绝
  { "pattern": ["iex"], "decision": "forbidden" },
  { "pattern": ["powershell"],               "decision": "prompt" },
  // Excel 主流场景放行
  { "pattern": ["python", "python3"],        "decision": "allow" },
]
```

- 用户可在设置页增减规则；设置文件位置：`settingsManager` 已用的 `electron-store` 加一项 `execPolicy: PrefixRule[]`
- 流程：`executeShellCommand` → `parseCommand` → `ExecPolicy.evaluate(tokens)` →
  - `forbidden`：直接返回错误，记审计
  - `prompt`：走现有 `requestToolApproval`（在确认对话框里把 justification 展示出来）
  - `allow`：直接进 spawn
  - 未命中 + `confirm_all/normal`：维持现状审批

**workdir 白名单**（对应 Codex `FileSystemSandboxPolicy.get_writable_roots`）：

- `shell.execute` 的 `workdir` 必须在以下白名单内：当前打开的工作簿目录、`os.tmpdir()`、用户文档目录、桌面 —— 违者改写到 `os.tmpdir()`，并在审计里标记 `workdir_redirected`
- 阶段 2 再加可由用户配置的额外可写根

### 阶段 2 ── Job Object 子进程隔离 + utility process 托管（无需 native）

**新增** `desktop/electron/agent/sandbox/jobObject.ts`：

- 通过 Node native - **无原生依赖**：用 `child_process.spawn` 包装外加一个 PowerShell 启动器，启动器自己往 Job Object 注册（PowerShell 难调 `AssignProcessToJobObject`，所以此处要补一个最小 native 模块，或借助 `process.kill` 树状终止 + 周期回收监控）
- 折中方案（不引入 native）：在 `shell.execute` spawn 时把 `child` 实例注册到 AgentLoop 的 `childProcessRegistry`（`Map<id, ChildProcess>`），增加：
  - 进程组终止：`child.kill('SIGKILL')` 同时对 `--windows-kill-tree` 等价的实现 —— 用 PS wrapper 启动子进程并 `taskkill /T /F /PID` 在超时时调用
  - 资源监控：spawn 时启动 `pslist`/`Get-Process` 周期采集，CPU/内存超阈值（可配置）即告警
- **utility process 化**：用 Electron 的 `utilityProcess.fork` 编一个 `desktop/electron/agent/sandbox/shellRunnerWorker.ts`，把 `shell.execute` 的 spawn 移出主进程；worker 崩溃不影响 UI；通过 `MessageChannel` 回传 stdout/stderr/exit。AgentLoop 用 `MessengerPort` 通信。
  - 类比 Codex：它们的命令是直接 spawn，进程级强化放在 `process-hardening`；这一层对应到 Electron 主进程外的物理隔离

### 阶段 3 ── Windows AppContainer / 受限令牌（OS 级沙箱，需 native）

这是真正对应 Codex `windows-sandbox-rs` 的一层，**有原生依赖、需要分阶段评估**：

**方案 A：复用 Codex 已有的 Rust 二进制**
- Codex 提供了 `codex-windows-sandbox` helper：`windows-sandbox-rs/src/wrapper.rs::create_windows_sandbox_command_args_for_permission_profile(inner_command, ...) -> Vec<String>` 能把内层命令打包成强制走受限令牌的 spawn argv
- 集成方式：把 `codex-windows-sandbox` 编一个 exe 放进 `desktop/resources/`，`shell.execute` 改为 `execFile(codexWindowsSandboxExe, [...transformedArgv])`
- 优点：直接复用 Codex 已被审查的 ACL/WFP/code，跨桌面用户配 SID/deny-ACE 都已就绪
- 缺点：依赖非自身 Rust 二进制、需要保 license & 同步上游；同时 AppContainer/WRITE_RESTRICTED 会让 `GetObject` 失败 —— **所以只对 A 类工具上**，前面的分级保证 `shell.execute` 不连 Excel

**方案 B：自己写最小 native addon**
- 只做 `CreateRestrictedToken + AssignProcessToJobObject(jobObject)`
- Node 的 `node-ffi-napi` 或自写 N-API 模块
- 写白名单 ACL 沿用阶段 1 的 cwd 白名单转换成 `denied write ACE`
- 工作量比 A 大、能力比 A 弱，仅在不愿意引外部 exe 时考虑

**网络封禁**：
- 阶段 3 内、方案 A 引入 Codex 时顺带拿到 `WFP` 网络封禁（`wfp.rs::install_wfp_filters_for_account`）
- 自写方案 B 时退化为"策略层禁止网络 cmdlet"（阶段 1 默认规则已覆盖）

### 阶段 4 ── 进程加固（对应 `process-hardening`）

短平快，可独立做：

- 在 `windowsHide: true` 之外，往上 spawn PS 时强制清掉 `PSMODULEPATH / PROFILE / XDG_*` 危险 env（Codex `add_windows_sandbox_wrapper_setup_env_from_vars` 的 allowlist 思路：只保留 `USERNAME/USERPROFILE`，其余剥光）
- Excel 工具 spawn 路径 **不动**（要保持 PSModulePath 等 COM 环境）
- `shell.execute` 路径强制 `-NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -EncodedCommand <base64>`，并把 `-Command` 模式标记为 legacy（除非用户显式选）
- 主进程本身不需要 `pre_main_hardening`（Electron 已在它自己引擎做的范围内）

### 阶段 5 ── 统一审计日志（对应 `windows-sandbox-rs/src/logging.rs`）

- 新增 `desktop/electron/agent/sandbox/audit.ts`：
  - 每次 `shell.execute` 记 `command / workdir / decision / userId / outcome / duration / stdoutHead / stderrHead`
  - 落到 `session-log-<yyyyMMdd>.jsonl`，与 `sessionStore.ts` 的 rollout 同目录便于排查
  - `forbidden` 决策**永远**记一条告警事件并推给 ChatPage，让用户看到 AI 触发了被拒的命令

---

## 六、写入位置与改动清单

新增文件：

- `desktop/electron/agent/sandbox/parseCommand.ts`
- `desktop/electron/agent/sandbox/execPolicy.ts`
- `desktop/electron/agent/sandbox/defaultRules.ts`
- `desktop/electron/agent/sandbox/jobObject.ts`（阶段 2）
- `desktop/electron/agent/sandbox/shellRunnerWorker.ts`（阶段 2 utility worker）
- `desktop/electron/agent/sandbox/audit.ts`
- `desktop/electron/agent/sandbox/index.ts`（导出 `evaluateCommand(cmd, ctx)` 一站式入口）

改动文件：

- `desktop/electron/agent/toolRegistry/executors.ts:80 executeShellCommand` → 改为先走 `sandbox.evaluateCommand`
- `desktop/electron/agent/agentLoop/toolExecutor.ts:109 shouldRequireApproval` → forbidden/prompt 决策与审批机制对齐（prompt 即走审批）
- `desktop/electron/main-modules/ipcHandlers.ts` / `eventForwarder.ts`：审批对话框增 `justification` 字段、增加 `sandbox:config` IPC 暴露规则管理
- `desktop/electron/main-modules/settingsManager.ts`：扩 `execPolicy: PrefixRule[]` 与 `shellExecuteAllowedWorkroots: string[]`、`shellExecutionMode: "encoded"|"legacy"`
- `desktop/src/components/settings/` 新增 `ExecPolicySettings.tsx` 让进阶用户配置规则
- `desktop/src/components/chat/ToolConfirmDialog.tsx`：展示命令被命中的规则 + justification
- `desktop/electron/agent/systemPrompt.ts:158 permissionRules`：给模型加 "`shell.execute` 受 execpolicy 约束，破坏性命令会被拒" 的提示，让模型避免重试

可选：

- `desktop/resources/codex-windows-sandbox.exe`（阶段 3 方案 A 拷入）+ `LicenseRef-Codex` 声明
- `desktop/electron/agent/sandbox/windowsToken.ts`（阶段 3 方案 B）

---

## 七、分阶段里程碑

| 阶段 | 工作量 | 安全收益 | 依赖 | 风险 |
|------|--------|----------|------|------|
| 0 命令接口重构 | ~1d | 为后续铺路；顺手消除 PS 注入 | 无 | PS tokenize 边界 |
| 1 execpolicy | ~3d | 立即拦截破坏性命令；可灰可白 | 阶段 0 | 误杀要给逃生口（per-thread 暂时禁用） |
| 2 Job + utility worker | ~3d | 失败隔离、资源限制、超时强杀 | 阶段 0 | Electron `utilityProcess.fork` 与 IPC 上手 |
| 3 受限令牌（A 复用 codex exe / B 自写） | 5–10d | 真 OS 沙箱 | 阶段 1 | 需引入非自研二进制 / native；仅 A 类生效 |
| 4 加固 + 编码化 | ~1d | 防 PS profile 注入 | 阶段 0 | 模型感官不会变 |
| 5 审计 | 1d | 可观测、合规 | 各阶段 | 无 |

**先做组合：阶段 0 + 1 + 5 + 4**，可在 1 周内交付一个"任何模式都不会让 `Remove-Item -Recurse C:\` 跑出"的最小可用版本。阶段 2、3 视用户反馈渐推。

---

## 八、与 Codex 沙箱的能力对照

| 能力 | Codex | 本方案 |
|------|-------|--------|
| 跨平台抽象 | sandboxing crate（cfg 分平台） | 条件编译改为运行时分支（process.platform） |
| macOS seated sandbox | seatbelt .sbpl | 不在范围（项目主线 Windows） |
| Linux bubblewrap | bwrap + landlock + seccomp | 不在范围 |
| Windows 受限令牌 | windows-sandbox-rs | 阶段 3：方案 A 移植、方案 B 精简自写 |
| 命令策略 | execpolicy Starlark | 阶段 1：TS + JSON，不到 1000 行 |
| 文件系统策略 | FileSystemSandboxPolicy | 阶段 1：cwd 白名单；阶段 3：ACL |
| 网络封禁 | `--unshare-net` / WFP | 阶段 1：策略禁 cmdlet；阶段 3：WFP（A） |
| 进程加固 | pre_main_hardening | 阶段 4：env allowlist + 编码化 |
| 审批缓存 | ApprovalStore + 审批 | 复用现有 `alwaysAllowedTools`，不影响 |
| 私有桌面 | CreateDesktop | ❌（破坏 COM 回连，禁用） |
| 审计 | logging.rs + setup_error | 阶段 5 JSONL |

**不适用的项**（避免无意义照搬）：私有桌面 / `--unshare-user` / seccomp / macOS seatbelt / `deny-read ACL`（受限令牌的 deny-read 需要 elevated backend，第一版不做，靠策略层兜底）。

---

## 九、对 Excel 主链路的兼容性声明

- 阶段 1（execpolicy）：仅挡 `shell.execute` 和 `script.execute` 的纯数据分支，对 `range.* / sheet.* / workbook.* / vba.* / ui.*` 等 COM 路径**完全透明**
- 阶段 2（utility worker）：只把 `executeShellCommand` 搬到 worker，Excel 脚本 spawn 仍由主进程发起（必要时 `script.execute` 的 COM 分支也保留主进程）
- 阶段 3（受限令牌）：仅当工具属于 A 类才进入受限令牌 shell；B/C/D 类继续原 spawn
- `GetObject("Excel.Application/Ket.Application")` 的可用性在每次发版前做手工回归：开一个 Excel → 跑「读 A1 → 写 B1 → 跑宏 → 切表」用例，确保 COM 回连未被沙箱意外影响

---

## 十、验收标准

1. 默认策略下以下命令在**所有 permissionMode**下被拒并提示理由：`Remove-Item -Recurse -Force C:\`、`rm -rf /`、`format C:`、`Stop-Computer`、`reg delete HKLM\... /f`、`iex (New-Object Net.WebClient).DownloadString(...)`
2. `Invoke-WebRequest https://...`、`curl https://...`、`powershell -c ...` 在未"始终允许"下进入审批对话框，对话框显示命中规则与 justification
3. 改 `workdir` 到非白名单根时自动重定向到 `os.tmpdir()`，审计中可见 `workdir_redirected: true`
4. 子进程超时不再留尾，能强杀整条进程树（阶段 2 后）
5. 审计 JSONL 可被 `jq` 解析，能统计每日 forbidden 数
6. Excel 主链路回归用例（见上节）通过

---

## 十一、不在本期范围

- 把 execpolicy 升到完整 Starlark DSL / `host_executable` 元数据 / `network_rule`
- 实现 Windows elevated backend（deny-read ACL）—— 需要管理员权限 + elevated runner IPC，代价高、收益边际低（策略层已能挡大部分破坏性命令）
- macOS/Linux 沙箱（平台已是 fallback，用户面小）
- 给 AI 模型加上类似 Codex `additional_permissions` 的细粒度提权谈判
---

## 十二、落地状态（2026-06-26）

本方案已落地阶段 **0 + 1 + 4 + 5**，阶段 2、3 留作后续可演进项。详细记录见 `docs/session-2026-06-26-sandbox-and-prompt-refinement.md`。

| 阶段 | 状态 | 文件位置 |
|------|------|----------|
| 0 命令接口重构 | ✅ 已落地 | `electron/agent/sandbox/parseCommand.ts` / `index.ts` |
| 1 execpolicy 策略层 | ✅ 已落地 | `electron/agent/sandbox/execPolicy.ts` / `defaultRules.ts` |
| 2 Job + utility worker | ⏳ 留待后续 | — |
| 3 OS 级受限令牌 | ⏳ 留待后续 | — |
| 4 进程加固 | ✅ 已落地 | `electron/agent/sandbox/index.ts::runShellSpawn`（env allowlist + `-EncodedCommand`） |
| 5 审计 | ✅ 已落地 | `electron/agent/sandbox/audit.ts` |
| UI 安全策略设置页 | ✅ 已落地 | `desktop/src/components/settings/ExecPolicySettings.tsx` |
| UI 审批对话框理由展示 | ✅ 已落地 | `desktop/src/components/chat/ToolConfirmDialog.tsx` |

单元测试：`electron/agent/sandbox/sandbox.test.ts` 14 个用例全过。验收标准 §十 已对默认策略负面用例全部覆盖。
