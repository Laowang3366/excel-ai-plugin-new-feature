# 企业级上线深度审查报告

> 审查日期：2026-07-15
> 审查对象：`D:\excel-ai-plugin-new-feature`
> 基线分支：`master`
> 基线提交：`660c0597 fix(excel): harden formulas charts and pivots`
> 审查结论：**No-Go，当前版本不应进入企业生产发布**

## 0. 整改回写（2026-07-15）

> 状态口径：“已实现”表示代码与自动化负向测试已落地；涉及真实 Office、生产证书、服务器 ACL 或密钥轮换的项目，在完成外部验收前不标记为关闭。总体结论仍为 **No-Go**。

| ID | 当前状态 | 代码与测试证据 | 尚待验收/剩余工作 |
|---|---|---|---|
| C-01 | 已关闭 | `trustedIpc.ts`、`windowNavigationPolicy.ts`、Markdown 外链 `preventDefault`+`openExternal`；IPC sender/子 frame/协议混淆负向测试；Electron E2E `navigation-external`：真实 Markdown `<a>` 点击走 shell.openExternal 且主窗口 URL 不变、`location=` 与 `window.open` 被拒 | 无 |
| C-02 | 已关闭 | 审批缺失默认拒绝；危险/未知/删除/外传强制审批；线程+工具+operation+目标+TTL 授权；全工具元数据表驱动测试 | 无 |
| C-03 | 代码整改完成，外部动作未完成 | Provider、OCR、远程压缩凭据及自定义请求头由 `settingsSecrets.ts` 使用 `safeStorage` 保护，Renderer 仅接收掩码，设置 key 白名单及迁移测试 | 轮换所有可能暴露的真实凭据；签名私钥移出开发机；核验 Windows ACL |
| H-01/H-02 | 已实现 | 路径授权改为 realpath/最近存在父目录解析；移除系统 Temp 信任；知识库索引、删除、重建均重新授权；junction 逃逸测试 | Windows 安装包内路径选择实机回归 |
| H-03 | 已实现，待第三方联调 | `egressPolicy.ts` 统一默认关闭策略；OCR 本地优先；搜索/Embedding/发票模型抽取关闭时零请求；高置信凭据发送前阻断；返回目标服务与数据摘要；搜索响应按解压后 2 MiB 流式限额拒绝异常大正文 | 企业服务 allowlist 与真实 MinerU/搜索/Embedding 联调验收 |
| H-04 | 已实现 | 工具结果以结构化不可信 JSON 回送；系统提示明确外部数据边界；`memory.write` 必须引用当前轮用户原文并记录哈希/线程/turn/citation；未确认旧记忆不注入 system prompt；提示注入负向测试 | 红队场景与多模型遵循性实测 |
| H-05 | 已实现 | `outboundUrlPolicy.ts` 强制 HTTPS/精确本地 allowlist，阻断私网、DNS rebinding 与跨源重定向；Provider/Embedding 负向测试 | 企业本地模型 allowlist 配置验收 |
| H-06 | 已实现，待安装包实测 | 新安装默认 `%LOCALAPPDATA%` 用户隔离目录并自动迁移旧安装目录；全根目录 staging 复制、稳定文件 SHA-256、SQLite 实际打开校验、空目标原子改名、UNC 默认拒绝；失败恢复旧 store/runtime 并清理目标 | 安装包跨盘迁移、断电/磁盘满与旧目录停止写入实测 |
| H-07 | 已关闭 | `ExcelRangeWriteTransaction` 整区回滚 + 单元测试；Excel 365 矩阵 run 29457636677 / job 87494253080 验证第二项失败整区回滚；Codex 本地真实 WPS：`WENGGE_EXCEL_DYNAMIC_ARRAY_HOST=wps npm run test:excel-dynamic-array` 通过（`multiFormulaRollback=true`，`rollbackError=宿主无法通过 Formula2 写入公式`，SEQUENCE spill `[[1],[2]]`，C2:D2 写前/写后全等） | 无 |
| H-08 | 已关闭 | COM `Formula2` + 四类表达式 .NET 测试；`ExcelFormulaService` COM dynamic 边界修复（Formula* xUnit 55/55）；Codex 真实 Excel 365：四类表达式 spill/无 `@`/保存关闭重开；Codex 真实 WPS：owned PID 退出 + 新 Worker 重开 SEQUENCE spill | 无 |
| H-09 | 已实现 | Open XML 不再先删除后跳过占位；覆盖写入克隆原单元格属性并保留样式；样式/非 spill 数据测试 | 大型真实工作簿兼容性回归 |
| H-10 | 已关闭 | Worker 协议升至 v2；`excel.range.write` 结果运行时字段校验；旧结果负向测试；`scripts/smoke-office-worker-protocol.ts` + `npm run test:office-worker-protocol`（强制绝对路径 `WENGGE_OFFICE_WORKER_PATH`，断言 `protocol_mismatch`）；Codex 基线 `660c0597` v1 Worker exit 0（应用=2 Worker=1，残留 0），当前 v2 Worker 按预期 exit 1 | 无 |
| H-11 | 已实现 | Fastify 仅信任本地代理；Nginx 覆盖 XFF；轮换伪造 XFF 第 9 次触发 429 | 生产 Nginx 配置上线与告警验证 |
| H-12 | 部分完成 | Release 显式依赖可复用完整 CI；第三方 Actions 固定完整 SHA；Syft 版本固定并发布 SPDX SBOM；构建/发布权限隔离；两阶段 Authenticode 校验；产品站 release 目录只读；工作流静态防回归测试 | 配置受保护证书/HSM 与 Environment approval；隔离发布账户执行产品站 Ed25519 最终清单生成和端到端验签 |
| H-13 | 已实现，待安装包实测 | 64 KiB 流式 ZIP 解压与逐文件边界/哈希校验；每次启动 Renderer health pending/ack，30 秒超时及下次启动自动回退；签名清单支持吊销 ID 和最低安全序列并立即停用 | 打包 Renderer 白屏/硬崩溃、生产签名吊销清单端到端演练 |
| M-01 | 已关闭 | 模型与运行时共用严格 Schema；所有模型可见 Office operation 均为独立 `const` 分支，params 终端对象强制 `additionalProperties:false`；工作流变量/步骤结果路径限制 32 段、只读自有属性并拒绝原型链与保留段；审批前、执行前和全 operation 覆盖测试共同防回归 | 无 |
| M-02 | 已实现 | 聊天/恢复文本、附件、OCR 文件、Excel 矩阵、单元格文本、路径和 Base64 文件传输均设上限；settings 按 key 判别值 Schema；开放 JSON 参数限制深度、节点、集合和序列化字节；高成本通道按 sender 令牌桶限流 | 打包应用压力与正常高频交互误伤回归 |
| M-03 | 已实现 | AES-256-GCM 会话/SQLite/知识密文落盘；safeStorage 保护密钥；首次启动原子迁移与可轮换密钥；受管副本登记；路径迁移旧根清理/待擦除登记；导出登记；全副本擦除前预创建 replacement 密钥并单次原子清除旧密钥+bootstrap 删除证明 | 打包应用实机密钥轮换与跨盘迁移体验回归 |
| M-04 | 已实现 | 可见输出前保留瞬时故障重试；正文、推理或工具 item 发出后关闭透明整体重试；正文和工具事件断线测试 | 打包应用真实弱网/断网交互回归 |
| M-05 | 已关闭 | `test:excel-advanced-intent`：7 个负向调用在 COM 前拒绝、bridgeCalls=0、SHA-256 不变；Codex 本地 Excel PQ/透视/切片器全绿，三同名文件以 3 个 PID/FullName/instanceId 正确激活；WPS 基础操作、透视表、切片器全绿，Power Query 返回结构化 `power_query_unavailable`；结束后无宿主/Worker 残留 | 无 |
| M-06 | 已实现，待生产配置验收 | 下载统计 best-effort；90 天默认留存与 6 小时清理；IP 标识每 30 天轮换；UA/Referer 最小化迁移；故障与边界测试 | 上线环境确认周期配置并接入维护失败告警 |
| M-07 | 已实现 | 生产 secret 强度/格式校验；密码哈希脚本只从 stdin 接收；HTTP 配置仅重定向 HTTPS | 生产环境配置验收 |
| M-08 | 已关闭 | 测试依赖升级后 NuGet 高危漏洞为 0；`global.json`、NuGet lockfile；CI run `29490566136` / desktop job `87595394532` 依次通过 locked restore、`office:audit`、`office:test` 109/109 与 build | 无 |
| M-09 | 已关闭 | 删除工作表 `finally` 恢复用户原 `DisplayAlerts`；`test:excel-display-alerts` 经 Codex 本地 Excel+WPS 实机矩阵通过（false 成功恢复、结构保护失败 true 恢复、最后可见表失败 true 恢复；无宿主/Worker 残留） | 无 |
| M-10 | 已关闭 | `test:excel-dynamic-array` + `test:e2e-electron` + `office-matrix-and-e2e.yml`；E2E run 29456041445 / job 87489319942 success；Excel run 29457636677 / job 87494253080 success（spill/回滚/重开/Formula2）；WPS run 29458142234 / job 87495761434 success（formula2_spill_ok，SEQUENCE spill [[1],[2]]）；Runner `wengge-office-local-01` labels `wengge-office-excel-365`/`wengge-office-wps`，测试前宿主空、无残留 | 无 |
| M-11 | 已关闭（代码整改范围） | `SECURITY.md` 私密披露渠道与响应目标；`CONTRIBUTING.md` 双人审查和门禁；敏感路径 `CODEOWNERS`；基于运行代码的数据处理/远程流向/留存/导出与已登记副本擦除边界清单及防回归测试；应用层加密与已登记旧根/导出删除证明已落地；所有者决定按代码整改目标关闭本项 | 非代码后续（不阻塞本项关闭）：LICENSE/NOTICE、正式隐私主体/法律条款、真实第二审查人、GitHub ruleset 强制 |
| M-12 | 部分完成 | 产品站 SQLite 在线备份、SHA-256 元数据、完整性校验、安全恢复、14 份轮换与 timer；桌面日志/Office 备份/事务/工作流按 TTL、条目和字节配额周期清理并保护活动记录；用户可导出或擦除已登记活动根/旧根/应用导出副本并生成 bootstrap 删除证明；更新/Worker/迁移/后台登录/5xx/analytics 等稳定失败 `event` 字段已落地 | 生产负责人确认 RPO/RTO 并完成恢复演练；启用 timer；配置生产告警接收端；异机/外部未登记副本清理 |
| L-01 | 已关闭 | 生产模块超限已清零（`legacyOversized=0`）；全量受治理源码 Prettier 债务已清零（`legacyFormatting=0`）；`desktop/scripts/smoke-*.ts` 按规则排除行数扫描、仍受 Prettier 检查 | 无 |
| L-02 | 已关闭 | Node/.NET 版本事实源、当前/历史文档分层、实际 CI 门禁与防漂移测试；旧 code review/会话/实施计划已清理 | 无 |
| L-03 | 已实现 | 设置、Office 自动化和任务面板按需加载；当前首屏入口 448.41 KB；CI/打包共用 480 KiB budget | 打包应用冷启动与低速磁盘体验回归 |

## 1. 执行摘要

本次审查覆盖桌面应用、Agent Runtime、Electron IPC、.NET Office Worker、Excel/WPS 自动化、Open XML、产品站、更新发布链、数据治理、测试与工程规范。共归并出：

| 级别 | 数量 | 上线判断 |
|---|---:|---|
| Critical | 3 | 必须立即隔离并修复，任一未关闭都禁止上线 |
| High | 13 | 必须在生产发布前关闭 |
| Medium | 12 | 应进入发布前整改或形成有期限的正式风险接受 |
| Low | 3 | 纳入近期工程治理 |

原始审查发现中最紧急的阻断项是：

1. 模型输出的外链可把 Electron 主窗口导航到远端页面，而远端页面继续获得完整 preload 和高权限 IPC。
2. “全部确认”模式实际无条件自动批准宏、删除和未知工具；“始终允许”还会跨线程、跨工作簿扩大授权。
3. 工作区存在明文凭据材料，应用设计也会把 API Key、Token 和自定义请求头保存为普通 JSON 并完整下发 Renderer。
4. Excel 公式写入仍存在失败后部分提交、动态数组漏判和 Open XML 数据/样式破坏风险。
5. Windows 安装包未建立 Authenticode 签名门禁，Release、热补丁和产品站发布目录的供应链隔离不足。

自动化测试全部通过说明当前代码具有较好的基础回归能力，但安全测试已经把部分危险行为固化为“预期”，因此“测试全绿”不能作为上线依据。

## 2. 范围、方法与限制

### 2.1 审查范围

- `desktop/src/`：React Renderer、Zustand 状态、聊天和设置界面。
- `desktop/electron/`：Electron 主进程、preload、IPC、Agent Runtime、知识库、OCR、文件与更新服务。
- `desktop/dotnet/Wengge.OfficeWorker/`：COM、Open XML、Excel/WPS、Worker 协议和 STA 调度。
- `product-site/`：Fastify 产品站、后台认证、下载统计、Nginx/systemd 配置和发布脚本。
- `.github/workflows/`、`desktop/package.json`：CI、桌面打包和发布链。
- 当前架构、开发、Office、发布和部署文档。

### 2.2 执行方法

- 静态审查运行代码、配置、测试和发布脚本。
- 检查 Electron 信任边界、IPC 来源和输入校验、路径授权、凭据生命周期、网络出站和 Agent 工具审批。
- 检查 Office 数据一致性、COM 状态恢复、动态数组、图表/透视表验证和 Worker 协议兼容性。
- 执行 Node 与 NuGet 依赖漏洞检查。
- 执行桌面端、.NET Worker、产品站测试和构建门禁。
- 对产品站代理限流进行定向对抗验证。
- 只检查敏感文件的存在、模式和 ACL；**没有读取或回显任何真实密钥、Token 或密码值**。

### 2.3 本次未执行

- 未启动真实 Excel、WPS、Word 或 PowerPoint，避免干扰用户现有 Office 进程。
- 未运行 NSIS 完整安装、升级、卸载和 Windows SmartScreen/Authenticode 实机验证。
- 未对生产服务器执行渗透、负载、故障注入或灾难恢复演练。
- 未执行真实第三方 OCR、模型、搜索和更新下载请求。

## 3. 质量门禁结果

| 检查 | 结果 | 说明 |
|---|---|---|
| Desktop `npm audit --audit-level=high` | 通过 | 0 个高危 npm 漏洞 |
| Desktop ESLint | 通过 | 本轮基线通过 |
| Desktop TypeScript typecheck | 通过 | Renderer 与 Electron 主进程通过 |
| Desktop Vitest | 通过 | 整改后 211 个测试文件、1114 项测试全部通过 |
| Desktop Vite build | 通过 | Renderer 首屏入口 448.41 KB（437.91 KiB），9 个异步 chunk；480 KiB entry budget 通过 |
| Desktop `format:check` | 通过 | 受治理源码全量 Prettier 一致；`governance:check` 为 legacyFormatting=0、legacyOversized=0 |
| .NET Worker test | 通过 | CI run `29490566136` / desktop job `87595394532`：locked restore、NuGet audit 与 109 项 xUnit 测试全部通过 |
| NuGet vulnerability scan | 通过 | Worker 与测试项目均未发现已知漏洞包 |
| Product-site `npm audit` | 通过 | 0 个高危 npm 漏洞 |
| Product-site test | 通过 | 14/14 通过，包含 XFF 绕过、统计故障、周期标识、留存清理、旧数据迁移及备份恢复 |
| Git 历史凭据扫描 | 通过 | 未发现敏感目录或高置信真实凭据被提交 |
| 真实 Office/WPS 冒烟 | 部分通过 | **M-10**、**M-09**、**M-05**、图表专项、legacy CSE 与 **H-10** 已通过。CSE：Codex `WENGGE_EXCEL_DYNAMIC_ARRAY_HOST=excel|wps` + `test:excel-dynamic-array` 全绿。H-10：Codex 基线 `660c0597` v1 Worker + 绝对路径 `WENGGE_OFFICE_WORKER_PATH` + `test:office-worker-protocol` exit 0（`protocol_mismatch`，应用=2 Worker=1）；当前 v2 Worker 按预期 exit 1 |

## 4. Critical — 上线硬阻断

### C-01 外部页面可继承完整 Electron preload 权限

> 整改完成：主窗口 `will-navigate`/`will-redirect` 拒绝非应用源导航；`setWindowOpenHandler` 默认 `deny`；Markdown `a` 统一 `preventDefault`，仅规范化 `http:`/`https:` 后经主进程 `shell.openExternal` 打开；敏感 IPC 经 `trustedIpc` 校验 sender frame。单元测试覆盖导航策略与远端/子 frame/协议混淆负向路径。Electron E2E `navigation-external` 通过真实 Agent 流式事件渲染 Markdown 外链并点击生产 `<a>`，断言 `shell.openExternal` 收到目标 URL 且主窗口 URL 不变；同场景断言 `location=` 与远端 `window.open` 不新增 BrowserWindow、不改变主页面 URL。以下“证据”为整改前基线。

**证据**

- `desktop/src/components/chat/MarkdownContent.tsx:17-45` 使用 `react-markdown` 默认链接行为，没有统一拦截并交给系统浏览器。
- `desktop/electron/main-modules/windowManager.ts:65-77` 为主窗口配置 preload，且 `sandbox:false`。
- 同文件没有 `will-navigate`、`will-redirect` 或 `setWindowOpenHandler` 防护。
- `desktop/electron/preload.ts:10` 向每次加载的页面暴露完整 `electronAPI`。
- `desktop/electron/main-modules/ipcHandlers.ts:302-304` 的 `settings:getAll` 返回完整设置。
- 生产代码约有 68 个 `ipcMain.handle` 和 1 个 `ipcMain.on`，未发现统一的 sender origin 校验。

**影响**

模型、网页、OCR 或知识库内容只需生成一个恶意 HTTPS 链接。用户点击后，主窗口可能直接导航到攻击者页面；该页面可调用 `window.electronAPI`，读取凭据和会话、修改权限模式和 Provider、访问知识库与文件、触发 Agent、Office 或更新操作。这是一条从“不可信文本”到“本地高权限能力”的完整攻击链。

**解决方案**

1. 主窗口生产环境只允许加载应用自身固定 `file:` 入口；开发环境只允许精确的本地开发源。
2. 在 `will-navigate`/`will-redirect` 中拒绝所有非应用导航。
3. `setWindowOpenHandler` 默认返回 `deny`。
4. Markdown 的 `a` 组件统一 `preventDefault`，只接受规范化后的 `http:`/`https:`，并通过主进程 `shell.openExternal` 打开。
5. 所有敏感 IPC 校验 `event.senderFrame.url`、主窗口身份和顶层 frame；默认拒绝未知 sender。
6. 开启 `sandbox:true`，按域拆分最小 preload，敏感凭据永不进入 Renderer。

**验收标准**

- 点击恶意链接后系统浏览器打开，主 `BrowserWindow` URL 不变。
- `window.open`、重定向、表单导航和 `location=` 均不能离开应用源。
- 使用伪造远端 sender 调用设置、文件、知识、Agent、Office、更新 IPC 时全部返回授权错误。
- 自动化测试覆盖生产 `file:`、开发 localhost、子 frame、重定向和协议混淆。

**验收状态**：C-01 已关闭。生产导航/IPC 负向测试 + Electron E2E（Markdown 点击 openExternal、blocked location、blocked window.open）已落地。`sandbox:true` 与 preload 按域最小化不在本项关闭范围内，属后续硬化。

### C-02 工具审批与“始终允许”整体 fail-open

**证据**

- `desktop/electron/agent/core/agentLoop/toolApproval.ts:20-33` 在 `confirm_all` 下直接返回 `false`，即所有工具均无需审批。
- 同文件 `39-52` 在没有审批回调时默认 `{ approved: true }`。
- `desktop/electron/agent/core/agentLoop/agentLoop.test.ts:1245-1262` 明确把宏、删除和未知工具自动批准固化为测试预期。
- `desktop/src/i18n.ts:125-130,330-335` 将该模式显示为“全部确认 / Confirm all”，与实际“无需确认”相反。
- `requiresApproval`、`isFileDeletion` 元数据没有成为统一审批决策依据。
- `toolApproval.ts:14` 使用进程级 `Set<string>` 记录“始终允许”；只按工具名授权。
- `sheet.operation` 等单一工具同时包含新增、移动、复制、删除，多种操作共享同一个授权键。

**影响**

宏写入、宏执行、清空区域、删除工作表和未知工具可无提示执行。用户曾对 `sheet.operation(add)` 选择“始终允许”后，后续 `sheet.operation(delete)` 也可能跨线程、跨工作簿自动执行。

**解决方案**

1. 审批策略改为 fail-closed；未配置审批通道时拒绝执行需要审批的工具。
2. 统一消费 `riskLevel`、`requiresApproval`、`isFileDeletion`、operation、目标对象和网络外传属性。
3. 未知工具、宏、删除、数据外传和更新操作在所有模式下都必须明确审批。
4. 如保留完全自动模式，应命名为“完全访问（无需确认）”，默认关闭，并增加二次确认和企业策略开关。
5. “始终允许”授权键至少包含工具、operation、目标工作簿/文件、线程、风险类别和有效期；危险操作禁止永久授权。

**验收标准**

- 宏、删除、未知工具在所有普通模式下必定弹出审批。
- 缺少审批 callback 时请求被拒绝而不是自动通过。
- `add` 的授权不能覆盖 `delete`，不能跨工作簿、跨线程或无限期复用。
- 新增表驱动测试枚举所有模型可见工具和每个 operation 的审批结果。

### C-03 明文凭据、更新私钥与宽松 ACL 构成现实泄露面

> 代码整改进展：设置 key 已改为白名单与按 key Schema 校验；Provider API Key、自定义请求头、OCR Token 和远程压缩 API Key 统一使用 Windows `safeStorage` 加密，启动时迁移旧明文值，Renderer 只接收掩码，掩码回写保留原密文，Agent 仅从主进程取得解密后的运行时配置。以下“证据”为整改前基线；真实凭据轮换、发布私钥迁出开发机和 Windows ACL 验收仍属于未完成外部动作。

**证据**

- `desktop/electron/main-modules/settingsManager.ts:32-77` 使用未配置 `encryptionKey` 的 `electron-store` 保存 Provider 配置和 MinerU Token。
- `settingsManager.ts:281-303` 会读取 `apiKey` 和 `customHeaders`。
- `ipcHandlers.ts:240-248,302-304` 允许 Renderer 按任意字符串 key 读写设置并获取全量 store。
- 本机存在且被 Git 忽略的敏感材料：
  - `desktop/data/settings/excel-ai-settings.json`
  - `product-site/.local/admin-password.txt`
  - `desktop/.secrets/update-private.pem`
- 上述文件继承的 ACL 允许普通 `BUILTIN\Users` 读取，并允许 `Authenticated Users` 修改。

审查只确认了文件存在、非空敏感模式和 ACL，没有读取或记录其值。更新私钥是否为当前生产私钥尚未确认，因此应按“可能已暴露”处置。

**影响**

本机其他普通用户、被攻陷的 Renderer、恶意插件或同权限进程可窃取模型 Key、OCR Token、后台密码；若更新签名私钥曾用于生产，攻击者可伪造受信任更新。

**解决方案**

1. 立即盘点并轮换所有曾使用过的 API Key、MinerU Token、后台密码和签名密钥。
2. Provider、OCR 与远程压缩凭据迁移到 Electron `safeStorage`、Windows Credential Manager 或企业密钥库；Renderer 只持掩码和 `secretRef`。
3. 更新签名私钥只能存在于受限发布环境/HSM/受保护 CI Secret，不得留在开发工作区。
4. 缩紧目录和文件 ACL；敏感文件只允许当前服务账户、Administrators 和 SYSTEM。
5. 引入提交前与 CI secret scanning，并检查构建产物、日志、会话和崩溃报告。

**验收标准**

- 设置 JSON、IPC `getAll`、Zustand、日志和会话中扫描不到明文 canary secret。
- 旧凭据全部失效；新凭据完成最小权限和轮换记录。
- 若轮换更新签名密钥，客户端公钥迁移和回滚方案经过端到端验证。
- 不同 Windows 普通用户不能读取或修改凭据材料。

## 5. High — 生产发布前必须关闭

### H-01 知识库 IPC 可绕过路径授权读取任意本地文件

> 整改结果：知识库文件/目录索引、删除与全量重建均通过 `PathAuthorizer` 重新校验；知识库 IPC 已从 Agent/线程生命周期注册器分离到 `ipcKnowledgeHandlers.ts`，独立持有运行时初始化、迁移阻断和路径授权依赖。以下“证据”为整改前基线。

**证据**

- `desktop/electron/agent/interaction/ipcAgentHandlers.ts:391-426` 仅校验路径是字符串，随后直接调用 `indexFile/indexFolder`。
- `desktop/electron/agent/knowledge/knowledgeIndexer.ts:69-115` 直接解析文件，`222-247,297-315` 可递归读取目录。
- 该 IPC 没有注入 `PathAuthorizer`；已索引内容可通过知识搜索读取，并可能发送给 embedding Provider。

**影响**

被攻陷或越权的 Renderer 可以索引用户未授权的本地文件、配置目录或整个磁盘，再通过搜索结果提取内容。

**解决方案与验收**

- 文件索引必须绑定用户选择产生的 canonical grant；目录索引必须绑定授权根。
- `reindexAll` 对历史来源重新校验授权，不得把旧记录当永久权限。
- 使用 `realpath.native`，拒绝 junction/reparse-point 逃逸。
- 验收应覆盖未授权 `C:\`、用户配置目录、UNC、junction 逃逸均拒绝；经对话框授权的文件正常工作。

### H-02 路径授权采用词法比较且默认信任整个系统 Temp

**证据**

- `desktop/electron/main-modules/ipcPathSecurity.ts:16-23` 只使用 `path.resolve().toLowerCase()`。
- `ipcPathSecurity.ts:36-41` 把 `os.tmpdir()` 作为永久可信根。
- 根授权后可通过文件 IPC 读取或移入回收站，未解析 Windows reparse point 的真实目标。

**影响**

应用可访问其他程序写入系统 Temp 的文件；攻击者可在可信目录放置 junction，将读、写、删除操作引向授权根之外。

**解决方案与验收**

- 移除系统 Temp 根级授权，只授权应用自身创建的精确临时目录/文件并设置过期时间。
- 读取使用最终 `realpath`；写入检查最近存在父目录的 `realpath`，并拒绝越界 reparse point。
- 自动化测试创建 junction 指向授权根外，确认读、写、删除全部失败。

### H-03 OCR 与联网搜索缺乏统一数据外传治理

> 整改结果：新增统一的远程数据处理开关，默认关闭；OCR 本地优先，搜索、远程 OCR、发票模型抽取和 Embedding 均在网络请求前执行策略与高置信凭据检查，并通过工具审批呈现外传风险。Web 搜索进一步将工具参数/策略编排与搜索源适配分离，并对所有 JSON/HTML 正文施加解压后 2 MiB 流式上限，声明长度或实际读取超限时在解析前拒绝。以下“证据”为整改前基线。

**证据**

- `desktop/electron/main-modules/ipcOcrHandlers.ts` 与 `desktop/electron/agent/tools/executors/ocrExecutors.ts` 默认先执行本地解析，仅把未解决文件依次交给 MinerU Token 和免费 MinerU Agent。
- `desktop/electron/main-modules/mineruOcr.ts:30-50,231-242` 会把完整本地文件上传到 MinerU 或其签名上传地址。
- 发票 OCR 文本还可能继续发送给当前模型 Provider。
- `desktop/electron/agent/tools/registry/web.ts:3-21` 把 `web.search` 标记为 safe 且无需审批。
- `webSearchExecutors.ts:57-121` 可能把查询依次发送给 Tavily、Bing、SerpAPI、百度、搜狗、360 和 DuckDuckGo。

**影响**

机密文档、单元格内容或提示注入诱导生成的敏感搜索词可能在用户未充分知情时发送给多个第三方。

**解决方案与验收**

- 建立统一 egress policy：企业模式默认本地、服务 allowlist、目的地显示、逐文件/逐查询预览和 DLP。
- OCR 本地能力优先；启用第三方上传必须由管理员策略和用户明确授权。
- `web.search` 改为“数据外传”风险，而不是 safe；离线模式不向模型暴露该工具。
- 用 canary 文档和搜索词验证：任何 HTTP 之前必须被阻断或明确确认，离线策略下网络请求数为 0。

### H-04 外部内容可形成持久化系统提示注入

> 整改结果：工具结果现以结构化不可信数据回送；长期记忆只接受当前轮用户原文证据，记录来源与哈希，且只有用户确认记录可进入 system prompt。以下“证据”为整改前基线。

**证据**

- `desktop/electron/agent/prompts/templates/system/security.zh-CN.md` 没有明确把网页、OCR、文档和工具结果定义为不可信数据。
- `desktop/electron/agent/shared/messageBuilder.ts:131` 原样回送工具结果。
- `desktop/electron/agent/tools/registry/memory.ts` 将 `memory.write` 标记为 safe、无需审批。
- `desktop/electron/agent/core/agentLoop/buildStreamParams.ts:25` 会把长期记忆拼入 system prompt。

**影响**

恶意文档或网页可诱导模型把攻击指令写入长期记忆，使其在后续会话持续获得更高优先级影响，并进一步触发外传或破坏性工具。

**解决方案与验收**

- 对所有外部内容增加来源、信任级别和不可执行边界；工具结果只作为数据块处理。
- 禁止外部内容直接触发长期记忆；记忆写入要求明确用户来源或审批。
- 结构化存储事实与偏好，不保存自由形式“指令”；过滤伪 system/user 标签和越界格式。
- 恶意 OCR/网页指令不能创建记忆；包含换行、伪系统指令的记忆不得改变工具审批或系统策略。

### H-05 Provider `baseUrl` 可形成 SSRF 和凭据转发

**证据**

- `desktop/electron/shared/ipcSchemas.ts:184-195` 仅把 `baseUrl` 校验为普通字符串。
- `desktop/electron/main-modules/ipcAiHandlers.ts:12-25,52-89` 直接拼接 URL 并携带用户 API Key 发起主进程请求。
- Provider 客户端同样使用可配置 `baseUrl`，缺少协议、私网、DNS 和 redirect 复验。

**影响**

恶意 Renderer 或错误配置可探测本机/内网服务，也可把真实 API Key 作为 Authorization 或 `x-api-key` 发送到攻击者地址。

**解决方案与验收**

- 只允许 `https:`；本地模型例外必须是管理员显式 allowlist。
- 拒绝 URL 凭据、非标准协议、环回/链路本地/私网/云 metadata 地址；解析 DNS 后检查所有地址。
- 禁止自动跨域 redirect，或对每次 redirect 重新执行完整策略。
- 测试 DNS rebinding、IPv6、十进制 IP、redirect 到私网和恶意 URL；API Key 不得到达未授权主机。

### H-06 数据目录默认、迁移和运行时重绑定不完整

> 整改结果：默认目录已切换为用户隔离路径；旧安装目录通过同父目录 staging 后原子迁入；手动迁移覆盖完整数据根、逐文件校验并验证 SQLite，可在 Runtime 重建失败时回滚。以下“证据”为整改前基线。

**证据**

- `desktop/electron/main-modules/settingsDataPath.ts:15-21,59-71` 优先使用可写的安装目录 `data/`，而非用户隔离目录。
- `settingsManager.ts:166-205` 迁移只复制 `sessions`、`knowledge`、`logs`。
- 未迁移 `office-backups`、`office-automation/workflows` 和 `office-automation/transactions`。
- `desktop/electron/agent/runtime/agentRuntime.ts:181-187` 在 Runtime 首次创建时固化 Office 自动化和备份路径；迁移逻辑只重绑部分 store。
- `app:migrateDataPath` 接收任意路径字符串，未限制 UNC/网络目标；目录选择对话框不是 IPC 信任边界。

**影响**

凭据、会话和日志可能落到 ACL 宽松的安装目录。迁移后新旧目录同时被写入，事务/备份与会话分裂；恶意 Renderer 还可把数据迁移到网络共享造成外传。

**解决方案与验收**

- 默认只使用 `%LOCALAPPDATA%`/`app.getPath("userData")` 下的用户隔离目录。
- 用版本化迁移清单覆盖所有数据域，复制后校验数量、哈希和 SQLite integrity，再原子切换。
- 迁移期间停止并销毁 Agent Runtime，切换后重新创建所有路径依赖。
- 默认拒绝 UNC 和远程卷；如企业确需网络目录，要求管理员策略、加密和显式提示。
- 验收确认迁移后旧目录无新增写入，所有 Office 备份/事务/工作流可恢复，失败时完整回滚。

### H-07 Excel `range.write` 失败会留下部分提交

> 整改完成：`ExcelRangeWriteTransaction` 在任一公式失败时恢复整区，并有第二项失败/回滚失败单元测试。生产写入对 `plan.Formulas` 顺序逐格执行，故合法第一项会先于失败第二项进入写入序列。Excel 365 隔离矩阵 run 29457636677 / job 87494253080 已验证 `C2:D2` 多公式第二项 overlong 失败后整区与写前一致。`smoke-excel-dynamic-array` WPS 路径同样做快照→真实 `writeRange`→必须失败→逐值回读；Codex 本地真实命令 `WENGGE_EXCEL_DYNAMIC_ARRAY_HOST=wps npm run test:excel-dynamic-array` 通过：`multiFormulaRollback=true`，`rollbackError=宿主无法通过 Formula2 写入公式`，Formula2 SEQUENCE spill `[[1],[2]]`，C2:D2 写前/写后全等。以下“证据”为整改前基线。

**证据**

- `desktop/dotnet/Wengge.OfficeWorker/Excel/ExcelRangeService.cs:67-81` 先整块执行 `Value2 = BulkValues`，再逐单元格写公式。
- `ExcelRangeWritePlan.cs:26-31` 会把公式位置替换为 `null`。
- `ExcelFormulaWriter.cs:43-77` 任一 setter 或回读失败都会抛错，但调用方没有原值快照和回滚。

**影响**

旧版 Excel/WPS 缺少 `Formula2`、公式无效或矩阵中第二个公式失败时，工具虽然返回失败，普通值已写入、原公式单元格已清空，前面成功的公式也会保留。

**解决方案与验收**

- 写入前预检公式 API 能力并保存目标范围值、公式类型和公式文本。
- 任一写入或回读失败时，在同一 STA 内恢复完整目标范围。
- 构造“第二个公式失败”的测试，断言整个区域与写前完全一致；Excel 和 WPS 都要真实冒烟。

**验收状态**：H-07 已关闭。代码事务回滚、Excel 365 矩阵实机与 Codex 本地真实 WPS 多公式第二项失败整区回滚均已通过。

### H-08 动态数组识别仍依赖函数白名单

> 整改完成：现代公式统一 `Formula2`，分类识别范围引用/范围运算/IF/TRANSPOSE 等表达式；COM 与 Open XML 有对应单元测试。`excel.formula.context` 将 COM Formula2/Value2 在获取边界隔离为 `object?` 再 `ToRows`，公式前缀用静态 `string.StartsWith("=", Ordinal)`，避免 dynamic binder 崩溃（ExcelFormula* xUnit 55/55）。`smoke-excel-dynamic-array` Excel 路径写入四类表达式 spill，`formula.context` 断言宿主公式无任何 `@`，保存关闭本任务拥有进程后重开再验。WPS 路径：SEQUENCE + H-07 回滚后 save→`disposeOfficeWorker`→`assertOwnedStopped(ownedWpsPids)`→新 Worker 重开读 spill。Codex 真实 Excel 365 与 WPS 均已绿。以下“证据”为整改前基线。

**证据**

- `ExcelFormulaClassification.cs:18-41` 仅根据现代函数名白名单判断动态数组。
- `ExcelFormulaWriter.cs:45-50` 漏判后会走 `Formula` 而非 `Formula2`。
- Open XML 路径也复用该分类，漏判后不会写动态数组 metadata。
- `=A1:A3`、`=A1:A3*2`、`=IF(A1:A3>0,A1:A3,"")`、`=TRANSPOSE(A1:A3)` 都可能 spill，但不会命中现有白名单。

**影响**

原始“动态数组被 `@` 降级”问题只覆盖了 `FILTER/LET/...` 等函数入口，表达式型数组仍可能回归。

**解决方案与验收**

- 支持 `Formula2` 的宿主对现代公式优先使用 `Formula2`，不要尝试仅凭函数名猜测数组语义。
- 不支持 `Formula2` 的宿主只对明确标量公式使用 `Formula`，其余明确失败或要求显式兼容策略。
- 为上述四类表达式补 COM/Open XML 测试，并在 Excel 365 实机确认没有 `=@` 且 spill 范围正确。

**验收状态**：H-08 已关闭。Codex 真实 Excel 365：四类表达式宿主公式恰好为 `=B2:B4` / `=B2:B4*2` / `=IF(B2:B4>85,B2:B4,"")` / `=TRANSPOSE(B2:B4)`，全程无 `@`；spill `[[90],[80],[70]]` / `[[180],[160],[140]]` / `[[90],[""],[""]]` / `[[90,80,70]]`；保存关闭重开后一致；H-07 回滚 true。Codex 真实 WPS：owned PID 退出后新 Worker 重开，SEQUENCE spill 前后 `[[1],[2]]`，`saveCloseReopenSpill=true`。生产 `ExcelFormulaService` dynamic 边界修复与 55/55 测试已落地。

### H-09 Open XML 动态公式写入会误删单元格并丢失样式

**证据**

- `OpenXmlExcelActionService.cs:170-172` 只要矩阵含任一动态公式，就对整个矩阵设置 `hasDynamicFormula`。
- `186-190` 先删除现有 Cell，再对所有空字符串直接 `continue`。
- 非空单元格也删除后重建，没有保留原 `StyleIndex` 等属性。

**影响**

一个动态公式可导致同批次中无关空字符串单元格消失，并丢失数字格式、边框、填充和其他单元格 metadata。

**解决方案与验收**

- 不用空字符串隐式表示 spill 占位符；增加明确的 blank/clear/spill-region 意图。
- 只处理对应动态锚点声明的区域，更新值时保留现有样式和不相关属性。
- 测试同一矩阵内“动态锚点 + 带样式非 spill 空字符串”，后者必须存在且样式不变。

### H-10 Worker 合约变化但协议版本仍为 1

**证据**

- `desktop/electron/agent/officeWorker/officeWorkerClient.ts:7` 和 `OfficeWorkerHost.cs:14` 都固定 `ProtocolVersion = 1`。
- 新请求增加 `legacyCse`，`RangeWriteResult` 新增强制计数字段，但协议没有升级。
- TypeScript 泛型没有对 Worker 响应进行运行时 schema 校验。

**影响**

残留旧 Worker、安装覆盖失败或 `WENGGE_OFFICE_WORKER_PATH` 指向旧版本时仍可通过握手，并静默忽略新语义，重新引入 `@` 或返回缺失字段。

**解决方案与验收**

- 协议升级到 v2，或进行 capability negotiation。
- 每个 RPC 结果执行运行时 schema 校验。
- 新客户端连接旧 Worker 时必须明确返回 `protocol_mismatch`，不能继续写入。

### H-11 产品站后台限流可通过伪造 X-Forwarded-For 绕过

**证据**

- `product-site/src/server.mjs:16` 使用 `trustProxy:true`。
- Nginx `plugin.shelelove.top.conf:46-52` 使用 `$proxy_add_x_forwarded_for`，会保留客户端提供的 XFF。
- 后台登录限流按请求 IP 计算。
- 定向验证：固定 XFF 时第 9 次登录返回 429；每次轮换 XFF 时连续 10 次均返回 401，未触发限流。

**影响**

攻击者可绕过 8 次/15 分钟限制持续尝试后台密码。

**解决方案与验收**

- Nginx 覆盖而不是追加客户端 XFF：只把可信边界上的 `$remote_addr` 传给应用。
- Fastify 仅信任明确的本地反向代理地址/跳数。
- 后台登录叠加账户级、设备/会话级限流、指数退避和告警。
- 轮换任意客户端 XFF 也必须在第 9 次前触发统一限流。

### H-12 桌面 Release 与下载资产的供应链隔离不足

> 整改结果：仓库侧已完成 CI 显式依赖、第三方 Action/Syft 不可变版本、SPDX SBOM、构建发布权限隔离、两阶段 Authenticode 校验及产品站 release 目录只读。生产证书/HSM、Environment approval 和产品站 Ed25519 最终清单端到端验签仍需外部环境验收，因此保持“部分完成”。以下“证据”为整改前基线。

**证据**

- `desktop/package.json:89-100` 的 Windows 构建没有 Authenticode 证书和验签配置。
- `.github/workflows/release-desktop.yml:9-29` 整个构建任务持有 `contents:write`，使用可变 action tag，且未执行 npm audit、.NET tests 或最终签名清单验签。
- Tag Release 与完整 CI 没有显式依赖关系，可并行运行。
- `product-site/deploy/wenge-product.service:18` 允许 Web 服务账户写 `/opt/wenge-product/releases`。

**影响**

用户无法通过 Windows 发布者签名确认安装包来源；被攻陷的构建步骤或产品站进程可以影响发布资产，发布流程也可能绕过更完整的门禁。

**解决方案与验收**

- 使用企业代码签名证书对 EXE/卸载器签名并加 RFC3161 时间戳；私钥置于 HSM 或受保护签名服务。
- 构建与发布拆为不同 job；构建 job 只读，发布 job 只下载已验证 artifact，并受 Environment approval 保护。
- Actions 固定完整 commit SHA；Release 必须依赖通过的 CI，并运行 npm/NuGet audit、.NET tests、SBOM、最终清单验签。
- 产品站运行账户对 release 目录只读；发布由独立账户原子切换版本目录。
- `Get-AuthenticodeSignature` 必须为 `Valid`，失败 CI 不得上传 Release。
- 自动化测试必须拒绝工作流中新出现的可变第三方 Action 引用，并确认 Release artifact 与 GitHub Release 均包含版本化 SPDX SBOM。

### H-13 热补丁缺少防重放、启动复验和有界解压

> 整改结果：补丁现执行流式有界解压、启动文件复验、序列防回退、签名策略吊销及 Renderer 健康确认回滚。以下“证据”为整改前基线。

**证据**

- `hotPatchManager.ts:93-99` 只检查基础版本和归档 SHA，没有补丁序列单调性或吊销机制。
- `146-163` 启动激活只检查状态文件、路径和存在性，不重新校验已安装文件哈希。
- `101` 使用 `unzipSync(await readFile(...))`，在检查 200 MB 解压后大小前已把归档和所有解压内容载入内存。
- 主窗口只在 `loadFile` 失败时回退；React 运行时白屏或启动后异常不会触发健康回滚。

**影响**

旧的有效补丁可被重放；已安装补丁被本地篡改后仍会加载；ZIP bomb 可先耗尽内存；逻辑损坏补丁可能长期白屏。

**解决方案与验收**

- 补丁清单增加单调版本、发布时间、过期时间、最低安全版本和吊销列表。
- 启动前按签名清单复验每个文件哈希。
- 使用流式、限条目、限单文件、限总大小和限压缩比的解压器。
- Renderer 启动成功后向主进程发送有超时的 health acknowledgement，否则自动回滚。
- 测试旧补丁重放、文件篡改、超高压缩比和白屏补丁均被拒绝或自动回滚。

## 6. Medium — 发布前整改或正式风险接受

### M-01 模型可见工具 Schema 与工作流占位符边界已关闭

> 整改完成：规范化 JSON Schema 同时提供给模型并用于审批前、executor 前两次运行时校验。所有模型可见 `office.action.inspect/apply/validate` 和 `office.workflow.run.steps` operation 都是独立 `const` 分支，params 的每个终端对象都强制 `additionalProperties:false`；缺少必填项、错误类型/枚举、未知嵌套字段和资源超限会在进入 Worker 前拒绝。覆盖测试直接遍历最终模型工具定义，若任何 operation 重新合并为通用 enum 分支或开放 params，门禁立即失败。

工作流变量保留任意业务 JSON 的必要能力，但不再形成路径逃逸：顶层键、步骤 ID、并行组和条件选择器有格式与长度约束；`{{vars.customer.name}}`、`{{vars.files.0}}` 和步骤结果 `dataPath` 最多 32 段。路径解析逐段验证，只读取对象或数组的自有属性，拒绝空段、`__proto__`、`prototype`、`constructor` 和原型链继承值。全局 JSON 深度、节点、集合、字符串和序列化字节预算继续生效，因此无需为不同企业模板预设固定业务字段。

**验收状态**：全 operation 严格分支、未知 params、变量键、选择器、条件路径、数组索引、原型链和超长路径均有自动化正负测试；M-01 关闭。

### M-02 IPC 缺少统一资源预算

> 整改进展：聊天正文 50,000 字、恢复上下文 200,000 字、附件和 OCR 文件各 20 个；Excel 直接写入最多 10,000 行、16,384 列且总计 100,000 个单元格，单元格文本不超过 32,767 字；Base64 与文件回读统一限制为 50MB，并分别在 `Buffer.from` 和 `readFile` 前拒绝。`settings:set` 已按 key 校验 Provider、固定文件夹、压缩配置、枚举、布尔值和数值范围；模板变量顶层限制 128 个安全键，所有开放 JSON 参数统一限制 16 层、100,000 节点、单字符串 1 MiB、数组 20,000 项、单对象 512 字段和转义后 4 MiB 序列化体积。Agent turn、日志、文件传输、Excel 写入、OCR、模板运行和设置写入在可信 sender 校验后按 `processId + frameId` 使用令牌桶限流，且不限制任务执行时长。

- 设置 key/value、矩阵、聊天 content、attachments、base64 和 OCR 文件数组缺少统一总量限制。
- 文件 IPC 可直接 `Buffer.from(base64)` 后写盘，可能先分配大内存。

**整改**：限制字符串、数组、对象深度、矩阵 cell 数、序列化字节、base64 预估解码大小和每 sender 速率；超限必须在大内存分配和写盘前快速拒绝。

**验收状态**：代码级整改与负向测试已完成；仍需在打包应用中对连续拖动透明度、批量附件、连续排队 turn 和大文件/OCR 请求做压力与误伤回归。

### M-03 工具、会话和日志重复保存原始敏感内容

> 整改进展：新增共享敏感数据模块，远程外传检测、通用日志、工具执行日志和 rollout 检索共用同一组高置信规则。日志消息和结构化字段会在控制台与文件输出前递归脱敏并限制单行长度；工具执行审计改为类型、字段名、集合数量和结构指纹，不再保存参数/结果原文。`rollout_events_fts` 已通过迁移移除 `item_json` 副本，工具调用/结果只索引结构摘要，reasoning 只索引 summary；旧 FTS 会重建，损坏投影也不会回退复制原始 JSON。单会话删除现同时清除活动 JSONL、旧 gzip 快照、zstd 冷归档、SQLite 投影、FTS 与工具日志；文件删除失败会中止而不会继续删除数据库状态。设置页提供本地数据隐私导出：活动 turn 或维护期间拒绝，导出前暂停 JSONL 写入并关闭 SQLite/知识运行时，复用 staging 把核心文件逐一哈希校验后复制到空目录（仍变化的日志/临时文件为 best-effort 快照），替换原始加密设置文件为不含凭据的安全快照，并写入类别和边界清单，随后恢复运行时。设置页同时提供当前活动数据根擦除：要求精确确认短语，清空内存会话和功能设置，保留用于继续定位活动数据根的 bootstrap 路径，只删除固定白名单目录并拒绝根目录、符号链接或联接；设置、临时文件、OCR/Provider 检查、直接 Office 自动化和周期留存共用维护锁与活动操作计数，维护不会和已开始的相关操作重叠，运行时关闭/恢复和逐类别失败都会显式报告。恢复运行时后只会按默认值创建新的设置/脱敏日志，不恢复历史内容。canary、删除、导出和擦除测试覆盖字段、文本、循环对象、私钥块、日志文件、SQLite FTS、旧库迁移、归档残留、失败一致性、凭据排除、路径边界及并发互斥。

- `toolExecutionLog.ts:28-40` 直接 JSON 序列化前 2000 字。
- 工具参数/结果还进入 SQLite、JSONL 和 FTS；通用 logger 无字段级脱敏。
- 日志按天分文件但没有清理策略，并使用同步 `appendFileSync`。

**整改**：工具级声明 `redactedFields/summaryFields`；默认只记数量、类型、状态和哈希；引入加密、留存、导出、线程删除、活动数据根擦除和跨副本删除策略。用 canary 扫描 logs/SQLite/JSONL/FTS，均不得出现原文。

**验收状态**：派生日志与索引的明文复制已关闭；会话 JSONL/冷归档与 state-runtime/knowledge 敏感字段已应用层 AES-GCM 加密，FTS 不再保留可还原正文；首次启动与密钥轮换使用 staging 原子迁移与 journal 恢复；设置页可导出（登记 privacy_export）、轮换本地数据密钥，并在精确确认后擦除已登记活动根/旧根/导出副本；擦除前预创建 replacement 密钥，删除后单次原子写切换为唯一 active 并清除旧密钥材料（无空 keystore 窗口），删除证明写入 bootstrap 审计位置。目录外原始 Office 文档、未登记外部副本与产品站数据仍不在擦除范围。M-03 标记为已实现，待打包实机回归。

### M-04 流式请求在已输出部分内容后仍会整体重试

> 整改完成：采样请求在尚未产生可见输出时保留瞬时故障重试；一旦正文、推理或工具 item 已通过回调发送到 UI，重试守卫立即关闭，后续可重试错误也会直接透出。新增测试覆盖“首个错误发生在可见输出前时可重试”“正文 partial 后断线不重试”“工具气泡出现后断线不重试”，不会再向同一 round 追加第二次尝试。

- 重试包装整个 stream operation，而 collector 已向 UI 发出 delta 和 tool item。
- UI 会继续向同一 round 追加第二次尝试数据。

**影响**：网络中断后可能出现重复正文、重复推理和孤立工具气泡。

**整改**：首次可见事件后禁止透明重试，或增加 `attemptId + stream_reset` 原子回滚。测试“首轮 partial 后断线、第二轮成功”，UI 和历史只能保留一份结果。

### M-05 高级 Excel 工具按任务意图动态开放

> **整改与 Codex 本地 Excel/WPS 实机验收均已完成。**
>
> - 执行层硬边界：Power Query 须 `advancedIntent:"refreshable-etl"` + `sourceKind`；透视/切片器须 `advancedIntent:"interactive-pivot"`；模型 Schema/提示词按轮意图动态开放。
> - 专用冒烟：`desktop/scripts/smoke-excel-advanced-intent.ts`，`npm run test:excel-advanced-intent`；宿主 `WENGGE_EXCEL_ADVANCED_INTENT_HOST=excel|wps|both`。
> - **负向实测**：7 个缺失/错误 intent 或 sourceKind 调用均经生产 `office.action.apply` 在 COM 前拒绝；`bridgeCalls=0`，工作簿 SHA-256 不变。
> - **Excel 实测**：基础写入、外部 CSV Power Query 创建与检查、透视表、切片器全部成功；三个不同目录的 `collision.xlsx` 分别得到 3 个独立 Excel PID、3 个完整路径、3 个不同 instanceId，并能按 `instanceId + fullPath` 激活。
> - **WPS 实测**：基础文件操作成功；透视表和切片器成功；Power Query 明确返回 `power_query_unavailable`，具体错误与结构化码沿生产工具链透传，不再伪报成功。
> - **生命周期实测**：Excel 与 WPS 两轮结束后 `excel/wps/et/Wengge.OfficeWorker` 均无新增残留，临时夹具已删除；未清理 WPS 后台 helper/cloud 服务。

- 历史缺口：高级工具曾常驻暴露，且 WPS Power Query 会伪成功；现已由动态开放、执行层硬边界和宿主能力结论共同关闭。

**整改**：动态暴露 + operation 前置条件 + 专用真实矩阵入口。

**验收状态**：M-05 已关闭。定向 Vitest 42/42、.NET Worker 109/109；Codex 本地 Excel/WPS 专项均 `ok:true`，且无宿主、Worker 或专项临时目录残留。

### M-06 下载可用性与统计可靠性耦合，统计数据过度保留

> 整改完成：下载统计写入保持 best-effort；IP 不落原文，改为基于主盐和 UTC 周期派生密钥的 HMAC 标识，默认每 30 天轮换。下载记录默认保留 90 天，启动时及此后每 6 小时执行有界清理，清理异常只记录告警，不影响新增记录或安装包响应。UA 上限从 500 缩短到 200 字符，Referer 只保存 Origin；既有数据在一次性迁移中同步最小化。自动化测试覆盖周期内/跨周期标识、留存边界、旧数据迁移、清理触发器故障和下载可用性。

- `product-site/src/server.mjs:64-84` 在发送安装包前同步写统计，整个流程共用一个 `catch`；SQLite 写失败会把正常下载降级为 404。
- GET 自动暴露的 HEAD 请求也会执行统计逻辑。
- IP 使用稳定 HMAC，UA 和 Referer 保存到 500 字符，没有 TTL。

**整改**：统计写入 best-effort，不得影响下载；只统计成功 GET；缩短或移除 UA/Referer，IP 标识按周期轮换；定义 30/90 天留存和删除作业。故障注入数据库只读/锁死时，下载仍必须成功。

**验收状态**：代码、迁移和自动化故障测试已完成；生产上线仍需确认 `/etc/wenge-product.env` 的留存/轮换值，并将 `download analytics maintenance failed` 纳入告警。跨轮换周期的“独立下载”可能重复计数，属于主动降低长期关联能力的隐私取舍。

### M-07 产品站生产配置只校验“存在”，且保留易误用的明文路径

- `product-site/src/config.mjs:3-8` 生产 secret 只检查环境变量是否存在，一字符值也能启动。
- `scripts/hash-password.mjs` 通过命令行参数接收明文，可能进入 shell history/进程列表。
- `plugin.shelelove.top.http.conf` 会把完整站点和后台登录代理在 HTTP 上，误部署会暴露密码。

**整改**：验证 cookie secret/analytics salt 的最小熵和格式；密码从 stdin/交互或密钥管理器读取；80 端口除 ACME 外只做 HTTPS redirect；启动时检测默认值和弱 secret 并拒绝运行。

### M-08 .NET 测试依赖存在 High 漏洞且构建不可完全复现

> **已关闭。** 测试依赖已升级，`global.json` 与 NuGet lockfile 已固定还原；CI run `29490566136` / desktop job `87595394532` 已真实通过 `dotnet restore --locked-mode`、`npm run office:audit`、`npm run office:test`（109/109）及后续 build。

NuGet 扫描在 `Wengge.OfficeWorker.Tests` 发现：

- `System.Net.Http 4.3.0`，High，GHSA-7jgj-8wvc-jh57。
- `System.Text.RegularExpressions 4.3.0`，High，GHSA-cmhx-cq75-c4mj。

二者是测试项目传递依赖，不属于当前 Worker 发布包，但仍属于开发/CI 供应链风险。仓库也没有 `global.json` 和 NuGet lockfile，CI 不运行 .NET build/test/audit。

**整改**：升级 xUnit/Test SDK/coverlet 依赖链；固定 .NET SDK；启用 lockfile/locked restore；把 .NET build、test 和 vulnerable package scan 加入 CI 与 Release 硬门禁。

### M-09 删除工作表异常会永久改变用户 Excel 的 DisplayAlerts

> **已关闭。** 代码与 Codex 本地真实 Excel/WPS 矩阵均已验收。
>
> - 生产路径：`ExcelWorkbookService.SheetOperation` 删除分支先保存 `DisplayAlerts`，`finally` 恢复原值（不再强制 `true`）。
> - 专用冒烟：`desktop/scripts/smoke-excel-display-alerts.ts`，`npm run test:excel-display-alerts`；宿主 `WENGGE_EXCEL_DISPLAY_ALERTS_HOST=excel|wps|both`；真实 `DotNetExcelBridge.sheetOperation("delete")` 命中生产路径。
> - **Excel（Codex 本地）**：`ok=true hostMode=excel`；(a) false 成功删除后仍 false，再恢复 true；(b) 结构保护删除失败 COM「不能取得类 Worksheet 的 Delete 属性」，alerts 仍 true，表 S1/S3；(c) 最后可见表删除失败（宿主消息），alerts 仍 true，表 S1；final true。
> - **WPS（Codex 本地）**：`ok=true hostMode=wps`；owned PIDs `[44880,46304,60336,67372,67560]` 全部退出；(a) false→false 再 true；(b) 结构错误 `0x8FE30C29`，alerts true，表 S1/S3；(c) 最后可见 `E_ACCESSDENIED`，alerts true，表 S1；final true。
> - 无 Excel/et/wps/wpp/`Wengge.OfficeWorker` 残留。门禁：typecheck、lint、governance 0/0、.NET 108/108、Vitest 211/1112、build entry 437.91 KiB asyncChunks 9。
> - 探针：`office.smoke.excel.*`（`WENGGE_OFFICE_SMOKE=1`）；COM dynamic 边界 `object?`；`wpsAll` 仅 et/wps；dispose → runningProcesses → dispose（无 `process.kill`）。

- 历史缺口：`ExcelWorkbookService` 删除路径曾在异常时不恢复 `DisplayAlerts`，成功时强制恢复为 `true`。

**整改（已完成）**：读取旧值并在 `finally` 中恢复；专用冒烟覆盖最后可见表、结构保护与用户原本关闭提示；Codex 双宿主实机通过。

### M-10 缺少真实 Office 动态数组冒烟和 UI E2E

> **已关闭。** 代码与专用 self-hosted Runner 矩阵均已验收。
>
> - **Electron E2E**：workflow run `29456041445` / job `87489319942` 全部 success（导航/外链、审批、设置 IPC、更新 UI、流式中断恢复顺序）。
> - **Excel 365**：run `29457636677` / job `87494253080` success；日志 `ok=true, hostMode=excel, host=excel, spill=true, multiFormulaRollback=true, saveReopenSpill=true, formula2=true`；无 Excel/`Wengge.OfficeWorker` 残留。
> - **WPS**：run `29458142234` / job `87495761434` success；日志 `ok=true, hostMode=wps, formula2Supported=true, capability=formula2_spill_ok, written=1, dynamicCells=1, spillValues=[[1],[2]]`；`excel.workbook.save` done；无 wps/et/wpp/`Wengge.OfficeWorker` 残留。
> - **Runner**：`wengge-office-local-01`，labels `wengge-office-excel-365` 与 `wengge-office-wps`；测试开始前宿主进程为空；脚本硬拒绝预存 Office/WPS。
> - 实现入口：`smoke-excel-dynamic-array.ts`、`e2e-electron.ts`、`office-matrix-and-e2e.yml`。

- 历史缺口（Excel 365 表达式 spill、WPS Formula2、多公式第二项失败回滚、保存重开 spill、桌面 UI E2E）已由上述矩阵与 E2E 关闭。
- 不覆盖 M-05（高级意图）等其它 Office 专项实机项；M-09 DisplayAlerts 由独立冒烟关闭。

**整改（已完成）**：专用 self-hosted Runner 小型矩阵 + Playwright Electron E2E；缺宿主 fail 不 skip-green。

### M-11 安全、所有权、许可和隐私治理文件缺失

> 部分整改：新增 `SECURITY.md`，使用仓库 Private Vulnerability Reporting 作为私密渠道，并定义确认、分级和分级修复目标；新增 `CONTRIBUTING.md`，规定敏感变更双人审批、CODEOWNER、必需门禁和禁止作者自合并；新增 `.github/CODEOWNERS`，按当前 GitHub 仓库所有者覆盖发布、preload/IPC、Agent、Worker、产品站和部署路径。`docs/data-handling-and-privacy.md` 按运行代码列出本地数据、远程 OCR/搜索/Embedding/模型、产品站统计及留存控制，并明确已登记活动根/旧根/应用导出的擦除与删除证明边界；应用层加密已落地。仍未完成的是法律主体/条款、身份级或外部未登记副本编排，以及 GitHub ruleset 强制审批。自动化治理测试约束私密报告链接、敏感路径所有者、审查要求和不得夸大的数据删除口径。

**审计当时（历史缺口）**：仓库尚未建立 SECURITY.md、CODEOWNERS、CONTRIBUTING.md、LICENSE/NOTICE 和正式隐私政策；当前工程治理文件与事实型数据清单已补齐，剩余缺口见下方验收状态。

**所有者决定（代码整改范围）**：工程侧漏洞披露（`SECURITY.md`）、仓库所有权（`CODEOWNERS`）、贡献与双人审查规则（`CONTRIBUTING.md`）、数据处理事实清单（`docs/data-handling-and-privacy.md`）、应用层加密/受管副本擦除与治理测试已落地；按当前“聚焦代码上存在问题”的整改目标，**关闭 M-11 的代码整改范围**。LICENSE/NOTICE、正式隐私主体与法律条款、真实第二审查人、GitHub ruleset 强制审批仍属所有者/法律/平台后续事项，**不得视为已完成**，也不得由代码虚构。

**整改**：建立漏洞披露 SLA、敏感目录所有者、合并审批规则、依赖许可清单、OCR/模型/搜索/统计数据流说明和用户删除机制。

**验收状态**：M-11 按当前代码整改目标**已关闭**。工程治理文件、事实型数据清单、应用层加密与已登记活动根/旧根/导出副本的受保护擦除及删除证明已落地。仍诚实保留且不声称完成：由法律责任人确认的 LICENSE/NOTICE 与正式隐私政策（主体、联系方式、处理依据、第三方处理者、跨境、DSR、未成年人等）、真实第二审查人配置，以及 GitHub branch protection/ruleset 的平台级强制。

### M-12 备份、留存、恢复和运行告警未形成可验证闭环

> 部分整改：产品站新增 SQLite 原生在线备份命令，归档转换为自包含 journal、生成 SHA-256/大小元数据并执行 `quick_check`；恢复命令再次执行完整性校验且拒绝覆盖现有目标。新增每日 systemd timer、14 份轮换、RPO≤25 小时/RTO≤4 小时运维基线和季度恢复演练步骤。桌面端新增启动及每 6 小时本地维护：日志 30 天/30 文件/100 MiB，Office 备份 30 天/500 份/2 GiB，事务 30 天/200 份/2 GiB，工作流 90 天/500 份/100 MiB；运行中记录和当前日志受保护，清理失败不阻止启动。备份元数据路径限定在受控目录，避免清理跟随伪造路径。自动化测试覆盖活动 WAL 库一致性快照、备份篡改/轮换/恢复、桌面 TTL/配额、活动记录保护、目录边界和分类故障隔离。更新检查/下载、Office Worker 停止、旧目录自动迁移、用户主动迁移、后台登录失败、最终 5xx 与 analytics 维护/写入失败日志已带稳定 `event` 字段；无内置告警发送端。

- 产品站使用 SQLite WAL，但没有仓库内可验证的备份、恢复和完整性演练流程。
- Office 事务快照、备份、日志和自动化状态缺少统一清理策略。
- 发布回滚主要依赖人工恢复文件。

**整改**：定义 RPO/RTO；执行在线 SQLite 备份和定期恢复演练；对 Office 备份/事务设置 TTL、配额和可审计清理；建立更新失败率、后台登录异常、下载 5xx、Worker 崩溃和数据迁移失败告警。

**验收状态**：产品站备份与桌面本地数据 TTL/配额代码已完成；本机已登记活动根/旧根/应用导出可导出、擦除并生成 bootstrap 删除证明。代码侧稳定失败事件已落地（`desktop.update.check_failed` / `desktop.update.download_failed` / `desktop.office_worker.stopped` / `desktop.data_path.legacy_auto_migrate_failed` / `desktop.data_path.user_migrate_failed` / `product_site.admin.login_failed` / `product_site.http.5xx` / `product_site.analytics.maintenance_failed` / `product_site.analytics.write_failed`），仅写日志字段、无内置告警发送端。生产仍需确认 RPO/RTO、启用 timer、在生产平台配置事件接收端并提交一次真实恢复演练记录；异机/外部未登记副本清理仍未完成，因此 M-12 不关闭。

## 7. Low — 工程治理观察项

### L-01 格式和文件规模规范没有落地

> 整改完成：新增 `scripts/check-source-governance.cjs` 与换行归一化 SHA-256 基线。CI 对桌面源码执行 Prettier 与生产模块 300/400/500 行分类上限；`desktop/scripts/smoke-*.ts` 真实冒烟场景脚本不按生产模块限行，但仍受 Prettier 检查。全量受治理源码已通过一次机械 Prettier 清零，`formatting` 基线为空；新文件与被修改文件继续强制通过 Prettier，生产超限继续拒绝。

- 全量 `npm run format:check`（desktop 受治理源码）已通过；`legacyFormatting=0`。
- 权威治理结果以 `npm run governance:check` 为准：生产模块 `legacyOversized=0`；真实 Office/WPS 冒烟脚本 `desktop/scripts/smoke-*.ts` 按规则排除行数上限（非“已拆分”），仍计入 Prettier 检查；格式债务 `legacyFormatting=0`。`CodeTaskComposerPanel` 已把草稿类型与选项常量提取到独立模型模块，组件从 303 行降至 277 行；`chatStore` 已把公共状态和动作类型提取到独立模块，Store 从 405 行降至 287 行；`i18n` 已把功能、权限、简单任务与时间文案提取为类型化叶子资源，入口从 421 行降至 376 行；`App` 已把窗口 IPC/resize/blur 生命周期集中到 `useWindowDisplayState`，标题栏变为纯展示组件，热补丁健康确认保持独立生命周期，根组件从 365 行降至 191 行；`ChatPage` 已把任务侧栏的懒加载与面板选择集中到 `ChatFeatureSidebar`，把侧栏焦点状态和当前文件夹加载分别交给专用 hook，主页面从 374 行降至 225 行，简单任务草稿更新也回收到 `useTaskDrafts`；`hotPatchManager` 已把 ZIP allowlist、压缩炸弹限制、流式解压与逐文件哈希集中到 `hotPatchArchive`，管理器保留安装事务、反回滚状态、激活/回退和健康确认，从 444 行降至 323 行；`ipcOcrHandlers` 已把本地解析、MinerU 和 Agent fallback 编排与发票识别、模型抽取、字段合并及结果归一化分离，IPC 编排器从 513 行降至 245 行，结果构建模块为 306 行；`ocrExecutors` 已把 provider fallback 执行与模型可见结果裁剪、远程摘要、告警和后续工具建议分离，执行器从 485 行降至 333 行，结果模块为 241 行，并修复单文档已裁剪但顶层截断标志仍为 false 的误报；`webSearchExecutors` 已把参数校验、远程数据策略和工具结果封装与搜索源 fallback、超时、错误归一化及响应读取分离，执行器从 423 行降至 94 行，provider 模块为 389 行，并新增解压后 2 MiB 响应上限；`ipcAgentHandlers` 已把 Agent/线程/统计 IPC 与知识库运行时初始化、路径授权、检索和索引 IPC 分离，主注册器从 496 行降至 336 行，知识库模块为 179 行，重建测试覆盖未授权来源跳过与单来源失败后继续处理；`ipcHandlers` 已把设置校验/持久化与 Agent、知识库、窗口运行时同步分离，主聚合器从 441 行降至 358 行，设置 IPC 模块为 58 行，纯运行时同步模块为 83 行，并覆盖 Provider、压缩、权限、窗口和动态数组副作用；`sqliteStore` 已把连接生命周期、事务、CRUD、来源摘要和维护与向量/关键词查询算法分离，存储类从 423 行降至 356 行，查询模块为 97 行，组合 source/path 过滤与损坏 embedding JSON 隔离均有定向测试；`officeActionAdapter` 已把单动作备份/恢复和独立跨 Office 事务、Open XML/COM 能力路由、validate 结果解释分成三个单向依赖模块，路由适配器从 550 行降至 359 行，并补充工作流上下文禁止嵌套事务及缺少协调器时执行前拒绝测试；`excelExecutors` 已把宏运行时（VBA/JSA/宿主能力）和 Excel UI 控件注册提取为 `excelMacroExecutors` 与 `excelUiExecutors` 两个能力域子注册器，组合入口仍为 `addExcelExecutors` 且保持原注册顺序，WPS 公式字符校验仍属于 `range.write`，组合器从 448 行降至 304 行，`excelMacroExecutors` 126 行、`excelUiExecutors` 107 行；`workflow` 已把 `OfficeWorkflowStatus`/`StepRecord`/`Record`/`Result`/`RunOptions` 抽到 `workflowTypes.ts`，并把校验/结果映射等纯辅助函数抽到 `workflowHelpers.ts` 以在 Prettier 展开后仍满足 400 行上限；`workflowRecordStore` 直接依赖领域类型而不再反向导入编排器；`transactionJournal` 已把 `OfficeTransactionStatus`/`Snapshot`/`Conflict`/`RestoreFile`/`RestoreOptions`/`Record` 抽到纯类型 `transactionTypes.ts`，路径/产物推导抽到 `transactionPaths.ts`，save/get/list 与受控目录/ID 校验抽到 `transactionRecordStore.ts`，begin/finalize/undo/redo 与原子恢复仍在门面，`workflowTypes`/`workflow` 直接依赖 `transactionTypes`；`ipcSchemas` 已把 IPC 通用限制/路径/Base64/`validateInput` 抽到 `ipcSchemaPrimitives.ts`，设置键与值 schema 抽到 `ipcSettingsSchemas.ts`，门面继续兼容既有导入路径；`shared/types` 已把 `Compaction*`/`ThreadCompactStartParams` 抽到 `compactionTypes.ts`，把 `Tool*` 执行契约抽到 `tools/contracts/toolExecutor.ts` 并再导出；Renderer `electronApi.d.ts` 已把 Agent/Thread 投影抽到 `electronApiAgentTypes.ts`、Update/Office/AI/文件夹等领域类型抽到 `electronApiDomainTypes.ts`、`ElectronAPI` 抽到 `electronApiInterface.ts`，门面保留 `Window` 扩展与全部类型 re-export。相关源码通过 Prettier、定向测试、类型检查、Lint、完整 Vitest、生产构建和治理门禁。

按运行时职责拆分大型生产文件，避免把格式噪声与安全修复混在同一个 PR。普通功能 PR 禁止刷新基线哈希以豁免漂移。

**验收状态**：L-01 已关闭。受治理源码 Prettier 债务与生产模块超限均为 0；`format:check` 与 `governance:check` 均通过。冒烟脚本不按 400 行强拆，但仍须保持 Prettier 一致。

### L-02 当前文档基线漂移

> 整改完成：两个 Node 项目在 `package.json#engines` 统一声明 Node 22.12+，根 README 改用 `npm ci` 并以 `global.json` 作为 .NET SDK 事实源。`overview.md` 和开发规范不再固定易失真的测试/代码行数量；过期的 `dev-log`、会话记录、早期 code review 手册/报告和实施计划已从工作树移除。现行门禁与审查口径以实际 `ci.yml`、开发规范和运行代码为准；未安装的 Husky/CommitLint/coverage/EditorConfig 不再作为当前文档要求。新增自动测试约束版本、动态数字、CI 门禁描述和历史文档分类。

- 根 `README.md:21` 仍写 Node.js 20+，实际 Electron/Vite 环境要求 Node 22.12+。
- `overview.md`、`development-standards.md` 等文档曾记录易失真的测试数量与实际不一致。
- 早期 code review 手册中的 CI/工具示例已过期（文件已从工作树移除）。

建议只在当前文档维护动态基线，不把易变化的精确测试数量复制到多个入口。

**验收状态**：当前文档、包元数据与 CI 已对齐，并有防回归测试；L-02 关闭。历史数字仅存在于 Git 历史，不纳入当前能力口径。

### L-03 Renderer 主包偏大

> 整改完成：`SettingsPage` 整体以及 Office 自动化、公式、代码、OCR、报告和简单任务面板改为 `React.lazy` 按需加载；首屏入口首次由 566.92 KB 降至 445.02 KB，当前构建为 448.41 KB，设置页形成独立异步块，其余功能形成更小异步块。热补丁健康确认从“应用壳出现”调整为当前首屏路由真实挂载后确认，避免设置页异步块尚未加载就提前 ACK。新增 bundle budget 脚本与边界测试，限制首屏 480 KiB、任一 chunk 500 KiB，并要求至少两个异步块；`electron:build` 复用同一 build 门禁。

生产构建主 Renderer chunk 约 566.92 KB，超过 Vite 500 KB 建议值。建议对设置、Office 自动化、知识库和统计面板做路由/组件级 lazy loading，并建立 bundle budget。

**验收状态**：代码拆分、自动预算和生产构建已完成，Vite 大 chunk 告警消失；L-03 代码层关闭，仍需打包应用冷启动和低速磁盘体验回归。

## 8. 已验证的正向控制

- Electron 已启用 `contextIsolation:true`、`nodeIntegration:false`。
- 多数 IPC 已开始使用 Zod 输入校验，文件/OCR 面板也已有路径授权框架。
- COM 操作集中到带 Windows 消息泵的 STA 调度器，没有把 COM 移出线程模型。
- 图表创建后已经检查对象、系列、可见性、尺寸和锚点；透视表创建后检查缓存、字段和目标范围。
- Codex 本地真实 Excel 图表专项已通过：`insertChart`、`formatChart`、`inspectCharts` 均为 `done`，回读 `RevenueChart` 标题与非空系列成功，结束后无 Excel/Worker 残留。
- Office Worker 超时会明确标注结果未知，并终止失控 Worker。
- Open XML 动态数组 metadata 已使用独立 metadata part，并复用描述。
- 完整更新清单已有 Ed25519 签名、size 和 SHA-256 校验基础。
- 产品站使用 Helmet CSP、scrypt 密码哈希、签名 HttpOnly/SameSite Cookie，systemd 也启用了多项沙箱选项。
- Desktop 和 Product-site npm 高危漏洞检查为 0；Vitest、xUnit 和产品站测试全部通过。
- 敏感本地目录已被 Git 忽略，Git 历史未发现已提交的高置信凭据。

这些控制是良好基础，但当前信任边界、审批和凭据问题会绕过其中多项防护。

## 9. 30/60/90 天整改路线

### 0-7 天：立即止血

1. 暂停对外发布当前构建。
2. 修复 C-01 导航/IPC sender 校验，先切断远端页面继承 preload 的攻击链。
3. 修复 C-02 审批语义，删除 fail-open 默认值，禁用危险操作永久授权。
4. 完成凭据盘点与轮换；隔离签名私钥并缩紧 ACL。
5. 为 Excel `range.write` 增加完整回滚，停止发布会导致部分提交的版本。
6. 修复产品站可信代理配置和后台限流绕过。

### 8-30 天：关闭全部 High

1. 完成知识库/Temp/junction 路径授权改造。
2. 上线统一 egress policy、离线模式、OCR/搜索明确授权和提示注入隔离。
3. 完成数据目录全量事务迁移和 Runtime 重建。
4. 修复动态数组分类、Open XML 样式保留和 Worker v2/capability negotiation。
5. 引入 Authenticode、隔离发布 job、只读下载目录和热补丁防重放/健康回滚。
6. ~~在隔离 Runner 执行 Excel 365/WPS 真实冒烟~~ **已完成（M-10）**：E2E 29456041445、Excel 29457636677、WPS 29458142234；Runner `wengge-office-local-01`。

### 31-60 天：工程与合规闭环

1. 统一工具 Schema、IPC 资源预算、日志脱敏和留存删除。
2. 补 UI E2E、流式断线测试、迁移故障注入和产品站下载故障测试。
3. 引入 `global.json`、NuGet lockfile、SBOM、provenance、CodeQL、Dependabot 和 secret scanning。
4. 建立 `SECURITY.md`、CODEOWNERS、许可清单和隐私政策。

### 61-90 天：运营成熟度

1. 完成 SQLite/Office 数据恢复演练并记录 RPO/RTO。
2. 建立更新、Worker、迁移、后台登录、下载和 OCR 出站监控告警。
3. 进行一次独立 Electron/Agent/更新链渗透测试和隐私影响评估。
4. 清理格式债务、文档漂移和 bundle 体积。

## 10. 上线前硬门槛

满足以下全部条件前，结论保持 **No-Go**：

- [ ] 3 个 Critical 全部关闭，并有负向安全测试。
- [ ] 13 个 High 全部关闭；如确需接受，必须由安全负责人和业务负责人共同签字、写明补偿控制和到期日。
- [ ] 所有可能使用过的凭据完成轮换，旧凭据失效，签名私钥离开开发工作区。
- [ ] Electron 主窗口不能导航到远端；所有敏感 IPC 有 sender/origin 验证。
- [ ] 宏、删除、未知工具、外传工具在策略要求下必定审批，审批缺失时默认拒绝。
- [ ] 安装包 Authenticode 为 `Valid`，更新 manifest、size、SHA-256 和发布资产全部端到端验签。
- [ ] CI/Release 通过 npm audit、NuGet audit、lint、typecheck、Vitest、.NET tests、产品站 tests、build、SBOM 和签名验证。
- [x] **M-10** Excel 365/WPS 动态数组矩阵与 Electron E2E（spill、多公式回滚、保存重开、Formula2 spill）已在专用 self-hosted Runner 跑绿（测试前无宿主进程）（见 M-10 章节 run/job）。
- [x] **M-09** DisplayAlerts 恢复矩阵已由 Codex 本地 Excel+WPS 实机通过（见 M-09 章节）。
- [x] **M-05** 高级意图专项已由 Codex 本地 Excel+WPS 实机通过（见 M-05 章节）。
- [x] legacy CSE 已由 Codex 在真实 Excel/WPS 分别通过：`N1:N3` 写入 `=B2:B4` 后 `currentArray`、值与公式回读正确；超长 FormulaArray 失败返回结构化错误且 `O1:O3` 完整恢复；测试后宿主、Worker 与临时目录无残留。
- [x] **H-10** Worker 协议不匹配专项已由 Codex 实机通过：基线 `660c0597` v1 Worker + 绝对路径 `WENGGE_OFFICE_WORKER_PATH` + `npm run test:office-worker-protocol` exit 0（`ok:true` / `code:protocol_mismatch` / 应用=2 Worker=1，残留 0）；当前 v2 Worker 同命令按预期 exit 1。
- [ ] 产品站伪造 XFF 无法绕过限流，统计数据库失败不影响下载。
- [ ] 数据目录迁移、SQLite 备份和 Office 事务恢复均完成故障演练。
- [ ] 隐私政策明确 OCR、模型、搜索和下载统计的数据流、目的地、留存和删除方式。

## 11. 最终结论

代码整改已关闭原报告中的 Electron 导航/IPC、工具审批、明文设置凭据、路径越界、Excel 部分提交、动态数组写入、Open XML 样式破坏、产品站代理信任、数据外传、提示注入、数据目录事务迁移、热补丁回滚/吊销、模型可见 Office 参数边界与 .NET 供应链复现等主要缺口，并建立全 operation 严格 Schema 与源码治理棘轮；当前自动化门禁为 211 个 Vitest 文件、1114 项测试全部通过，最近一次 .NET Worker 门禁为 109 项测试通过。

总体结论仍保持 **No-Go / Request Changes**，原因已从“存在可直接利用的代码攻击链”转为“生产外部验收和治理门槛尚未完成”：真实凭据轮换与 ACL、受保护 Authenticode 证书/HSM、Environment approval、SBOM 与最终发布清单端到端验签、打包 Electron 导航/热补丁白屏回滚、生产 Nginx/告警、备份恢复演练，以及 SECURITY/隐私/事件响应制度仍需落地（**M-05、M-09、M-10、H-10、legacy CSE 与真实 Excel 图表专项已关闭**）。完成这些外部证据或经正式风险接受前，不应发布企业生产版本。
