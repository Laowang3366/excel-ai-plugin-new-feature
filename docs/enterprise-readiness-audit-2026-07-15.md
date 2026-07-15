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
| C-01 | 已实现，待 Electron E2E | `trustedIpc.ts`、`windowNavigationPolicy.ts`、Markdown 外链拦截；远端 sender/子 frame/协议混淆负向测试 | 打包应用实机导航与系统浏览器联调 |
| C-02 | 已关闭 | 审批缺失默认拒绝；危险/未知/删除/外传强制审批；线程+工具+operation+目标+TTL 授权；全工具元数据表驱动测试 | 无 |
| C-03 | 代码整改完成，外部动作未完成 | `settingsSecrets.ts` 使用 `safeStorage`，Renderer 仅接收掩码，设置 key 白名单及迁移测试 | 轮换所有可能暴露的真实凭据；签名私钥移出开发机；核验 Windows ACL |
| H-01/H-02 | 已实现 | 路径授权改为 realpath/最近存在父目录解析；移除系统 Temp 信任；知识库索引、删除、重建均重新授权；junction 逃逸测试 | Windows 安装包内路径选择实机回归 |
| H-03 | 已实现，待第三方联调 | `egressPolicy.ts` 统一默认关闭策略；OCR 本地优先；搜索/Embedding/发票模型抽取关闭时零请求；高置信凭据发送前阻断；返回目标服务与数据摘要 | 企业服务 allowlist 与真实 MinerU/搜索/Embedding 联调验收 |
| H-04 | 已实现 | 工具结果以结构化不可信 JSON 回送；系统提示明确外部数据边界；`memory.write` 必须引用当前轮用户原文并记录哈希/线程/turn/citation；未确认旧记忆不注入 system prompt；提示注入负向测试 | 红队场景与多模型遵循性实测 |
| H-05 | 已实现 | `outboundUrlPolicy.ts` 强制 HTTPS/精确本地 allowlist，阻断私网、DNS rebinding 与跨源重定向；Provider/Embedding 负向测试 | 企业本地模型 allowlist 配置验收 |
| H-06 | 已实现，待安装包实测 | 新安装默认 `%LOCALAPPDATA%` 用户隔离目录并自动迁移旧安装目录；全根目录 staging 复制、稳定文件 SHA-256、SQLite 实际打开校验、空目标原子改名、UNC 默认拒绝；失败恢复旧 store/runtime 并清理目标 | 安装包跨盘迁移、断电/磁盘满与旧目录停止写入实测 |
| H-07 | 已实现，待真实 Office | `ExcelRangeWriteTransaction` 在任一公式失败时恢复整区；第二个公式失败与回滚失败测试 | Excel/WPS 当前窗口实机验证 |
| H-08 | 已实现，待真实 Office | COM 所有现代公式统一 `Formula2`；表达式型 spill 识别覆盖范围表达式、IF、TRANSPOSE；COM/Open XML 测试 | Excel 365/WPS 保存重开与 spill 冒烟 |
| H-09 | 已实现 | Open XML 不再先删除后跳过占位；覆盖写入克隆原单元格属性并保留样式；样式/非 spill 数据测试 | 大型真实工作簿兼容性回归 |
| H-10 | 已实现 | Worker 协议升至 v2；`excel.range.write` 结果运行时字段校验；旧结果负向测试 | 安装覆盖失败/旧 Worker 实机握手测试 |
| H-11 | 已实现 | Fastify 仅信任本地代理；Nginx 覆盖 XFF；轮换伪造 XFF 第 9 次触发 429 | 生产 Nginx 配置上线与告警验证 |
| H-12 | 部分完成 | CI 增加 .NET test/NuGet audit；Release 拆分构建签名与受审批发布；两阶段 Authenticode 校验；产品站 release 目录只读 | 配置受保护证书/HSM、Environment approval；固定第三方 Action SHA；SBOM/最终清单端到端验签 |
| H-13 | 已实现，待安装包实测 | 64 KiB 流式 ZIP 解压与逐文件边界/哈希校验；每次启动 Renderer health pending/ack，30 秒超时及下次启动自动回退；签名清单支持吊销 ID 和最低安全序列并立即停用 | 打包 Renderer 白屏/硬崩溃、生产签名吊销清单端到端演练 |
| M-05 | 部分完成 | Power Query、透视表、切片器新增显式高级意图参数和执行层前置条件；直接 action、工作流及模板均在 Worker 前拒绝越界请求 | 依据任务分类动态裁剪模型可见高级 operation；真实 Excel/WPS 回归 |
| M-06 | 已实现 | 下载统计改为 best-effort，数据库故障时安装包仍返回 200；故障注入测试 | 留存清理与周期盐轮换 |
| M-07 | 已实现 | 生产 secret 强度/格式校验；密码哈希脚本只从 stdin 接收；HTTP 配置仅重定向 HTTPS | 生产环境配置验收 |
| M-08 | 已实现 | 测试依赖升级后 NuGet 高危漏洞为 0；`global.json`、NuGet lockfile、CI audit/test 门禁 | CI 首次运行确认锁定还原 |
| M-09 | 已实现，待真实 Office | 删除工作表用 `finally` 恢复用户原 `DisplayAlerts` | 最后可见表、保护表与原关闭状态实机验证 |

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
| Desktop Vitest | 通过 | 整改后 157 个测试文件、822 项测试全部通过 |
| Desktop Vite build | 通过但有告警 | 主 Renderer chunk 约 566.92 KB，超过 500 KB 建议值 |
| Desktop `format:check` | **失败** | 限定源码检查仍约有 482 个格式不匹配文件 |
| .NET Worker test | 通过 | 整改后 97 项 xUnit 测试全部通过 |
| NuGet vulnerability scan | 通过 | Worker 与测试项目均未发现已知漏洞包 |
| Product-site `npm audit` | 通过 | 0 个高危 npm 漏洞 |
| Product-site test | 通过 | 4/4 通过，包含 XFF 绕过与统计故障注入 |
| Git 历史凭据扫描 | 通过 | 未发现敏感目录或高置信真实凭据被提交 |
| 真实 Office/WPS 冒烟 | 未执行 | 上线前必须在隔离 Windows Runner 补齐 |

## 4. Critical — 上线硬阻断

### C-01 外部页面可继承完整 Electron preload 权限

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
2. Provider/OCR 凭据迁移到 Electron `safeStorage`、Windows Credential Manager 或企业密钥库；Renderer 只持掩码和 `secretRef`。
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

> 整改结果：新增统一的远程数据处理开关，默认关闭；OCR 本地优先，搜索、远程 OCR、发票模型抽取和 Embedding 均在网络请求前执行策略与高置信凭据检查，并通过工具审批呈现外传风险。以下“证据”为整改前基线。

**证据**

- `desktop/electron/main-modules/ipcOcrHandlers.ts:72-107` 的默认顺序为 MinerU Token、免费 MinerU Agent、最后本地解析。
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

### H-08 动态数组识别仍依赖函数白名单

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

### M-01 模型可见工具 Schema 没有统一运行时执行

- `ToolDefinition.parameters` 主要用于告诉模型，不是执行信任边界。
- `toolExecutor.ts` 只做 JSON parse；executor 多为零散的必填基础类型检查。
- 部分 operation enum 和未知字段会继续下传。

**整改**：由同一份严格 Zod/JSON Schema 同时生成模型描述并在审批前、执行前校验；`.strict()`、enum、深度、长度和数值边界统一实施。验收应对每个工具做 malformed/fuzz 测试。

### M-02 IPC 缺少统一资源预算

- 设置 key/value、矩阵、聊天 content、attachments、base64 和 OCR 文件数组缺少统一总量限制。
- 文件 IPC 可直接 `Buffer.from(base64)` 后写盘，可能先分配大内存。

**整改**：限制字符串、数组、对象深度、矩阵 cell 数、序列化字节、base64 预估解码大小和每 sender 速率；超限必须在大内存分配和写盘前快速拒绝。

### M-03 工具、会话和日志重复保存原始敏感内容

- `toolExecutionLog.ts:28-40` 直接 JSON 序列化前 2000 字。
- 工具参数/结果还进入 SQLite、JSONL 和 FTS；通用 logger 无字段级脱敏。
- 日志按天分文件但没有清理策略，并使用同步 `appendFileSync`。

**整改**：工具级声明 `redactedFields/summaryFields`；默认只记数量、类型、状态和哈希；引入加密、留存、导出、线程删除和全局擦除策略。用 canary 扫描 logs/SQLite/JSONL/FTS，均不得出现原文。

### M-04 流式请求在已输出部分内容后仍会整体重试

- 重试包装整个 stream operation，而 collector 已向 UI 发出 delta 和 tool item。
- UI 会继续向同一 round 追加第二次尝试数据。

**影响**：网络中断后可能出现重复正文、重复推理和孤立工具气泡。

**整改**：首次可见事件后禁止透明重试，或增加 `attemptId + stream_reset` 原子回滚。测试“首轮 partial 后断线、第二轮成功”，UI 和历史只能保留一份结果。

### M-05 高级 Excel 工具已有执行层边界，模型侧动态裁剪仍待完成

> 整改进展：已增加执行层硬边界。Power Query 创建/更新必须声明 `advancedIntent:"refreshable-etl"` 和外部/多来源 `sourceKind`；透视表/切片器必须声明 `advancedIntent:"interactive-pivot"`，创建透视表还必须给出明确源区域和字段。直接 action、工作流和模板解析均在 Worker 调用前拒绝不满足条件的请求。动态按任务裁剪模型可见高级 operation 仍待完成。

- `office-tools.zh-CN.md:8-9` 的文本边界清楚，这是正向改进。
- 但高级工具仍常驻注册并向模型暴露；现有测试主要断言提示词包含指定字符串。

**整改**：根据任务分类动态暴露高级工具；Power Query 只在明确外部/多来源可刷新 ETL 条件下开放，透视表只在明确要求交互对象时开放；增加 operation-specific 前置条件、审批和执行层拒绝测试。

### M-06 下载可用性与统计可靠性耦合，统计数据过度保留

- `product-site/src/server.mjs:64-84` 在发送安装包前同步写统计，整个流程共用一个 `catch`；SQLite 写失败会把正常下载降级为 404。
- GET 自动暴露的 HEAD 请求也会执行统计逻辑。
- IP 使用稳定 HMAC，UA 和 Referer 保存到 500 字符，没有 TTL。

**整改**：统计写入 best-effort，不得影响下载；只统计成功 GET；缩短或移除 UA/Referer，IP 标识按周期轮换；定义 30/90 天留存和删除作业。故障注入数据库只读/锁死时，下载仍必须成功。

### M-07 产品站生产配置只校验“存在”，且保留易误用的明文路径

- `product-site/src/config.mjs:3-8` 生产 secret 只检查环境变量是否存在，一字符值也能启动。
- `scripts/hash-password.mjs` 通过命令行参数接收明文，可能进入 shell history/进程列表。
- `plugin.shelelove.top.http.conf` 会把完整站点和后台登录代理在 HTTP 上，误部署会暴露密码。

**整改**：验证 cookie secret/analytics salt 的最小熵和格式；密码从 stdin/交互或密钥管理器读取；80 端口除 ACME 外只做 HTTPS redirect；启动时检测默认值和弱 secret 并拒绝运行。

### M-08 .NET 测试依赖存在 High 漏洞且构建不可完全复现

NuGet 扫描在 `Wengge.OfficeWorker.Tests` 发现：

- `System.Net.Http 4.3.0`，High，GHSA-7jgj-8wvc-jh57。
- `System.Text.RegularExpressions 4.3.0`，High，GHSA-cmhx-cq75-c4mj。

二者是测试项目传递依赖，不属于当前 Worker 发布包，但仍属于开发/CI 供应链风险。仓库也没有 `global.json` 和 NuGet lockfile，CI 不运行 .NET build/test/audit。

**整改**：升级 xUnit/Test SDK/coverlet 依赖链；固定 .NET SDK；启用 lockfile/locked restore；把 .NET build、test 和 vulnerable package scan 加入 CI 与 Release 硬门禁。

### M-09 删除工作表异常会永久改变用户 Excel 的 DisplayAlerts

- `ExcelWorkbookService.cs:225-228` 先 `DisplayAlerts=false`，删除抛错时不会恢复；成功时也强制恢复为 `true`，没有保留用户原状态。

**整改**：读取旧值，并在 `finally` 中恢复。测试删除最后可见工作表、受保护工作表和用户原本关闭提示三种情况。

### M-10 缺少真实 Office 动态数组冒烟和 UI E2E

- 公式修复测试主要使用 fake COM 或检查 Open XML 包结构。
- 没有覆盖 Excel 365 表达式型 spill、WPS `Formula2` 能力、多公式第二项失败回滚、保存重开后的 spill。
- 桌面端缺少自动化 UI E2E；真实 Office smoke 不在 CI。

**整改**：建立隔离的 Windows/Office/WPS 预发布 Runner；按宿主版本执行小型稳定矩阵；补 Playwright/Electron E2E 覆盖外链、审批、设置、更新和流式恢复。

### M-11 安全、所有权、许可和隐私治理文件缺失

仓库未发现 `SECURITY.md`、`CODEOWNERS`、`CONTRIBUTING.md`、LICENSE/NOTICE 和隐私政策。对于会处理本地文档、OCR 上传、模型调用、搜索查询和下载统计的企业产品，这会造成漏洞响应、代码所有权、第三方许可和数据告知不清晰。

**整改**：建立漏洞披露 SLA、敏感目录所有者、合并审批规则、依赖许可清单、OCR/模型/搜索/统计数据流说明和用户删除机制。

### M-12 备份、留存、恢复和运行告警未形成可验证闭环

- 产品站使用 SQLite WAL，但没有仓库内可验证的备份、恢复和完整性演练流程。
- Office 事务快照、备份、日志和自动化状态缺少统一清理策略。
- 发布回滚主要依赖人工恢复文件。

**整改**：定义 RPO/RTO；执行在线 SQLite 备份和定期恢复演练；对 Office 备份/事务设置 TTL、配额和可审计清理；建立更新失败率、后台登录异常、下载 5xx、Worker 崩溃和数据迁移失败告警。

## 7. Low — 工程治理观察项

### L-01 格式和文件规模规范没有落地

- `npm run format:check` 失败，源码限定检查约 482 个文件不匹配。
- 约 11 个生产 TS/TSX 文件超过当前文档规定的 400 行上限，另有多个大型测试和脚本文件。

建议先固定换行/Prettier 基线，再分批机械格式化；按运行时职责拆分大型生产文件，避免把格式噪声与安全修复混在同一个 PR。

### L-02 当前文档基线漂移

- 根 `README.md:21` 仍写 Node.js 20+，实际 Electron/Vite 环境要求 Node 22.12+。
- `overview.md`、`development-standards.md`、`dev-log.md` 等当前文档记录的测试数量与实际 148/758、79 不一致。
- `code-review-standards.md` 部分 CI 和工具示例已过期。

建议只在当前文档维护动态基线，历史报告明确归档，不把易变化的精确测试数量复制到多个入口。

### L-03 Renderer 主包偏大

生产构建主 Renderer chunk 约 563.73 KB，超过 Vite 500 KB 建议值。建议对设置、Office 自动化、知识库和统计面板做路由/组件级 lazy loading，并建立 bundle budget。

## 8. 已验证的正向控制

- Electron 已启用 `contextIsolation:true`、`nodeIntegration:false`。
- 多数 IPC 已开始使用 Zod 输入校验，文件/OCR 面板也已有路径授权框架。
- COM 操作集中到带 Windows 消息泵的 STA 调度器，没有把 COM 移出线程模型。
- 图表创建后已经检查对象、系列、可见性、尺寸和锚点；透视表创建后检查缓存、字段和目标范围。
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
6. 在隔离 Runner 执行 Excel 365/WPS 真实冒烟。

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
- [ ] Excel 365 与 WPS 真实冒烟覆盖 spill、回滚、CSE、图表、透视表、保存重开和 Worker 协议不匹配。
- [ ] 产品站伪造 XFF 无法绕过限流，统计数据库失败不影响下载。
- [ ] 数据目录迁移、SQLite 备份和 Office 事务恢复均完成故障演练。
- [ ] 隐私政策明确 OCR、模型、搜索和下载统计的数据流、目的地、留存和删除方式。

## 11. 最终结论

代码整改已关闭原报告中的 Electron 导航/IPC、工具审批、明文设置凭据、路径越界、Excel 部分提交、动态数组写入、Open XML 样式破坏、产品站代理信任、数据外传、提示注入、数据目录事务迁移和热补丁回滚/吊销等主要缺口；当前自动化门禁为 157 个 Vitest 文件、822 项测试和 97 项 .NET 测试全部通过。

总体结论仍保持 **No-Go / Request Changes**，原因已从“存在可直接利用的代码攻击链”转为“生产外部验收和治理门槛尚未完成”：真实凭据轮换与 ACL、受保护 Authenticode 证书/HSM、Environment approval、SBOM 与最终发布清单端到端验签、Excel/WPS 实机矩阵、打包 Electron 导航/热补丁白屏回滚、生产 Nginx/告警、备份恢复演练，以及 SECURITY/隐私/事件响应制度仍需落地。完成这些外部证据或经正式风险接受前，不应发布企业生产版本。
