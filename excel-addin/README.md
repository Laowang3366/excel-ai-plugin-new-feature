# 文格 Excel 加载项验证项目

独立 `excel-addin/` 包：Office.js + WPS JSA 任务窗格骨架，**不**依赖 Electron / COM / .NET Worker，也不是根 workspace 成员。

> **交付状态（分层）**：
> - **代码/单测/打包**：Linux CI 可验证（test/typecheck/build/manifest/package）。
> - **Microsoft Excel 真实侧载**：**未验收**。
> - **WPS 安装与加载**：Windows 真机已见 `package:wps` → `wps:install` → `status current`、authaddin `enable/isload=true`、目录 `WenggeExcelAiAddin_`（安装包 gitSha `c46362f8`，WPS 12.1.0.26885，`wps:status` current=true、drift=[]）。
> - **WPS Ribbon / 任务窗格**：同一安装状态**冷启动后**「文格 AI」Ribbon 已恢复；真机点击「打开助手」任务窗格完整渲染（无代码/包变更）。此前短暂缺失属 WPS 加载/缓存**瞬态**，**不能**定为代码回归。
> - **WPS `selection.get`（真机单点）**：空白工作簿 Sheet1 选中 G17，工具页执行返回 `ok:true`、`tool:"selection.get"`、`sheetName:"Sheet1"`、`address:"G17"`、`values:[[null]]`（gitSha `c46362f8`）。**仅此证据**；不得扩大为其它 WPS 工具全部真机通过。其它 implemented* 仍为 member-probe/mock。
>
> 工具总数 98。本仓库 Linux 环境 ≠ 宿主验收；`isload:true` / `status current` ≠ 功能全通过。**Phase60** 证据收口见 [`docs/excel-parity-audit.md`](./docs/excel-parity-audit.md)。

## 命令

```powershell
cd excel-addin
npm ci
npm run sync:prompts   # 从 desktop 同步 Excel 相关提示词 + SHA-256 manifest
npm run typecheck
npm test
npm run build
```

### 开发 HTTPS（Office 侧载）

Microsoft Excel 开发侧载要求任务窗格来源为 **HTTPS**。本包默认用官方 `office-addin-dev-certs`：

```powershell
npm run certs:install   # 生成并安装/信任开发 CA（Windows 首次需要）
npm run certs:verify
npm run dev             # https://localhost:3000 （strictPort 3000）
npm run certs:uninstall # 卸载开发证书
```

纯浏览器调试（不启用 HTTPS，**不能**当作 Office 侧载来源）：

```powershell
npm run dev:http        # http://localhost:3000
```

说明：

- `npm run build` / `typecheck` / `test` **不**依赖本机已安装证书，也不交互安装。
- 仅 `vite` serve/preview（`dev`/`preview`）在未走 HTTP 模式时读取/生成 HTTPS 证书。
- HTTP 模式：`npm run dev:http`（通过 `npm_lifecycle_event=dev:http` 判定，Windows cmd 可用），或任意环境设置 `VITE_DEV_HTTP=1`。
- 服务器 Linux 上可以生成证书文件；**Windows 上的“信任 CA”必须在目标开发机执行**，本仓库不宣称已完成。

### Manifest 生成与校验

```powershell
npm run manifest:dev
# 写出 manifest/office-excel-manifest.xml（https://localhost:3000）

npm run manifest:prod -- --base-url https://your.cdn.example/excel-addin --version 0.1.0.0 --out ./dist/office-excel-manifest.xml
# prod 必须显式 HTTPS，且禁止 localhost/127.0.0.1

npm run manifest:check
# 校验 checked-in dev manifest 与模板渲染一致，且满足 UUID/TaskPaneApp/HTTPS 等规则
```

侧载：用 Excel「插入 → 我的加载项 → 上传我的加载项」选择 `manifest/office-excel-manifest.xml`（开发机需已信任证书并运行 `npm run dev`）。

### 生产 base path

```powershell
# 默认 base "/"
npm run build

# 子路径部署（前后斜杠会规范化）
# PowerShell: $env:VITE_BASE="excel-addin"; npm run build
VITE_BASE=excel-addin npm run build
```

非法 `VITE_BASE`（绝对 URL、query、hash、路径穿越）会在配置阶段失败。

## 结构

| 路径 | 职责 |
|------|------|
| `shared/host` | HostAdapter、Office.js / WPS JSA 适配 |
| `shared/tools` | 模型可见工具合同与执行（随 registry 扩展；写后回读可显式开启；**schema 默认 `additionalProperties:false`，拒绝未知字段**） |
| `shared/provider` | 供应商模板 CRUD / active / apiFormat；三协议 stream provider + factory |
| `shared/agent` / `shared/agentChat` | 离线 AgentLoop、聊天控制器、审批 gate（非 alwaysAllow） |
| `shared/prompts` | 同步生成的提示词与 manifest |
| `scripts/` | HTTPS 证书 helpers、Vite base、Office manifest 渲染/校验 |
| `manifest/` | Office manifest 模板 + 可侧载 dev 产物；`manifest/wps-jsa/` 含正式本地 jsaddons 源（publish/ribbon/entry） |
| `docs/capability-matrix.md` | 能力矩阵与交付/侧载状态 |
| `src/` | Task pane UI（聊天默认 Tab、工具演示、供应商配置） |

## 安全存储说明

供应商名称、模型、协议、Base URL、Gateway 配置和活动供应商会以版本化文档保存到浏览器 `localStorage`，损坏或未知版本的数据会安全忽略。API key 只存在 `MemorySecretStore`（当前任务窗格内存），**不会**写入持久化文档；刷新或重开任务窗格后，直连模式需要重新输入 API key。

## 失败分类

- **typed unsupported**（`unsupported === true`）：仅用于官方 requirement precheck 失败/缺失/抛错、缺 `Excel.run` runtime、或 WPS 无已验证能力等“宿主能力不可用”路径。
- **ordinary failure**（`unsupported` 不得为 `true`）：requirement precheck 已通过并进入 `Excel.run` 之后的 load/sync/缺 API 成员/坏回读/业务错误；`detail` 保留原 `capability` / `host` / `reason`（及 evidence）。
- Tool 层把 HostResult 映射为 ToolResult 时：仅 `result.unsupported === true` 才保留 typed unsupported；其它 fail 一律 ordinary failure。


## CI 与生产静态包

仓级 CI（`.github/workflows/ci.yml` → job `excel-addin`，`working-directory: excel-addin`）：

```text
npm ci
npm audit --audit-level=high
npm run manifest:check
npm run typecheck
npm test
npm run build
```

手动生产打包（不部署）：`.github/workflows/excel-addin-package.yml`（仅 `workflow_dispatch`）。

Inputs：

| input | 说明 |
|------|------|
| `base_url` | 必填 HTTPS 生产 base（如 `https://addin.example.com/excel-addin`） |
| `version` | 可选四段版本；空则 `package.json` 的 `x.y.z` → `x.y.z.0` |
| `vite_base` | 可选；空则从 `base_url` pathname 推导；显式值必须与推导结果一致 |

本地：

```bash
npm run package:prod -- --base-url https://example.com/excel-addin
# 可选: --version 0.1.0.0 --vite-base /excel-addin/ --git-sha <sha>
```

`package:prod` 会以对应 `VITE_BASE` 执行 `npm run build`，并向 `dist/` 写入：

- Vite 产物：`index.html`、`assets/**`、public 图标
- `office-excel-manifest.xml`（**prod**，不提交到 Git；仅入包）
- `BUILD_INFO.json`（gitSha / packageVersion / manifestVersion / baseUrl / viteBase；无 secrets）
- `SHA256SUMS.txt`（稳定排序相对路径哈希）

GitHub Actions artifact 名形如 `excel-addin-<version>-<shortSha>`，内容仅为 `excel-addin/dist/**`（Actions 自带压缩，本批不另造 zip）。**Artifact ≠ 真实 Excel/WPS 宿主验收**；下载后仍需在 Windows 上信任/部署 HTTPS 来源并侧载验证。

生产 `base_url` 必须为 HTTPS、非 localhost；尾斜杠会被规范化。`SourceLocation` / `AppDomains` / 图标 URL 与 `VITE_BASE` 均由同一 base 推导，支持根域或子路径部署。打包门禁会扫描 `dist` 的 HTML/JS/manifest，拒绝 localhost/dev-server/`http://` 残留（Office.js 官方 CDN 除外）。

**同源 AI Gateway（可选，不强制）**：若加载项使用 Gateway 模式，建议与静态资源同 origin 部署，浏览器请求固定路径 `/api/ai/v1/:upstreamId/...`；`gatewayBaseUrl` 留空即同源。Gateway 是独立服务（`ai-gateway/`），产品站无需变成 AI 代理。


### WPS JSA 本地 jsaddons 包与可重复安装

**真实 WPS 侧载证据（12.1.0.26885，包目录 `WenggeExcelAiAddin_`，安装包 gitSha `c46362f8`）：**
- 安装：`wps:install` + `status current=true` + `drift=[]` + authaddin `enable/isload=true` 已见。
- Ribbon「文格 AI」：同一安装状态冷启动后已恢复；真机点击「打开助手」任务窗格完整渲染。此前缺失为 WPS 加载/缓存瞬态，**非**代码回归。
- `selection.get` 真机单点（Sheet1!G17）：`ok:true`、`tool:"selection.get"`、`sheetName:"Sheet1"`、`address:"G17"`、`values:[[null]]`。Phase59/59.3 Address 方法/属性与 `$G$17`→`G17` 规范化已在此证据闭合。
- **边界**：不得把上述证据扩大为其它 WPS 工具全部真机通过；implemented* 仍 member-probe/mock（见 capability-matrix / excel-parity-audit）。
- 包更新后可能弹出「加载项已被修改」；若 Ribbon 暂缺，优先**完整退出所有 WPS/ET 进程**再冷启动。仅 isload 不能证明 Ribbon 已绘制。
- Ribbon tab **不使用** `getVisible`（与可加载的 ExcelAIWps 一致）；恒真 getVisible 在回调未就绪时可能导致整 tab 不显示。
- **Microsoft Excel 侧载仍未验收**。
- **任务窗格布局（真机测量，未修）**：WPS CEF 布局 viewport ~1428px、可见 child 宽 ~646px；居中 `.app`（`max-width:720;margin:0 auto`）导致内容右裁。Playwright 1428 viewport 可复现 ~354px 左边距。下一批优先 WPS 专用左对齐布局（`hostKind=wps-jsa`，禁 UA/新依赖）。详见 [`docs/excel-parity-audit.md`](./docs/excel-parity-audit.md) §1.1 / §6.1。

可生成**正式本地 file:// jsaddons 包**，并用 **install-time 纯 Node CLI** 安全合并到用户 `jsaddons`（不覆盖其他插件的 `publish.xml`）。安装与 `status current` / isload 已在真机见过；**不等于**全部能力或当前会话 Ribbon 一定可见。安装器**不会**启动/结束/附加任何 WPS 进程。

```bash
npm run manifest:wps:check
npm run package:wps -- --git-sha 0123456789abcdef

# 推荐：先打包，再只读预演（保证 AppData/jsaddons 零写入；可对真实 %APPDATA% 安全检查）
npm run package:wps -- --git-sha 0123456789abcdef
npm run wps:install -- --package-dir ./dist --app-data /path/to/AppData/Roaming --dry-run

# 确认计划后真实安装（Windows 默认 %APPDATA%；非 Windows 必须显式 --app-data）
npm run wps:install -- --package-dir ./dist --app-data /path/to/AppData/Roaming

安装并**完整退出/重启 WPS** 后，功能区「文格 AI」提供「打开助手 / 模型配置 / 宿主状态」。
任务窗格 URL 从本地 index 派生（`?page=chat|providers|host`），不会写死外网域名。
本仓库 CI/服务器**不得**把 authaddin `isload=true` 或 `status current` 单独等同于 Ribbon/全部工具已通过；包更新后若 Ribbon 缺失，优先完整结束 WPS/ET 进程后冷启动，并确认 `index.html` 中 `wps-entry.js` 先于 module 脚本加载。

# 无 --package-dir 时会先重建项目 dist/ 再安装（dry-run 亦可能写 dist，但不写 AppData）
npm run wps:install -- --git-sha 0123456789abcdef --app-data /path/to/AppData/Roaming

npm run wps:status -- --app-data /path/to/AppData/Roaming
npm run wps:uninstall -- --app-data /path/to/AppData/Roaming
```

`package:wps` 执行 `build:wps`（`vite build --base ./`），整理 `dist/` 为：

- `publish.xml`（包内单插件注册样例；安装时会 **upsert 本插件条目** 并保留其他 jsplugin）
- `WenggeExcelAiAddin_/`：相对路径任务窗格、`manifest.xml` / `ribbon.xml` / `wps-entry.js`（无 Office.js CDN）
- `BUILD_INFO.json` / `SHA256SUMS.txt`

安装行为摘要：

- 目标：`<appData>/kingsoft/wps/jsaddons/WenggeExcelAiAddin_` + 同目录 `publish.xml` 中 `name=WenggeExcelAiAddin` 条目
- 写前完整校验包哈希、拒绝 symlink/路径穿越；addon 目录 staging → atomic swap；`publish.xml` 原子写 + 本工具前缀备份（`publish.xml.wengge-excel-ai.bak.*`，最多 10；**不清理** 第三方 `publish.xml.bak.*`）
- 状态文件：`wengge-excel-ai-addin-install-state.json`（无密钥）；`status` 以 publish/目录/哈希为准，不靠状态文件伪装成功
- 安装/卸载后均 `restartRequired: true`：**请完整退出并重启 WPS** 后再加载
- 源布局见 `manifest/wps-jsa/`

## 非目标

- 不替换 `desktop/` Electron 产品
- 不引入 COM / .NET / Electron 运行时
- 不伪造宿主不支持的能力（返回 typed `unsupported`）
- 不宣称 Microsoft Excel 侧载已通过；WPS 仅记录已见的安装/加载与曾验证入口，不宣称全部能力真机通过（Linux CI ≠ 宿主验收）
