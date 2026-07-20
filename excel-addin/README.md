# 文格 Excel 加载项验证项目

独立 `excel-addin/` 包：Office.js + WPS JSA 任务窗格骨架，**不**依赖 Electron / COM / .NET Worker，也不是根 workspace 成员。

> **交付状态**：代码与单测可验证；**尚未**在真实 Windows Microsoft Excel / WPS 完成侧载验收。
>
> Phase55：`workbook.template.apply` / `workbook.template.capture`（Office.js；WPS unsupported；工具总数 98）。Phase56：WPS JSA `wps:install`/`status`/`uninstall` 可重复安装 CLI（install-time only）。Phase58：Ribbon→CreateTaskPane 任务窗格与 page 深链（**真实点击验收仍待本机重装/重启**；`isload:true` ≠ 功能全通过）。本仓库 Linux 环境不代表本机证书信任或真实宿主已通过。

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

可生成**正式本地 file:// jsaddons 包**，并用 **install-time 纯 Node CLI** 安全合并到用户 `jsaddons`（不覆盖其他插件的 `publish.xml`）。**真实 WPS 侧载仍未验收**，不得宣称宿主已通过。安装器**不会**启动/结束/附加任何 WPS 进程。

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
本仓库 CI/服务器**不得**把 authaddin isload=true 或 status current 等同于真实 UI 已通过。

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
- 不宣称真实 Microsoft Excel / WPS 侧载已通过（Linux CI 与本地打包 ≠ 宿主验收）
