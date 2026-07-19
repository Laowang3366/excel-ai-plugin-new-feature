# 文格 Excel 加载项验证项目

独立 `excel-addin/` 包：Office.js + WPS JSA 任务窗格骨架，**不**依赖 Electron / COM / .NET Worker，也不是根 workspace 成员。

> **交付状态**：代码与单测可验证；**尚未**在真实 Windows Microsoft Excel / WPS 完成侧载验收。本仓库 Linux 环境不代表本机证书信任或真实宿主已通过。

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
- 仅 `vite` serve/preview（`dev`/`preview`）在未设置 `VITE_DEV_HTTP=1` 时读取/生成 HTTPS 证书。
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
| `shared/tools` | Phase 1–36 工具合同与执行（写后回读可显式开启；**schema 默认 `additionalProperties:false`，拒绝未知字段**） |
| `shared/provider` | 供应商模板 CRUD / active / apiFormat；三协议 stream provider + factory |
| `shared/agent` / `shared/agentChat` | 离线 AgentLoop、聊天控制器、审批 gate（非 alwaysAllow） |
| `shared/prompts` | 同步生成的提示词与 manifest |
| `scripts/` | HTTPS 证书 helpers、Vite base、Office manifest 渲染/校验 |
| `manifest/` | Office manifest 模板 + 可侧载 dev 产物；WPS JSA 说明（WPS 包下一阶段） |
| `docs/capability-matrix.md` | 能力矩阵与交付/侧载状态 |
| `src/` | Task pane UI（聊天默认 Tab、工具演示、供应商配置） |

## 安全存储说明

API key 默认只存在 `MemorySecretStore`（进程内存）。**禁止**写入 `localStorage`。跨会话持久化需后续本地安全存储方案。

## 失败分类

- **typed unsupported**（`unsupported === true`）：仅用于官方 requirement precheck 失败/缺失/抛错、缺 `Excel.run` runtime、或 WPS 无已验证能力等“宿主能力不可用”路径。
- **ordinary failure**（`unsupported` 不得为 `true`）：requirement precheck 已通过并进入 `Excel.run` 之后的 load/sync/缺 API 成员/坏回读/业务错误；`detail` 保留原 `capability` / `host` / `reason`（及 evidence）。
- Tool 层把 HostResult 映射为 ToolResult 时：仅 `result.unsupported === true` 才保留 typed unsupported；其它 fail 一律 ordinary failure。

## 非目标

- 不替换 `desktop/` Electron 产品
- 不引入 COM / .NET / Electron 运行时
- 不伪造宿主不支持的能力（返回 typed `unsupported`）
- 本提交不实现 WPS jsaddons 打包；真实 Excel/WPS 侧载验收仍属后续
