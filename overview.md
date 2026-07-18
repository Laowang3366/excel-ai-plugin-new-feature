# 项目概览

更新时间：2026-07-18

文格 AI 助手是一个 Windows Electron 桌面应用，通过模型工具调用操作 Excel、WPS、Word 和 PowerPoint。项目同时包含独立产品站，用于安装包分发、远程更新和匿名下载统计。

## 目录

```text
desktop/       Electron 桌面端、Agent Runtime、.NET Office Worker 和安装包配置
product-site/  产品页、更新 API、下载统计后台和生产部署配置
release-notes/ 面向用户的版本功能说明
docs/          当前架构、开发、发布和历史设计文档
```

## 桌面端链路

```text
用户输入
  -> 场景提示词与工具选择
  -> Agent Loop
  -> range / macro / word / presentation / office.action 等工具
  -> TypeScript 类型化 Office Worker 薄桥
  -> .NET 8 Worker
  -> COM / WPS JSA / C# Open XML
  -> Office 当前窗口或磁盘文件
  -> 回读结果并展示
```

Office 自动化只有一条生产主链路：Electron 不拼接 PowerShell、Python、JScript 或任意外部脚本，而是向 Worker 发送类型化请求。COM 负责当前窗口和动态对象模型，C# `DocumentFormat.OpenXml` 负责无需启动 Office 的文件级处理及少量 WPS 兼容修正。

复杂 Office 任务由持久化工作流执行：每步记录输入、输出、修改和失败位置；失败后可从当前步骤继续。Excel 区域/图表可链接到 Word/PPT 并原位刷新。多窗口操作使用完整文件路径和对象 locator，跨文件任务使用组快照整体撤销或重做。

完整权限模式显示为“完整权限（自动执行）”，不会弹出工具确认窗口；参数校验、路径授权、沙箱限制、执行层硬拒绝和高置信凭据外发阻断仍保留。应用不再提供“远程数据处理”总开关，OCR、联网搜索、发票抽取和知识 Embedding 按各自能力直接执行。

## 更新链路

- 完整版本：使用 NSIS 安装包和 `electron-updater`，在应用内下载后覆盖安装。
- 热补丁：下载签名 ZIP，仅允许覆盖前端资源、内置知识库和 WPS JSA 桥接资源，重启生效。
- 安全性：更新清单使用 Ed25519 签名，安装包和补丁均校验 SHA-256。
- 发布源：`https://plugin.shelelove.top`。

## 产品站

- 产品页和更新日志从当前发布 API 自动读取版本信息。
- 下载由 Nginx 直接传输，Node 服务只记录匿名化统计。
- 后台提供周期下载量、独立访客、版本分布和最近下载记录。
- 生产部署使用独立目录、系统用户、systemd 服务、端口和 Nginx 站点。

## 当前验证基线

- 测试数量和代码行数属于易变化运行结果，不在当前架构文档固定；以对应命令输出和 CI 记录为准。
- 桌面 CI：`npm ci`、高危 npm/NuGet 审计、ESLint、TypeScript、Vitest、.NET Worker 测试和 Vite build。
- 产品站 CI：`npm ci`、高危 npm 审计和 Node 测试。
- `format:check` 与 coverage 当前不是 CI 门禁；coverage 脚本存在，但仓库尚未声明其 provider 依赖。
- Office 冒烟按变更范围执行，默认单动作超时 30 秒并输出等待探测；生产调用默认超时 120 秒。
