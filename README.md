# 文格 AI 助手

面向 Excel、Word、PowerPoint 与 WPS Office 的 Windows 桌面 AI 助手。应用通过模型工具调用读取、编辑和验证办公内容；文件级处理优先使用 Open XML，当前窗口交互使用 Office/WPS 桥接。

## 主要功能

- Excel/WPS 表格：选区读取与写入、公式生成与动态数组验证、工作表和工作簿操作。
- Word/WPS 文字：长文档统一排版、目录/题注/交叉引用、规则化审阅、修订式 AI 编辑、文档对比、批量邮件合并和智能模板。
- PowerPoint/WPS 演示：主题母版和品牌模板、元素级排版与问题诊断、四类动画与自动放映、演讲者备注和讲义 PDF。
- 文件级 Office 处理：通过 `office.action.inspect/apply/validate` 统一处理 `.xlsx`、`.docx`、`.pptx`。
- 高级 Office 自动化：支持 Excel Power Query 完整生命周期、图表深度编辑、工作簿对象、专业模板、打印/PDF、公式依赖与恢复、透视表/切片器，Word 审阅/邮件合并/内容控件，PPT 母版品牌/元素诊断/动画放映/备注讲义，以及跨应用报告和失败回滚。
- Python 扩展：保留兼容性较强的 `python.execute`，用于复杂文件处理和批量任务。
- 本地知识与长期记忆：内置 Excel 解题方法论、知识检索、偏好和纠错记忆。
- 多模型接入：支持 OpenAI 兼容协议、Anthropic 及常见国内模型平台。
- 远程更新：支持应用内检查、下载并覆盖安装完整版本，以及受签名和路径约束的热补丁。

项目已移除卡密授权、桌面激活、悬浮球、文件监控和独立授权后台，当前为开放使用版本。

## 快速开始

环境要求：Windows、Node.js 20+。当前窗口自动化需要安装对应的 Microsoft Office 或 WPS Office；仅处理 Open XML 文件时不要求 Office 进程运行。

```powershell
cd desktop
npm install
npm run dev
```

常用命令：

| 命令 | 说明 |
| --- | --- |
| `npm run typecheck` | 检查渲染进程和 Electron 主进程类型 |
| `npm run lint` | 执行 ESLint |
| `npm test` | 执行桌面端单元测试 |
| `npm run build` | 构建渲染进程 |
| `npm run electron:build` | 生成 Windows NSIS 安装包 |
| `npm run patch:build -- --id <id> --base-version <version>` | 创建受限热补丁 |

## 项目结构

```text
.
|- desktop/                         Electron + React 桌面应用
|  |- electron/main-modules/        窗口、IPC、更新和热补丁管理
|  |- electron/agent/               Agent、工具、知识、记忆与安全策略
|  |- src/                          React 界面与状态管理
|  |- public/                       内置知识、WPS 桥接和更新公钥
|  `- scripts/                      构建与热补丁脚本
|- product-site/                    产品页、下载统计后台和发布服务
|- release-notes/                   面向用户的版本更新日志
|- docs/                            当前文档与历史设计归档
|- CHANGELOG.md                     用户可感知版本变化
`- overview.md                      简明架构总览
```

## 工具边界

- 当前 Excel/WPS 窗口：`range.*`、`formula.*`、`sheet.*`、`workbook.*`。
- Office 磁盘文件：`office.action.inspect/apply/validate`，优先 Open XML，必要时使用 COM。
- 多窗口与批处理：`office.documents.*` + `office.objects.*` 按完整路径和稳定 locator 选择文件、工作表、页面、幻灯片与对象；`office.workflow.*` 执行可暂停续跑的多步任务。
- 跨软件联动：Excel 区域或图表可作为链接对象写入 Word/PPT，数据变化后原位刷新；工作流事务支持修改清单、整体撤销和重做。
- Word/PPT 当前窗口：`word.*`、`presentation.*`。
- 通用复杂处理：`python.execute`。
- 系统命令：`shell.execute`，受安全策略、目录约束、审批和审计限制。
- 外部 `script.execute`、任意 PowerShell Office 脚本和 JScript 写入入口已移除。

## 更新与发布

桌面端从 `https://plugin.shelelove.top` 获取 Ed25519 签名更新清单。完整版本由 `electron-updater` 下载 NSIS 安装包并覆盖安装；热补丁只允许更新界面资源、内置知识和 WPS JSA 桥接资源，主进程、preload、原生依赖与 Python 运行时必须发布完整安装包。

产品页服务位于 `product-site/`，提供产品介绍、安装包下载、更新 API 和带登录保护的下载统计后台。生产服务只监听 `127.0.0.1:18120`，由独立 Nginx 站点代理。

发布与回滚步骤见：

- [更新与发布](docs/update-and-release.md)
- [产品站部署](docs/product-site-deployment.md)
- [Office 高级自动化](docs/office-advanced-automation.md)
- [文档索引](docs/README.md)

## 验证基线

桌面端当前基线为 174 个测试文件、968 项测试；产品站另有接口与下载统计测试。发布前必须通过：

```powershell
cd desktop
npm run typecheck
npm run lint
npm test
npm run build
```

安装包输出为 `desktop/release/Wengge-AI-Assistant-Setup-<version>.exe`，安装包、blockmap 和 `latest.yml` 作为 GitHub Release 资产发布，不提交为 Git 大文件。

## 文档维护

当前规范以 [开发规范](docs/development-standards.md) 和 [架构图](docs/architecture-map.md) 为准。`docs/superpowers/`、审查报告与阶段方案是历史记录，不代表当前仍存在对应模块或工具。

## License

Private - 仅供项目所有者使用。
