# 安装包 UI 与命名设计

日期：2026-07-07

## 目标

为 Windows 安装包建立统一品牌名称和安装器视觉，使用户在下载、安装、桌面快捷方式、开始菜单中看到一致的产品识别。

## 产品名称

采用名称：**文格 AI 助手**

命名理由：

- “文”覆盖 Word、PPT、文档写作、报告生成等办公场景。
- “格”保留 Excel、表格、公式、数据处理的核心心智。
- “AI 助手”保持功能直观，适合安装包、桌面快捷方式和非技术用户理解。

预计映射：

- `build.productName`：`文格 AI 助手`
- `build.nsis.shortcutName`：`文格 AI 助手`
- 安装包文件名：沿用 electron-builder 默认模式，以 productName 和版本号生成。

## 安装器 UI 方向

选择方向：A，绿色办公风格。

视觉原则：

- 克制、专业、办公软件友好，避免花哨营销页风格。
- 主色使用偏深绿色，呼应表格、效率、可靠感。
- 辅色使用浅背景和低饱和灰蓝，提高安装向导文本可读性。
- 安装器界面展示明确品牌：“安装 文格 AI 助手”。

## 技术接入方案

使用现有 `electron-builder` + NSIS 链路，不更换打包工具。

拟新增资源：

- `desktop/build/installer-sidebar.bmp`：NSIS 安装向导左侧品牌图。
- `desktop/build/installer-header.bmp`：NSIS 页面头部图。
- 可选 `desktop/build/icon.ico`：如果当前 `public/icon.png` 在 Windows 安装包图标上表现不稳定，再生成 ico。

拟修改配置：

- `desktop/package.json`
  - `build.productName`
  - `build.nsis.shortcutName`
  - `build.nsis.installerSidebar`
  - `build.nsis.installerHeader`

实现方式：

- 通过本地脚本生成 BMP 资源，避免手工二进制资源不可复现。
- 资源脚本放在 `desktop/scripts/`，输出到 `desktop/build/`。
- 图片尺寸按 electron-builder NSIS 常用约束生成：
  - Sidebar：`164x314`
  - Header：`150x57`

## 兼容性与边界

- 不改应用主界面 UI。
- 不改 Agent、Office、知识库、OCR、数据库等业务链路。
- 不改安装流程行为：仍保留非一键安装、允许修改安装目录、创建桌面快捷方式。
- 不强制引入新 npm 依赖；优先用已有 Node/系统能力生成静态资源。
- 若中文 productName 在打包产物名中出现编码问题，保留配置回滚点，并优先通过 electron-builder 官方字段解决。

## 验收标准

- 打包后的安装器名称、标题、快捷方式名称均显示为“文格 AI 助手”。
- 安装器左侧和头部出现新的品牌视觉，不再是默认空白/默认 NSIS 视觉。
- 安装路径选择、安装、卸载、桌面快捷方式创建流程保持可用。
- `npm run electron:build` 能完成 Windows NSIS 打包。
- Git diff 仅包含安装器命名、安装器资源、生成脚本、文档/变更记录相关文件。

## 后续实现计划入口

确认本设计后，再编写实现计划并实施：

1. 生成 NSIS 安装器图片资源。
2. 更新 `desktop/package.json` 打包配置。
3. 更新 CHANGELOG。
4. 执行类型/构建校验。
5. 打包安装包并提交。
