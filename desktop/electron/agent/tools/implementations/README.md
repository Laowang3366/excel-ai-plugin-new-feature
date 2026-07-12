# 工具实现层

职责：实现工具契约中的本机能力，包括 Excel/WPS、Word、PowerPoint、Office 脚本桥接和不依赖 COM 的 Open XML 文件编辑。

模块说明：

- `excel/`: Excel/WPS 表格 COM 桥接、内部 VBA/WPS JSA 宏、UI 控件和工作簿/区域/公式/工作表操作。
- `office/`: Word、PowerPoint 和 Office 脚本 COM 桥接；`officeComActionBridge.ts` 作为统一 Office action 的 COM 兜底，覆盖目录刷新、快照、图表、PPT 删除页等需要应用对象模型的场景。
- `officeOpenXml/`: 直接读写 `.docx` / `.pptx` / `.xlsx` 的 ZIP + XML 文件结构，用于不依赖 Office 进程的文件级编辑；当前覆盖 Excel 高级编辑、Word 标题/表格/页眉页脚、PPT 创建/主题/删除页等操作。

关联模块：

- `../contracts`: 本层实现契约接口，不把 COM、PowerShell、Open XML 包结构等细节泄漏给执行器。
- `../executors`: 执行器通过注入的桥接实例调用本层能力。
- `../../automation`: 复用 Python、PowerShell 和 JSON 解析基础能力。
- `../../runtime/bridgeRegistry.ts`: 创建并复用 Office/Excel bridge 实例。

实现约定：

- 文件级 Word/PPT/Excel 操作优先接入 `officeOpenXml/`，并通过 `officeCore/officeActionAdapter.ts` 暴露给 `office.action.*`。
- 只有 Open XML 不适合处理的动态对象、当前窗口状态、预览导出或应用内刷新，才进入 `office/` COM 兜底。
- 新增操作必须同步更新能力表、工具注册表、提示词说明和测试。
