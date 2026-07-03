# 工具执行器层

职责：处理工具调用的参数校验、路由分发和结果封装。

模块说明：

- `createToolExecutors.ts`: 组装所有工具执行器。
- `validation.ts`: 工具参数校验和错误格式化。
- `shellExecutor.ts`: Shell 命令执行工具，依赖安全沙箱评估结果。
- `excelExecutors.ts`: Excel/WPS 工作簿、区域、公式、工作表、脚本和 UI 工具执行器。
- `fileExecutors.ts`: 文件读写工具执行器。
- `knowledgeExecutors.ts`: 知识库检索工具执行器。
- `ocrExecutors.ts`: MinerU OCR 解析工具执行器，把图片/PDF 转成文本和表格。
- `officeExecutors.ts`: Word、PowerPoint、统一 Office action 和 Office 脚本工具执行器。
- `pythonExecutor.ts`: 通用 Python 脚本执行器，通过临时脚本文件避免 shell 引号转义问题。

关联模块：

- `../contracts`: 通过契约调用注入的 bridge。
- `../registry`: 执行器名称必须与工具定义保持一致。
- `../../security/sandbox`: Shell 工具执行前的策略评估和 spawn 包装。

执行约定：

- Word/PPT/Excel 文件级编辑优先路由到 `office.action.*`，由 `officeCore/officeActionAdapter.ts` 决定 Open XML 或 COM。
- `office.script.execute` 只用于统一 action 和专用工具无法覆盖的复杂自动化，不作为常规删除、替换、样式调整的首选。
- Shell/Python 执行器不得承担 Office 专用能力的常规路径，避免权限、依赖和编码问题扩大失败率。
