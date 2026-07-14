# 工具执行器层

职责：处理工具调用的参数校验、路由分发和结果封装。

模块说明：

- `createToolExecutors.ts`: 组装所有工具执行器。
- `validation.ts`: 工具参数校验和错误格式化。
- `excelExecutors.ts`: Excel/WPS 工作簿、区域、公式、工作表、内部宏和 UI 工具执行器。
- `fileExecutors.ts`: 文件读写工具执行器。
- `knowledgeExecutors.ts`: 知识库检索工具执行器。
- `ocrExecutors.ts`: MinerU OCR 解析工具执行器，把图片/PDF 转成文本和表格。
- `officeExecutors.ts`: Word、PowerPoint、统一 Office action、多窗口选择和可回滚工作流执行器。

关联模块：

- `../contracts`: 通过契约调用注入的 bridge。
- `../registry`: 执行器名称必须与工具定义保持一致。

执行约定：

- Word/PPT/Excel 文件级编辑优先路由到 `office.action.*`，由 `officeCore/officeActionAdapter.ts` 决定 Open XML 或 COM。
- `macro.*` 只路由到工作簿内部 VBA/WPS JSA，不能回退为外部脚本执行。
- Office 与文件处理能力必须通过类型化工具和 .NET Office Worker 暴露，不提供任意 Shell/Python 脚本入口。
