## Excel/WPS 内部宏执行规则

- 创建、修改或修复宏时，不要只输出代码让用户手工粘贴。先调 `office.connection.status`，再用 `macro.detect` 检查当前宿主真正可写入的内部宏语言。
- 统一使用 `macro.write({ language, code, entryPoint, ... })`；VBA 还要提供 `moduleName`。WPS JavaScript 指 WPS JSA 内部宏，不是桌面端 cscript；按工具返回的写入或源码错误修正重试。
- `macro.write` 只写内部宏；公式、数值和文本用 `range.write`。Python 或 PowerShell 的执行结果不得冒充内部宏。
- 宏按钮调用 `ui.addControl({ controlType:"button", macroName:"模块名.入口名", ... })`，随后调用 `ui.listControls`，确认对应按钮的 `onAction` 正确。需要识别点击来源时在 VBA 中使用 `Application.Caller`。
- 普通 VBA 入口写入后用 `macro.run({ language:"vba", ... })` 试运行，再用 `range.read` 验证结果。WPS JSA 只验收写入和源码回读，不声称已远程运行。依赖 `Application.Caller` 时，按钮入口只读取标题并调用可传参的公开过程。
- 不要给 ActiveX 控件设置 `OnAction`。`macro.detect` 没有返回可用语言时才说明当前环境不能创建内部宏，不得声称已写入。
