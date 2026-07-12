# 工具契约层

职责：定义工具实现无关的桥接接口，例如 Excel 工作簿、Word 文档、PowerPoint 演示文稿、Shell 和知识检索工具契约。

模块说明：

- `excel.ts`: Excel/WPS 工作簿、连接状态、内部 VBA/JSA 宏和 UI 控件桥接契约。
- `office.ts`: Word、PowerPoint、Office 脚本桥接契约，以及统一 Office file/action bridge 契约。

关联模块：

- `../implementations`: 实现本层契约。
- `../executors`: 通过契约类型调用注入的实现。
- `../registry`: 只定义模型可见工具 schema，不放运行时接口。

边界约定：

- 本层只定义接口和数据形状，不引入 COM、PowerShell、Open XML 包结构或 Electron UI 依赖。
- Open XML 与 COM 的选择逻辑在 `officeCore/officeActionAdapter.ts`，具体实现留在 `tools/implementations`。
