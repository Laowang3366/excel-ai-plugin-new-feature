# Excel/WPS 连接方案评估：COM 自动化 vs 插件方案

## 当前方案：PowerShell COM 自动化

### 工作原理
通过 `child_process.execFile` 启动 PowerShell 进程，使用 `-EncodedCommand` 传递 Base64 编码的 COM 脚本，与 Excel/WPS 的 COM 接口交互。

### 优点
1. **零部署成本** — 无需在 Excel/WPS 侧安装任何插件，用户打开桌面端即可使用
2. **通用性强** — 同一套代码同时支持 Excel (`Excel.Application`) 和 WPS (`Ket.Application`)
3. **功能覆盖广** — COM 接口暴露了 Excel 几乎全部能力（读写单元格、操作工作表、VBA 等）
4. **开发迭代快** — 修改 PowerShell 脚本即可调整行为，无需重新编译/安装插件

### 缺点与风险
1. **每次操作启动新进程** — 每个 `executePowerShell()` 都创建新的 PowerShell 进程（约 200-500ms 启动开销），频繁操作时延迟明显
2. **COM 对象生命周期不可控** — 每次调用都通过 `GetActiveObject` 获取 COM 引用，无法持有持久引用，在 Excel 弹出模态对话框（如保存提示）时 COM 调用会挂起
3. **超时风险** — PowerShell 进程有 30s 超时，Excel 处于编辑模式/弹窗时 COM 调用阻塞，导致超时失败
4. **WPS 兼容性不稳定** — WPS 的 COM 实现与 Excel 存在差异：
   - `Ket.Application` 部分属性/方法缺失或不一致
   - WPS 多实例时 `GetActiveObject` 可能返回错误的实例
5. **权限要求** — VBA 操作需要用户在 Excel 中启用"信任对 VBA 工程对象模型的访问"，多数用户未开启
6. **安全性顾虑** — 通过 PowerShell 执行 COM 代码可能触发安全软件告警
7. **无法实时监听事件** — 无法监听 Excel 的事件（如单元格变化、工作表切换），只能轮询
8. **编码/转义复杂** — 中文内容、特殊字符需要 Base64 编码传递，增加复杂度

---

## 替代方案 A：VSTO/COM Add-in 插件

### 工作原理
开发一个 Excel/WPS 的 COM Add-in 插件（.dll），在 Excel 进程内运行，通过 Named Pipes / WebSocket 与桌面端通信。

### 优点
1. **进程内运行** — 插件在 Excel 进程内，直接访问 COM 对象，无进程启动开销
2. **可持有持久引用** — 可缓存 COM 对象引用，避免反复 `GetActiveObject`
3. **可监听 Excel 事件** — 可实时监听单元格变化、选区变化、工作表切换等
4. **操作延迟极低** — 进程内调用，无 PowerShell 启动开销
5. **更稳定** — 不受 PowerShell 超时限制，Excel 弹窗时可等待

### 缺点
1. **需要安装注册** — COM Add-in 需要注册到 Windows 注册表（`regsvr32` 或 `regasm`），增加部署复杂度
2. **WPS 需要单独适配** — WPS 的 COM Add-in 接口与 Excel 不完全兼容，可能需要维护两套插件
3. **开发成本高** — 需要 C#/.NET 开发，增加技术栈复杂度
4. **版本兼容** — 需要同时支持 x86/x64 的 Excel，.NET Framework 版本兼容问题
5. **更新困难** — 插件更新需要关闭 Excel 才能替换 DLL
6. **安全审查** — COM Add-in 需要数字签名，否则可能被安全策略拦截

---

## 替代方案 B：Office Web Add-in（JavaScript 插件）

### 工作原理
使用 Office.js API 开发 Web Add-in，通过 Excel 的内置插件机制加载，通过 WebSocket/HTTP 与桌面端通信。

### 优点
1. **跨平台** — 同一套插件可运行在 Excel Desktop、Excel Online、Mac
2. **沙箱安全** — 在沙箱中运行，不需要 COM 注册
3. **安装简单** — 通过 Excel 的"插入 > 加载项"安装，或侧载 sideload
4. **WPS 支持** — WPS 也支持 Office Web Add-in 格式

### 缺点
1. **API 限制严重** — Office.js API 功能远少于 COM：
   - 无 VBA 操作能力
   - 部分高级格式化操作不支持
   - 事件监听有限
2. **需要 Excel 保持运行** — 插件的生命周期绑定到 Excel
3. **通信延迟** — 通过 HTTP/WebSocket 与桌面端通信，增加一跳
4. **本项目已有 add-in** — 当前项目本身就是 add-in，再创建另一个 add-in 会架构混乱

---

## 替代方案 C：Named Pipe + 驻留 PowerShell 进程

### 工作原理
启动一个常驻 PowerShell 进程（而非每次新建），通过 Named Pipe / stdin-stdout 与桌面端通信，在常驻进程中维护 COM 对象引用。

### 优点
1. **零部署** — 不需要安装任何插件
2. **减少进程启动开销** — 常驻进程省去每次 200-500ms 的启动时间
3. **可持有 COM 引用** — 在常驻进程中缓存 `$excel` 对象，避免反复 `GetActiveObject`
4. **实现成本较低** — 在当前架构基础上改进，不需要新技术栈
5. **可监听事件** — 通过 COM 事件在 PowerShell 中监听，通过 Pipe 通知桌面端

### 缺点
1. **进程管理复杂** — 需要处理常驻进程的崩溃、重启、超时
2. **COM 对象失效** — Excel 关闭后缓存的 COM 引用失效，需要检测和重建
3. **并发控制** — 多个工具调用同时发来时需要串行化执行
4. **仍受 COM 限制** — Excel 弹窗阻塞、WPS 兼容性等问题仍然存在
5. **调试困难** — 常驻 PowerShell 进程的错误更难排查

---

## 推荐方案

### 短期（当前阶段）：优化 COM 自动化 + 引入常驻进程

采用 **方案 C**，在当前 PowerShell COM 架构基础上引入常驻进程：

1. **实现常驻 PowerShell 进程管理器**
   - 启动时创建常驻 PowerShell 进程
   - 通过 stdin/stdout JSON-RPC 通信
   - 进程崩溃时自动重启
   - 优雅关闭时发送退出命令

2. **缓存 COM 对象引用**
   - 首次连接时获取 `$excel` 对象并缓存
   - 每次操作前验证引用有效性
   - 引用失效时自动重建

3. **操作队列化**
   - 所有操作通过队列串行执行
   - 避免并发 COM 调用冲突

4. **健康检查**
   - 定期 ping 常驻进程
   - 检测 Excel 进程退出后自动清理状态

### 预期收益
- 操作延迟从 ~500ms 降至 ~50ms（省去进程启动开销）
- 可靠性提升（缓存的 COM 引用减少 `GetActiveObject` 失败概率）
- 为后续事件监听能力奠定基础

### 中期（如果 COM 稳定性仍然不足）：开发 VSTO 插件

如果优化后 COM 方案仍不能满足稳定性需求，则开发方案 A 的 VSTO 插件：
- 只针对 Excel（WPS 继续用 COM）
- 使用 Named Pipes 通信
- 核心逻辑保持在桌面端，插件只做 COM 代理

---

## 实施优先级

| 优先级 | 任务 | 预估工时 |
|--------|------|----------|
| P0 | 优化当前 COM：引入 psVar 安全传参、统一错误处理 | ✅ 已完成 |
| P1 | 常驻 PowerShell 进程管理器 | 2-3天 |
| P1 | COM 引用缓存与自动重建 | 1天 |
| P2 | 操作队列化与并发控制 | 1天 |
| P2 | 健康检查与自动重连 | 1天 |
| P3 | VSTO 插件原型（如需要） | 5-7天 |
