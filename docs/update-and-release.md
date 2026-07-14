# 更新与发布

## 更新类型

### 完整安装包

完整版本使用 electron-builder 生成 NSIS x64 安装包、blockmap 和 `latest.yml`。`npm run electron:build` 会先执行 `office:publish`，把 win-x64 self-contained .NET Worker 一并打入安装包。桌面端通过 `electron-updater` 下载，校验签名清单中的 SHA-256 后调用 NSIS 覆盖安装。

适用范围：

- Electron 主进程和 preload
- 原生依赖或 Node 依赖
- .NET Office Worker 及其自包含运行时
- 安装参数和系统集成

### 热补丁

热补丁是 ZIP 文件，只允许以下目录：

```text
dist/**
public/knowledge/**
public/wps-jsa-bridge/**
```

界面补丁必须包含 `dist/index.html`；知识库补丁必须包含 `builtin-knowledge.json`；WPS 桥接补丁必须包含 `index.html`。补丁声明精确基础版本、单调递增序列、发布时间、过期时间和逐文件 SHA-256/size 清单。客户端在解压前检查归档、声明总量、单文件大小和压缩比，安装后原子切换；每次启动会重新校验已安装文件，篡改、过期或低序列补丁不会激活。

热补丁不能更新 `dist-electron/`、preload、Node 依赖、原生模块或 `.NET Office Worker`。只要 `desktop/dotnet/`、`desktop/electron/agent/officeWorker/` 的协议或主进程执行逻辑发生变化，就必须发布完整安装包。

## 签名

- 算法：Ed25519
- 公钥：随桌面安装包发布在 `public/update-public.pem`
- 私钥：仅存放在本地 `.secrets/` 或服务器 `/opt/wenge-product/secrets/`
- Git：`.secrets/` 已忽略，禁止提交私钥

更新清单格式由 `electron/main-modules/updateManifest.ts` 校验，产品站 `publish-release.mjs` 使用相同的稳定 JSON 排序规则签名。

Windows 安装包还必须使用企业 Authenticode 证书签名并带 RFC3161 时间戳。Release workflow 分离构建签名与发布权限，两个阶段都会用 `Get-AuthenticodeSignature` 验证状态为 `Valid`；未配置受保护的签名凭据时禁止生成可发布资产。

## 发布步骤

1. 更新 `release-notes/<version>.json`，只填写用户可感知变化。
2. 更新 `CHANGELOG.md`、README 和当前架构文档。
3. 在 `desktop` 执行类型检查、Lint、TypeScript 测试和 `.NET Worker` 测试；真实 Office 冒烟只按变更范围定向执行。
4. 递增 `desktop/package.json` 与锁文件版本。
5. 执行 `npm run electron:build`。
6. 使用产品站 `publish-release` 生成 `release.json` 和签名 `manifest.json`。
7. 使用桌面端 `release:verify` 对真实发布目录执行客户端验签、大小和 SHA-256 校验。
8. 上传安装包、blockmap、`latest.yml`、两个 JSON 文件到服务器发布目录。
9. 创建同版本 GitHub Release 并上传安装包资产。
10. 验证产品页下载、桌面端更新检查、SHA-256 和下载统计。

```powershell
cd desktop
npm run release:verify -- `
  --manifest ../product-site/.local/releases/manifest.json `
  --public-key public/update-public.pem `
  --artifact-dir ../product-site/.local/releases
```

## 回滚

- 完整版本：把服务器 `release.json`、`manifest.json` 和 `latest.yml` 恢复到上一份备份，旧版本安装包保留在发布目录。
- 热补丁：发布不含 `hotPatch` 的重新签名清单；客户端补丁加载失败时自动回退内置资源。
- 产品站：服务器每次部署前备份当前应用目录和 Nginx 配置，systemd 单元保持不变。
