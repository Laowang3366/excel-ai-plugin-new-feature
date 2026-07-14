# 更新与发布

## 更新类型

### 完整安装包

完整版本使用 electron-builder 生成 NSIS x64 安装包、blockmap 和 `latest.yml`。桌面端通过 `electron-updater` 下载，校验签名清单中的 SHA-256 后调用 NSIS 覆盖安装。

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

界面补丁必须包含 `dist/index.html`；知识库补丁必须包含 `builtin-knowledge.json`；WPS 桥接补丁必须包含 `index.html`。补丁声明精确基础版本，下载后校验签名、文件大小和 SHA-256，再写入临时目录并原子切换。加载失败会删除激活状态并回退安装包内置界面。

## 签名

- 算法：Ed25519
- 公钥：随桌面安装包发布在 `public/update-public.pem`
- 私钥：仅存放在本地 `.secrets/` 或服务器 `/opt/wenge-product/secrets/`
- Git：`.secrets/` 已忽略，禁止提交私钥

更新清单格式由 `electron/main-modules/updateManifest.ts` 校验，产品站 `publish-release.mjs` 使用相同的稳定 JSON 排序规则签名。

## 发布步骤

1. 更新 `release-notes/<version>.json`，只填写用户可感知变化。
2. 更新 `CHANGELOG.md`、README 和当前架构文档。
3. 在 `desktop` 执行完整测试、类型检查和 Lint。
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
