# 产品站部署

## 生产拓扑

```text
plugin.shelelove.top:443
  -> Nginx
     -> /releases/windows/*  静态安装包与 electron-updater 元数据
     -> /download/windows    Fastify 记录统计后 X-Accel-Redirect
     -> 其他请求             127.0.0.1:18120
```

服务目录：

```text
/opt/wenge-product/app
/opt/wenge-product/data
/opt/wenge-product/releases
/opt/wenge-product/secrets
```

服务使用独立系统用户 `wenge-product`，systemd 单元为 `wenge-product.service`。Nginx 配置为 `/etc/nginx/conf.d/plugin.shelelove.top.conf`，日志为 `/var/log/nginx/wenge-product.*.log`。

## 隔离原则

- 不停止、删除或修改服务器已有 Docker 容器。
- 不占用已有的 8787、18080-18082、18789、18888、19899 等端口。
- 不修改其他域名的 Nginx 文件和证书。
- 上线前必须执行 `nginx -t`；失败时不 reload。
- 应用只监听 `127.0.0.1:18120`，公网只开放 Nginx 80/443。

## 运行配置

生产环境变量存放在 `/etc/wenge-product.env`，权限 `0600`。必须设置：

- `ADMIN_PASSWORD_HASH`
- `COOKIE_SECRET`
- `ANALYTICS_SALT`
- `ANALYTICS_RETENTION_DAYS=90`（允许 1-3650 天）
- `ANALYTICS_IP_ROTATION_DAYS=30`（允许 1-365 天）
- `DATA_DIR`
- `RELEASES_DIR`
- `DATABASE_PATH`
- `USE_ACCEL_REDIRECT=true`

## 下载统计数据治理

- 不保存原始 IP；服务使用 `ANALYTICS_SALT` 派生周期密钥，按 UTC 周期生成不可逆 HMAC 标识。
- 同一 IP 在一个轮换周期内保持一致，跨周期会生成不同标识。因此跨周期查询的“独立下载”可能重复计数，这是降低长期关联能力的预期取舍。
- 下载记录默认保留 90 天。服务启动时立即清理过期记录，并每 6 小时再次清理；清理定时器不阻塞进程退出。
- 清理失败只写告警，不阻塞新的统计写入或安装包下载。生产监控应对 `download analytics maintenance failed` 告警。
- UA 最多保存 200 字符；Referer 只保存 `scheme://host[:port]`，不保留路径、查询参数或片段。首次升级会对既有记录执行同样的数据最小化迁移。

## 验收

```bash
curl -fsS http://127.0.0.1:18120/healthz
curl -fsS https://plugin.shelelove.top/healthz
curl -I https://plugin.shelelove.top/download/windows
systemctl status wenge-product --no-pager
nginx -t
```

还需要在浏览器验证产品页、后台登录、统计周期切换，以及实际下载后后台计数增加。
