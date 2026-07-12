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
- `DATA_DIR`
- `RELEASES_DIR`
- `DATABASE_PATH`
- `USE_ACCEL_REDIRECT=true`

## 验收

```bash
curl -fsS http://127.0.0.1:18120/healthz
curl -fsS https://plugin.shelelove.top/healthz
curl -I https://plugin.shelelove.top/download/windows
systemctl status wenge-product --no-pager
nginx -t
```

还需要在浏览器验证产品页、后台登录、统计周期切换，以及实际下载后后台计数增加。
