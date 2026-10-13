# Web 部署与维护

[返回项目首页](../README.md) · [使用指南](使用指南.md)

### Docker 镜像与 1Panel

```sh
corepack yarn docker:export
```

需要安装并启动 Docker。默认生成 Linux amd64 镜像包 `release/docker/mydoc-web-linux-amd64.tar.gz`，可在 1Panel 的「容器 → 镜像 → 导入镜像」中导入，选择 `mydoc-web:latest` 创建容器。映射 `8080:8080`，将服务器数据目录挂载到 `/app/data`，使用默认启动命令即可运行。通过环境变量 `MYDOC_ADMIN_EMAIL` 和 `MYDOC_ADMIN_PASSWORD` 初始化管理员；环境变量优先于挂载目录中的 `server-config.json`，但只在数据库首次初始化时创建管理员，不会覆盖已有账号密码。

ARM 服务器使用 `corepack yarn docker:export --platform linux/arm64`。构建、1Panel 创建容器、HTTPS、更新和备份步骤见 [Docker 与 1Panel 部署](Docker与1Panel部署.md)。

### 本机预览构建产物

在项目根目录执行以下命令生成 Web 生产产物：

```sh
corepack yarn install --immutable
corepack yarn build
```

产物位于 `packages/frontend/apps/web/dist`。构建完成后，可用 Vite 预览：

```sh
corepack yarn preview
```

默认访问 `http://127.0.0.1:5173`。修改源码后需要重新执行 `corepack yarn build`。同一端口已有服务运行时，可以换端口：

```sh
corepack yarn workspace @mydoc/web exec vite preview --host 127.0.0.1 --port 4173
```

`preview` 适合本机验收，不适合作为公共团队服务。仅上传 `packages/frontend/apps/web/dist` 到静态主机不能提供项目 SQLite 存储接口。

### 部署到 Linux 服务器

独立服务必须在目标服务器上安装依赖并构建，不能直接把其他操作系统的 `node_modules` 复制过去。以下示例以 Debian/Ubuntu、部署目录 `/srv/mydoc` 和运行用户 `mydoc` 为例；路径和域名请替换为实际值。

1. 安装 Node.js 22 LTS、Git，并准备运行用户：

   ```sh
   sudo apt update
   sudo apt install -y ca-certificates curl git
   curl -fsSL https://deb.nodesource.com/setup_22.x | sudo -E bash -
   sudo apt install -y nodejs
   node --version                  # 应为 22.x
   sudo corepack enable
   sudo useradd --system --home /srv/mydoc --shell /usr/sbin/nologin mydoc
   sudo mkdir -p /srv
   sudo chown mydoc:mydoc /srv
   ```

2. 获取代码并在项目根目录安装依赖、生成 Web 和服务端产物：

   ```sh
   sudo -u mydoc git clone <仓库地址> /srv/mydoc
   cd /srv/mydoc
   sudo -u mydoc corepack yarn install --immutable
   sudo -u mydoc corepack yarn server:build
   sudo -u mydoc mkdir -p data
   ```

   `server:build` 会先生成 `packages/frontend/apps/web/dist`，再生成服务端入口 `packages/frontend/apps/web/server-dist/main.cjs`。服务必须从项目根目录启动，因为页面产物和默认数据库路径按当前工作目录解析。

3. 首次启动前创建服务配置文件。复制 `data/server-config.example.json` 为 `data/server-config.json`，填写管理员邮箱和密码；配置文件不会提交到 Git：

   ```sh
   cp data/server-config.example.json data/server-config.json
   chmod 600 data/server-config.json
   ```

   邮件通知可继续设置 `MYDOC_SMTP_HOST`、`MYDOC_MAIL_FROM`、`MYDOC_SMTP_PORT`、`MYDOC_SMTP_SECURE`、`MYDOC_SMTP_USER` 和 `MYDOC_SMTP_PASS`。不配置 SMTP 时，账户相关的一次性口令不会通过邮件发送。也可用 `MYDOC_CONFIG_PATH` 指定配置文件；旧环境变量仍兼容。服务支持的全部配置项和安全约束见 [工作空间与 Web 协作](工作空间与协作.md)。

4. 前台启动并检查服务：

   ```sh
   cd /srv/mydoc
   corepack yarn server:start
   ```

   默认访问 `http://127.0.0.1:8080`。另开终端检查：

   ```sh
   curl -I http://127.0.0.1:8080/
   ```

   需要直接监听服务器网卡时设置 `MYDOC_HOST=0.0.0.0`；公共部署建议仍监听 `127.0.0.1`，由 HTTPS 反向代理对外提供服务。启动日志出现 `MyDoc workspace server listening at` 即表示 HTTP 服务已监听。

### 使用 systemd 常驻运行

生产服务器建议使用 systemd 管理进程。先确认 `corepack` 的绝对路径：

```sh
command -v corepack
```

创建只允许 root 读取的环境文件（把值替换为实际配置）：

```sh
sudo install -d -m 750 /etc/mydoc
sudo install -m 600 /dev/null /etc/mydoc/mydoc.env
sudoedit /etc/mydoc/mydoc.env
```

文件内容示例：

```dotenv
NODE_ENV=production
MYDOC_CONFIG_PATH=/etc/mydoc/server-config.json
```

创建 `/etc/systemd/system/mydoc.service`（将 `ExecStart` 中的 `/usr/bin/corepack` 替换为上一步 `command -v corepack` 的结果）：

```ini
[Unit]
Description=MyDoc Web service
After=network.target

[Service]
Type=simple
User=mydoc
Group=mydoc
WorkingDirectory=/srv/mydoc
EnvironmentFile=/etc/mydoc/mydoc.env
ExecStart=/usr/bin/corepack yarn server:start
Restart=on-failure
RestartSec=5
KillSignal=SIGINT
TimeoutStopSec=15
NoNewPrivileges=true
PrivateTmp=true

[Install]
WantedBy=multi-user.target
```

加载并启动：

```sh
sudo systemctl daemon-reload
sudo systemctl enable --now mydoc
sudo systemctl status mydoc
sudo journalctl -u mydoc -f
```

停止、重启和查看最近日志：

```sh
sudo systemctl stop mydoc
sudo systemctl restart mydoc
sudo journalctl -u mydoc -n 100 --no-pager
```

### Nginx HTTPS 反向代理

安装 Nginx 并申请证书后，让 MyDoc 只监听 `127.0.0.1:8080`，由 Nginx 负责 TLS：

```sh
sudo apt install -y nginx
```

将 `docs.example.com` 和证书路径替换为实际值，创建站点配置：

```nginx
server {
    listen 80;
    server_name docs.example.com;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name docs.example.com;

    ssl_certificate     /etc/letsencrypt/live/docs.example.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/docs.example.com/privkey.pem;

    location / {
        proxy_pass http://127.0.0.1:8080;
        proxy_http_version 1.1;
        proxy_buffering off;
        proxy_read_timeout 1h;
        proxy_send_timeout 1h;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto https;
    }
}
```

`proxy_buffering off` 和较长的读取超时用于保持 SSE 协作连接；`X-Forwarded-Proto https` 用于让服务端正确设置 Secure 会话 Cookie。配置完成后执行 `sudo nginx -t && sudo systemctl reload nginx`，再通过 `https://docs.example.com` 访问。防火墙只需开放 80/443，不要把 8080 直接暴露到公网。

### 更新、回滚与数据备份

更新代码前先备份 `data/cloud.sqlite`；服务运行时还可能有同目录的 `cloud.sqlite-wal` 和 `cloud.sqlite-shm`，直接复制数据库前必须先停止服务。更新和重启命令如下：

```sh
cd /srv/mydoc
sudo systemctl stop mydoc
sudo -u mydoc cp -a data/cloud.sqlite "data/cloud.sqlite.$(date +%Y%m%d%H%M%S).bak"
sudo -u mydoc git pull --ff-only
sudo -u mydoc corepack yarn install --immutable
sudo -u mydoc corepack yarn server:build
sudo systemctl start mydoc
sudo systemctl status mydoc
```

同一个 `cloud.sqlite` 只能运行一个独立协作服务实例；不要同时用 `corepack yarn dev`、`corepack yarn preview` 或另一份 `server:start` 连接生产数据库。SQLite 数据库不应放在 NFS 等不保证文件锁语义的共享目录中。

