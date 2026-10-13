# Docker 与 1Panel 部署

镜像包含 Web 页面、独立协作服务和 Linux SQLite 运行库，导入后无需安装 Node.js、Yarn 或另建数据库。容器默认监听 `0.0.0.0:8080`，首次启动创建空库。

## 生成可导入镜像

构建机器需要安装并启动 Docker，且支持 `docker buildx`。在项目根目录执行：

```sh
corepack yarn docker:export
```

默认生成适用于 Intel / AMD Linux 服务器的文件：

- 文件：`release/docker/mydoc-web-linux-amd64.tar.gz`
- 镜像名：`mydoc-web:latest`
- 容器端口：`8080/TCP`
- 持久化目录：`/app/data`

ARM Linux 服务器使用：

```sh
corepack yarn docker:export --platform linux/arm64
```

生成 `release/docker/mydoc-web-linux-arm64.tar.gz`。可通过 `--tag mydoc-web:2026-10-13` 指定版本，通过 `--output release/docker/mydoc-web.tar.gz` 指定文件。跨架构构建要求 Docker 构建器支持目标架构；Mac 的 `node_modules` 不会复制进镜像，SQLite 原生库在 Linux 构建阶段安装。

镜像包采用 `docker image save` 格式并 gzip 压缩。测试、开发数据库、管理员配置和桌面安装包不会打入运行镜像。

## 1Panel 导入并运行

1. 将镜像包上传到服务器。在 **容器 → 镜像 → 导入镜像** 中选择该 `.tar.gz` 文件并导入。若面板版本只接受 `.tar`，先执行 `gzip -dk mydoc-web-linux-amd64.tar.gz` 再导入解压文件。
2. 在 **容器 → 容器 → 创建容器** 中选择本地镜像 `mydoc-web:latest`，容器名填 `mydoc`，重启策略选择 `unless-stopped`（除非手动停止）。启动命令和工作目录沿用镜像默认值。
3. 添加端口映射：服务器端口 `8080` → 容器端口 `8080`，协议 TCP。服务器端口被占用时可改为 `8081`，容器端口保持 `8080`。
4. 添加目录挂载：服务器目录 `/opt/1panel/apps/mydoc/data` → 容器目录 `/app/data`，设为读写。启动时自动处理该目录的所有权，应用以容器内 `node` 用户（UID/GID 1000）运行。
5. 首次部署建议添加以下环境变量，然后创建并启动容器：

   | 环境变量 | 示例值 |
   | --- | --- |
   | `MYDOC_ADMIN_EMAIL` | 自己的有效邮箱，例如 `admin@example.com` |
   | `MYDOC_ADMIN_PASSWORD` | 自行设置的 10 至 256 位强密码，首尾不能有空白 |
   | `MYDOC_ADMIN_NAME` | `系统管理员`，可省略 |

6. 日志出现 `MyDoc workspace server listening at`、容器健康状态变为 `healthy` 后，通过 `http://服务器IP:8080` 访问并使用配置的管理员账号登录。需要直接访问时在服务器防火墙放行映射的端口。

不配置管理员也能直接启动并注册普通账户，但普通注册不会获得服务器管理员权限。管理员邮箱和密码必须一起设置；只设置其中之一会拒绝启动。初始化完成后重启不会覆盖已修改的账号或密码，调整初始化环境变量也不是重置密码的方法。

挂载目录中的 `cloud.sqlite` 保存账户、工作空间、文档和附件。更新或重建容器时必须继续挂载原目录，否则会打开新的空数据库。不要让多个容器同时操作这份数据库。

## 使用 1Panel 编排

也可以导入镜像后，在 1Panel 的编排界面使用下列配置。请替换邮箱和密码后运行：

```yaml
services:
  mydoc:
    image: mydoc-web:latest
    pull_policy: never
    restart: unless-stopped
    ports:
      - "8080:8080"
    environment:
      MYDOC_ADMIN_EMAIL: admin@example.com
      MYDOC_ADMIN_PASSWORD: replace-with-your-strong-password
      MYDOC_ADMIN_NAME: 系统管理员
    volumes:
      - /opt/1panel/apps/mydoc/data:/app/data
```

仓库根目录也提供了 `compose.yaml`，在导入镜像后执行 `docker compose up -d` 即可启动，默认挂载项目 `data` 目录。首次使用前可取消管理员环境变量的注释并填写账号；它不自带默认密码。

## HTTPS 与配置文件

公网使用建议在 1Panel **网站** 中创建反向代理站点，将上游设置为 `http://127.0.0.1:8080`，申请证书并启用 HTTPS。代理应保留原始 `Host`，传递 `X-Forwarded-Proto https`，关闭 `proxy_buffering`，将读取超时设为 `1h`，以维持 SSE 协作连接。完整 Nginx 示例见根目录 README。

镜像内端口固定按 `8080` 提供健康检查，建议仅修改服务器映射端口。管理员相关的 `MYDOC_ADMIN_EMAIL`、`MYDOC_ADMIN_PASSWORD`、`MYDOC_ADMIN_NAME` 环境变量优先于 `server-config.json` 中的 `admin` 字段；其余 `host`、`port` 和 `cloudDbPath` 仍以配置文件为优先。如果采用该文件，必须将 `host` 设置为 `0.0.0.0`、`port` 设置为 `8080`，避免使用本机示例中的 `127.0.0.1`：

```json
{
  "host": "0.0.0.0",
  "port": 8080,
  "cloudDbPath": "/app/data/cloud.sqlite",
  "admin": {
    "email": "admin@example.com",
    "password": "replace-with-your-strong-password",
    "name": "系统管理员"
  }
}
```

SMTP 通过 1Panel 容器环境变量配置：`MYDOC_SMTP_HOST`、`MYDOC_MAIL_FROM`、`MYDOC_SMTP_PORT`、`MYDOC_SMTP_SECURE`、`MYDOC_SMTP_USER`、`MYDOC_SMTP_PASS`。具体说明见 [工作空间与协作](工作空间与协作.md)。

管理员环境变量只用于数据库首次初始化：如果 `/app/data/cloud.sqlite` 已经存在并完成过管理员初始化，重启或重建容器不会覆盖现有账号密码。需要更换已有管理员密码时，使用网页账户设置中的改密功能；需要重新初始化测试环境时，先停止容器并备份后清空数据目录中的数据库文件。

## 更新与备份

更新时重新生成并导入镜像，在 1Panel 中重建容器，保留原端口、环境变量和目录挂载。使用不同版本标签可以保留旧镜像用于回滚；数据库版本高于旧程序时，旧版本可能无法直接启动，回滚前应保留相应数据库备份。

直接备份数据库前先停止容器，再备份整个服务器数据目录，完成后启动容器。不要在运行时仅复制 `cloud.sqlite` 而遗漏尚未写回的 WAL 数据。

```sh
docker stop mydoc
tar -czf mydoc-data-backup.tar.gz -C /opt/1panel/apps/mydoc data
docker start mydoc
```

排查启动问题时检查容器日志、管理员变量是否完整、映射端口是否占用以及数据目录是否可写。
