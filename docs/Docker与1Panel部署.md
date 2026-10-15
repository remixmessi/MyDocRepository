# Docker 与 1Panel 部署

镜像包含 Web 页面、独立协作服务和 Linux SQLite 运行库，导入后无需安装 Node.js、Yarn 或另建数据库。容器默认监听 `0.0.0.0:8080`。从 `2026.10.14-1` 起必须显式挂载 `/app/data`；默认只打开已有数据库，缺库时拒绝启动。只有首次安装可临时设置 `MYDOC_INIT_DATABASE=1`，初始化后删除该变量。

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

生成 `release/docker/mydoc-web-linux-arm64.tar.gz`。可通过 `--tag mydoc-web:2026.10.15-1` 指定版本（同时写入镜像版本标签和 `MYDOC_VERSION`），通过 `--output release/docker/mydoc-web.tar.gz` 指定文件。跨架构构建要求 Docker 构建器支持目标架构；Mac 的 `node_modules` 不会复制进镜像，SQLite 原生库在 Linux 构建阶段安装。

镜像包采用 `docker image save` 格式并 gzip 压缩，同时生成 `.tar.gz.sha256` 校验文件。测试、开发数据库、管理员配置和桌面安装包不会打入运行镜像。

## 1Panel 导入并运行

1. 将镜像包上传到服务器。在 **容器 → 镜像 → 导入镜像** 中选择该 `.tar.gz` 文件并导入。若面板版本只接受 `.tar`，先执行 `gzip -dk mydoc-web-linux-amd64.tar.gz` 再导入解压文件。
2. 在 **容器 → 容器 → 创建容器** 中选择本地镜像 `mydoc-web:latest`，容器名填 `mydoc`，重启策略选择 `unless-stopped`（除非手动停止）。启动命令和工作目录沿用镜像默认值。
3. 添加端口映射：服务器端口 `8080` → 容器端口 `8080`，协议 TCP。服务器端口被占用时可改为 `8081`，容器端口保持 `8080`。
4. 添加目录挂载：服务器目录 `/opt/1panel/apps/mydoc/data` → 容器目录 `/app/data`，设为读写。启动时自动处理该目录的所有权，应用以容器内 `node` 用户（UID/GID 1000）运行。
5. 全新安装时临时设置 `MYDOC_INIT_DATABASE=1`；已有数据升级时不要设置。建议添加以下管理员环境变量，然后创建并启动容器：

   | 环境变量 | 示例值 |
   | --- | --- |
   | `MYDOC_ADMIN_EMAIL` | 自己的有效邮箱，例如 `admin@example.com` |
   | `MYDOC_ADMIN_PASSWORD` | 自行设置的 10 至 256 位强密码，首尾不能有空白 |
   | `MYDOC_ADMIN_NAME` | `系统管理员`，可省略 |

6. 日志会打印实际数据库路径。首次安装成功后移除 `MYDOC_INIT_DATABASE`，以后保持关闭。日志出现 `MyDoc workspace server listening at`、容器健康状态变为 `healthy` 后，通过 `http://服务器IP:8080` 访问并使用配置的管理员账号登录。需要直接访问时在服务器防火墙放行映射的端口。

不配置管理员也能直接启动并注册普通账户，但普通注册不会获得服务器管理员权限。管理员邮箱和密码必须一起设置；只设置其中之一会拒绝启动。初始化完成后重启不会覆盖已修改的账号或密码，调整初始化环境变量也不是重置密码的方法。

挂载目录中的 `cloud.sqlite` 保存账户、工作空间、文档和附件。更新或重建容器时必须复用原来的服务器绝对路径，或原来的 Docker 卷名。仅保持容器端路径 `/app/data` 相同并不够。旧版镜像可能自动分配匿名卷，创建新容器不会自动复用它；新版取消自动匿名卷并检查挂载，找不到原库时拒绝启动。不要让多个容器同时操作这份数据库。

## 使用 1Panel 编排

也可以导入镜像后，在 1Panel 的编排界面使用下列配置。请替换邮箱和密码后运行：

```yaml
services:
  mydoc:
    image: mydoc-web:2026.10.15-1
    pull_policy: never
    restart: unless-stopped
    ports:
      - "8080:8080"
    environment:
      MYDOC_INIT_DATABASE: "0"
      MYDOC_ADMIN_EMAIL: admin@example.com
      MYDOC_ADMIN_PASSWORD: replace-with-your-strong-password
      MYDOC_ADMIN_NAME: 系统管理员
    volumes:
      - type: bind
        source: /opt/1panel/apps/mydoc/data
        target: /app/data
        bind:
          create_host_path: false
```

仓库根目录也提供了 `compose.yaml`，默认复用服务器目录 `/opt/1panel/apps/mydoc/data`，不再使用随编排位置变化的相对路径。目录必须已存在，错误路径不会自动创建。可用 `MYDOC_DATA_DIR` 指定原来的绝对路径。首次安装先创建空目录，再临时设置 `MYDOC_INIT_DATABASE=1` 启动；已有数据升级保持 `0`。管理员环境变量按示例填写；它不自带默认密码。

## HTTPS 与配置文件

公网使用建议在 1Panel **网站** 中创建反向代理站点，将上游设置为 `http://127.0.0.1:8080`，申请证书并启用 HTTPS。代理应保留原始 `Host`，传递 `X-Forwarded-Proto https`，关闭 `proxy_buffering`，将读取超时设为 `1h`，以维持 SSE 协作连接。完整 Nginx 示例见根目录 README。

镜像内端口固定按 `8080` 提供健康检查，建议仅修改服务器映射端口。所有显式设置的 `MYDOC_*` 部署配置均优先于 `server-config.json`。镜像默认固定数据库为 `/app/data/cloud.sqlite`，旧配置文件不会把数据库切换到另一条路径；确需自定义时设置 `MYDOC_CLOUD_DB_PATH`，且必须位于 `/app/data` 内。配置文件可作为未设置环境变量时的回退：

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

管理员环境变量用于初始化管理员：如果 `/app/data/cloud.sqlite` 已经完成过管理员初始化，重启或重建容器不会覆盖现有账号密码。需要更换已有管理员密码时，使用网页账户设置中的改密功能。升级时不要删除数据库；`MYDOC_INIT_DATABASE=1` 也不会覆盖已有数据库，目录内有孤立 WAL 或其他数据但缺少主库时会拒绝初始化。

## 更新与备份

`2026.10.15-1` 在打开带 WAL 的原库前，先把主库、WAL、SHM 和回滚日志原样保存到 `/app/data/upgrade-inputs/input-*`（包含文件大小和 SHA256 清单），再在临时副本上比较账号 ID。应用 WAL 后若主库已有账号 ID 消失，立即停止，不打开原库。应用没有删除账号或重分配账号 ID 的功能；此检查不拦截正常注册、改名、改邮箱及删除团队空间。没有 WAL 时跳过原始文件归档。

随后在任何数据库迁移前生成一致的 SQLite 整库快照，保存到 `/app/data/upgrade-backups/before-upgrade-*.sqlite`，旁边的 `.json` 记录镜像版本、升级前结构版本和记录数量。快照包含已提交到 WAL 的数据，保留旧结构，可用于旧镜像回滚。相同镜像、结构版本与数据库身份复用已完成的快照；结构升级后的首次重启可能再产生一份。保护快照和原始文件归档均不自动删除，需自行归档并监控磁盘空间。快照失败则停止启动，不执行迁移。

迁移在同一个写事务中逐值核对原文档、附件、偏好、账号及空间成员；发生删改则整体回滚。已有库缺少基础业务表或没有账号及业务数据时拒绝升级，不通过补建空表伪装成功。正式升级保持 `MYDOC_INIT_DATABASE=0`。

恢复时先停止所有使用数据库的容器，将当前目录完整留存，然后把升级前快照复制到**新的空目录**并命名为 `cloud.sqlite`，让旧镜像挂载该目录。不要把快照覆盖到留有其他数据库 `-wal`／`-shm` 的目录，混用主库与日志可能导致恢复结果错误，甚至完整性检查通过但账号和空间都不见；见 [SQLite 官方对数据库与日志混配的说明](https://www.sqlite.org/howtocorrupt.html#_mispairing_database_files_and_hot_journals)。发布文件目录 `releases` 也应从升级前目录复制保留。

更新前先用 `docker inspect` 确认旧容器 `/app/data` 对应的 `Source`（服务器目录）或 `Name`（Docker 卷名），停止旧容器并备份后，新容器必须挂载同一个来源。升级期间不要开启 `MYDOC_INIT_DATABASE`，不要勾选删除存储卷。容器路径相同不表示来源相同。使用不同版本标签可以保留旧镜像用于回滚；数据库版本高于旧程序时，旧版本会拒绝启动，需要用升级前备份恢复到单独目录后再回滚，不要删除现有数据库。

直接备份数据库前先停止容器，再备份整个服务器数据目录，完成后启动容器。不要在运行时仅复制 `cloud.sqlite` 而遗漏尚未写回的 WAL 数据。

```sh
docker stop mydoc
tar -czf mydoc-data-backup.tar.gz -C /opt/1panel/apps/mydoc data
docker start mydoc
```

排查启动问题时检查容器日志、管理员变量是否完整、映射端口是否占用以及数据目录是否可写。

## 升级后数据不可见：先找原数据卷

镜像文件只含程序，回滚镜像不会恢复旧卷或已删除的数据库。先停止继续写入，保留新旧容器和全部存储卷，不执行 `docker volume prune`、`docker system prune --volumes`，不删除或初始化目录。

以下只读取容器挂载信息，不输出密码。对新旧容器分别执行（替换容器名）：

```sh
docker ps -a --format '{{.Names}}'
docker inspect 容器名 --format '{{json .Mounts}}'
docker volume ls
```

查看 `Destination` 为 `/app/data` 的记录：

- `Type=bind`：比对新旧容器 `Source`，应为完全相同的服务器目录。
- `Type=volume`：比对 `Name`，不同名称表示不同数据卷。旧容器已删但未删除存储卷时，原匿名卷仍可能保留；用 `docker volume inspect 卷名` 查询其 `Mountpoint`，检查 `cloud.sqlite` 的大小、时间及备份目录。

若旧容器仍在，先停止所有使用该目录的容器，再复制到**新的恢复目录**留存（不要覆盖当前目录）：

```sh
docker stop 旧容器名
docker cp 旧容器名:/app/data ./mydoc-recovered-data
```

应一起保留主库、`-wal`、`-shm`、备份和附件／发布目录。找到原数据后，先完整备份，再让新镜像挂载原卷，或已核验的恢复副本；保持初始化开关关闭。如果原库和卷已被删除，需要从独立备份或磁盘恢复，重新导入镜像无法恢复文档。

## 本地升级回归验证

构建新版后，可在隔离的临时 Docker 卷中验证旧版到新版的完整页面流程：

```sh
MYDOC_OLD_IMAGE=mydoc-web:latest MYDOC_UPGRADE_IMAGE=mydoc-web:2026.10.15-1 node scripts/docker/verify-upgrade.mjs
```

验证包含真实私人文档、团队文档和附件下载、强制终止后的 WAL 恢复、迁移前后业务记录摘要一致，以及两次重建容器后的页面内容。需要安装开发依赖及 Playwright Chromium。脚本结束后清理自身的临时容器和卷。

异常 WAL 回归验证（只操作脚本创建的临时卷）：

```sh
MYDOC_OLD_IMAGE=mydoc-web:latest MYDOC_UPGRADE_IMAGE=mydoc-web:2026.10.15-1 node scripts/docker/verify-wal-recovery.mjs
```

该用例复现：恢复的主库配上另一份 WAL，完整性检查仍通过，但原账号和空间消失。验证最终镜像在打开原库前拒绝启动，原主库和 WAL 的 SHA256 不变，原始文件归档可读。账号 ID 对比依据应用不删除／重分配账号 ID 的约定，不能识别所有日志混配；恢复仍应使用正确的整组文件或一致性快照。程序不会通过删除 WAL 来猜测恢复结果。
