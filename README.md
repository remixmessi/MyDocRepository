![MyDoc 项目图标](packages/frontend/apps/web/public/favicon.png)

# MyDoc · 文档库

**把文档、白板和知识管理放进同一个工作空间。**

本地优先 · Markdown 编辑 · 无限画布 · 团队协作 · 自托管

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![项目状态：开发中](https://img.shields.io/badge/状态-开发中-f59e0b)
![TypeScript 5.8](https://img.shields.io/badge/TypeScript-5.8-3178c6?logo=typescript&logoColor=white)
![React 18](https://img.shields.io/badge/React-18-149eca?logo=react&logoColor=white)
![SQLite 存储](https://img.shields.io/badge/存储-SQLite-003b57?logo=sqlite&logoColor=white)
![Electron 桌面应用](https://img.shields.io/badge/桌面-Electron-47848f?logo=electron&logoColor=white)

[项目介绍](#项目介绍) · [核心功能](#核心功能) · [快速开始](#快速开始) · [部署与打包](#部署与打包) · [文档导航](#文档导航) · [参与贡献](#参与贡献)

---

## 项目介绍

MyDoc 是一个将 Markdown 文档、白板和知识管理结合起来的文档应用，适合个人笔记、项目资料整理和团队知识协作。你可以在同一篇文档中切换页面与白板视图，用双向链接串联内容，通过搜索、文件夹和标签管理资料。

项目提供 Web 与 macOS / Windows 桌面入口。个人本地模式使用 SQLite 保存文档与附件；团队模式通过自托管服务提供账号、工作空间和实时协作。正文与白板数据使用 Yjs 管理，编辑器基于 ProseMirror 构建。

> 项目持续开发中。桌面开发包尚未完成正式签名与公证配置，自动更新渠道尚未启用；具体实现进度见[功能清单](docs/TODO-功能清单.md)。

## 核心功能

| 功能 | 支持能力 |
| --- | --- |
| 📝 文档编辑 | Markdown 快捷输入、源码模式、表格、代码高亮、任务列表、标注框、LaTeX 公式与 Mermaid 图表 |
| 🎨 白板画布 | 图形与连线、手绘、思维导图、图片、富文本笔记、分组、图层与画框 |
| 🔗 知识连接 | 文档提及、双向链接、文档嵌入、白板引用、标题大纲与悬浮预览 |
| 🗂️ 内容管理 | 多工作区、嵌套文件夹、标签、收藏、集合、模板、日记与回收站 |
| 🔍 查找与工作台 | 全文搜索、页内查找、统一命令面板、多标签页、分屏与导航历史 |
| 📦 导入与导出 | Markdown / Notion 导入、文件夹与 ZIP 导入、Markdown / HTML / PDF 导出、工作区备份 |
| 💾 数据与备份 | SQLite 持久化、自动整库快照、恢复预览、恢复前保护备份 |
| 🤝 团队协作 | 自托管账号与工作空间、成员邀请、角色管理、实时编辑、协作光标与共享附件 |
| 🖥️ 桌面体验 | macOS / Windows 打包、文件关联、拖放导入、`mydoc://` 深链与远程工作空间连接 |

完整操作与快捷键见[使用指南](docs/使用指南.md)。团队成员可读取空间内全部文档，文档授权目前控制编辑权限；详细范围与离线行为见[工作空间与协作](docs/工作空间与协作.md)。

## 快速开始

### 环境准备

- Node.js 22 LTS。
- Corepack 与 Yarn 4（项目固定使用 `yarn@4.9.2`）。
- Git；Docker 部署时另需安装并启动 Docker。

获取仓库代码后，在项目根目录执行：

```sh
corepack enable
corepack yarn install --immutable
corepack yarn dev
```

浏览器访问 [http://127.0.0.1:5173](http://127.0.0.1:5173)。开发服务同时监听 `0.0.0.0`，同一局域网设备可通过本机 IP 的 5173 端口访问。

首次启动创建空数据库，不写入示例文档。个人本地数据默认保存在 `data/mydoc.sqlite`。

所有端每次启动都会检查数据库结构与完整性，并按版本增量升级表和字段，保留已有数据。迁移失败会回滚并停止启动；新功能的数据库变更规范见[数据库迁移](docs/数据库迁移.md)。

### 启动桌面应用

```sh
corepack yarn desktop:dev
```

该命令先构建 Web 与 Electron，再打开桌面应用。源码修改后重新执行即可；已有构建可使用 `corepack yarn desktop:start`。Electron 运行时下载与平台要求见[桌面打包说明](docs/桌面打包.md)。

## 部署与打包

### 自托管 Web 服务

```sh
corepack yarn server:build
corepack yarn server:start
```

默认访问 [http://127.0.0.1:8080](http://127.0.0.1:8080)，注册或登录后使用。可复制 `data/server-config.example.json` 为 `data/server-config.json`，修改管理员邮箱和密码后首次启动；管理员密码至少 10 位，后续启动不会覆盖已有账号。

公共部署使用 HTTPS 反向代理。配置项、systemd、Nginx 与备份步骤见 [Web 部署与维护](docs/Web部署与维护.md)。

### Docker / 1Panel

```sh
# 默认导出 Linux amd64 镜像
corepack yarn docker:export

# ARM 服务器
corepack yarn docker:export --platform linux/arm64
```

默认产物为 `release/docker/mydoc-web-linux-amd64.tar.gz`。在 1Panel 导入镜像后使用 `mydoc-web:latest`，映射端口 `8080:8080`，将持久化数据目录挂载到 `/app/data`。

详细配置与更新步骤见 [Docker 与 1Panel 部署](docs/Docker与1Panel部署.md)。

### 桌面安装包

```sh
# 在 macOS 上构建 Apple Silicon 与 Intel 安装包
corepack yarn desktop:dist:mac

# 在 Windows 上构建 x64 安装包
corepack yarn desktop:dist:win
```

产物位于 `release/desktop/`。本地开发包、正式签名构建与发布配置分别见[桌面打包](docs/桌面打包.md)和[桌面发布](docs/桌面发布.md)。

### 存储与备份

| 运行方式 | 默认数据位置 |
| --- | --- |
| Web 本地开发 / 预览 | 项目内 `data/mydoc.sqlite` |
| 独立 Web 协作服务 | 项目内 `data/cloud.sqlite`；Docker 内为 `/app/data/cloud.sqlite` |
| 桌面源码开发 | 项目内 `data/mydoc.sqlite` |
| 已安装桌面应用 | 应用所在目录的 `data/mydoc.sqlite`，可从「文件 → 打开数据目录」查看 |

Web 服务需要 SQLite 接口，不能只将前端构建产物上传到静态主机。浏览器在整个服务不可达时没有独立持久化通道，关闭页面前应导出未保存草稿。

迁移单个工作区可使用 `.mydoc.json` 备份；整库恢复会替换全部工作区数据。直接复制 SQLite 文件前先停止服务；安装后的桌面应用卸载会删除应用目录中的数据库，请先保留独立备份。具体步骤见[使用指南](docs/使用指南.md#自动备份与整库恢复)。

## 开发与验证

| 命令 | 用途 |
| --- | --- |
| `corepack yarn dev` | 启动 Web 开发服务 |
| `corepack yarn build` | 构建 Web，输出到 `packages/frontend/apps/web/dist` |
| `corepack yarn preview` | 本机预览 Web 构建产物 |
| `corepack yarn typecheck` | TypeScript 类型检查 |
| `corepack yarn test` | 单元测试 |
| `corepack yarn test:e2e` | 浏览器端到端测试 |
| `corepack yarn test:cloud` | 团队协作端到端测试 |
| `corepack yarn test:desktop` | 运行 Electron 桌面测试，需先执行 `desktop:build` |

首次运行浏览器测试前安装 Chromium：

```sh
corepack yarn exec playwright install chromium
```

### 项目结构

```text
MyDocRepository/
├── assets/branding/          # 项目图标与品牌素材
├── packages/
│   ├── common/
│   │   ├── crdt/             # Yjs 文档模型
│   │   └── nbstore/          # 存储、同步与备份
│   ├── editor/
│   │   ├── framework/        # 编辑器框架
│   │   └── blocks/           # 内容节点与编辑能力
│   └── frontend/
│       ├── core/             # Web 与桌面共用界面
│       └── apps/
│           ├── web/          # Web 应用与服务端
│           └── desktop/      # Electron 桌面应用
├── docs/                     # 使用、部署与设计文档
├── scripts/                  # 构建、打包与发布脚本
├── tests/                    # 单元与端到端测试
└── data/                     # 本地运行数据与配置
```

## 文档导航

| 文档 | 内容 |
| --- | --- |
| [使用指南](docs/使用指南.md) | 编辑器、白板、文档管理、导入导出与备份操作 |
| [Web 部署与维护](docs/Web部署与维护.md) | 独立服务、systemd、Nginx、更新与数据备份 |
| [Docker 与 1Panel 部署](docs/Docker与1Panel部署.md) | 镜像导出、容器配置、数据挂载与升级 |
| [桌面打包](docs/桌面打包.md) | Electron 运行时、平台架构与安装包构建 |
| [桌面发布](docs/桌面发布.md) | 签名、公证、版本维护与更新配置 |
| [工作空间与协作](docs/工作空间与协作.md) | 账号、成员权限、实时同步与离线限制 |
| [功能清单](docs/TODO-功能清单.md) | 已实现功能与后续规划 |
| [实现方案](docs/实现方案.md) | 架构与实现设计 |
| [开发批次](docs/开发批次.md) | 分批开发与验收记录 |

## 参与贡献

欢迎通过仓库的 Issue 或 Pull Request 提交问题与改进：

- **报告问题**：说明运行方式、系统与版本、复现步骤、预期结果和实际结果；附上必要日志或截图。
- **建议功能**：描述使用场景与希望解决的问题，可先参考功能清单中的计划。
- **提交代码**：保持改动聚焦，更新相关文档，运行类型检查与单元测试，并按改动范围执行浏览器、协作或桌面测试。

请勿在提交中包含运行数据库、账号配置、证书或密钥。

## 致谢

项目在工作空间、编辑器和白板设计上参考了 [AFFiNE](https://github.com/toeverything/AFFiNE)，并使用 React、ProseMirror、Yjs、SQLite、Electron 等开源技术。

## 许可证

本项目采用 [MIT License](LICENSE) 开源许可。第三方依赖与引用代码遵循各自的许可证。
