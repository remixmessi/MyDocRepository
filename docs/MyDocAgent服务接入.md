# MyDocAgent 服务接入（A1 部分实现）

当前范围：个人提示词问答、持久运行与事件分页。每个 Run 独立输入，不自动载入前次会话、文档、附件或团队内容。无编辑工具、SSE、模型配置 UI；不能当成 A1 整体验收完成。

## 配置与构建

宿主启动前设置 `MYDOC_AGENT_ENDPOINT`（完整 chat/completions 地址）、`MYDOC_AGENT_MODEL`，可选 `MYDOC_AGENT_API_KEY`。密钥只用于宿主请求，不写数据库或传给子进程。模型配置源目前只有宿主环境，不支持 renderer 任意指定地址；未配置时 status.enabled=false，运行接口返回 503。

公网地址必须 HTTPS，HTTP 仅允许明确配置的回环模型服务。当前单次输出上限 2048、输入预算 24000、运行限时 120 秒、并发会话上限 4、账户最多 100 个会话/100 个排队任务。真实供应商能力与系统密钥存储仍待后续验收。

`server:build` 与桌面 build 同时生成 `agent-worker.cjs`，位于服务端 bundle 或桌面 main.cjs 同级。路径由应用固定，客户端不能提供。桌面 fork 使用 ELECTRON_RUN_AS_NODE；真实 Electron 安装包运行另行验收，Node 子进程测试不代替实机。

## API

统一前缀 `/api/agent`。Web 复用 CloudServer Cookie 会话，scope=`user:<id>`；桌面复用本地服务 Bearer token，scope=`main`（与现有桌面存储一致，不冒充独立本地账户隔离）。请求提供 `X-Mydoc-Account` 可拒绝旧账户页面；POST 还必须携带当前 `X-Mydoc-Generation`。跨站请求拒绝，所有请求体有大小限制。

| 方法与路径 | 行为 |
| --- | --- |
| GET /status | 是否配置模型及当前范围，不返回密钥 |
| GET /sessions | 当前账户个人会话，最多 100 条 |
| POST /sessions | `{title}`，创建个人会话，返回 id |
| POST /runs | `{session,requestKey,prompt}`，先持久入队再调度，返回 202；同 key 同内容复用 Run |
| GET /runs?id=… | 本账户运行状态；其他账户返回不可用 |
| POST /cancel | `{id}`，取消本账户活动/排队 Run |
| GET /events?session=…&after=… | 按会话 cursor 递增返回最多 200 条，客户端推进 cursor；当前是 JSON 分页而非 SSE |

输入不能指定 user/scope、worker 路径、供应商地址或工具。运行中每次模型调用/输出都检查原 Cookie 会话，并定时检查注销/过期；失效取消活动 Run。已发往模型的内容无法撤回。不会把原 Cookie 或 API key 传给 worker。

## 生命周期与恢复

AgentService 在数据库已完成公共/协作升级后、服务监听前创建；仅这一启动 owner 恢复活动 Run，不在每个请求重新恢复。待运行任务保留，需该账户重新认证请求触发调度；不凭保存的 Cookie 自动运行。运行中断不盲重放写操作。

关闭先停止准入、取消并等待子进程，再关闭办公服务与数据库。事件事务提交后才回复 worker，跨账户事件不可读；恢复数据库会通过既有 generation 拒绝旧写入。当前不宣称支持跨服务多实例共享运行 owner。

## 验证范围

真实 HTTP 测试复用 CloudServer 注册/登录 Cookie，覆盖账户隔离、跨站/错账户/旧代际拒绝、本地合成模型流、实际子进程执行、持久事件、请求去重及注销失效。桌面服务测试验证已有 Bearer token 与 main scope，无 UI 或真实 Electron 启动。
