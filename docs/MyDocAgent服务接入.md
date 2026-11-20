# MyDocAgent 服务接入（A1 部分实现）

当前范围：个人提示词问答、持久运行与事件分页。每个 Run 独立输入，不自动载入前次会话、文档、附件或团队内容。已接入 Agent 面板、Web SSE 与桌面分页订阅；无文档编辑工具，不能当成 A1 整体验收完成。

## 配置与构建

共享服务端只用管理员环境配置：`MYDOC_AGENT_ENDPOINT`（所选协议的完整推理端点地址）、`MYDOC_AGENT_MODEL`，可选 `MYDOC_AGENT_API_TYPE`（`openai-chat-completions`，默认；或 `openai-responses` / `anthropic-messages`）、`MYDOC_AGENT_API_KEY`、`MYDOC_AGENT_MAX_OUTPUT_TOKENS`（默认 2048）、`MYDOC_AGENT_CONTEXT_WINDOW`（默认 32000）、`MYDOC_AGENT_TOOLS=1`、`MYDOC_AGENT_USAGE=1`，以及逗号分隔的 `MYDOC_AGENT_ALLOWED_HOSTS`（设置后环境地址不在列表内则启动失败）。用户不能在服务端保存模型。

桌面可经 `POST /api/agent/model` 保存本机模型；设置了环境模型时以环境为准并拒绝修改。非密钥字段存于 `preferences` 的 `agent-config` scope，密钥用 Electron safeStorage 加密写入 userData/agent-secrets，数据库只存 secretRef。系统加密不可用时拒绝保存密钥。备份不含密钥，换机恢复后显示密钥不可用、需重新输入。保留旧密钥时更换主机必须重新输入。未配置时 status.enabled=false，运行接口返回 503。

公网地址必须 HTTPS，HTTP 仅允许明确配置的回环模型服务。当前单次输出上限 2048、输入预算 24000、运行限时 120 秒、并发会话上限 4、账户最多 100 个会话/100 个排队任务。真实供应商能力与系统密钥存储仍待后续验收。

`server:build` 与桌面 build 同时生成 `agent-worker.cjs`，位于服务端 bundle 或桌面 main.cjs 同级。路径由应用固定，客户端不能提供。桌面 fork 使用 ELECTRON_RUN_AS_NODE；真实 Electron 安装包运行另行验收，Node 子进程测试不代替实机。

## API

统一前缀 `/api/agent`。Web 复用 CloudServer Cookie 会话，scope=`user:<id>`；桌面复用本地服务 Bearer token，scope=`main`（与现有桌面存储一致，不冒充独立本地账户隔离）。请求提供 `X-Mydoc-Account` 可拒绝旧账户页面；POST 还必须携带当前 `X-Mydoc-Generation`。跨站请求拒绝，所有请求体有大小限制。

| 方法与路径 | 行为 |
| --- | --- |
| GET /status | 是否配置模型、是否管理员托管、能否用户配置、系统加密是否可用、公共模型字段（hasKey/keyAvailable）、外发说明；不返回密钥或 secretRef |
| POST /model | 仅桌面：`{endpoint,model,maxOutputTokens,contextWindow,supportsTools,includeUsage,apiKey?|keepKey?}` 保存，`{clear:true}` 清除 |
| GET /sessions | 当前账户个人会话，最多 100 条 |
| POST /sessions | `{title}`，创建个人会话，返回 id |
| POST /runs | `{session,requestKey,prompt}`，先持久入队再调度，返回 202；同 key 同内容复用 Run |
| GET /runs?id=… | 本账户运行状态及 `unconfirmed`（结果未确认的调用 id/name，重试前须核对领域回执）；其他账户返回不可用 |
| POST /cancel | `{id}`，取消本账户活动/排队 Run |
| POST /sessions/delete | `{id}`，删除会话及其 Run/事件/消息/调用/变更集；有活动 Run 时 409，排队中的一并删除 |
| GET /events?session=…&after=… | 按会话 cursor 递增返回最多 200 条，客户端推进 cursor；当前是 JSON 分页而非 SSE |

输入不能指定 user/scope、worker 路径、供应商地址或工具。运行中每次模型调用/输出都检查原 Cookie 会话，并定时检查注销/过期；失效取消活动 Run。已发往模型的内容无法撤回。不会把原 Cookie 或 API key 传给 worker。

## 保留与幂等

每个会话保留持有最新 cursor 的 50 个终态 Run，更早的 Run 连同事件、消息、调用、批准一并删除；会话 cursor 不回退。终态 Run 已并入 assistant 消息的文本 delta 会删除，崩溃前未保存的尾部文本保留。requestKey 去重只覆盖仍保留的 Run：被清理的旧 key 重试会产生新 Run。模型错误只在首个事件前对限流/5xx/网络/首包超时最多重试 2 次，不切换供应商。

## 生命周期与恢复

AgentService 在数据库已完成公共/协作升级后、服务监听前创建；仅这一启动 owner 恢复活动 Run，不在每个请求重新恢复。待运行任务保留，需该账户重新认证请求触发调度；不凭保存的 Cookie 自动运行。运行中断不盲重放写操作。

关闭先停止准入、取消并等待子进程，再关闭办公服务与数据库。事件事务提交后才回复 worker，跨账户事件不可读；恢复数据库会通过既有 generation 拒绝旧写入。当前不宣称支持跨服务多实例共享运行 owner。

## 验证范围

真实 HTTP 测试复用 CloudServer 注册/登录 Cookie，覆盖账户隔离、跨站/错账户/旧代际拒绝、本地合成模型流、实际子进程执行、持久事件、请求去重及注销失效。桌面服务测试验证已有 Bearer token 与 main scope，无 UI 或真实 Electron 启动。


## Agent 界面（2026-11-02）

在工作台点击右下角“Agent”悬浮按钮（可访问名称“Agent 助手”），或通过命令面板“打开 Agent 助手”进入悬浮窗。桌面可拖动标题调整位置；模型配置统一位于“设置 → Agent”，也可通过聊天标题齿轮跳转。可选择历史/新建对话、发送问题、停止执行、删除空闲会话；关闭面板保留后台运行，重新选择会话可读取持久记录。当前每个 Run 仍为独立提问，不会自动将历史或打开的文档发送到模型。

Web 使用带账户头的 fetch SSE `/subscribe?session=…&after=…`，断开自动重连，401/409 提示重新登录或载入。桌面使用受限 IPC 的事件分页更新；流式 IPC 尚未实现。桌面模型设置可保存接口/名称/密钥，服务器模型继续由管理员配置。未配置或密钥不可用时禁用发送。

界面适用于独立 Web 服务、正常 Vite 开发入口与桌面；Web 使用现有在线 Cookie 身份，需先登录。测试存储模式明确禁用 Agent 服务，浏览器 UI 测试通过合成 API 验证。资料来源、选区、文档工具及全文历史上下文仍未开放。


## Web 本地账号与服务器账号隔离（2026-11-02）

个人本地空间使用 `/api/local-agent`，从项目 SQLite 的本地会话校验身份；团队工作空间使用 `/api/agent`，只认服务器 Cookie。相同邮箱也不是同一身份，没有互相回退。聊天和模型配置分别位于 `local:<id>` 与服务器 `user:<id>` 命名空间，旧服务器记录不自动迁移为本地记录。

本地登录签发独立 HttpOnly/SameSite=Strict 的 `mydoc_local_session`，与 `mydoc_session` 并存。登录/退出任一身份不会清除另一种 Cookie；切换本地账号会撤销旧浏览器本地会话，旧订阅失效。Web 只接受独立 Cookie，不向脚本返回会话令牌、不写共享项目偏好，也不接受旧令牌头；启动时在事务中撤销并删除旧偏好里暴露的令牌，因此旧共享登录需要重新登录。密码校验用途的 verifyOnly 不签发或注销会话。桌面受主进程保护的本地账号接口保持兼容。

开发入口在 ProjectDatabase 上装配本地 Agent，独立服务在其现有数据库上装配；不新增数据库打开路径、表或版本。本地 Agent 只接受本机回环连接/Host，避免把本机模型配置权限公开到共享部署。桌面保留已有受信任主进程凭证，不改成本地账号强制登录。

Web 本地模型可按本地账号单独保存；配置写 `preferences` 的 `agent-config:local:<id>`，密钥以 AES-256-GCM 保存在数据库旁 `local-agent-secrets`，目录0700/文件0600，主密钥与密文均不写数据库或整库备份。此为本机服务端文件加密，不冒充 Electron safeStorage；换目录恢复须重新配置密钥。目录需随本机服务数据妥善保管，不要提交到 Git。

客户端按当前空间选择入口，状态接口提供各自的数据库 generation，避免将项目库代际误用于团队库。跨服务器/桌面团队 Agent 尚未接入时明确提示，不回退到本机身份或模型。本次没有开放团队文档读写工具，也没有改变独立托管模式自身的首页登录策略。


### API 类型

“设置 → Agent → 模型连接”可选择 OpenAI Chat Completions、OpenAI Responses、Anthropic Messages。接口地址填写完整端点，分别通常以 `/v1/chat/completions`、`/v1/responses`、`/v1/messages` 结尾；网关路径以服务商文档为准，程序不自动改写已填写地址。

旧配置缺少 apiType 时继续使用 Chat Completions。apiType 保存在现有 agent.model JSON 中，无 SQL 结构变化。Responses 使用 Bearer 认证、input/max_output_tokens 和原生响应事件；Messages 使用 x-api-key、anthropic-version、独立 system 和内容块事件。两种适配均支持文本流、工具调用与工具结果、用量、取消、完整性检查及响应大小上限。工具只在终态完整且调用参数验证通过后交给宿主。

协议参考：[Responses 流式事件](https://developers.openai.com/api/reference/resources/responses/streaming-events)、[Anthropic Messages 流式事件](https://platform.claude.com/docs/en/build-with-claude/streaming)。已验证合成协议流、配置保存/恢复与 UI，尚未进行真实供应商或 Electron 实机联网验收。多连接管理见下节；自动模型清单获取和连接探测仍属后续工作。

### 多连接和默认模型

2026-11-03 起，本地账号与桌面的“设置 → Agent”支持多个模型服务连接。每个连接填写名称、启用状态、API 类型、完整接口地址、模型名称及其他模型（换行或逗号分隔）。模型共用连接的协议、地址与凭据；首次添加的启用连接成为默认，后续添加不会替换默认值。

在“默认模型”选择“连接 / 模型”并保存，新请求使用此模型。删除、停用默认连接或移除当前默认模型前，先更换默认值；有排队/运行任务时不能修改配置。非默认连接可二次确认删除，删除后清理其密钥。

存储为现有 preferences 中独立账号 scope 的 `agent.connections`（JSON version=1）；旧 `agent.model` 兼容读取，首次配置修改时在事务中升级并维护旧默认快照，不涉及数据库结构迁移。恢复备份携带配置及 secretRef，不携带秘密；缺失密钥显示不可用。团队环境连接仍由管理员管理，普通成员只读。
