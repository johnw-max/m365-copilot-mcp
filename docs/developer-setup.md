# 开发者接入指南

本指南说明如何在自己的电脑上运行单账号 Microsoft 365 Copilot MCP 服务，并从兼容的 MCP 客户端连接。默认可与 Copilot 对话；管理员完成历史读取授权并配置证书后，同一个 MCP 连接还会提供旧会话工具。使用者需要 Microsoft 365 Copilot 附加许可；组织需评估 [Chat API 预览条款](https://learn.microsoft.com/en-us/legal/m365-copilot-apis/terms-of-use)。微软将当前 Chat `/beta` 接口标注为不支持生产应用。

## 1. 在 Microsoft Entra 注册应用

在 Microsoft Entra 为自己的组织租户注册应用，记下**租户 ID** 和**应用（客户端）ID**。将 `http://localhost:8788` 配为公共客户端的本机登录回调；这是微软登录回到本机的地址，与 MCP 客户端的 OAuth 回调不同。本服务使用授权码 + PKCE，不需要客户端密钥。

给此应用配置 Microsoft Graph **委托权限**，七项必须齐全：`Sites.Read.All`、`Mail.Read`、`People.Read.All`、`OnlineMeetingTranscript.Read.All`、`Chat.Read`、`ChannelMessage.Read.All`、`ExternalItem.Read.All`。按组织策略完成管理员同意；用户以自己的身份登录。权限可涉及该用户有权访问的文件、邮件、聊天等数据。[微软的创建会话文档](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/copilotroot-post-conversations)列出了当前必需权限。

普通问答不需要 `AiEnterpriseInteraction.Read.All`。只有组织决定在这个 MCP 服务中启用历史会话读取时，才按下文配置该应用权限。

## 2. 配置服务与 MCP 客户端

需要 Node.js 22+。在仓库根目录运行：

```powershell
npm ci
Copy-Item .env.example .env
```

填写 `.env`：

| 字段 | 填写方式 |
| --- | --- |
| `M365_MCP_TENANT_ID` / `M365_MCP_ENTRA_CLIENT_ID` | 组织租户 ID / Entra 应用 ID |
| `M365_MCP_ALLOWED_USERNAME` | 唯一允许登录的用户邮箱；代码核对租户和账号 |
| `M365_MCP_OAUTH_CLIENT_ID` | MCP 客户端向本服务发起 OAuth 时使用的客户端标识 |
| `M365_MCP_REDIRECT_URIS` | MCP 客户端的**精确** OAuth 回调 URL，多个用逗号分隔；允许 HTTPS 或 `localhost` / `127.0.0.1` 的 HTTP 回环地址 |
| `M365_MCP_TIME_ZONE` | 用户所在的 IANA 时区，例如 `Asia/Singapore`；Copilot Chat API 要求在每次问答中提供时区 |
| `M365_MCP_EXPIRES_AT` | 未来的 UTC 截止时间；到时服务停止，可在部署时设定所需期限 |
| `M365_MCP_MAX_GRAPH_REQUESTS` | 本进程 Chat API Graph 请求总上限，1–32；新会话首轮通常消耗两次请求，重连不会重置 |
| `M365_MCP_PORT` / `M365_MCP_ORIGIN` | 默认 `8787` / `http://127.0.0.1:8787`，两者端口须一致 |

先运行 `npm test`，再运行 `node --env-file=.env src/server.mjs`。浏览器打开 `http://127.0.0.1:8787/health` 应返回 `status: ok`；这只表示本地服务可访问，不表示微软授权或 Copilot 问答已成功。

在支持 Streamable HTTP MCP 与 OAuth 的客户端中添加连接：地址 `http://127.0.0.1:8787/mcp`，客户端标识使用 `M365_MCP_OAUTH_CLIENT_ID`；回调 URL 必须列在 `M365_MCP_REDIRECT_URIS`。从客户端发起连接，在连接页登录允许的微软账号，再返回客户端完成授权。回环地址只供同一台电脑上的客户端访问。若客户端是云服务，应将 MCP 服务部署在受控 HTTPS 入口之后，并将该 HTTPS 地址设为 `M365_MCP_ORIGIN`；同时设计访问控制、凭据保护和撤销机制。

## 3. 验收一个真实来回

1. 调 `microsoft_connection_status`，核对登录邮箱、`serviceGraphRequestsUsed` / `serviceGraphRequestsMax` 及历史读取状态。`graphRequestsUsed` 只表示当前连接的消耗，重连后可归零；服务总消耗不会归零。
2. 调 `microsoft_ask_copilot`，提出一个可核对的问题；保留返回的 `conversationHandle`。
3. 再调一次 `microsoft_ask_copilot`，带上该句柄追问。核对两次回答确实来自微软响应；本地离线测试不能替代此验收。
4. 如要使用已有背景，先由用户选定文字，调 `microsoft_import_context`；再把返回的 ID 放入 `microsoft_ask_copilot.contextIds`。这一步传的是用户选定快照，不会自动读取旧 Copilot 会话。

示例工具参数：

```json
{"question":"根据我有权访问的工作资料，概括本周项目风险。","contextIds":[]}
```

普通问答的远程会话句柄仅在当前个人连接内有效；进程重启需要重新连接。它不会恢复 Copilot 网页中的旧会话。若启用下述历史工具，可以在同一 MCP 连接中读取旧会话的提问和回答。指定 Copilot Studio Agent、Pages / Notebooks 结构化迁移及 Cowork 续跑均不由当前服务提供。

## 可选：启用历史会话读取

历史读取使用的是 Microsoft Graph 的 Interaction Export API。它返回微软已记录的用户提问与 Copilot 回答，并带有 `sessionId`；本服务可先列出会话，再读取选中的会话，**无需预先填写会话 ID**。它不会恢复 Copilot 网页的会话界面，也不保证涵盖所有 Copilot 体验或附件正文。[微软接口文档](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions)说明了返回范围和许可要求。

1. 在 [Microsoft Entra 管理中心](https://entra.microsoft.com/)打开同一个应用注册，进入 **Entra ID → App registrations → 应用 → API permissions → Add a permission → Microsoft Graph → Application permissions**，添加 `AiEnterpriseInteraction.Read.All`。由有权授予该权限的管理员点击 **Grant admin consent**，在权限状态中确认已授权。这里是租户级应用授权，普通用户的 OAuth 登录不能代替它；服务只查询当前允许登录的用户。[微软的配置与管理员同意说明](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-configure-app-access-web-apis)
2. 在该应用的 **Certificates & secrets → Certificates** 上传公钥证书；把对应私钥作为 PEM 文件保存在运行服务的机器上，且不要提交到仓库。记录证书的 SHA-256 thumbprint。服务通过证书取得应用令牌，私钥不得交给 MCP 客户端。[微软的证书配置说明](https://learn.microsoft.com/en-us/entra/msal/javascript/node/certificate-credentials)
3. 在 `.env` 中增加以下配置，然后重启服务。查询范围由用户、起止时间和最大请求数共同限定；本项目不会自行在 Entra 中申请管理员同意。

```dotenv
M365_MCP_HISTORY_ENABLED=true
M365_MCP_HISTORY_PRIVATE_KEY_PATH=path/to/private.key
M365_MCP_HISTORY_CERT_THUMBPRINT_SHA256=your-64-character-sha256-thumbprint
M365_MCP_HISTORY_FROM=2026-01-01T00:00:00Z
M365_MCP_HISTORY_TO=2026-10-01T00:00:00Z
M365_MCP_HISTORY_MAX_REQUESTS=10
```

4. 从同一 MCP 客户端重新连接，先看 `microsoft_connection_status`，再调用 `microsoft_list_copilot_history` 列出日期范围内的会话，选一个返回的 `sessionId` 调用 `microsoft_read_copilot_history`。列表工具可带 `from`、`to` 缩小日期，也可带 `query` 在已读取的对话文字中查找项目词。下面的例子只查询配置范围内的一周：

```json
{"from":"2026-09-18T00:00:00Z","to":"2026-09-25T00:00:00Z","query":"项目名称"}
```

读取会话时默认每页返回 10 条消息，不含微软原始 JSON；按 `nextOffset` 继续读取，最多一次取 20 条。需要核对原始记录时可传 `includeRaw:true`，例如 `{"sessionId":"返回的会话 ID","offset":0,"limit":10,"includeRaw":true}`。原始记录可能含附件与上下文元数据；单次返回超过 2 MiB 会报 `HISTORY_RESULT_TOO_LARGE`，应缩小 `limit`，仍无法容纳单条记录时须通过服务端安全导出处理，不能把截断内容当作完整记录。

每次列表查询最多读取四页；如果返回 `HISTORY_QUERY_TOO_WIDE`，直接用更短的 `from` / `to` 重试，无需改 `.env` 或重新登录。`M365_MCP_HISTORY_MAX_REQUESTS` 是本服务进程的历史接口**总请求上限**，即使重新连接也不会重置；`HISTORY_CALL_LIMIT` 表示本次运行的额度已用完。读取会话前必须先成功列出包含它的日期范围。服务不会把未读完的日期窗口结果当作完整结果；跨出日期窗口的同一会话消息仍可能缺失，响应会标记 `sessionMayBePartial`。`query` 是在读回内容后本地查找，不是微软提供的全文搜索接口。私钥配置或管理员同意缺失时，历史调用不会成功。

## 常见排查

| 现象 | 先检查 |
| --- | --- |
| 启动时报 `CONFIG_*` | 租户 / 应用 ID 格式、允许邮箱、客户端回调、IANA 时区、UTC 截止时间和请求上限 |
| 客户端连不上 MCP | 服务是否可从客户端访问；客户端是否支持 Streamable HTTP 和 OAuth；端口、地址及回调是否精确一致 |
| 微软登录完成但连接失败 | 登录账号与 `M365_MCP_ALLOWED_USERNAME`、租户 ID 是否完全匹配；本机 `8788` 端口是否可用 |
| `GRAPH_HTTP_403` | 用户许可、七项委托权限与组织同意状态；以微软响应和组织管理员日志进一步判别 |
| `CALL_LIMIT_REACHED` | 当前进程请求预算已用尽；不要自动切换到付费接口 |
| `HISTORY_QUERY_TOO_WIDE` | 在历史列表工具中缩小 `from` / `to` 后重试 |
| `HISTORY_CALL_LIMIT` | 历史接口进程总请求上限已用完；由部署者检查额度并决定是否重新启动 |
| `HISTORY_RESULT_TOO_LARGE` | 缩小读取消息的 `limit`，或关闭 `includeRaw`；单条超限需另行安全导出 |
| `HISTORY_HTTP_403` | 核查目标用户许可、应用权限、管理员同意和证书；以微软租户日志判别 |

这个服务的 OAuth 客户端、令牌、会话、背景都只在进程内存中。它没有跨用户隔离所需的生产数据层，也没有长期运行所需的审计和撤销机制。请先阅读[安全与部署说明](../SECURITY.md)。
