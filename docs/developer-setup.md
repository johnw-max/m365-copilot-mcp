# 开发者接入指南

本指南说明如何在自己的电脑上运行单账号 Microsoft 365 Copilot MCP 服务，并从兼容的 MCP 客户端连接。使用者需要 Microsoft 365 Copilot 附加许可；组织需评估 [Chat API 预览条款](https://learn.microsoft.com/en-us/legal/m365-copilot-apis/terms-of-use)。微软将当前 `/beta` 接口标注为不支持生产应用。

## 1. 在 Microsoft Entra 注册应用

在 Microsoft Entra 为自己的组织租户注册应用，记下**租户 ID** 和**应用（客户端）ID**。将 `http://localhost:8788` 配为公共客户端的本机登录回调；这是微软登录回到本机的地址，与 MCP 客户端的 OAuth 回调不同。本服务使用授权码 + PKCE，不需要客户端密钥。

给此应用配置 Microsoft Graph **委托权限**，七项必须齐全：`Sites.Read.All`、`Mail.Read`、`People.Read.All`、`OnlineMeetingTranscript.Read.All`、`Chat.Read`、`ChannelMessage.Read.All`、`ExternalItem.Read.All`。按组织策略完成管理员同意；用户以自己的身份登录。权限可涉及该用户有权访问的文件、邮件、聊天等数据。[微软的创建会话文档](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/copilotroot-post-conversations)列出了当前必需权限。

`AiEnterpriseInteraction.Read.All` 是**另一条历史导出路径**的应用权限，普通问答不需要它。不要为了跑通本指南顺手添加该权限。

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
| `M365_MCP_EXPIRES_AT` | 未来的 UTC 截止时间；到时服务停止，可在部署时设定所需期限 |
| `M365_MCP_MAX_GRAPH_REQUESTS` | 本进程 Graph 请求总上限，1–32；新会话首轮通常消耗两次请求 |
| `M365_MCP_PORT` / `M365_MCP_ORIGIN` | 默认 `8787` / `http://127.0.0.1:8787`，两者端口须一致 |

先运行 `npm test`，再运行 `node --env-file=.env src/server.mjs`。浏览器打开 `http://127.0.0.1:8787/health` 应返回 `status: ok`；这只表示本地服务可访问，不表示微软授权或 Copilot 问答已成功。

在支持 Streamable HTTP MCP 与 OAuth 的客户端中添加连接：地址 `http://127.0.0.1:8787/mcp`，客户端标识使用 `M365_MCP_OAUTH_CLIENT_ID`；回调 URL 必须列在 `M365_MCP_REDIRECT_URIS`。从客户端发起连接，在连接页登录允许的微软账号，再返回客户端完成授权。回环地址只供同一台电脑上的客户端访问。若客户端是云服务，应将 MCP 服务部署在受控 HTTPS 入口之后，并将该 HTTPS 地址设为 `M365_MCP_ORIGIN`；同时设计访问控制、凭据保护和撤销机制。

## 3. 验收一个真实来回

1. 调 `microsoft_connection_status`，核对登录邮箱、剩余请求次数以及历史读取状态。
2. 调 `microsoft_ask_copilot`，提出一个可核对的问题；保留返回的 `conversationHandle`。
3. 再调一次 `microsoft_ask_copilot`，带上该句柄追问。核对两次回答确实来自微软响应；本地离线测试不能替代此验收。
4. 如要使用已有背景，先由用户选定文字，调 `microsoft_import_context`；再把返回的 ID 放入 `microsoft_ask_copilot.contextIds`。这一步传的是用户选定快照，不会自动读取旧 Copilot 会话。

示例工具参数：

```json
{"question":"根据我有权访问的工作资料，概括本周项目风险。","contextIds":[]}
```

普通问答的远程会话句柄仅在当前个人连接内有效；进程重启需要重新连接。它不会恢复 Copilot 网页中的旧会话。历史原文要走单独的 [Interaction Export 路径](architecture.md)，由组织管理员授权并由部署方接入。指定 Copilot Studio Agent、Pages / Notebooks 结构化迁移及 Cowork 续跑均不由当前服务提供。

## 常见排查

| 现象 | 先检查 |
| --- | --- |
| 启动时报 `CONFIG_*` | 租户 / 应用 ID 格式、允许邮箱、客户端回调、UTC 截止时间和请求上限 |
| 客户端连不上 MCP | 服务是否可从客户端访问；客户端是否支持 Streamable HTTP 和 OAuth；端口、地址及回调是否精确一致 |
| 微软登录完成但连接失败 | 登录账号与 `M365_MCP_ALLOWED_USERNAME`、租户 ID 是否完全匹配；本机 `8788` 端口是否可用 |
| `GRAPH_HTTP_403` | 用户许可、七项委托权限与组织同意状态；以微软响应和组织管理员日志进一步判别 |
| `CALL_LIMIT_REACHED` | 当前进程请求预算已用尽；不要自动切换到付费接口 |

这个服务的 OAuth 客户端、令牌、会话、背景都只在进程内存中。它没有跨用户隔离所需的生产数据层，也没有长期运行所需的审计和撤销机制。请先阅读[安全与部署说明](../SECURITY.md)。
