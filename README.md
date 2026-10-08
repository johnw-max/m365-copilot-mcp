# Microsoft 365 Copilot MCP

这个独立 MCP 服务让兼容 Streamable HTTP MCP 与 OAuth 的客户端连接 Microsoft 365 Copilot。用户登录组织账号后，可以在客户端中向 Copilot 提问、追问；组织管理员如另外启用历史读取，用户还可以列出自己过去的 Copilot 会话并读取其中的提问和回答。两种能力都由**同一个 MCP 服务**提供，不依赖特定的 Agent 平台或客户端品牌。

历史读取默认关闭。启用时须在 Microsoft Entra 为该应用授予 `AiEnterpriseInteraction.Read.All` **应用权限**并由管理员同意，配置证书和查询日期范围；仅完成个人登录并不会取得历史记录。具体步骤见[开发者接入指南](docs/developer-setup.md#可选启用历史会话读取)。

```mermaid
flowchart LR
  U[用户] --> H[兼容的 MCP 客户端]
  H -->|OAuth| S[本 MCP 服务]
  S -->|用户委托权限| C[Copilot Chat API]
  C --> S
  A[Entra 管理员同意历史读取权限] --> S
  K[部署方配置证书] --> S
  S -->|启用后使用应用权限| I[Interaction Export API]
  I --> S --> H
```

## 能力与边界

| 能力 | 本仓库状态 | 关键边界 |
| --- | --- | --- |
| Copilot 问答与同一 API 会话内追问 | 已接入 MCP 服务 | Chat API 使用 `/beta`，微软标注不支持用于生产应用；用户需有 Microsoft 365 Copilot 附加许可 |
| 用户选定的文字背景 | 已接入 MCP 服务 | 只保存本进程内的快照，不自动抓取文件或旧聊天 |
| 旧 Copilot 会话中的提问和回答 | 可选 MCP 工具，配置后启用 | 需 Entra 管理员同意应用权限、配置证书与日期范围；可发现会话 ID、按文字查找，并按消息分页读取整理文字或微软原始记录 |
| 指定 Copilot Studio Agent、Pages / Notebooks 结构化导出、Cowork 续跑 | 未实现 | 普通 Chat API 不提供这些能力 |

微软目前要求 Chat API 使用七项 Graph 委托权限；对持有 Microsoft 365 Copilot 附加许可的用户，微软说明该 API 无额外费用。客户端模型、基础许可与部署仍可能产生费用。历史读取需要同一个服务额外取得 `AiEnterpriseInteraction.Read.All` 应用权限；它在微软侧是租户级授权，本服务仅查询当前允许登录的用户。详见[接口与边界](docs/architecture.md)。

默认 MCP 服务提供五个工具：`microsoft_connection_status` 查询连接状态，`microsoft_capability_status` 查看可用范围，`microsoft_ask_copilot` 提问或追问，`microsoft_import_context` 保存用户选定的文字背景，`microsoft_list_contexts` 列出已保存的背景。启用历史读取后，同一 MCP 服务增加 `microsoft_list_copilot_history` 和 `microsoft_read_copilot_history`。背景和会话都只存在本进程内存中。

## 启动

需要 Node.js 22+、组织租户中的 Entra 应用注册、用户的 Microsoft 365 Copilot 附加许可，以及组织核准的 Chat API 委托权限。使用前阅读 [Microsoft 预览条款](https://learn.microsoft.com/en-us/legal/m365-copilot-apis/terms-of-use)。

```powershell
npm ci
Copy-Item .env.example .env
# 编辑 .env：填入自己的租户、应用、允许账号、MCP 客户端回调及服务截止时间
node --env-file=.env src/server.mjs
```

Entra 应用的本机登录回调是 `http://localhost:8788`；MCP 客户端的 OAuth 回调是另一条地址，必须精确列入 `M365_MCP_REDIRECT_URIS`。客户端回调可以是 HTTPS，或仅限本机的 HTTP 回环地址。还须配置用户的 IANA 时区 `M365_MCP_TIME_ZONE`。服务默认监听 `127.0.0.1:8787`，MCP 地址为 `http://127.0.0.1:8787/mcp`。完整配置与首次验收见[开发者接入指南](docs/developer-setup.md)。远端客户端需要受控的 HTTPS 入口；不要将本地服务直接暴露在公网。

运行 `npm test` 可离线检查 OAuth / MCP 契约、会话隔离、调用上限、历史授权接线及分页与筛选；测试不调用微软。历史查询每次最多读取四页，范围过大时可在工具参数中缩小日期重试；全进程另有请求总上限。正式使用前仍需在自己的租户验收权限和实际返回内容。

## 代码与文档

- `src/server.mjs`、`src/oauth.mjs`：本地 MCP 端点、OAuth / PKCE、令牌与连接绑定。
- `src/microsoft.mjs`：Copilot Chat API 问答、追问和用户选定背景。
- `src/interaction-history.mjs`、`src/history-auth.mjs`：历史会话读取、证书令牌和同一服务内的可选启用。
- [`docs/architecture.md`](docs/architecture.md)：权限、数据流、成本与接口边界。
- [`docs/developer-setup.md`](docs/developer-setup.md)：从空白租户应用到首轮 MCP 问答的配置与验收。

本仓库采用 [MIT License](LICENSE)。微软、Copilot、Microsoft 365 及相关商标属于各自权利人。
