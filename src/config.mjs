const list = value => (value ?? '').split(',').map(part => part.trim()).filter(Boolean);
const maxGraphRequests = Number(process.env.M365_MCP_MAX_GRAPH_REQUESTS ?? 12);

export const config = Object.freeze({
  tenantId: process.env.M365_MCP_TENANT_ID ?? '',
  clientId: process.env.M365_MCP_ENTRA_CLIENT_ID ?? '',
  allowedUsername: (process.env.M365_MCP_ALLOWED_USERNAME ?? '').toLowerCase(),
  mcpClientId: process.env.M365_MCP_OAUTH_CLIENT_ID ?? 'm365-copilot-local',
  redirects: list(process.env.M365_MCP_REDIRECT_URIS),
  timeZone: process.env.M365_MCP_TIME_ZONE ?? '',
  maxGraphRequests,
  expiresAt: process.env.M365_MCP_EXPIRES_AT ?? '1970-01-01T00:00:00Z',
  graphScopes: [
    'Sites.Read.All', 'Mail.Read', 'People.Read.All',
    'OnlineMeetingTranscript.Read.All', 'Chat.Read',
    'ChannelMessage.Read.All', 'ExternalItem.Read.All',
  ],
});

export const graphBase = 'https://graph.microsoft.com';

export function isAllowedMcpRedirect(uri) {
  try {
    const url = new URL(uri);
    if (url.username || url.password || url.hash) return false;
    return url.protocol === 'https:' || (url.protocol === 'http:' && ['127.0.0.1','localhost','[::1]'].includes(url.hostname));
  } catch { return false; }
}

export function assertRunnableConfig() {
  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!guid.test(config.tenantId) || !guid.test(config.clientId)) throw new Error('CONFIG_MICROSOFT_ID_REQUIRED');
  if (!/^[^\s@]+@[^\s@]+$/.test(config.allowedUsername)) throw new Error('CONFIG_ALLOWED_USER_REQUIRED');
  if (!config.redirects.length || config.redirects.some(uri => !isAllowedMcpRedirect(uri))) throw new Error('CONFIG_MCP_REDIRECT_REQUIRED');
  try { if (!config.timeZone) throw new Error(); new Intl.DateTimeFormat('en',{timeZone:config.timeZone}); }
  catch { throw new Error('CONFIG_TIME_ZONE_REQUIRED'); }
  if (!Number.isInteger(maxGraphRequests) || maxGraphRequests < 1 || maxGraphRequests > 32) throw new Error('CONFIG_REQUEST_CAP_INVALID');
  if (!Number.isFinite(Date.parse(config.expiresAt)) || Date.now() >= Date.parse(config.expiresAt)) throw new Error('CONFIG_EXPIRY_REQUIRED');
}
