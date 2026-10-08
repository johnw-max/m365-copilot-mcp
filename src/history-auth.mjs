import { readFileSync } from 'node:fs';
import { createPrivateKey } from 'node:crypto';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { config, graphBase } from './config.mjs';
import { createInteractionHistoryReader } from './interaction-history.mjs';

/** Enable history tools in this same MCP server only when the host explicitly configures app access. */
export function createHistoryReaderFactory({
  env = process.env,
  readKey = readFileSync,
  ConfidentialClient = ConfidentialClientApplication,
  makeReader = createInteractionHistoryReader,
} = {}) {
  if (env.M365_MCP_HISTORY_ENABLED !== 'true') return undefined;
  const keyPath = env.M365_MCP_HISTORY_PRIVATE_KEY_PATH;
  const thumbprintSha256 = (env.M365_MCP_HISTORY_CERT_THUMBPRINT_SHA256 ?? '').replace(/[:\s]/g, '');
  const from = env.M365_MCP_HISTORY_FROM;
  const to = env.M365_MCP_HISTORY_TO;
  const maxRequests = Number(env.M365_MCP_HISTORY_MAX_REQUESTS ?? 10);
  if (!keyPath || !/^[a-f0-9]{64}$/i.test(thumbprintSha256) || !from || !to || !Number.isInteger(maxRequests) || maxRequests < 1 || maxRequests > 32) {
    throw new Error('CONFIG_HISTORY_REQUIRED');
  }
  const start = Date.parse(from), end = Date.parse(to);
  if (!Number.isFinite(start) || !Number.isFinite(end) || start >= end || end - start > 5 * 366 * 86400000) {
    throw new Error('CONFIG_HISTORY_WINDOW_INVALID');
  }
  const privateKey = readKey(keyPath, 'utf8');
  if (typeof privateKey !== 'string' || !/-----BEGIN (?:RSA )?PRIVATE KEY-----/.test(privateKey)) throw new Error('CONFIG_HISTORY_PRIVATE_KEY_INVALID');
  try { createPrivateKey(privateKey); }
  catch { throw new Error('CONFIG_HISTORY_PRIVATE_KEY_INVALID'); }
  const client = new ConfidentialClient({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
      clientCertificate: { thumbprintSha256, privateKey },
    },
    system: { loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} } },
  });
  async function getToken() {
    const result = await client.acquireTokenByClientCredential({ scopes: [`${graphBase}/.default`] });
    if (!result?.accessToken) throw new Error('HISTORY_TOKEN_UNAVAILABLE');
    return result.accessToken;
  }
  const budget = {used:0,max:maxRequests};
  return identity => makeReader({
    tenantId: config.tenantId,
    userId: identity.subject.split(':')[1],
    boundSubject: identity.subject,
    from, to,
    deadline: Date.parse(config.expiresAt),
    maxRequests,
    budget,
    getToken,
  });
}
