import { test } from 'node:test';
import assert from 'node:assert/strict';
import { generateKeyPairSync } from 'node:crypto';
import { createHistoryReaderFactory } from '../src/history-auth.mjs';

test('history stays absent without explicit opt-in', () => {
  assert.equal(createHistoryReaderFactory({ env: {} }), undefined);
});
test('invalid history configuration fails before a Microsoft sign-in', () => {
  assert.throws(()=>createHistoryReaderFactory({env:{
    M365_MCP_HISTORY_ENABLED:'true',
    M365_MCP_HISTORY_PRIVATE_KEY_PATH:'ignored.key',
    M365_MCP_HISTORY_CERT_THUMBPRINT_SHA256:'a'.repeat(64),
    M365_MCP_HISTORY_FROM:'2026-09-18T00:00:00Z',
    M365_MCP_HISTORY_TO:'2026-09-17T00:00:00Z',
  }}),/CONFIG_HISTORY_WINDOW_INVALID/);
});

test('the same MCP host binds app-only history to its signed-in user', async () => {
  let clientConfig, readerOptions;
  const seen=[];
  const privateKey=generateKeyPairSync('rsa',{modulusLength:2048}).privateKey.export({format:'pem',type:'pkcs8'});
  class FakeClient {
    constructor(value) { clientConfig = value; }
    async acquireTokenByClientCredential({ scopes }) {
      assert.deepEqual(scopes, ['https://graph.microsoft.com/.default']);
      return { accessToken: 'test-app-token' };
    }
  }
  const factory = createHistoryReaderFactory({
    env: {
      M365_MCP_HISTORY_ENABLED: 'true',
      M365_MCP_HISTORY_PRIVATE_KEY_PATH: 'ignored.key',
      M365_MCP_HISTORY_CERT_THUMBPRINT_SHA256: 'a'.repeat(64),
      M365_MCP_HISTORY_FROM: '2026-01-01T00:00:00Z',
      M365_MCP_HISTORY_TO: '2026-10-01T00:00:00Z',
      M365_MCP_HISTORY_MAX_REQUESTS: '10',
    },
    readKey: () => privateKey,
    ConfidentialClient: FakeClient,
    makeReader: value => { readerOptions = value; seen.push(value); return { status: () => ({ configured: true }) }; },
  });
  const reader = factory({ subject: '00000000-0000-4000-8000-000000000001:00000000-0000-4000-8000-000000000003' });
  assert.deepEqual(reader.status(), { configured: true });
  assert.equal(readerOptions.userId, '00000000-0000-4000-8000-000000000003');
  assert.equal(readerOptions.allowedSessionIds, undefined);
  assert.equal(await readerOptions.getToken(), 'test-app-token');
  assert.equal(clientConfig.auth.clientId, '00000000-0000-4000-8000-000000000002');
  factory({ subject: '00000000-0000-4000-8000-000000000001:00000000-0000-4000-8000-000000000003' });
  assert.equal(seen[0].budget,seen[1].budget);
});
