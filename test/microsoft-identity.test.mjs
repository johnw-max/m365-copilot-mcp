import { test } from 'node:test';
import assert from 'node:assert/strict';
import { boundMicrosoftSubject } from '../src/microsoft.mjs';
import { config } from '../src/config.mjs';

test('binds history to the authenticated tenant user object ID', () => {
  const oid = '00000000-0000-4000-8000-000000000003';
  const auth = {
    accessToken:'synthetic',
    tenantId:config.tenantId,
    account:{username:config.allowedUsername,localAccountId:oid},
    idTokenClaims:{oid},
  };
  assert.equal(boundMicrosoftSubject(auth), `${config.tenantId}:${oid}`);
  assert.throws(()=>boundMicrosoftSubject({...auth,idTokenClaims:{oid:undefined},account:{...auth.account,localAccountId:'pairwise-sub'}}),/MICROSOFT_USER_ID_REQUIRED/);
  assert.throws(()=>boundMicrosoftSubject({...auth,tenantId:'11111111-1111-1111-1111-111111111111'}),/UNAPPROVED_MICROSOFT_ACCOUNT/);
  assert.throws(()=>boundMicrosoftSubject({...auth,idTokenClaims:{oid,tid:'11111111-1111-1111-1111-111111111111'}}),/UNAPPROVED_MICROSOFT_ACCOUNT/);
});
