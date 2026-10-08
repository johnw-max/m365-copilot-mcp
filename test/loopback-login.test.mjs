import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { loopbackLogin } from '../src/loopback-login.mjs';

test('local OAuth rejects wrong state and method, redeems matching PKCE once, then closes',async()=>{
  let authRequest, exchanges=0;
  const pca={
    async getAuthCodeUrl(r){authRequest=r;return 'https://login.microsoftonline.com/approved';},
    async acquireTokenByCode(r){exchanges++;assert.equal(createHash('sha256').update(r.codeVerifier).digest('base64url'),authRequest.codeChallenge);assert.equal(r.code,'fixture-code');return {fixture:true};}
  };
  let ready;
  const started=new Promise(r=>ready=r);
  const result=loopbackLogin(pca,['approved.read'],ready,{port:0,timeoutMs:2000});
  await started;
  const base=authRequest.redirectUri.replace('localhost','127.0.0.1');
  assert.equal((await fetch(base+'?state=wrong&code=fixture-code')).status,400);
  assert.equal((await fetch(base+'?state='+authRequest.state+'&code=fixture-code',{method:'POST'})).status,400);
  assert.equal((await fetch(base+'?state='+authRequest.state+'&code=fixture-code')).status,200);
  assert.deepEqual(await result,{fixture:true});assert.equal(exchanges,1);
  await assert.rejects(fetch(base));
});

test('local OAuth timeout closes its listener without acquiring a token',async()=>{
  let callback;
  const pca={async getAuthCodeUrl(r){callback=r.redirectUri;return 'https://login.microsoftonline.com/approved';},async acquireTokenByCode(){throw new Error('must not redeem');}};
  await assert.rejects(loopbackLogin(pca,[],()=>{},{port:0,timeoutMs:20}),/INTERACTIVE_LOGIN_TIMEOUT/);
  await assert.rejects(fetch(callback.replace('localhost','127.0.0.1')));
});
