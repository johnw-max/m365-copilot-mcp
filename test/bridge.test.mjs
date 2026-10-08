import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash, randomBytes } from 'node:crypto';
import { createConnection } from '../src/microsoft.mjs';
import { createApp } from '../src/server.mjs';
import { config, isAllowedMcpRedirect } from '../src/config.mjs';

const identity = () => ({username:'fixture@example.invalid',subject:'fixture:user',token:async()=>'fixture-not-real',disconnect:async()=>{}});
test('OAuth callbacks allow exact loopback HTTP or HTTPS but reject external HTTP and URL credentials',()=>{
  assert.equal(isAllowedMcpRedirect('http://127.0.0.1:8788/callback'),true);
  assert.equal(isAllowedMcpRedirect('http://localhost:8788/callback'),true);
  assert.equal(isAllowedMcpRedirect('https://mcp-client.example/callback'),true);
  assert.equal(isAllowedMcpRedirect('http://mcp-client.example/callback'),false);
  assert.equal(isAllowedMcpRedirect('https://name:pass@mcp-client.example/callback'),false);
  assert.equal(isAllowedMcpRedirect('http://localhost:8788/callback#fragment'),false);
});
test('context and conversation remain personal; inline context, same remote continuation and call cap',async()=>{
  const requests=[];
  const options={deadline:Date.now()+60000,maxRequests:3,fetchImpl:async(url,init)=>{requests.push({url,body:JSON.parse(init.body)});return Response.json(requests.length===1?{id:'remote-1'}:{messages:[{text:'synthetic answer'}]});}};
  const a=createConnection(identity(),options),b=createConnection(identity(),options);
  const context=a.importContext({title:'selected',text:'synthetic committed 120000'});
  await assert.rejects(b.ask({question:'test',contextIds:[context.id]}),/CONTEXT_NOT_OWNED/);
  const first=await a.ask({question:'calculate',contextIds:[context.id]});
  assert.match(requests[1].body.message.text,/synthetic committed 120000/);
  assert.equal(requests[1].body.locationHint.timeZone,config.timeZone);
  assert.equal(requests[1].body.contextualResources.webContext.isWebEnabled,false);
  await assert.rejects(b.ask({question:'steal',conversationHandle:first.conversationHandle}),/CONVERSATION_NOT_OWNED/);
  await a.ask({question:'continue',conversationHandle:first.conversationHandle});
  assert.equal(requests[1].url,requests[2].url);
  await assert.rejects(a.ask({question:'over limit'}),/CALL_LIMIT_REACHED/);
  assert.equal(requests.length,3);
});
test('provider denial is not retried; revoked and expired connections fail closed',async()=>{
  let requests=0;
  const a=createConnection(identity(),{deadline:Date.now()+60000,fetchImpl:async()=>{requests++;return new Response('sensitive provider detail',{status:403});}});
  await assert.rejects(a.ask({question:'test'}),/^Error: GRAPH_HTTP_403$/);assert.equal(requests,1);
  await a.disconnect();await assert.rejects(a.ask({question:'test'}),/CONNECTION_REVOKED/);
  const expired=createConnection(identity(),{deadline:0});await assert.rejects(expired.ask({question:'test'}),/CONNECTION_EXPIRED/);
});

test('real HTTP OAuth and MCP contract: anonymous rejection, CSRF, PKCE, redirect, replay, audience, refresh, revoke',async t=>{
  // Fake Microsoft login is injected only by this test; the executable server has no fixture mode.
  // Exercise the configured service window while valid.
  t.mock.timers.enable({apis:['Date'],now:new Date(Date.parse(config.expiresAt)-60000)});
  const origin='http://127.0.0.1:18787';
  const {app,provider}=createApp(origin,{login:async()=>identity()});
  const listener=app.listen(18787,'127.0.0.1');await new Promise(r=>listener.once('listening',r));
  t.after(async()=>{await provider.close();await new Promise(r=>listener.close(r));});
  const req=(path,options={})=>fetch(origin+path,{redirect:'manual',...options});
  const form=(body,headers={})=>({method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded',...headers},body:new URLSearchParams(body)});
  const verifier=randomBytes(32).toString('base64url'),challenge=createHash('sha256').update(verifier).digest('base64url');
  const metadata=await (await req('/.well-known/oauth-authorization-server')).json();assert.equal(metadata.issuer,origin+'/');
  assert.equal((await req('/mcp',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'})).status,401);
  assert.equal((await req('/authorize?'+new URLSearchParams({client_id:config.mcpClientId,redirect_uri:'https://evil.invalid',response_type:'code',code_challenge:challenge,code_challenge_method:'S256'}))).status,400);
  const auth=await req('/authorize?'+new URLSearchParams({client_id:config.mcpClientId,redirect_uri:config.redirects[0],response_type:'code',code_challenge:challenge,code_challenge_method:'S256',scope:'copilot.chat',state:'fixture-state',resource:origin+'/mcp'}));
  assert.equal(auth.status,302);
  const cookie=auth.headers.get('set-cookie').split(';')[0],location=auth.headers.get('location');
  assert.equal((await req(location)).status,400);
  const html=await (await req(location,{headers:{cookie}})).text();
  const csrf=html.match(/name="csrf" value="([^"]+)"/)[1];
  assert.equal((await req(location,form({csrf,action:'start'},{cookie,Origin:'https://evil.invalid'}))).status,400);
  assert.equal((await req(location,form({csrf,action:'start'},{cookie,Origin:origin}))).status,303);
  const finish=await req(location,form({csrf,action:'finish'},{cookie,Origin:origin}));
  assert.equal(finish.status,200);const callback=new URL((await finish.text()).match(/href="([^"]+)"/)[1].replaceAll('&amp;','&'));assert.equal(callback.searchParams.get('state'),'fixture-state');
  const code=callback.searchParams.get('code');
  const exchange={grant_type:'authorization_code',client_id:config.mcpClientId,code,code_verifier:verifier,redirect_uri:config.redirects[0],resource:origin+'/mcp'};
  assert.equal((await req('/token',form({...exchange,code_verifier:'wrong'}))).status,400);
  assert.equal((await req('/token',form({...exchange,resource:'https://evil.invalid'}))).status,400);
  assert.equal((await req('/token',form({...exchange,redirect_uri:config.redirects[1]}))).status,400);
  const tokenReply=await req('/token',form(exchange));assert.equal(tokenReply.status,200);const tokens=await tokenReply.json();
  assert.equal((await req('/token',form(exchange))).status,400);
  const headers={'Content-Type':'application/json',Accept:'application/json, text/event-stream',Authorization:`Bearer ${tokens.access_token}`};
  const init=await req('/mcp',{method:'POST',headers,body:JSON.stringify({jsonrpc:'2.0',id:1,method:'initialize',params:{protocolVersion:'2025-03-26',capabilities:{},clientInfo:{name:'fixture',version:'1'}}})});
  assert.equal(init.status,200);assert.equal((await init.json()).result.serverInfo.name,'Microsoft 365 Copilot MCP');
  const listed=await (await req('/mcp',{method:'POST',headers,body:JSON.stringify({jsonrpc:'2.0',id:2,method:'tools/list'})})).json();
  assert.equal(listed.result.tools.length,5);
  assert.equal(listed.result.tools.some(x=>x.name==='microsoft_ask_agent'),false);
  const called=await (await req('/mcp',{method:'POST',headers,body:JSON.stringify({jsonrpc:'2.0',id:3,method:'tools/call',params:{name:'microsoft_connection_status',arguments:{}}})})).json();
  assert.equal(JSON.parse(called.result.content[0].text).graphRequestsMax,config.maxGraphRequests);
  const refreshed=await req('/token',form({grant_type:'refresh_token',client_id:config.mcpClientId,refresh_token:tokens.refresh_token,resource:origin+'/mcp'}));assert.equal(refreshed.status,200);const next=await refreshed.json();
  assert.equal((await req('/token',form({grant_type:'refresh_token',client_id:config.mcpClientId,refresh_token:tokens.refresh_token}))).status,400);
  assert.equal((await req('/revoke',form({client_id:config.mcpClientId,token:next.refresh_token}))).status,200);
  assert.equal((await req('/mcp',{method:'POST',headers,body:'{}'})).status,401);
});
