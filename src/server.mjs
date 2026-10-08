import express from 'express';
import { createMcpExpressApp } from '@modelcontextprotocol/sdk/server/express.js';
import { mcpAuthRouter } from '@modelcontextprotocol/sdk/server/auth/router.js';
import { requireBearerAuth } from '@modelcontextprotocol/sdk/server/auth/middleware/bearerAuth.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { createOAuthProvider, page } from './oauth.mjs';
import { buildMcpServer } from './mcp.mjs';
import { pathToFileURL } from 'node:url';
import { config, assertRunnableConfig } from './config.mjs';
import { createHistoryReaderFactory } from './history-auth.mjs';

const maxTimerDelay = 2 ** 31 - 1;
export function scheduleShutdown(deadline, stop, {now = Date.now, schedule = setTimeout} = {}) {
  function check() {
    const remaining = deadline - now();
    if (remaining <= 0) { void stop(); return; }
    schedule(check, Math.min(remaining, maxTimerDelay)).unref?.();
  }
  check();
}

export function createApp(origin,options={}) {
  const parsed=new URL(origin);
  if (parsed.origin!==origin || (parsed.protocol!=='https:' && parsed.hostname!=='127.0.0.1')) throw new Error('INVALID_PUBLIC_ORIGIN');
  const app=createMcpExpressApp({host:'127.0.0.1',allowedHosts:[parsed.hostname,'127.0.0.1','localhost']});
  app.disable('x-powered-by');
  app.use((req,res,next)=>{
    // Do not trust forwarded client IPs for this loopback-only server.
    delete req.headers['x-forwarded-for'];
    res.on('finish',()=>options.onAudit?.({event:'http',method:req.method,path:req.path,status:res.statusCode}));
    // same-origin suppresses cross-site referrers while retaining Origin on same-site POST forms.
    res.set({'Cache-Control':'no-store','Referrer-Policy':'same-origin','X-Content-Type-Options':'nosniff','Content-Security-Policy':"default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; frame-ancestors 'none'; base-uri 'none'"});
    if (req.headers.origin && req.headers.origin!==origin && req.path==='/mcp') return res.status(403).end();
    next();
  });
  const provider=createOAuthProvider(origin,options);
  app.use(mcpAuthRouter({provider,issuerUrl:parsed,resourceServerUrl:new URL(`${origin}/mcp`),scopesSupported:['copilot.chat'],resourceName:'Microsoft 365 Copilot MCP'}));
  provider.installUI(app,express.urlencoded({extended:false,limit:'16kb'}));
  app.get('/',(req,res)=>res.type('html').send(page('<p>从你的 MCP 客户端发起连接，并以自己的 Microsoft 365 组织账号登录。</p><p>MCP endpoint: /mcp</p>')));
  app.get('/health',(req,res)=>res.json({status:'ok',authenticatedBusinessCallsOnly:true,paidFallbackEnabled:false}));
  const auth=requireBearerAuth({verifier:provider,requiredScopes:['copilot.chat'],resourceMetadataUrl:`${origin}/.well-known/oauth-protected-resource/mcp`});
  app.post('/mcp',auth,async(req,res)=>{
    const server=buildMcpServer(provider.getConnection(req.auth.extra.connectionId),options.onAudit);
    const transport=new StreamableHTTPServerTransport({sessionIdGenerator:undefined,enableJsonResponse:true});
    res.on('close',()=>{void transport.close();void server.close();});
    try { await server.connect(transport); await transport.handleRequest(req,res,req.body); }
    catch { if (!res.headersSent) res.status(500).json({error:'MCP_REQUEST_FAILED'}); }
  });
  app.all('/mcp',auth,(req,res)=>res.status(405).end());
  app.use((error,req,res,next)=>{if(!res.headersSent)res.status(400).json({error:'INVALID_REQUEST'});});
  return {app,provider};
}
if (process.argv[1] && import.meta.url===pathToFileURL(process.argv[1]).href) {
  assertRunnableConfig();
  const port=Number(process.env.M365_MCP_PORT??8787);
  const origin=process.env.M365_MCP_ORIGIN??`http://127.0.0.1:${port}`;
  const audit=event=>console.log(JSON.stringify({time:new Date().toISOString(),...event}));
  const connectionOptions={onAudit:audit,historyReaderFactory:createHistoryReaderFactory()};
  const {app,provider}=createApp(origin,{onAudit:audit,connectionOptions});
  const listener=app.listen(port,'127.0.0.1',()=>console.log(JSON.stringify({event:'listening',origin,mcp:`${origin}/mcp`})));
  async function stop(){listener.close();await provider.close();process.exit(0);}
  process.on('SIGINT',stop);process.on('SIGTERM',stop);
  scheduleShutdown(Date.parse(config.expiresAt),stop);
}
