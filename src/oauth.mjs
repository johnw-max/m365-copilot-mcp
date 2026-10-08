import { randomBytes, createHash } from 'node:crypto';
import { InvalidGrantError, InvalidTokenError, InvalidRequestError } from '@modelcontextprotocol/sdk/server/auth/errors.js';
import { config } from './config.mjs';
import { createConnection, startMicrosoftLogin } from './microsoft.mjs';

const opaque = () => randomBytes(32).toString('base64url');
const digest = value => createHash('sha256').update(value).digest('hex');
const escape = value => String(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function page(body) { return `<!doctype html><html lang="zh"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Microsoft 365 Copilot MCP</title><style>body{font:17px/1.6 system-ui;max-width:660px;margin:60px auto;padding:24px;color:#172b3a}button,a{font:inherit}button{padding:12px 18px;background:#245d70;color:white;border:0;border-radius:8px}code{font-size:24px}small{color:#556}</style><h1>Microsoft 365 Copilot MCP</h1>${body}</html>`; }

export function createOAuthProvider(origin, {login = startMicrosoftLogin, connectionOptions = {}} = {}) {
  const requestBudget={used:0,max:config.maxGraphRequests};
  const resource = `${origin}/mcp`;
  const clients = new Map([[config.mcpClientId, {client_id:config.mcpClientId, redirect_uris:config.redirects, token_endpoint_auth_method:'none', grant_types:['authorization_code','refresh_token'], response_types:['code'], client_name:'Microsoft 365 Copilot MCP'}]]);
  const flows = new Map(), codes = new Map(), tokens = new Map(), refreshes = new Map(), connections = new Map();
  const life = 8 * 60 * 60 * 1000;
  const expires = (ms=life) => Math.min(Date.now()+ms, Date.parse(config.expiresAt));
  function validateResource(value) { if (value && value.toString() !== resource) throw new InvalidRequestError('Unexpected resource'); }
  function validateScope(scopes) { if (scopes?.some(s=>s!=='copilot.chat')) throw new InvalidRequestError('Unexpected scope'); }
  function issue(clientId, connectionId) {
    const access_token=opaque(), refresh_token=opaque();
    tokens.set(digest(access_token), {clientId,connectionId,expires:expires(3600000)});
    refreshes.set(digest(refresh_token), {clientId,connectionId,expires:expires()});
    return {access_token,refresh_token,token_type:'Bearer',expires_in:Math.max(0,Math.floor((expires(3600000)-Date.now())/1000)),scope:'copilot.chat'};
  }
  function validCode(client, code) { const item=codes.get(digest(code)); if (!item || item.expires<Date.now() || item.clientId!==client.client_id) throw new InvalidGrantError('Invalid or expired code'); return item; }
  function sweep() {
    for (const map of [flows,codes,tokens,refreshes]) for (const [key,value] of map) if (value.expires<Date.now()) map.delete(key);
    for (const [key,value] of connections) if (value.expires<Date.now()) { void value.connection.disconnect(); connections.delete(key); }
  }
  const provider = {
    clientsStore: {
      async getClient(id) { return clients.get(id); },
      // Dynamic client registration is limited to explicitly configured host callbacks.
      async registerClient(metadata) {
        if (clients.size>=20 || !metadata.redirect_uris?.length || metadata.redirect_uris.some(uri=>!config.redirects.includes(uri))) throw new InvalidRequestError('Only approved host callbacks are allowed');
        const client={...metadata, client_id:opaque(), token_endpoint_auth_method:'none', client_secret:undefined, grant_types:['authorization_code','refresh_token'], response_types:['code']};
        clients.set(client.client_id,client); return client;
      },
    },
    async authorize(client, params, res) {
      sweep(); validateResource(params.resource); validateScope(params.scopes);
      if (!config.redirects.includes(params.redirectUri) || !/^[A-Za-z0-9_-]{43}$/.test(params.codeChallenge)) throw new InvalidRequestError('Invalid redirect or PKCE challenge');
      if (flows.size>=10 || Date.now()>=Date.parse(config.expiresAt)) throw new InvalidRequestError('Connector unavailable');
      const id=opaque(), cookie=opaque();
      flows.set(id,{...params, clientId:client.client_id, cookie:digest(cookie), csrf:opaque(), expires:expires(600000), stage:'ready'});
      res.cookie('copilot_mcp_flow',cookie,{httpOnly:true,secure:origin.startsWith('https:'),sameSite:'lax',maxAge:600000,path:'/connect'});
      res.redirect(`/connect?flow=${id}`);
    },
    async challengeForAuthorizationCode(client,code) { return validCode(client,code).codeChallenge; },
    async exchangeAuthorizationCode(client, code, verifier, redirectUri, target) {
      validateResource(target); const item=validCode(client,code);
      if (redirectUri!==item.redirectUri) throw new InvalidGrantError('Redirect mismatch');
      codes.delete(digest(code)); return issue(client.client_id,item.connectionId);
    },
    async exchangeRefreshToken(client,token,scopes,target) {
      validateResource(target); validateScope(scopes);
      const item=refreshes.get(digest(token));
      if (!item || item.clientId!==client.client_id || item.expires<Date.now() || !connections.has(item.connectionId)) throw new InvalidGrantError('Invalid refresh token');
      refreshes.delete(digest(token)); return issue(client.client_id,item.connectionId);
    },
    async verifyAccessToken(token) {
      sweep(); const item=tokens.get(digest(token));
      if (!item || !connections.has(item.connectionId)) throw new InvalidTokenError('Invalid or expired connector token');
      return {token, clientId:item.clientId, scopes:['copilot.chat'],expiresAt:Math.floor(item.expires/1000),resource:new URL(resource),extra:{connectionId:item.connectionId}};
    },
    async revokeToken(client,{token}) {
      const item=tokens.get(digest(token)) ?? refreshes.get(digest(token));
      if (!item || item.clientId!==client.client_id) return;
      for (const map of [tokens,refreshes]) for (const [key,value] of map) if (value.connectionId===item.connectionId) map.delete(key);
      await connections.get(item.connectionId)?.connection.disconnect(); connections.delete(item.connectionId);
    },
    getConnection(id) { const entry=connections.get(id); if (!entry || entry.expires<Date.now()) throw new InvalidTokenError('Connection missing'); return entry.connection; },
    installUI(app, formParser) {
      function flowFor(req) {
        const id=req.query.flow;
        if (typeof id!=='string') throw new Error('INVALID_FLOW');
        const flow=flows.get(id);
        const cookieName='copilot_mcp_flow=';
        const cookie=req.headers.cookie?.split(';').map(x=>x.trim()).find(x=>x.startsWith(cookieName))?.slice(cookieName.length);
        if (!flow || flow.expires<Date.now() || !cookie || digest(cookie)!==flow.cookie) throw new Error('INVALID_FLOW');
        return {id,flow};
      }
      app.get('/connect',(req,res)=> {
        try {
          const {id,flow}=flowFor(req);
          const form=(action,label)=>`<form method="post" action="/connect?flow=${escape(id)}"><input type="hidden" name="csrf" value="${escape(flow.csrf)}"><input type="hidden" name="action" value="${action}"><button>${label}</button></form>`;
          let body='<p>用自己的 Microsoft 365 组织账号连接这个 MCP 服务。Copilot Chat 使用微软要求的七项 Graph 委托读取权限。服务不会改走按量计费接口；请核对组织许可与使用条款。连接仅保存在本进程内，并会在配置的截止时间过期。</p>';
          if (flow.stage==='ready') body+=form('start','开始微软登录');
          else if (flow.stage==='pending') body+=flow.device?.authorizationUrl ? `<p>使用微软官方浏览器登录。回调仅在这台电脑接收，沿用已批准权限。</p><p><a href="${escape(flow.device.authorizationUrl)}" target="_blank" rel="noreferrer">打开微软登录</a></p><a href="/connect?flow=${escape(id)}">检查登录结果</a>` : flow.device ? `<p>在微软官方页面输入此代码：</p><p><code>${escape(flow.device.userCode)}</code></p><p><a href="${escape(flow.device.verificationUri)}" target="_blank" rel="noreferrer">打开微软登录</a></p><p>完成后返回此页。</p><a href="/connect?flow=${escape(id)}">检查登录结果</a>` : `<p>正在准备微软登录。</p><a href="/connect?flow=${escape(id)}">检查登录结果</a>`;
          else if (flow.stage==='authenticated') body+=`<p>微软账号：${escape(flow.identity.username.replace('@',' [at] '))}</p><p>你选定的背景与微软返回结果会传给发起请求的 MCP 客户端及其模型。此服务调用普通 Copilot Chat；不调用指定 Agent。Graph 请求上限 ${config.maxGraphRequests} 次。</p>${connectionOptions.historyReaderFactory?'<p>此服务已配置历史读取，将按设定日期范围查询当前登录账号的 Copilot 历史交互。</p>':''}${form('finish','连接到 MCP 客户端')}`;
          else body+='<p>登录未完成或账号不在允许列表中。请从 MCP 客户端重新发起连接。</p>';
          res.type('html').send(page(body));
        } catch { res.status(400).send('Invalid or expired connection flow. Restart from your MCP client.'); }
      });
      app.post('/connect',formParser, (req,res)=> {
        try {
          const {id,flow}=flowFor(req);
          if (req.headers.origin!==origin || req.body.csrf!==flow.csrf) throw new Error('INVALID_CSRF');
          if (req.body.action==='start' && flow.stage==='ready') {
            flow.stage='pending';
            login(device=>{flow.device=device;}).then(identity=> {
              if (flow.expires<Date.now()) { void identity.disconnect?.(); return; }
              flow.identity=identity; flow.stage='authenticated';
            }).catch(()=>{flow.stage='failed';});
          } else if (req.body.action==='finish' && flow.stage==='authenticated') {
            const connection=createConnection(flow.identity,{...connectionOptions,budget:requestBudget});
            connections.set(connection.id,{connection,expires:expires()});
            const code=opaque(); codes.set(digest(code),{clientId:flow.clientId,connectionId:connection.id,codeChallenge:flow.codeChallenge,redirectUri:flow.redirectUri,expires:expires(120000)});
            const target=new URL(flow.redirectUri); target.searchParams.set('code',code); if(flow.state) target.searchParams.set('state',flow.state);
            flows.delete(id); res.clearCookie('copilot_mcp_flow',{path:'/connect'});
            // A top-level link works in embedded browsers with strict cross-origin form redirect policy.
            res.type('html').send(page(`<p>个人连接已准备好。请在两分钟内返回 MCP 客户端完成授权。</p><p><a href="${escape(target.toString())}" rel="noreferrer">返回 MCP 客户端</a></p>`)); return;
          } else throw new Error('INVALID_STAGE');
          res.redirect(303,`/connect?flow=${id}`);
        } catch { res.status(400).send('Invalid connection action.'); }
      });
    },
    async close() { for (const {connection} of connections.values()) await connection.disconnect(); connections.clear(); tokens.clear(); refreshes.clear(); flows.clear(); codes.clear(); },
  };
  return provider;
}
