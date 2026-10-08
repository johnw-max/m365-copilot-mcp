import http from 'node:http';
import { randomBytes, createHash } from 'node:crypto';

// A local-only authorization-code callback for the configured delegated app.
// No browser/session scraping, shared credentials, or tenant-policy changes.
export async function loopbackLogin(pca, scopes, onReady, {port=8788, timeoutMs=300000}={}) {
  const state=randomBytes(32).toString('base64url');
  const verifier=randomBytes(32).toString('base64url');
  const challenge=createHash('sha256').update(verifier).digest('base64url');
  let resolveCode, rejectCode, consumed=false;
  const codeResult=new Promise((resolve,reject)=>{resolveCode=resolve;rejectCode=reject;});
  // Attach a handler immediately, including while the authorization URL loads.
  void codeResult.catch(()=>{});
  const server=http.createServer((req,res)=>{
    res.setHeader('Cache-Control','no-store');
    res.setHeader('Referrer-Policy','no-referrer');
    res.setHeader('Content-Security-Policy',"default-src 'none'; frame-ancestors 'none'");
    const url=new URL(req.url,'http://localhost');
    if(req.method!=='GET'||url.pathname!=='/'||url.searchParams.get('state')!==state||consumed){res.writeHead(400);res.end('Invalid login callback.');return;}
    consumed=true;
    if(url.searchParams.has('error')){res.end('Microsoft sign-in did not complete. Return to your MCP client.');rejectCode(new Error('MICROSOFT_INTERACTIVE_DENIED'));return;}
    const code=url.searchParams.get('code');
    if(!code){res.writeHead(400);res.end('Missing authorization code.');rejectCode(new Error('MISSING_AUTHORIZATION_CODE'));return;}
    res.setHeader('Content-Type','text/plain; charset=utf-8');
    res.end('微软登录回调已收到。请返回连接页面检查登录结果。');
    resolveCode(code);
  });
  let timer;
  try {
    await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,'127.0.0.1',resolve);});
    const redirectUri=`http://localhost:${server.address().port}`;
    timer=setTimeout(()=>rejectCode(new Error('INTERACTIVE_LOGIN_TIMEOUT')),timeoutMs);
    const url=await pca.getAuthCodeUrl({scopes,redirectUri,state,responseMode:'query',codeChallenge:challenge,codeChallengeMethod:'S256'});
    onReady({authorizationUrl:url});
    const code=await codeResult;
    return await pca.acquireTokenByCode({scopes,redirectUri,code,codeVerifier:verifier});
  } finally {
    clearTimeout(timer);server.close();server.closeAllConnections();
  }
}
