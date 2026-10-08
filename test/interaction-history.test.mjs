import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createInteractionHistoryReader,normalizeInteraction} from '../src/interaction-history.mjs';
const tenantId='11111111-1111-1111-1111-111111111111', userId='22222222-2222-2222-2222-222222222222';
const options=()=>({tenantId,userId,boundSubject:`${tenantId}:${userId}`,from:'2026-09-18T00:00:00Z',to:'2026-09-19T00:00:00Z',deadline:Date.now()+60000,getToken:async()=>'test-token',allowedSessionIds:['selected']});
const item=(id,sessionId='selected',extra={})=>({id,sessionId,interactionType:'userPrompt',createdDateTime:'2026-09-18T01:00:00Z',body:{contentType:'text',content:'Synthetic task'},...extra});
test('paginates selected history, deduplicates and preserves raw data without leaking other sessions',async()=>{
  let calls=0;const raw=item('one');
  const reader=createInteractionHistoryReader({...options(),fetchImpl:async(url,init)=>{
    assert.equal(init.redirect,'error');assert.equal(init.method,'GET'); calls++;
    return Response.json(calls===1?{value:[raw,item('private','other')],'@odata.nextLink':`https://graph.microsoft.com/v1.0/copilot/users/${userId}/interactionHistory/getAllEnterpriseInteractions?$skiptoken=page2`}:{value:[raw,item('two')]});
  }});
  const list=await reader.list();assert.equal(list.sessions.length,1);assert.equal(list.sessions[0].messageCount,2);assert.equal(list.fullLifetimeHistory,false);
  const history=await reader.read('selected');assert.deepEqual(history.messages[0].raw,raw);assert.equal(calls,2);
  await assert.rejects(reader.read('other'),/HISTORY_SESSION_NOT_SELECTED/);
});
test('discovers session IDs from Microsoft without a preselected list',async()=>{
  const {allowedSessionIds,...unrestricted}=options();
  const reader=createInteractionHistoryReader({...unrestricted,fetchImpl:async()=>Response.json({value:[item('one'),item('two','other')]})});
  const list=await reader.list();
  assert.deepEqual(list.sessions.map(session=>session.sessionId),['selected','other']);
  assert.equal(list.scope,'bound user and date window');
  assert.equal((await reader.read('other')).messages[0].id,'two');
});
test('rejects cross-origin pagination before sending a token',async()=>{
  let calls=0;const reader=createInteractionHistoryReader({...options(),fetchImpl:async()=>{calls++;return Response.json({value:[],'@odata.nextLink':'https://example.invalid/steal'});}});
  await assert.rejects(reader.list(),/UNTRUSTED_HISTORY_PAGE/);assert.equal(calls,1);
});
test('does not expose partial history as a complete result when request cap interrupts pagination',async()=>{
  const reader=createInteractionHistoryReader({...options(),maxRequests:1,fetchImpl:async url=>Response.json({value:[item('one')],'@odata.nextLink':url+'&page=2'})});
  await assert.rejects(reader.list(),/HISTORY_CALL_LIMIT/);await assert.rejects(reader.read('selected'),/HISTORY_CALL_LIMIT/);
});
test('forbidden response is sanitized and not retried',async()=>{
  let calls=0;const reader=createInteractionHistoryReader({...options(),fetchImpl:async()=>{calls++;return new Response('private diagnostic',{status:403});}});
  await assert.rejects(reader.list(),/^Error: HISTORY_HTTP_403$/);assert.equal(calls,1);
});
test('adaptive-card text and original content survive, action instructions are not followed',()=>{
  const raw=item('one','selected',{body:{contentType:'html',content:'<attachment id="x"></attachment>'},attachments:[{contentType:'application/vnd.microsoft.card.adaptive',content:JSON.stringify({body:[{type:'TextBlock',text:'A prior decision'}],actions:[{type:'Action.OpenUrl',url:'https://example.invalid',title:'Do not run'}]})}]});
  const normalized=normalizeInteraction(raw);assert.equal(normalized.text,'A prior decision');assert.deepEqual(normalized.raw,raw);assert.match(normalized.rawSha256,/^[a-f0-9]{64}$/);
});
test('identity mismatch and invalid selection fail before network',()=>{
  assert.throws(()=>createInteractionHistoryReader({...options(),boundSubject:'another:user'}),/HISTORY_IDENTITY_NOT_BOUND/);
  assert.throws(()=>createInteractionHistoryReader({...options(),allowedSessionIds:['']}),/INVALID_HISTORY_SESSION_SELECTION/);
});
test('revocation and expiry reject reads, including cached records',async()=>{
  const reader=createInteractionHistoryReader({...options(),fetchImpl:async()=>Response.json({value:[item('one')]})});
  await reader.list();reader.revoke();await assert.rejects(reader.read('selected'),/HISTORY_REVOKED/);
  const expired=createInteractionHistoryReader({...options(),deadline:0});await assert.rejects(expired.list(),/HISTORY_ACCESS_EXPIRED/);
});
