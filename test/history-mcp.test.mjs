import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { buildMcpServer } from '../src/mcp.mjs';

test('one MCP connection exposes optional history and forwards its bounded search options',async()=>{
  let received;
  const connection={
    status:()=>({history:{configured:true}}),
    importContext:()=>({}),
    listContexts:()=>[],
    ask:async()=>({}),
    history:{
      list:async filters=>{received=filters;return {sessions:[],paginationComplete:true};},
      read:async sessionId=>({sessionId,messages:sessionId==='oversized'?[{id:'large',text:'x'.repeat(2*1024*1024),raw:{id:'large'}}]:[{id:'one',text:'first',raw:{id:'one'}},{id:'two',text:'second',raw:{id:'two'}}]}),
    },
  };
  const server=buildMcpServer(connection);
  const client=new Client({name:'history-test',version:'1.0.0'});
  const [clientTransport,serverTransport]=InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  try {
    const tools=(await client.listTools()).tools.map(tool=>tool.name);
    assert.equal(tools.length,7);
    assert.ok(tools.includes('microsoft_list_copilot_history'));
    assert.ok(tools.includes('microsoft_read_copilot_history'));
    const result=await client.callTool({name:'microsoft_list_copilot_history',arguments:{
      from:'2026-09-18T00:00:00Z',to:'2026-09-25T00:00:00Z',query:'project',
    }});
    assert.equal(result.isError,undefined);
    assert.deepEqual(received,{from:'2026-09-18T00:00:00Z',to:'2026-09-25T00:00:00Z',query:'project'});
    const first=await client.callTool({name:'microsoft_read_copilot_history',arguments:{sessionId:'selected',limit:1}});
    const firstPage=JSON.parse(first.content[0].text);
    assert.equal(firstPage.totalMessages,2);
    assert.equal(firstPage.nextOffset,1);
    assert.equal(firstPage.rawIncluded,false);
    assert.equal(firstPage.messages[0].raw,undefined);
    const second=await client.callTool({name:'microsoft_read_copilot_history',arguments:{sessionId:'selected',offset:1,includeRaw:true}});
    const secondPage=JSON.parse(second.content[0].text);
    assert.deepEqual(secondPage.messages[0].raw,{id:'two'});
    assert.equal(secondPage.nextOffset,null);
    const oversized=await client.callTool({name:'microsoft_read_copilot_history',arguments:{sessionId:'oversized'}});
    assert.equal(oversized.isError,true);
    assert.equal(JSON.parse(oversized.content[0].text).error,'HISTORY_RESULT_TOO_LARGE');
  } finally { await client.close(); await server.close(); }
});
