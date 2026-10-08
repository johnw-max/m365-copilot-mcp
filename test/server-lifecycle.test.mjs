import {test} from 'node:test';
import assert from 'node:assert/strict';
import {scheduleShutdown} from '../src/server.mjs';

test('a distant expiry is scheduled in bounded intervals and stops at the deadline',()=>{
  const max=2**31-1;
  let now=0,stops=0;
  const queued=[];
  const schedule=(run,delay)=>{queued.push({run,delay});return {unref(){}};};
  scheduleShutdown(max*2+1000,()=>{stops++;},{now:()=>now,schedule});
  const first=queued.shift();assert.equal(first.delay,max);
  now=max;first.run();
  const second=queued.shift();assert.equal(second.delay,max);
  now=max*2;second.run();
  const third=queued.shift();assert.equal(third.delay,1000);
  now=max*2+1000;third.run();
  assert.equal(stops,1);
  assert.equal(queued.length,0);
});
