import {describe,it,expect,vi} from 'vitest';
import {runAgent,canTransition} from '@mydoc/agent-core';
import {AgentHost} from '@mydoc/agent-host';
import {validateCall,validateLimits,DEFAULT_RUN_LIMITS,validateSession,validateStep} from '@mydoc/agent-contracts';
import type {AgentEvent,ProviderEvent,ProviderPort} from '@mydoc/agent-contracts';
function fixture(rounds:ProviderEvent[][]){
 const events:AgentEvent[]=[];let index=0;
 const execute=vi.fn(async()=>({status:'ok' as const,text:'found'}));
 const options={sessionId:'s',runId:'r',prompt:'Read',tools:[{name:'resource.read',effect:'read' as const,description:'Read resource'}],
  provider:{async *stream(){yield* rounds[index++]??[{type:'done'}];}} as ProviderPort,
  gateway:{execute},events:{async append(event:AgentEvent){events.push(event);}},clock:{now:()=>123}};
 return {events,execute,options};
}
const call={id:'c1',name:'resource.read',arguments:{resource:'d'}};
describe('agent framework',()=>{
 it('runs model/tool/model and orders accepted events',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}],[{type:'text',text:'Answer'},{type:'done'}]]);
  const result=await runAgent(f.options);
  expect(result.status).toBe('completed');expect(f.execute).toHaveBeenCalledTimes(1);
  expect(result.messages.map(m=>m.role)).toEqual(['user','assistant','tool','assistant']);
  expect(f.events.map(e=>e.seq)).toEqual(f.events.map((_,i)=>i+1));
  expect(f.events.every(e=>e.protocolVersion===1&&e.sessionId==='s'&&e.timestamp===123)).toBe(true);
 });
 it('never executes calls from an interrupted provider stream',async()=>{
  const f=fixture([[{type:'tool',call}]]);const result=await runAgent(f.options);
  expect(result.status).toBe('failed');expect(f.execute).not.toHaveBeenCalled();
  expect(result.messages.at(-1)).toMatchObject({role:'tool',callId:'c1',status:'skipped'});
 });
 it('denies undeclared tools',async()=>{
  const f=fixture([[{type:'tool',call:{...call,name:'shell'}},{type:'done'}],[{type:'done'}]]);
  const result=await runAgent(f.options);expect(f.execute).not.toHaveBeenCalled();
  expect(result.messages).toContainEqual(expect.objectContaining({role:'tool',status:'denied'}));
 });
 it('rejects duplicate call IDs before executing the batch',async()=>{
  const f=fixture([[{type:'tool',call},{type:'tool',call},{type:'done'}]]);
  const result=await runAgent(f.options);expect(result.status).toBe('failed');expect(f.execute).not.toHaveBeenCalled();
  expect(result.messages.filter(m=>m.role==='tool')).toHaveLength(1);
 });
 it('feeds tool errors back without losing call pairing',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}],[{type:'done'}]]);
  f.execute.mockRejectedValue(new Error('conflict'));
  const result=await runAgent(f.options);expect(result.status).toBe('completed');
  expect(result.messages).toContainEqual({role:'tool',callId:'c1',status:'error',text:'conflict'});
 });
 it('enforces round and output budgets',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}]]);
  expect((await runAgent({...f.options,limits:{maxRounds:1}})).status).toBe('budget_exhausted');
  const g=fixture([[{type:'text',text:'too long'},{type:'done'}]]);
  expect((await runAgent({...g.options,limits:{maxOutputChars:2}})).status).toBe('budget_exhausted');
 });
 it('bounds and explicitly truncates tool output',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}],[{type:'done'}]]);
  const result=await runAgent({...f.options,limits:{maxToolResultChars:2}});
  expect(result.messages).toContainEqual({role:'tool',callId:'c1',status:'ok',text:'fo\n[truncated; result=r:c1; offset=2]'});
 });
 it('times out even when a provider ignores abort',async()=>{
  const f=fixture([]);f.options.provider={async *stream(){await new Promise(()=>{});yield {type:'done'};}};
  expect((await runAgent({...f.options,limits:{timeoutMs:10}})).status).toBe('budget_exhausted');
 });
 it('cancels an active tool as unconfirmed and skips remaining calls',async()=>{
  const f=fixture([[{type:'tool',call},{type:'tool',call:{...call,id:'c2'}},{type:'done'}]]);
  const controller=new AbortController();f.execute.mockImplementation(async()=>{controller.abort();return new Promise(()=>{});});
  const result=await runAgent({...f.options,signal:controller.signal});
  expect(result.status).toBe('cancelled');expect(f.execute).toHaveBeenCalledTimes(1);
  expect(result.messages).toContainEqual(expect.objectContaining({callId:'c1',status:'error',text:expect.stringContaining('unconfirmed')}));
  expect(result.messages).toContainEqual(expect.objectContaining({callId:'c2',status:'skipped'}));
 });
 it('does not call provider when already cancelled',async()=>{
  const f=fixture([]),controller=new AbortController();controller.abort();
  expect((await runAgent({...f.options,signal:controller.signal})).status).toBe('cancelled');expect(f.execute).not.toHaveBeenCalled();
 });
 it('fences concurrent runs and stale cancellation in the host',async()=>{
  const f=fixture([]),host=new AgentHost();f.options.provider={async *stream(){await new Promise(()=>{});yield {type:'done'};}};
  const running=host.start(f.options);
  await expect(host.start({...f.options,runId:'r2'})).rejects.toThrow('already running');
  expect(host.cancel('s','old')).toBe(false);expect(host.cancel('s','r')).toBe(true);
  expect((await running).status).toBe('cancelled');
  const next=fixture([[{type:'done'}]]);expect((await host.start(next.options)).status).toBe('completed');
 });
 it('propagates persistence failure without executing tools',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}]]);
  f.options.events.append=async()=>{throw new Error('disk full');};
  await expect(runAgent(f.options)).rejects.toThrow('disk full');expect(f.execute).not.toHaveBeenCalled();
 });
 it('validates contract limits and keeps terminal states closed',()=>{
  expect(()=>validateCall({...call,arguments:undefined})).toThrow();
  expect(()=>validateLimits({...DEFAULT_RUN_LIMITS,maxRounds:0})).toThrow();
  expect(canTransition('completed','streaming')).toBe(false);expect(canTransition('streaming','executing_tools')).toBe(true);
 });
});

describe('agent scheduling and context',()=>{
 it('bounds read concurrency and places writes behind a barrier',async()=>{
  const calls=[1,2,3,4].map(i=>({...call,id:'c'+i,name:i===3?'document.write':'resource.read'}));
  const f=fixture([[...calls.map(call=>({type:'tool' as const,call})),{type:'done'}],[{type:'done'}]]);
  let active=0,peak=0;const order:string[]=[];
  const gateway={async execute(item:typeof call){
   if(item.name==='document.write')expect(active).toBe(0);
   active++;peak=Math.max(peak,active);order.push('start:'+item.id);
   await new Promise(resolve=>setTimeout(resolve,2));active--;order.push('end:'+item.id);
   return {status:'ok' as const,text:'done'};
  }};
  const result=await runAgent({...f.options,gateway,tools:[...f.options.tools,{name:'document.write',effect:'write',description:'Write'}],limits:{maxReadConcurrency:2}});
  expect(result.status).toBe('completed');expect(peak).toBe(2);
  expect(order.indexOf('start:c3')).toBeGreaterThan(order.indexOf('end:c1'));
  expect(order.indexOf('start:c3')).toBeGreaterThan(order.indexOf('end:c2'));
  expect(order.indexOf('start:c4')).toBeGreaterThan(order.indexOf('end:c3'));
  expect(f.events.map(e=>e.seq)).toEqual(f.events.map((_,i)=>i+1));
 });
 it('rechecks admission after approval instead of trusting the UI answer',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}],[{type:'done'}]]);
  const prepare=vi.fn(async(_call:unknown,_signal:unknown,response?:unknown)=>response?{status:'denied' as const,reason:'Permission revoked'}:{status:'wait' as const,request:{id:'q',kind:'approval' as const,prompt:'Apply?'}});
  const wait=vi.fn(async()=>{expect(f.execute).not.toHaveBeenCalled();return 'yes';});
  const result=await runAgent({...f.options,gateway:{execute:f.execute,prepare},interaction:{wait}});
  expect(result.status).toBe('completed');expect(f.execute).not.toHaveBeenCalled();expect(prepare).toHaveBeenCalledTimes(2);
  expect(f.events).toContainEqual(expect.objectContaining({payload:{type:'status',status:'waiting_approval'}}));
  expect(result.messages).toContainEqual(expect.objectContaining({status:'denied',text:'Permission revoked'}));
 });
 it('supports waiting for input and cancels an unresponsive input port',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}]]),controller=new AbortController();
  const result=await runAgent({...f.options,signal:controller.signal,gateway:{execute:f.execute,async prepare(){return {status:'wait',request:{id:'q',kind:'input',prompt:'Which document?'}};}},interaction:{async wait(){controller.abort();return new Promise(()=>{});}}});
  expect(result.status).toBe('cancelled');expect(f.execute).not.toHaveBeenCalled();
  expect(f.events).toContainEqual(expect.objectContaining({payload:{type:'status',status:'waiting_input'}}));
  expect(result.messages.at(-1)).toMatchObject({status:'skipped'});
 });
 it('denies a waiting tool when there is no interaction port',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}],[{type:'done'}]]);
  const result=await runAgent({...f.options,gateway:{execute:f.execute,async prepare(){return {status:'wait',request:{id:'q',kind:'approval',prompt:'Approve?'}};}}});
  expect(result.status).toBe('completed');expect(f.execute).not.toHaveBeenCalled();
  expect(result.messages).toContainEqual(expect.objectContaining({status:'denied',text:'Interaction unavailable'}));
 });
 it('stops before provider invocation when context exceeds hard budget',async()=>{
  const f=fixture([]);const stream=vi.fn(f.options.provider.stream);f.options.provider.stream=stream;
  const result=await runAgent({...f.options,context:{inputBudget:1}});
  expect(result.status).toBe('budget_exhausted');expect(stream).not.toHaveBeenCalled();
 });
 it('does not rewrite the journal during context compaction',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}],[{type:'done'}]]);
  let nextContext:readonly unknown[]=[];let turns=0;
  const provider:ProviderPort={async *stream(messages){if(turns++===0){yield {type:'tool',call};yield {type:'done'};}else{nextContext=messages;yield {type:'done'};}}};
  const compact=vi.fn(async()=> 'Summary of source and completed read');
  const result=await runAgent({...f.options,provider,context:{inputBudget:100,keepRecentMessages:0,estimate:messages=>messages.some(m=>m.role==='tool')?101:10,compact}});
  expect(result.status).toBe('completed');expect(compact).toHaveBeenCalledTimes(1);
  expect(nextContext).toEqual([{role:'user',text:'Read'},{role:'summary',text:'Summary of source and completed read'}]);
  expect(result.messages.some(m=>m.role==='tool')).toBe(true);expect(result.messages.some(m=>m.role==='summary')).toBe(false);
 });
 it('fails safely if compaction fails or cannot reduce the hard budget',async()=>{
  for(const failure of [true,false]){
   const f=fixture([[{type:'tool',call},{type:'done'}]]);
   const result=await runAgent({...f.options,context:{inputBudget:100,keepRecentMessages:0,estimate:messages=>messages.length===1?1:101,compact:async()=>{if(failure)throw new Error('Summary unavailable');return 'short';}}});
   expect(result.status).toBe(failure?'failed':'budget_exhausted');expect(f.execute).toHaveBeenCalledTimes(1);
  }
 });
 it('serializes concurrent event writes and stops after journal failure',async()=>{
  const f=fixture([[{type:'tool',call},{type:'tool',call:{...call,id:'c2'}},{type:'done'}]]);
  let active=0,peak=0;
  f.options.events.append=async(event)=>{
   active++;peak=Math.max(peak,active);await new Promise(resolve=>setTimeout(resolve,1));active--;
   if(event.payload.type==='message'&&event.payload.message.role==='tool')throw new Error('journal unavailable');
  };
  await expect(runAgent(f.options)).rejects.toThrow('journal unavailable');expect(peak).toBe(1);
 });
});

describe('agent context boundaries',()=>{
 it('retains entire recent assistant/tool exchanges',async()=>{
  const {ContextManager}=await import('@mydoc/agent-core');
  const compact=vi.fn(async()=> 'older exchange');
  const manager=new ContextManager({inputBudget:100,keepRecentMessages:1,estimate:messages=>messages.length>4?90:10,compact});
  const messages:import('@mydoc/agent-contracts').Message[]=[{role:'user',text:'goal'},{role:'assistant',text:'',calls:[call]},{role:'tool',callId:'c1',status:'ok',text:'old'},
   {role:'assistant',text:'',calls:[{...call,id:'c2'}]},{role:'tool',callId:'c2',status:'ok',text:'recent'}];
  const context=await manager.prepare(messages,[],new AbortController().signal);
  expect(compact.mock.calls).toHaveLength(1);
  expect(context.map(m=>m.role)).toEqual(['user','summary','assistant','tool']);
  expect(context.at(-1)).toMatchObject({callId:'c2'});expect(messages).toHaveLength(5);
 });
 it('bounds compaction requests before calling the summarizer',async()=>{
  const {ContextManager}=await import('@mydoc/agent-core');const compact=vi.fn(async()=> 'short');
  const manager=new ContextManager({inputBudget:10,keepRecentMessages:0,maxCompactionInputChars:5,estimate:()=>9,compact});
  await expect(manager.prepare([{role:'user',text:'goal'},{role:'assistant',text:'long history',calls:[]}],[],new AbortController().signal)).rejects.toThrow('Compaction input budget');
  expect(compact).not.toHaveBeenCalled();
 });
 it('rejects lossy or cyclic arguments and oversized catalog schemas',async()=>{
  const {validateTools}=await import('@mydoc/agent-contracts');
  for(const argumentsValue of [{value:undefined},{value:NaN},new Date(),{fn:()=>{}}])expect(()=>validateCall({...call,arguments:argumentsValue})).toThrow();
  const cycle:Record<string,unknown>={};cycle.self=cycle;expect(()=>validateCall({...call,arguments:cycle})).toThrow();
  expect(()=>validateTools([{name:'tool',effect:'read',description:'read',inputSchema:{description:'x'.repeat(128_000)}}])).toThrow();
 });
});

describe('A0 lifecycle acceptance',()=>{
 it('accounts for usage and prevents tools when total token budget is exceeded',async()=>{
  const f=fixture([[{type:'tool',call},{type:'usage',usage:{inputTokens:10,outputTokens:5,totalTokens:15}},{type:'done'}]]);
  const result=await runAgent({...f.options,limits:{maxTotalTokens:14}});
  expect(result.status).toBe('budget_exhausted');expect(result.usage.totalTokens).toBe(15);expect(f.execute).not.toHaveBeenCalled();
 });
 it('bounds result artifacts and supports continuation pages',async()=>{
  const f=fixture([[{type:'tool',call},{type:'done'}],[{type:'done'}]]);
  const result=await runAgent({...f.options,limits:{maxToolResultChars:2}});
  expect(await result.resultPages.read('r:c1',2,2)).toEqual({text:'un',nextOffset:4});
  expect(await result.resultPages.read('r:c1',4,2)).toEqual({text:'d',nextOffset:null});
  await expect(result.resultPages.read('other',0,10)).rejects.toThrow();
 });
 it('fences interrupted recovery by advancing generation and preserves terminal records',async()=>{
  const {recoverRun,validateRun,validateSourceRef,validateChangeSet,validateEvent}=await import('@mydoc/agent-contracts');
  const run={sessionId:'s',runId:'r',requestKey:'key',generation:3,status:'waiting_approval' as const};
  expect(recoverRun(run)).toMatchObject({generation:4,status:'interrupted'});expect(run.generation).toBe(3);
  const terminal={...run,status:'completed' as const};expect(recoverRun(terminal)).toEqual(terminal);
  expect(()=>validateRun({...run,generation:-1})).toThrow();expect(()=>validateSourceRef({})).toThrow();expect(()=>validateChangeSet({})).toThrow();expect(()=>validateEvent({})).toThrow();
 });
 it('writes a single terminal event through cancellation state',async()=>{
  const f=fixture([]),controller=new AbortController();controller.abort();
  await runAgent({...f.options,signal:controller.signal});
  const statuses=f.events.flatMap(e=>e.payload.type==='status'?[e.payload.status]:[]);
  expect(statuses).toEqual(['preparing','cancelling','cancelled']);
 });
 it('bounds artifact memory and validates pagination arguments',async()=>{
  const {MemoryResultPages}=await import('@mydoc/agent-core');const pages=new MemoryResultPages(3);
  await pages.put('x','abc');await expect(pages.put('y','d')).rejects.toThrow('budget');
  await expect(pages.read('x',-1,1)).rejects.toThrow();await expect(pages.put('x','abc')).rejects.toThrow('Duplicate');
 });
});

it('validates lifecycle payloads and distinguishes unavailable usage',async()=>{
 const {validateEvent}=await import('@mydoc/agent-contracts');
 validateSession({id:'s',ownerId:'u',workspaceId:'w',createdAt:1});
 validateStep({sessionId:'s',runId:'r',id:'step',ordinal:0,kind:'model'});
 expect(()=>validateStep({sessionId:'s',runId:'r',id:'step',ordinal:-1,kind:'model'})).toThrow();
 expect(()=>validateEvent({protocolVersion:1,seq:1,timestamp:0,sessionId:'s',runId:'r',payload:{type:'message',message:{role:'assistant',text:'',calls:[{}]}}})).toThrow();
 expect(DEFAULT_RUN_LIMITS).toMatchObject({maxRounds:20,timeoutMs:600000,maxReadConcurrency:4,maxToolResultChars:8000});
 const f=fixture([[{type:'done'}]]);expect((await runAgent(f.options)).usageReports).toBe(0);
});
