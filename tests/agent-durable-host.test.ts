import {it,expect} from 'vitest';
import {build} from 'esbuild';
import Database from 'better-sqlite3';
import {mkdtempSync,rmSync,writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {SQLiteProjectDatabase} from '../packages/common/nbstore/server/database';
import {SQLiteAgentStore} from '../packages/common/agent-host/src/sqlite-store';
import {DurableAgentHost} from '../packages/common/agent-host/src/durable-host';
import {ChildAgentRuntime} from '../packages/common/agent-host/src/child-runtime';
function fixture(){const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-host-')),path=join(dir,'db.sqlite');const db=new SQLiteProjectDatabase(path,p=>new Database(p));const store=new SQLiteAgentStore(db,'a');store.createSession('s','w','test');return {dir,path,db,store,close(){db.close();rmSync(dir,{recursive:true,force:true});}};}
const config=()=>({provider:{async *stream(){yield {type:'text' as const,text:'ok'};yield {type:'done' as const};}},gateway:{async execute(){return {status:'ok' as const,text:'result'};}},tools:[],clock:{now:()=>1}});
it('executes a durable queue in order, deduplicates requests and isolates active sessions',async()=>{
 const f=fixture();try{
  const host=new DurableAgentHost(f.store);host.enqueue('s','r','key','first');host.enqueue('s','r2','key2','second');expect(host.enqueue('s','other','key','first').runId).toBe('r');
  expect(await host.runNext('s',config)).toBe(true);expect(f.store.load('r')?.status).toBe('completed');expect(f.store.load('r2')?.status).toBe('queued');expect(await host.runNext('s',config)).toBe(true);expect(await host.runNext('s',config)).toBe(false);expect(f.store.replay('s').length).toBeGreaterThan(6);await host.close();expect(await host.runNext('s',config)).toBe(false);
 }finally{f.close();}
});
it('runs in a real child process with host-only model/tools and persisted events',async()=>{
 const f=fixture();try{
  const entry=join(f.dir,'worker.cjs');await build({entryPoints:['packages/common/agent-host/src/worker-child.ts'],outfile:entry,bundle:true,platform:'node',format:'cjs'});
  const host=new DurableAgentHost(f.store,new ChildAgentRuntime(entry));host.enqueue('s','r','key','read');let turn=0;
  await host.runNext('s',()=>({...config(),tools:[{name:'read',effect:'read' as const,description:''}],provider:{async *stream(){if(turn++===0)yield {type:'tool' as const,call:{id:'c',name:'read',arguments:{}}};else yield {type:'text' as const,text:'saved'};yield {type:'done' as const};}}}));
  expect(f.store.load('r')?.status).toBe('completed');expect(f.db.db.prepare('SELECT result FROM agent_tool_calls').get()).toMatchObject({result:expect.stringContaining('result')});await host.close();
 }finally{f.close();}
});
it('fences a crashed child and preserves queued work across actual database reopen',async()=>{
 const f=fixture();try{
  const entry=join(f.dir,'crash.cjs');writeFileSync(entry,'process.exit(3)');const host=new DurableAgentHost(f.store,new ChildAgentRuntime(entry));host.enqueue('s','r','k','one');host.enqueue('s','r2','k2','two');await expect(host.runNext('s',config)).rejects.toThrow('exited');expect(f.store.load('r')?.status).toBe('interrupted');
  f.db.close();const reopened=new SQLiteProjectDatabase(f.path,p=>new Database(p));try{const store=new SQLiteAgentStore(reopened,'a'),next=new DurableAgentHost(store);next.recover();expect(await next.runNext('s',config)).toBe(true);expect(store.load('r2')?.status).toBe('completed');await next.close();}finally{reopened.close();}
 }finally{try{f.close();}catch{rmSync(f.dir,{recursive:true,force:true});}}
});
it('cancels an active child, safely closes, and cancels queued work without executing',async()=>{
 const f=fixture();try{
  const entry=join(f.dir,'worker.cjs');await build({entryPoints:['packages/common/agent-host/src/worker-child.ts'],outfile:entry,bundle:true,platform:'node',format:'cjs'});
  const host=new DurableAgentHost(f.store,new ChildAgentRuntime(entry,100));host.enqueue('s','r','k','one');host.enqueue('s','r2','k2','two');expect(host.cancel('r2')).toBe(true);
  let entered!:()=>void;const ready=new Promise<void>(resolve=>{entered=resolve;});
  const running=host.runNext('s',()=>({...config(),provider:{async *stream(){entered();await new Promise(()=>{});yield {type:'done' as const};}}}));await ready;await host.close();await running;expect(['cancelled','interrupted']).toContain(f.store.load('r')?.status);expect(f.store.load('r2')?.status).toBe('cancelled');
 }finally{f.close();}
});
it('forwards source policy and human admission across worker RPC and keeps result pages in host',async()=>{
 const f=fixture();try{
  const entry=join(f.dir,'worker.cjs');await build({entryPoints:['packages/common/agent-host/src/worker-child.ts'],outfile:entry,bundle:true,platform:'node',format:'cjs'});
  f.store.enqueue('s','r','k','read');const lease=f.store.claim('r');let turn=0,approved=false,sourceChecks=0;
  const result=await new ChildAgentRuntime(entry).run({...config(),sessionId:'s',runId:'r',prompt:'read',events:f.store.events(lease),tools:[{name:'read',effect:'read',description:''}],
   sourceContext:{async snapshot(){return {policy:{serverId:'local',workspaceId:'w',revision:'1',instructions:'Read only',allowedTools:['read'],allowModelTransfer:true},sources:[{ref:{serverId:'local',workspaceId:'w',resourceId:'d',kind:'text',anchor:'p',revision:'1',excerptHash:'h'},text:'source'}]};},async authorize(){sourceChecks++;return true;}},
   interaction:{async wait(){approved=true;return 'yes';}},gateway:{async prepare(_call,_signal,response){return response?{status:'ready'}:{status:'wait',request:{id:'q',kind:'approval',prompt:'Read?'}};},async execute(){expect(approved).toBe(true);return {status:'ok',text:'x'.repeat(9000)};}},
   provider:{async *stream(messages){expect(messages[0].role).toBe('system');if(turn++===0)yield {type:'tool',call:{id:'c',name:'read',arguments:{}}};yield {type:'done'};}}});
  expect(result.status).toBe('completed');expect(sourceChecks).toBeGreaterThan(2);expect((await result.resultPages.read('r:c',8000,8000)).text).toHaveLength(1000);
 }finally{f.close();}
});
it('records unconfirmed call results when a runtime dies after journaling a tool request',async()=>{
 const f=fixture();try{
  const host=new DurableAgentHost(f.store,{async run(options){await options.events.append({protocolVersion:1,runId:options.runId,sessionId:options.sessionId,seq:1,timestamp:1,payload:{type:'message',message:{role:'assistant',text:'',calls:[{id:'c',name:'read',arguments:{}}]}}});throw new Error('worker lost');}});
  host.enqueue('s','r','k','read');await expect(host.runNext('s',config)).rejects.toThrow('worker lost');expect(f.store.load('r')?.status).toBe('interrupted');expect(f.db.db.prepare('SELECT result FROM agent_tool_calls').get()).toMatchObject({result:expect.stringContaining('unconfirmed')});await host.close();
 }finally{f.close();}
});
it('bounds an unresponsive worker without relying on its cancellation handler',async()=>{
 const f=fixture();try{
  const entry=join(f.dir,'hung.cjs');writeFileSync(entry,"process.on('message',()=>{});setInterval(()=>{},1000)");
  const host=new DurableAgentHost(f.store,new ChildAgentRuntime(entry,20));host.enqueue('s','r','k','read');
  await expect(host.runNext('s',()=>({...config(),limits:{timeoutMs:20}}))).rejects.toThrow('deadline');expect(f.store.load('r')?.status).toBe('interrupted');await host.close();
 }finally{f.close();}
});
