import {it,expect} from 'vitest';
import Database from 'better-sqlite3';
import {mkdtempSync,rmSync,readFileSync,writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {createHash} from 'node:crypto';
import {SQLiteProjectDatabase} from '../packages/common/nbstore/server/database';
import {DatabaseBackups} from '../packages/common/nbstore/server/backups';
import {SCHEMA_MIGRATIONS,SCHEMA_VERSION} from '../packages/common/nbstore/src/schema';
import {AGENT_COLUMNS} from '../packages/common/nbstore/src/agent-schema';
import {SQLiteAgentStore} from '../packages/common/agent-host/src/sqlite-store';
import {runAgent} from '@mydoc/agent-core';
const connect=(path:string)=>new Database(path);
function fixture(){const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-'));const path=join(dir,'db.sqlite');return {dir,path,open:()=>new SQLiteProjectDatabase(path,connect),clean:()=>rmSync(dir,{recursive:true,force:true})};}
it('upgrades version 4 atomically, retains data, reopens and refuses future versions',()=>{
 const f=fixture();try{
  let raw=new Database(f.path);for(let i=1;i<=4;i++)raw.exec(SCHEMA_MIGRATIONS[i]);raw.pragma('user_version=4');raw.exec("INSERT INTO preferences VALUES('main','keep','yes')");raw.close();
  expect(()=>new SQLiteProjectDatabase(f.path,connect,()=>{throw new Error('migration failed');})).toThrow('migration failed');
  raw=new Database(f.path);expect(raw.pragma('user_version',{simple:true})).toBe(4);expect(raw.prepare("SELECT name FROM sqlite_master WHERE name='agent_runs'").get()).toBeUndefined();raw.close();
  let db=f.open();expect(db.schemaVersion).toBe(5);expect(db.db.prepare('SELECT value FROM preferences').get()).toEqual({value:'yes'});db.close();db=f.open();expect(db.db.pragma('user_version',{simple:true})).toBe(5);db.db.exec('PRAGMA user_version=6');db.close();expect(()=>f.open()).toThrow('高于当前程序');
 }finally{f.clean();}
});
it('persists run messages, call/results and replay cursors through actual reopen',async()=>{
 const f=fixture();let db=f.open();try{
  let store=new SQLiteAgentStore(db,'user:a');store.createSession('s','w','test');store.enqueue('s','r','key','read');const lease=store.claim('r');let n=0;
  const result=await runAgent({sessionId:'s',runId:'r',prompt:'read',tools:[{name:'read',effect:'read',description:''}],provider:{async *stream(){if(n++===0)yield {type:'tool',call:{id:'c',name:'read',arguments:{}}};yield {type:'done'};}},gateway:{async execute(){return {status:'ok',text:'saved'};}},clock:{now:()=>1},events:store.events(lease)});
  expect(result.status).toBe('completed');expect(db.db.prepare('SELECT result FROM agent_tool_calls').get()).toMatchObject({result:expect.stringContaining('saved')});db.close();db=f.open();store=new SQLiteAgentStore(db,'user:a');expect(store.load('r')?.status).toBe('completed');expect(store.replay('s').length).toBeGreaterThan(5);expect(new SQLiteAgentStore(db,'user:b').listSessions('w')).toEqual([]);expect(()=>new SQLiteAgentStore(db,'user:b').replay('s')).toThrow();
 }finally{db.close();f.clean();}
});
it('deduplicates queue, fences recovery and rolls back orphan results',async()=>{
 const f=fixture(),db=f.open();try{
  const store=new SQLiteAgentStore(db,'a');store.createSession('s','w','test');store.enqueue('s','r','key','read');expect(store.enqueue('s','r2','key','read').runId).toBe('r');expect(()=>store.enqueue('s','r2','key','different')).toThrow('conflict');const lease=store.claim('r');store.enqueue('s','r2','key2','more');expect(()=>store.claim('r2')).toThrow('already running');
  const event={protocolVersion:1 as const,sessionId:'s',runId:'r',seq:1,timestamp:1,payload:{type:'message' as const,message:{role:'tool' as const,callId:'missing',status:'ok' as const,text:'bad'}}};
  await expect(store.events(lease).append(event)).rejects.toThrow('Orphan');expect(store.replay('s')).toEqual([]);
  store.recoverActive();expect(store.load('r')).toMatchObject({status:'interrupted',generation:2});expect(store.load('r2')?.status).toBe('queued');await expect(store.events(lease).append({...event,payload:{type:'text',text:'stale'}})).rejects.toThrow('Stale');store.deleteSession('s');expect(store.load('r')).toBeNull();
 }finally{db.close();f.clean();}
});
it('restores Agent snapshots with fencing, supports legacy snapshots without Agent tables',async()=>{
 const f=fixture(),db=f.open();try{
  const store=new SQLiteAgentStore(db,'a');store.createSession('s','w','test');store.enqueue('s','r','key','read');const lease=store.claim('r'),manager=new DatabaseBackups(db,connect),backup=await manager.create();
  store.createSession('later','w','later');await manager.restore(backup.id);expect(store.listSessions('w')).toHaveLength(1);expect(store.load('r')?.status).toBe('interrupted');await expect(store.events(lease).append({protocolVersion:1,sessionId:'s',runId:'r',seq:1,timestamp:1,payload:{type:'text',text:'stale'}})).rejects.toThrow();
  const legacy=await manager.create(),path=join(manager.settings.directory,legacy.id+'.sqlite'),raw=new Database(path);for(const table of Object.keys(AGENT_COLUMNS))raw.exec('DROP TABLE '+table);raw.pragma('user_version=4');raw.close();const bytes=readFileSync(path);writeFileSync(join(manager.settings.directory,legacy.id+'.json'),JSON.stringify({...legacy,schemaVersion:4,bytes:bytes.length,sha256:createHash('sha256').update(bytes).digest('hex')}));
  await manager.restore(legacy.id);expect(store.listSessions('w')).toEqual([]);expect(db.schemaVersion).toBe(SCHEMA_VERSION);
 }finally{db.close();f.clean();}
});
it('Agent restore failure rolls back prior deletes and database generation',async()=>{
 const f=fixture(),db=f.open();try{
  const store=new SQLiteAgentStore(db,'a');store.createSession('old','w','old');const manager=new DatabaseBackups(db,connect),backup=await manager.create();store.createSession('new','w','new');const generation=db.generation;
  db.db.exec("CREATE TRIGGER fail_agent_restore BEFORE INSERT ON agent_sessions BEGIN SELECT RAISE(ABORT,'agent restore failure'); END");
  await expect(manager.restore(backup.id)).rejects.toThrow('agent restore failure');expect(store.listSessions('w')).toHaveLength(2);expect(db.generation).toBe(generation);db.db.exec('DROP TRIGGER fail_agent_restore');
 }finally{db.close();f.clean();}
});
it('partial Agent structure fails startup without advancing the version',()=>{
 const f=fixture();try{
  const raw=new Database(f.path);for(let i=1;i<=4;i++)raw.exec(SCHEMA_MIGRATIONS[i]);raw.exec('CREATE TABLE agent_sessions(scope TEXT,id TEXT);PRAGMA user_version=4');raw.close();
  expect(()=>f.open()).toThrow();const check=new Database(f.path);expect(check.pragma('user_version',{simple:true})).toBe(4);expect(check.prepare("SELECT name FROM sqlite_master WHERE name='agent_runs'").get()).toBeUndefined();check.close();
 }finally{f.clean();}
});
