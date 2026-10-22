import {it,expect,vi} from 'vitest';
import Database from 'better-sqlite3';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {createServer} from 'node:http';
import * as Y from 'yjs';
import {SQLiteProjectDatabase} from '../packages/common/nbstore/server/database';
import {SQLiteSyncQueue} from '../packages/common/nbstore/server/sync-queue';
import {storageHandler} from '../packages/common/nbstore/server/http';
import {CloudStorage} from '../packages/common/nbstore/src/cloud';
import {setStorageAccount,setStorageGeneration} from '../packages/common/nbstore/src/desktop';
import {PersistentDocument} from '../packages/common/nbstore/src/index';

const target={scope:'user:a',server:'http://localhost',workspace:'ws',doc:'doc:a'};
function update(text:string){const doc=new Y.Doc();doc.getText('body').insert(0,text);const bytes=Y.encodeStateAsUpdate(doc);doc.destroy();return bytes;}
function read(bytes:Uint8Array){const doc=new Y.Doc();try{Y.applyUpdate(doc,bytes);return doc.getText('body').toString();}finally{doc.destroy();}}

it('队列与完整草稿原子落盘，重启保留顺序、重试信息，确认幂等且不复用序号',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-durable-')),path=join(dir,'queue.sqlite');let db=new SQLiteProjectDatabase(path,p=>new Database(p));
 try{
  let queue=new SQLiteSyncQueue(db);const seq=queue.enqueue(target,update('草稿'));queue.fail(target,seq,'503');
  expect(()=>queue.enqueue(target,new Uint8Array([255]))).toThrow();
  expect(queue.list(target)).toHaveLength(1);
  const generation=db.generation;db.close();db=new SQLiteProjectDatabase(path,p=>new Database(p));queue=new SQLiteSyncQueue(db);
  expect(db.generation).toBe(generation);expect(queue.list(target)[0]).toMatchObject({seq,attempts:1,error:'503'});
  expect(queue.list(target)[0].retryAt).toBeGreaterThan(Date.now());
  const draft=db.db.prepare('SELECT data FROM drafts').get() as {data:Uint8Array};expect(read(draft.data)).toBe('草稿');
  queue.acknowledge({...target,scope:'user:b'},seq);expect(queue.list(target)).toHaveLength(1);
  queue.acknowledge(target,seq);queue.acknowledge(target,seq);expect(queue.list(target)).toHaveLength(0);
  const next=queue.enqueue(target,update('新内容'));expect(next).toBeGreaterThan(seq);
  queue.acknowledge(target,seq);expect(queue.list(target)).toHaveLength(1);
  db.generation='restored';queue.acknowledge(target,next);expect(queue.list(target)).toHaveLength(1);
  expect(db.db.prepare('SELECT count(*) n FROM drafts').get()).toEqual({n:1});
 }finally{db.close();rmSync(dir,{recursive:true,force:true});}
});

it('草稿写入失败整个入队事务回滚，既有账户、服务器、空间与文档隔离',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-atomic-')),db=new SQLiteProjectDatabase(join(dir,'queue.sqlite'),p=>new Database(p));
 try{const queue=new SQLiteSyncQueue(db);
  db.db.exec("CREATE TRIGGER fail_draft BEFORE INSERT ON drafts BEGIN SELECT RAISE(ABORT,'disk full'); END");
  expect(()=>queue.enqueue(target,update('保留'))).toThrow('disk full');expect(queue.list(target)).toEqual([]);
  db.db.exec('DROP TRIGGER fail_draft');queue.enqueue(target,update('甲'));
  for(const other of [{...target,scope:'user:b'},{...target,server:'https://other.example'},{...target,workspace:'other'},{...target,doc:'doc:b'}])expect(queue.list(other)).toEqual([]);
 }finally{db.close();rmSync(dir,{recursive:true,force:true});}
});

async function fixture(){
 const dir=mkdtempSync(join(tmpdir(),'mydoc-replay-')),path=join(dir,'db.sqlite');let db=new SQLiteProjectDatabase(path,p=>new Database(p));
 const state={syncDown:true,storageDown:false,deny:0,viewer:false,loseAck:false,posts:0};
 const server=createServer((req,res)=>{
  const url=new URL(req.url??'/','http://local');
  if(url.pathname.startsWith('/api/storage/')){if(state.storageDown){res.writeHead(503);res.end();return;}
   if(state.loseAck&&req.method==='DELETE'){state.loseAck=false;res.writeHead(503);res.end();return;}
   req.url=req.url!.slice('/api/storage'.length);void storageHandler(db,false,undefined,'user:a')(req,res);return;
  }
  const json=(value:unknown,status=200)=>{res.writeHead(status,{'Content-Type':'application/json'});res.end(JSON.stringify(value));};
  if(req.headers['x-mydoc-account']!=='a'){json({error:'账户已变更'},409);return;}
  if(url.pathname==='/api/cloud/access'){json({role:state.viewer?'viewer':'member'});return;}
  if(url.pathname==='/api/cloud/document'){
   if(req.method==='GET'){const bytes=db.load('remote',target.doc);res.writeHead(bytes?200:204);res.end(bytes??undefined);return;}
   if(req.method==='POST'){state.posts++;if(state.deny||state.viewer){json({error:'权限失效或文档删除'},state.deny||403);return;}
    if(state.syncDown){json({error:'同步接口不可用'},503);return;}
    const parts:Buffer[]=[];req.on('data',part=>parts.push(part));req.on('end',()=>{db.append('remote',target.doc,Buffer.concat(parts));json({ok:true});});return;
   }
  }
  json({ok:true});
 });
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const fetchOriginal=globalThis.fetch;
 vi.stubGlobal('window',{});vi.stubGlobal('fetch',(path:string,init?:RequestInit)=>fetchOriginal(root+path,init));
 vi.stubGlobal('EventSource',class {static OPEN=1;readyState=0;close(){};onopen?:()=>void;onmessage?:()=>void;onerror?:()=>void;});
 setStorageAccount('a');setStorageGeneration(db.generation);
 const clients:CloudStorage[]=[];
 return {state,get db(){return db;},queue(){return new SQLiteSyncQueue(db);},client(){const cloud=new CloudStorage('ws');clients.push(cloud);return cloud;},restart(){db.close();db=new SQLiteProjectDatabase(path,p=>new Database(p));},async close(){clients.forEach(client=>client.close());setStorageAccount(undefined);vi.unstubAllGlobals();await new Promise<void>(resolve=>server.close(()=>resolve()));db.close();rmSync(dir,{recursive:true,force:true});}};
}

it('同步失败先落盘，客户端与 SQLite 重启后合并恢复，手动重试确认后清理',async()=>{
 const f=await fixture();try{
  const first=f.client();await first.append(target.doc,update('断线编辑'));expect(first.pendingCount).toBe(1);expect(first.status).toBe('offline');
  first.close();f.restart();const next=f.client();expect(read((await next.load(target.doc))!)).toBe('断线编辑');
  f.state.syncDown=false;await next.reconnect();expect(next.pendingCount).toBe(0);expect(next.lastSuccess).toBeGreaterThan(0);
  expect(read(f.db.load('remote',target.doc)!)).toBe('断线编辑');expect(f.db.db.prepare('SELECT count(*) n FROM drafts').get()).toEqual({n:0});
 }finally{await f.close();}
});

it('远端已确认但本地确认失败，重复发送幂等；服务器完全不可达时不声称落盘',async()=>{
 const f=await fixture();try{const client=f.client();f.state.syncDown=false;f.state.loseAck=true;
  await client.append(target.doc,update('仅一次正文'));expect(client.pendingCount).toBe(1);
  await client.reconnect();expect(read(f.db.load('remote',target.doc)!)).toBe('仅一次正文');expect(client.pendingCount).toBe(0);expect(f.state.posts).toBe(2);
  f.state.storageDown=true;await expect(client.append(target.doc,update('未落盘'))).rejects.toThrow('SQLite');expect(f.queue().list(target)).toEqual([]);
 }finally{await f.close();}
});

it.each([401,403,410])('会话失效、权限撤销或永久删除（%s）暂停且保留队列，不自动复活文档',async(status)=>{
 const f=await fixture();try{const client=f.client();f.state.syncDown=false;f.state.deny=status;
  await client.append(target.doc,update('被阻止草稿'));expect(client.status).toBe(status===403?'offline':'blocked');expect(client.isReadOnly(target.doc)).toBe(true);expect(client.pendingCount).toBe(1);
  await expect(client.append(target.doc,update('不允许编辑'))).rejects.toThrow();expect(f.queue().list(target)).toHaveLength(1);
  expect(f.db.load('remote',target.doc)).toBeNull();expect(f.db.db.prepare('SELECT count(*) n FROM drafts').get()).toEqual({n:1});
 }finally{await f.close();}
});

it('降权、账号切换、旧代际恢复均保留队列且停止重放',async()=>{
 const f=await fixture();try{const client=f.client();await client.append(target.doc,update('账户甲'));
  const posts=f.state.posts;f.state.viewer=true;await client.reconnect();expect(client.readOnly).toBe(true);expect(f.state.posts).toBe(posts);
  setStorageAccount('b');await expect(client.append(target.doc,update('不转投乙'))).rejects.toThrow('账户已变更');expect(client.status).toBe('blocked');expect(f.queue().list(target)).toHaveLength(1);
  setStorageAccount('a');f.state.viewer=false;f.state.syncDown=false;f.db.generation='restored';setStorageGeneration('restored');
  const restored=f.client();await restored.load(target.doc);await restored.reconnect();expect(restored.status).toBe('blocked');expect(f.state.posts).toBe(posts);expect(f.queue().list(target)).toHaveLength(1);
 }finally{await f.close();}
});

it('文档只读只暂停该文档队列，不影响其他文档，重新授权后重放草稿',async()=>{
 const f=await fixture();try{const client=f.client();f.state.syncDown=false;f.state.deny=403;
  await client.append(target.doc,update('只读草稿'));expect(client.isReadOnly(target.doc)).toBe(true);expect(client.readOnly).toBe(false);expect(client.pendingCount).toBe(1);
  const posts=f.state.posts;await expect(client.append(target.doc,update('不能续写'))).rejects.toThrow();expect(f.state.posts).toBe(posts);
  f.state.deny=0;await client.append('doc:other',update('其他文档'));expect(client.isReadOnly('doc:other')).toBe(false);expect(client.pendingCount).toBe(1);
  await client.reconnect();expect(client.isReadOnly(target.doc)).toBe(false);expect(client.pendingCount).toBe(0);expect(read(f.db.load('remote',target.doc)!)).toContain('只读草稿');
 }finally{await f.close();}
});

it('退避到期自动重试，不需要手动操作；暂存草稿 API 只允许绑定账户读取',async()=>{
 const f=await fixture();try{const client=f.client();await client.append(target.doc,update('自动恢复'));const first=f.queue().list(target)[0];expect(first.attempts).toBe(1);
  const path='/api/storage/sync-drafts';expect((await fetch(path,{headers:{'X-Mydoc-Account':'b'}})).status).toBe(409);
  const drafts=await (await fetch(path,{headers:{'X-Mydoc-Account':'a'}})).json() as {data:string}[];expect(drafts).toHaveLength(1);expect(read(Buffer.from(drafts[0].data,'base64'))).toBe('自动恢复');
  f.state.syncDown=false;await vi.waitFor(()=>expect(client.pendingCount).toBe(0),{timeout:5000,interval:50});expect(f.state.posts).toBe(2);expect(read(f.db.load('remote',target.doc)!)).toBe('自动恢复');
 }finally{await f.close();}
});

it('尚未落盘的后续改动不会被旧条目的成功确认清除；重试完整状态后恢复保存',async()=>{
 const f=await fixture();const doc=new Y.Doc();let persistence:PersistentDocument|undefined;
 try{const client=f.client();persistence=new PersistentDocument(target.doc,doc,client);await persistence.open();
  doc.getText('body').insert(0,'已落盘');await persistence.flush().catch(()=>{});expect(client.pendingCount).toBeGreaterThan(0);
  f.state.storageDown=true;doc.getText('body').insert(doc.getText('body').length,'未落盘');await expect(persistence.flush()).rejects.toThrow();expect(client.memoryOnly).toBe(true);
  f.state.storageDown=false;f.state.syncDown=false;await client.reconnect();expect(client.pendingCount).toBe(0);expect(persistence.state).toBe('error');expect(client.memoryOnly).toBe(true);
  await persistence.retry();expect(persistence.state).toBe('saved');expect(client.memoryOnly).toBe(false);expect(read(f.db.load('remote',target.doc)!)).toBe('已落盘未落盘');
 }finally{await persistence?.destroy();await f.close();}
});

it('办公代际封包与原始草稿同时落盘；归档旧队列保留旧草稿且不混入新代际',async()=>{
 const {packOfficeUpdate,unpackOfficeUpdate}=await import('../packages/common/nbstore/src/office-generation');const dir=mkdtempSync(join(tmpdir(),'mydoc-office-queue-')),db=new SQLiteProjectDatabase(join(dir,'db.sqlite'),p=>new Database(p));try{
  const queue=new SQLiteSyncQueue(db),old=update('旧代际草稿'),fresh=update('新代际内容');queue.enqueue(target,packOfficeUpdate(old,'old-generation-12345'),old);expect(unpackOfficeUpdate(queue.list(target)[0].data).generation).toBe('old-generation-12345');expect(read((db.db.prepare('SELECT data FROM drafts').get() as {data:Buffer}).data)).toBe('旧代际草稿');expect(queue.archiveOffice(target,'old-generation-12345')).toBe(1);expect(queue.list(target)).toHaveLength(0);
  const seq=queue.enqueue(target,packOfficeUpdate(fresh,'epoch-new-generation-12345'),fresh);expect(db.db.prepare('SELECT count(*) n FROM drafts').get()).toEqual({n:2});const newer=db.db.prepare("SELECT data FROM drafts WHERE generation LIKE '%@office:epoch-%'").get() as {data:Buffer};expect(read(newer.data)).toBe('新代际内容');queue.acknowledge(target,seq);expect(db.db.prepare('SELECT count(*) n FROM drafts').get()).toEqual({n:1});expect(read((db.db.prepare('SELECT data FROM drafts').get() as {data:Buffer}).data)).toBe('旧代际草稿');
 }finally{db.close();rmSync(dir,{recursive:true,force:true});}
});
