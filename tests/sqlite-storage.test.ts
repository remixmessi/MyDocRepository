import {it,expect} from 'vitest';
import Database from 'better-sqlite3';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {createServer} from 'node:http';
import * as Y from 'yjs';
import {ProjectDatabase,storageHandler,projectDatabasePath} from '../packages/frontend/apps/web/sqlite-server';
import {SCHEMA_VERSION,SCHEMA_MIGRATIONS} from '../packages/common/nbstore/src/schema';
it('本地账号密码哈希、重复邮箱、会话注销和数据库重启保留',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-local-account-')),path=join(dir,'db.sqlite');let db=new ProjectDatabase(path);
 const server=createServer((req,res)=>void storageHandler(db)(req,res));await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(value:unknown)=>fetch(root+'/local-account',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(value)});
 try{
  const input={email:'local@example.com',name:'本地用户',password:'local-password-123'};
  const created=await(await call({...input,action:'register'})).json();expect(created.user.email).toBe(input.email);
  const row=db.db.prepare('SELECT salt,password FROM local_accounts').get() as {salt:string;password:string};expect(row.password).not.toBe(input.password);expect(row.salt).toHaveLength(32);
  expect((await call({...input,action:'register'})).status).toBe(409);
  expect((await call({...input,password:'incorrect-password',action:'login'})).status).toBe(401);
  db.close();db=new ProjectDatabase(path);
  expect((await(await call({action:'session',token:created.token})).json()).user).toEqual(created.user);
  expect((await call({...input,action:'login'})).status).toBe(200);
  await call({action:'logout',token:created.token});expect((await(await call({action:'session',token:created.token})).json()).user).toBeNull();
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));db.close();rmSync(dir,{recursive:true});}
});
it('项目路径固定，真实 SQLite 重启恢复、并发增量合并和损坏更新不覆盖原数据',()=>{expect(projectDatabasePath).toBe(resolve('data/mydoc.sqlite'));const dir=mkdtempSync(join(tmpdir(),'mydoc-sqlite-')),path=join(dir,'data.sqlite');let db=new ProjectDatabase(path);const a=new Y.Doc(),b=new Y.Doc();a.getMap('body').set('a','中文正文');b.getMap('body').set('b','另一实例');db.append('main','doc',Y.encodeStateAsUpdate(a));const second=new ProjectDatabase(path);second.append('main','doc',Y.encodeStateAsUpdate(b));second.close();expect(()=>db.append('main','doc',new Uint8Array([255]))).toThrow();db.close();db=new ProjectDatabase(path);const restored=new Y.Doc();Y.applyUpdate(restored,db.load('main','doc')!);expect(restored.getMap('body').toJSON()).toEqual({a:'中文正文',b:'另一实例'});expect(db.load('other','doc')).toBeNull();db.close();for(const d of [a,b,restored])d.destroy();rmSync(dir,{recursive:true});});
it('偏好按更新时间合并不同文档阅读位置，保留最近值',()=>{const dir=mkdtempSync(join(tmpdir(),'mydoc-sqlite-')),db=new ProjectDatabase(join(dir,'data.sqlite'));const key='mydoc:reading-positions:v1';db.preference('main',key,JSON.stringify({a:{updatedAt:20,scroll:100}}));const result=db.preference('main',key,JSON.stringify({a:{updatedAt:10,scroll:0},b:{updatedAt:30,scroll:200}}));expect(JSON.parse(result)).toEqual({a:{updatedAt:20,scroll:100},b:{updatedAt:30,scroll:200}});db.close();rmSync(dir,{recursive:true});});
it('HTTP 二进制文档与附件、删除、来源限制、迁移原子回滚和重复导入保护',async()=>{const dir=mkdtempSync(join(tmpdir(),'mydoc-sqlite-')),db=new ProjectDatabase(join(dir,'data.sqlite')),server=createServer((req,res)=>void storageHandler(db)(req,res));await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const address=server.address() as {port:number},root=`http://127.0.0.1:${address.port}`;try{const doc=new Y.Doc();doc.getMap('meta').set('title','迁移文档');const update=Buffer.from(Y.encodeStateAsUpdate(doc)).toString('base64');const payload={documents:[['registry:local',update]],assets:[{workspace:'local',id:'a',name:'附件.txt',mime:'text/plain',data:Buffer.from('附件内容').toString('base64')}],preferences:[['mydoc:theme:v1','dark']]};let response=await fetch(root+'/migrate',{method:'POST',body:JSON.stringify({...payload,documents:[...payload.documents,['bad','/w==']]})});expect(response.status).toBe(500);expect(db.load('main','registry:local')).toBeNull();response=await fetch(root+'/migrate',{method:'POST',body:JSON.stringify(payload)});expect(await response.json()).toEqual({migrated:true});expect(await (await fetch(root+'/migrate',{method:'POST',body:JSON.stringify(payload)})).json()).toEqual({migrated:false});expect(new Uint8Array(await (await fetch(root+'/document?id=registry%3Alocal')).arrayBuffer())).toEqual(new Uint8Array(Buffer.from(update,'base64')));const asset=await (await fetch(root+'/asset?workspace=local&id=a')).json();expect(Buffer.from(asset.data,'base64').toString()).toBe('附件内容');expect((await fetch(root+'/document?id=x',{method:'POST',headers:{Origin:'https://evil.example'},body:Buffer.from(update,'base64')})).status).toBe(403);expect((await fetch(root+'/test-clear-assets',{method:'DELETE'})).status).toBe(404);await fetch(root+'/asset?workspace=local&id=a',{method:'DELETE'});expect((await fetch(root+'/asset?workspace=local&id=a')).status).toBe(404);doc.destroy();}finally{await new Promise<void>(resolve=>server.close(()=>resolve()));db.close();rmSync(dir,{recursive:true});}});

it('新文档读取返回 204 且不创建记录，保存后读取二进制，删除后恢复空响应',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-sqlite-')),db=new ProjectDatabase(join(dir,'data.sqlite'));
 const server=createServer((req,res)=>void storageHandler(db)(req,res));
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const root=`http://127.0.0.1:${(server.address() as {port:number}).port}`,doc=new Y.Doc();
 try{
  const url=root+'/document?id=doc%3Anew';
  let response=await fetch(url);expect(response.status).toBe(204);expect(await response.text()).toBe('');expect(db.load('main','doc:new')).toBeNull();
  doc.getMap('body').set('text','Markdown 正文');const bytes=Y.encodeStateAsUpdate(doc);
  expect((await fetch(url,{method:'POST',body:bytes})).status).toBe(200);
  response=await fetch(url);expect(response.status).toBe(200);expect(new Uint8Array(await response.arrayBuffer())).toEqual(bytes);
  await fetch(url,{method:'DELETE'});expect((await fetch(url)).status).toBe(204);
  expect((await fetch(root+'/unknown')).status).toBe(404);
 }finally{doc.destroy();await new Promise<void>(resolve=>server.close(()=>resolve()));db.close();rmSync(dir,{recursive:true});}
});

it('首次事务初始化空库，已初始化的空库和有数据的库重启均跳过',()=>{const dir=mkdtempSync(join(tmpdir(),'mydoc-init-')),path=join(dir,'project.sqlite');try{let db=new ProjectDatabase(path);expect(db.initializedThisStart).toBe(true);expect(db.db.pragma('user_version',{simple:true})).toBe(SCHEMA_VERSION);for(const table of ['documents','assets','preferences','pending_updates','drafts'])expect((db.db.prepare(`SELECT count(*) n FROM ${table}`).get() as {n:number}).n).toBe(0);db.close();db=new ProjectDatabase(path);expect(db.initializedThisStart).toBe(false);db.preference('main','theme','dark');const doc=new Y.Doc();doc.getMap('body').set('text','保留内容');db.append('main','doc',Y.encodeStateAsUpdate(doc));doc.destroy();db.close();db=new ProjectDatabase(path);expect(db.initializedThisStart).toBe(false);expect(db.load('main','doc')).not.toBeNull();expect(db.db.prepare('SELECT value FROM preferences WHERE key=?').get('theme')).toEqual({value:'dark'});db.close();}finally{rmSync(dir,{recursive:true,force:true});}});
it('版本 1 旧库升级到 2 补建离线队列与草稿表并保留原数据；迁移失败整体回滚不前进版本',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-migrate-')),path=join(dir,'project.sqlite');
 try{
  // 构造只有版本 1 结构的既有库，并写入业务数据。
  let raw=new Database(path);raw.exec(SCHEMA_MIGRATIONS[1]);raw.pragma('user_version = 1');
  const doc=new Y.Doc();doc.getMap('body').set('text','升级前正文');raw.prepare('INSERT INTO documents VALUES (?,?,?)').run('main','doc:a',Buffer.from(Y.encodeStateAsUpdate(doc)));
  raw.prepare('INSERT INTO assets VALUES (?,?,?,?,?,?)').run('main','local','asset-a','旧附件.txt','text/plain',Buffer.from('附件内容'));
  raw.prepare('INSERT INTO preferences VALUES (?,?,?)').run('main','theme','dark');
  // 迁移步骤 2 与既有对象冲突：同名表挡住索引创建，整体必须回滚。
  raw.exec('CREATE TABLE pending_updates_target(id INTEGER)');raw.close();
  expect(()=>new ProjectDatabase(path)).toThrow();
  raw=new Database(path);expect(raw.pragma('user_version',{simple:true})).toBe(1);
  expect(raw.prepare("SELECT name FROM sqlite_master WHERE name='drafts'").get()).toBeUndefined();
  expect(raw.prepare("SELECT name FROM sqlite_master WHERE name='pending_updates'").get()).toBeUndefined();
  raw.exec('DROP TABLE pending_updates_target');raw.close();
  // 清除冲突后升级：补建新表，业务数据保留，且不算首次初始化。
  const db=new ProjectDatabase(path);
  expect(db.initializedThisStart).toBe(false);
  expect(db.db.pragma('user_version',{simple:true})).toBe(SCHEMA_VERSION);
  for(const table of ['pending_updates','drafts'])expect((db.db.prepare(`SELECT count(*) n FROM ${table}`).get() as {n:number}).n).toBe(0);
  const restored=new Y.Doc();Y.applyUpdate(restored,db.load('main','doc:a')!);expect(restored.getMap('body').get('text')).toBe('升级前正文');
  expect(db.db.prepare('SELECT name FROM assets WHERE id=?').get('asset-a')).toEqual({name:'旧附件.txt'});
  expect(db.db.prepare('SELECT value FROM preferences WHERE key=?').get('theme')).toEqual({value:'dark'});
  db.close();doc.destroy();restored.destroy();
  // 再次启动已是最新版本，不重复迁移。
  const again=new ProjectDatabase(path);expect(again.initializedThisStart).toBe(false);expect(again.db.pragma('user_version',{simple:true})).toBe(SCHEMA_VERSION);again.close();
 }finally{rmSync(dir,{recursive:true,force:true});}
});
it('离线队列与草稿表按账户／服务器／空间隔离键写入，同一文档保持提交顺序',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-queue-')),db=new ProjectDatabase(join(dir,'project.sqlite'));
 try{
  const insert=db.db.prepare('INSERT INTO pending_updates(scope,server,workspace,doc,data,created,generation) VALUES (?,?,?,?,?,?,?)');
  insert.run('user:a','https://one.example','ws-1','doc:x',Buffer.from([1]),1,db.generation);
  insert.run('user:a','https://one.example','ws-1','doc:x',Buffer.from([2]),2,db.generation);
  insert.run('user:a','https://two.example','ws-1','doc:x',Buffer.from([3]),3,db.generation);
  insert.run('user:b','https://one.example','ws-1','doc:x',Buffer.from([4]),4,db.generation);
  insert.run('user:a','https://one.example','ws-2','doc:x',Buffer.from([5]),5,db.generation);
  const rows=db.db.prepare('SELECT seq,data FROM pending_updates WHERE scope=? AND server=? AND workspace=? AND doc=? ORDER BY seq').all('user:a','https://one.example','ws-1','doc:x') as {seq:number;data:Buffer}[];
  expect(rows.map(r=>r.data[0])).toEqual([1,2]);
  expect(rows[0].seq).toBeLessThan(rows[1].seq);
  expect((db.db.prepare('SELECT count(*) n FROM pending_updates').get() as {n:number}).n).toBe(5);
  expect((db.db.prepare('SELECT attempts,retryAt,error FROM pending_updates WHERE seq=?').get(rows[0].seq) as {attempts:number;retryAt:number;error:null})).toEqual({attempts:0,retryAt:0,error:null});
  const draft=db.db.prepare('INSERT INTO drafts(id,scope,server,workspace,doc,title,summary,data,missingAssets,created,generation) VALUES (?,?,?,?,?,?,?,?,?,?,?)');
  draft.run('draft-1','user:a','https://one.example','ws-1','doc:x','草稿标题','正文摘要',Buffer.from([9]),1,10,db.generation);
  expect(()=>draft.run('draft-1','user:a','https://one.example','ws-1','doc:x','重复','重复',Buffer.from([9]),0,11,db.generation)).toThrow();
  expect(db.db.prepare('SELECT title,missingAssets FROM drafts WHERE scope=? AND server=?').get('user:a','https://one.example')).toEqual({title:'草稿标题',missingAssets:1});
  expect(db.db.prepare('SELECT count(*) n FROM drafts WHERE scope=?').get('user:b')).toEqual({n:0});
 }finally{db.close();rmSync(dir,{recursive:true,force:true});}
});
it('初始化失败整体回滚，修复后能重试；拒绝降级覆盖高版本数据库',()=>{const dir=mkdtempSync(join(tmpdir(),'mydoc-init-')),path=join(dir,'project.sqlite');try{let raw=new Database(path);raw.exec('CREATE TABLE helper(id INTEGER); CREATE INDEX assets ON helper(id)');raw.close();expect(()=>new ProjectDatabase(path)).toThrow();raw=new Database(path);expect(raw.pragma('user_version',{simple:true})).toBe(0);expect(raw.prepare("SELECT name FROM sqlite_master WHERE name='documents'").get()).toBeUndefined();raw.exec('DROP INDEX assets');raw.close();const db=new ProjectDatabase(path);expect(db.initializedThisStart).toBe(true);db.close();raw=new Database(path);raw.pragma('user_version = 99');raw.close();expect(()=>new ProjectDatabase(path)).toThrow('数据库版本高于当前程序');raw=new Database(path);expect(raw.pragma('user_version',{simple:true})).toBe(99);raw.close();}finally{rmSync(dir,{recursive:true,force:true});}});
