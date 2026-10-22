import {it,expect} from 'vitest';
import {createServer,request} from 'node:http';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import Database from 'better-sqlite3';
import {DatabaseBackups} from '../packages/common/nbstore/server/backups';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {SQLiteProjectDatabase} from '../packages/common/nbstore/server/database';
import {upgradeCloudDatabase,CLOUD_SCHEMA_VERSION} from '../packages/frontend/apps/web/server/schema';
const admin={email:'admin@example.com',name:'服务器管理员',password:'admin-password-123'};
async function fixture(){
 const dir=mkdtempSync(join(tmpdir(),'mydoc-global-admin-')),path=join(dir,'cloud.sqlite');let cloud=new CloudServer(path,{defaultAdmin:admin});
 const server=createServer((req,res)=>void cloud.handler(req,res));await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(route:string,method='GET',input?:unknown,cookie?:string)=>fetch(root+route,{method,headers:{...(cookie?{Cookie:cookie}:{}),'Content-Type':'application/json'},body:input===undefined?undefined:JSON.stringify(input)});
 const account=async(email:string)=>{const response=await call('/register','POST',{email,name:email.split('@')[0],password:'member-password-123'});return {user:(await response.json()).user,cookie:response.headers.get('set-cookie')!.split(';')[0]};};
 const login=await call('/login','POST',admin),cookie=login.headers.get('set-cookie')!.split(';')[0];
 return {path,dir,root,call,account,cookie,get cloud(){return cloud;},restart(){cloud.close();cloud=new CloudServer(path);},async close(){server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}};
}
it('服务器管理员全局管理未加入的空间，成员权限与文档访问保持隔离，操作留痕并持久化',async()=>{
 const f=await fixture();try{
  const owner=await f.account('owner@example.com'),member=await f.account('member@example.com');
  const space=await(await f.call('/workspaces','POST',{name:'未加入的团队'},owner.cookie)).json(),q='?workspace='+space.id;
  expect(await(await f.call('/workspaces','GET',undefined,f.cookie)).json()).toEqual([]);
  expect((await f.call('/admin/workspaces')).status).toBe(401);
  for(const route of ['/admin/workspaces','/admin/workspace'+q,'/admin/workspace-audit'+q])expect((await f.call(route,'GET',undefined,owner.cookie)).status).toBe(403);
  expect(await(await f.call('/admin/workspaces?q=未加入','GET',undefined,f.cookie)).json()).toMatchObject({total:1,spaces:[{id:space.id,ownerEmail:'owner@example.com',memberCount:1}]});
  expect((await f.call('/admin/workspaces?offset=-1','GET',undefined,f.cookie)).status).toBe(400);
  expect((await f.call('/admin/workspace?workspace=missing','GET',undefined,f.cookie)).status).toBe(404);
  expect((await f.call('/admin/users'+q,'GET',undefined,f.cookie)).status).toBe(200);
  for(const route of ['/access'+q,'/members'+q,'/document'+q+'&id=doc:test','/asset'+q+'&id=asset','/sync'+q+'&id=doc:test'])expect((await f.call(route,'GET',undefined,f.cookie)).status).toBe(403);
  expect((await f.call('/admin/workspace'+q,'PATCH',{name:'越权'},owner.cookie)).status).toBe(403);
  expect((await f.call('/admin/workspace'+q,'PATCH',{name:'改名团队'},f.cookie)).status).toBe(200);
  expect((await f.call('/admin/workspace-member'+q,'PUT',{user:member.user.id,role:'admin'},f.cookie)).status).toBe(200);
  expect((await f.call('/admin/workspace-member'+q,'PATCH',{user:member.user.id,role:'viewer'},f.cookie)).status).toBe(200);
  expect((await f.call('/admin/workspace-member'+q,'DELETE',{user:owner.user.id},f.cookie)).status).toBe(400);
  expect((await f.call('/admin/workspace-owner'+q,'POST',{user:member.user.id,confirmName:'旧名称'},f.cookie)).status).toBe(400);
  expect((await f.call('/admin/workspace-owner'+q,'POST',{user:member.user.id,confirmName:'改名团队'},f.cookie)).status).toBe(200);
  const detail=await(await f.call('/admin/workspace'+q,'GET',undefined,f.cookie)).json();
  expect(detail.members.find((m:{id:string})=>m.id===member.user.id).role).toBe('owner');expect(detail.members.find((m:{id:string})=>m.id===owner.user.id).role).toBe('admin');
  expect((await f.call('/admin/workspace-member'+q,'DELETE',{user:owner.user.id},f.cookie)).status).toBe(200);
  expect((await f.call('/access'+q,'GET',undefined,owner.cookie)).status).toBe(403);
  f.restart();
  expect(await(await f.call('/admin/workspace-audit'+q,'GET',undefined,f.cookie)).json()).toMatchObject({total:5});
  expect((await f.call('/admin/workspace'+q,'DELETE',{confirmName:'错误'},f.cookie)).status).toBe(400);
  expect((await f.call('/admin/workspace'+q,'DELETE',{confirmName:'改名团队'},f.cookie)).status).toBe(200);
  expect(await(await f.call('/admin/workspaces','GET',undefined,f.cookie)).json()).toMatchObject({total:0});
  expect(await(await f.call('/admin/workspace-audit'+q,'GET',undefined,f.cookie)).json()).toMatchObject({total:6,entries:expect.arrayContaining([expect.objectContaining({action:'delete',actorName:admin.name})])});
 }finally{await f.close();}
});
it('重新指定未入空间用户为所有者，仅保留一个所有者；审计写入失败时管理修改一并回滚',async()=>{
 const f=await fixture();try{
  const owner=await f.account('owner@example.com'),next=await f.account('next@example.com');
  const space=await(await f.call('/workspaces','POST',{name:'所有权测试'},owner.cookie)).json(),q='?workspace='+space.id;
  f.cloud.database.db.exec("CREATE TRIGGER reject_audit BEFORE INSERT ON cloud_workspace_audit BEGIN SELECT RAISE(ABORT,'test audit failure'); END");
  expect((await f.call('/admin/workspace'+q,'PATCH',{name:'不应保存'},f.cookie)).status).toBe(500);
  expect(f.cloud.database.db.prepare('SELECT name FROM cloud_workspaces WHERE id=?').get(space.id)).toEqual({name:'所有权测试'});
  f.cloud.database.db.exec('DROP TRIGGER reject_audit');
  expect((await f.call('/admin/workspace-owner'+q,'POST',{user:next.user.id,confirmName:'所有权测试'},f.cookie)).status).toBe(200);
  expect(f.cloud.database.db.prepare("SELECT user FROM cloud_members WHERE workspace=? AND role='owner'").all(space.id)).toEqual([{user:next.user.id}]);
  expect((await f.call('/admin/workspace-member'+q,'PATCH',{user:owner.user.id,role:'owner'},f.cookie)).status).toBe(400);
 }finally{await f.close();}
});
it('读取请求体期间取消管理员权限，管理请求重新鉴权并拒绝执行',async()=>{
 const f=await fixture();try{
  const owner=await f.account('owner@example.com');const space=await(await f.call('/workspaces','POST',{name:'原名'},owner.cookie)).json();
  const db=f.cloud.database.db;const result=new Promise<number>(resolve=>{
   const req=request(f.root+'/admin/workspace?workspace='+space.id,{method:'PATCH',headers:{Cookie:f.cookie,'Content-Type':'application/json','Transfer-Encoding':'chunked'}},res=>{res.resume();res.on('end',()=>resolve(res.statusCode!));});
   req.flushHeaders();req.write('{"name":');
   setTimeout(()=>{db.prepare('UPDATE cloud_users SET server_admin=0 WHERE email=?').run(admin.email);req.end('"被拒绝"}');},50);
  });
  expect(await result).toBe(403);expect(db.prepare('SELECT name FROM cloud_workspaces WHERE id=?').get(space.id)).toEqual({name:'原名'});
 }finally{await f.close();}
});
it('版本 6 升级保留空间和账号，重复启动、迁移失败回滚、高版本拒绝、完整快照保留审计',async()=>{
 const f=await fixture();try{
  const owner=await f.account('owner@example.com'),space=await(await f.call('/workspaces','POST',{name:'迁移空间'},owner.cookie)).json();
  f.cloud.database.db.exec('DROP TABLE cloud_workspace_audit;DELETE FROM cloud_migrations WHERE version>=7');
  const before=f.cloud.database.db.prepare('SELECT * FROM cloud_users').all();
  expect(()=>new SQLiteProjectDatabase(f.path,p=>new Database(p),db=>{upgradeCloudDatabase(db);throw new Error('migration failed');})).toThrow('migration failed');
  expect(f.cloud.database.db.prepare("SELECT name FROM sqlite_master WHERE name='cloud_workspace_audit'").get()).toBeUndefined();
  expect(f.cloud.database.db.prepare('SELECT max(version) version FROM cloud_migrations').get()).toEqual({version:6});
  f.restart();f.restart();expect(f.cloud.database.db.prepare('SELECT * FROM cloud_users').all()).toEqual(before);
  expect(f.cloud.database.db.prepare('SELECT name FROM cloud_workspaces WHERE id=?').get(space.id)).toEqual({name:'迁移空间'});
  expect((await f.call('/admin/workspace?workspace='+space.id,'PATCH',{name:'升级后'},f.cookie)).status).toBe(200);
  const backup=join(f.dir,'snapshot.sqlite');f.cloud.database.db.prepare('VACUUM INTO ?').run(backup);
  const restored=new SQLiteProjectDatabase(backup,p=>new Database(p),upgradeCloudDatabase);expect(restored.db.prepare('SELECT count(*) n FROM cloud_workspace_audit').get()).toEqual({n:1});restored.close();
  f.cloud.database.db.prepare('INSERT INTO cloud_migrations VALUES(?)').run(CLOUD_SCHEMA_VERSION+1);
  expect(()=>new SQLiteProjectDatabase(f.path,p=>new Database(p),upgradeCloudDatabase)).toThrow('协作数据库版本高于');
 }finally{await f.close();}
});

it('全局列表和审计分页，删除清理所有空间数据并保留历史，旧业务备份恢复不覆盖审计',async()=>{
 const f=await fixture();const oldDb=new SQLiteProjectDatabase(join(f.dir,'old-business.sqlite'),p=>new Database(p));const oldBackups=new DatabaseBackups(oldDb,p=>new Database(p));const backups=new DatabaseBackups(f.cloud.database,p=>new Database(p));
 try{
  const owner=await f.account('owner@example.com');const space=await(await f.call('/workspaces','POST',{name:'清理空间'},owner.cookie)).json(),q='?workspace='+space.id;
  const db=f.cloud.database.db;
  for(let i=0;i<23;i++)db.prepare('INSERT INTO cloud_workspaces VALUES(?,?,?)').run('page-'+i,'分页团队 '+i,i);
  expect(await(await f.call('/admin/workspaces?q=分页团队','GET',undefined,f.cookie)).json()).toMatchObject({total:23,spaces:expect.any(Array),limit:20});
  const second=await(await f.call('/admin/workspaces?q=分页团队&offset=20','GET',undefined,f.cookie)).json();expect(second.spaces).toHaveLength(3);
  await f.call('/admin/workspace'+q,'PATCH',{name:'清理空间改名'},f.cookie);
  const backup=await oldBackups.create();backups.setDirectory(oldBackups.settings.directory);await backups.restore(backup.id);
  expect(db.prepare('SELECT count(*) n FROM cloud_workspace_audit').get()).toEqual({n:1});expect(db.prepare('SELECT name FROM cloud_workspaces WHERE id=?').get(space.id)).toEqual({name:'清理空间改名'});
  db.prepare('INSERT INTO cloud_document_permissions VALUES(?,?,?,?)').run(space.id,'doc:a',owner.user.id,'denied');
  db.prepare('INSERT INTO cloud_asset_owners VALUES(?,?,?,?)').run(space.id,'asset-a','doc:a',owner.user.id);
  db.prepare('INSERT INTO cloud_deleted VALUES(?,?)').run(space.id,'doc:deleted');
  db.prepare('INSERT INTO cloud_invites VALUES(?,?,?,?,?)').run('invite',space.id,'new@example.com','member',Date.now()+10000);
  db.prepare('INSERT INTO assets VALUES(?,?,?,?,?,?)').run('cloud:'+space.id,'local','asset-a','file.txt','text/plain',Buffer.from('file'));
  expect((await f.call('/admin/workspace'+q,'DELETE',{confirmName:'清理空间改名'},f.cookie)).status).toBe(200);
  for(const table of ['cloud_members','cloud_invites','cloud_document_permissions','cloud_asset_owners','cloud_deleted'])expect(db.prepare(`SELECT count(*) n FROM ${table} WHERE workspace=?`).get(space.id)).toEqual({n:0});
  expect(db.prepare('SELECT count(*) n FROM assets WHERE scope=?').get('cloud:'+space.id)).toEqual({n:0});
  expect(await(await f.call('/admin/workspace-audit','GET',undefined,f.cookie)).json()).toMatchObject({total:2,entries:expect.arrayContaining([expect.objectContaining({action:'delete'})])});
  expect(await(await f.call('/admin/workspace-audit?offset=20','GET',undefined,f.cookie)).json()).toMatchObject({entries:[],total:2});
 }finally{oldBackups.stop();backups.stop();oldDb.close();await f.close();}
});
