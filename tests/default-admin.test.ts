import {it,expect} from 'vitest';
import {createServer} from 'node:http';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {writeFileSync} from 'node:fs';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {defaultAdminConfig,type DefaultAdminConfig} from '../packages/frontend/apps/web/server/default-admin';
import {serverConfig} from '../packages/frontend/apps/web/server/config';

const admin={email:'admin@example.com',password:'initial-password-123',name:'服务器管理员'};
async function fixture(config?:DefaultAdminConfig){
 const dir=mkdtempSync(join(tmpdir(),'mydoc-default-admin-')),path=join(dir,'cloud.sqlite');let cloud=new CloudServer(path,{defaultAdmin:config});
 const server=createServer((req,res)=>void cloud.handler(req,res));await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(route:string,input?:unknown,cookie?:string,method?:string)=>fetch(root+route,{method:method??(input===undefined?'GET':'POST'),headers:{...(cookie?{Cookie:cookie}:{}),'Content-Type':'application/json'},body:input===undefined?undefined:JSON.stringify(input)});
 return {call,get cloud(){return cloud;},restart(config?:DefaultAdminConfig){cloud.close();cloud=new CloudServer(path,{defaultAdmin:config});},async close(){server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}};
}

it('管理员用户列表按所选空间筛选，搜索仍受成员范围限制',async()=>{
 const f=await fixture(admin);try{
  const login=await f.call('/login',{email:admin.email,password:admin.password}),cookie=login.headers.get('set-cookie')!.split(';')[0];
  const a=await (await f.call('/workspaces',{name:'空间 A'},cookie)).json(),b=await (await f.call('/workspaces',{name:'空间 B'},cookie)).json();
  const registered=await f.call('/register',{email:'member@example.com',name:'A 成员',password:'member-password-123'}),memberCookie=registered.headers.get('set-cookie')!.split(';')[0];
  const invite=await (await f.call('/invite?workspace='+a.id,{email:'member@example.com',role:'member'},cookie)).json();
  await f.call('/accept',{token:invite.token},memberCookie);
  const listA=await (await f.call('/admin/users?workspace='+a.id,undefined,cookie)).json();
  expect(listA.total).toBe(2);expect(listA.users.map((user:{email:string})=>user.email)).toContain('member@example.com');
  const listB=await (await f.call('/admin/users?workspace='+b.id+'&q=member',undefined,cookie)).json();
  expect(listB).toMatchObject({users:[],total:0});
  expect((await f.call('/admin/users?workspace='+a.id,undefined,memberCookie)).status).toBe(403);
  expect((await f.call('/admin/users?workspace=missing',undefined,cookie)).ok).toBe(false);
 }finally{await f.close();}
});

it('默认管理员配置可选，邮箱规范化，拒绝不完整或无效配置',()=>{
 expect(defaultAdminConfig({})).toBeUndefined();expect(defaultAdminConfig({MYDOC_ADMIN_EMAIL:' Admin@Example.com ',MYDOC_ADMIN_PASSWORD:admin.password})).toEqual({...admin,name:'管理员'});
 for(const env of [{MYDOC_ADMIN_EMAIL:admin.email},{MYDOC_ADMIN_PASSWORD:admin.password},{MYDOC_ADMIN_EMAIL:'invalid',MYDOC_ADMIN_PASSWORD:admin.password},{MYDOC_ADMIN_EMAIL:admin.email,MYDOC_ADMIN_PASSWORD:'short'},{MYDOC_ADMIN_EMAIL:admin.email,MYDOC_ADMIN_PASSWORD:' password-with-space '}])expect(()=>defaultAdminConfig(env)).toThrow();
});

it('首次创建可登录的服务器管理员，普通注册不能自授管理员，列表不返回凭据',async()=>{
 const f=await fixture(admin);try{
  const row=f.cloud.database.db.prepare('SELECT password,salt FROM cloud_users').get() as {password:string;salt:string};expect(row.password).not.toBe(admin.password);expect(row.password).toHaveLength(128);
  const login=await f.call('/login',{email:admin.email,password:admin.password}),cookie=login.headers.get('set-cookie')!.split(';')[0];expect(login.status).toBe(200);expect((await login.json()).user.serverAdmin).toBe(true);
  expect((await(await f.call('/session',undefined,cookie)).json()).user).toMatchObject({email:admin.email,name:admin.name,serverAdmin:true,verified:false});
  const ordinary=await f.call('/register',{email:'ordinary@example.com',name:'普通用户',password:'ordinary-password-123',serverAdmin:true}),ordinaryCookie=ordinary.headers.get('set-cookie')!.split(';')[0];expect((await ordinary.json()).user.serverAdmin).toBe(false);
  expect((await f.call('/admin/users',undefined,ordinaryCookie)).status).toBe(403);expect((await f.call('/admin/users')).status).toBe(401);
  const list=await (await f.call('/admin/users',undefined,cookie)).json();expect(list.total).toBe(2);for(const user of list.users){expect(user).not.toHaveProperty('salt');expect(user).not.toHaveProperty('password');}
  const ordinaryUser=list.users.find((user:{email:string})=>user.email==='ordinary@example.com')!;const grant=await f.call('/admin/user',{user:ordinaryUser.id,serverAdmin:true},cookie,'PATCH');expect(grant.status,await grant.clone().text()).toBe(200);const revoke=await f.call('/admin/user',{user:ordinaryUser.id,serverAdmin:false},cookie,'PATCH');expect(revoke.status,await revoke.clone().text()).toBe(200);const current=(await f.call('/session',undefined,cookie).then(response=>response.json())).user;expect((await f.call('/admin/user',{user:current.id,serverAdmin:false},cookie,'PATCH')).status).toBe(400);
  expect((await f.call('/admin/users?offset=-1',undefined,cookie)).status).toBe(400);
  expect((await(await f.call('/workspaces',undefined,cookie)).json())).toEqual([]);
 }finally{await f.close();}
});

it('重启不会重复创建或覆盖改密、改名，删除初始化配置仍保留管理员',async()=>{
 const f=await fixture(admin);try{
  const login=await f.call('/login',{email:admin.email,password:admin.password}),cookie=login.headers.get('set-cookie')!.split(';')[0],id=(await login.json()).user.id;
  expect((await f.call('/password',{current:admin.password,password:'changed-password-123'},cookie)).status).toBe(200);
  f.cloud.database.db.prepare('UPDATE cloud_users SET name=?,email=? WHERE id=?').run('已修改名称','changed@example.com',id);
  f.restart({...admin,password:'different-config-password',email:'new-config@example.com'});
  expect(f.cloud.database.db.prepare('SELECT count(*) count FROM cloud_users').get()).toEqual({count:1});
  const response=await f.call('/login',{email:'changed@example.com',password:'changed-password-123'});expect(response.status).toBe(200);expect((await response.json()).user).toMatchObject({id,name:'已修改名称',serverAdmin:true});
  expect((await f.call('/login',{email:'changed@example.com',password:admin.password})).status).toBe(401);
  f.restart();expect((await f.call('/login',{email:'changed@example.com',password:'changed-password-123'})).status).toBe(200);
 }finally{await f.close();}
});

it('未配置时保持空库，已有普通邮箱冲突时拒绝提升，版本升级保留旧用户',async()=>{
 const f=await fixture();try{
  expect(f.cloud.database.db.prepare('SELECT count(*) count FROM cloud_users').get()).toEqual({count:0});
  await f.call('/register',{email:admin.email,name:'原有普通用户',password:'ordinary-password-123'});
  // 模拟上一版数据库：只有协作迁移 1，没有服务器管理员字段。
  f.cloud.database.db.exec('DELETE FROM cloud_migrations WHERE version=2; DROP TABLE cloud_bootstrap; ALTER TABLE cloud_users DROP COLUMN server_admin;');
  f.restart();expect(f.cloud.database.db.prepare('SELECT name,server_admin FROM cloud_users').get()).toEqual({name:'原有普通用户',server_admin:0});
  expect(()=>f.restart(admin)).toThrow('邮箱已被现有账户使用');
  f.restart();expect(f.cloud.database.db.prepare('SELECT server_admin FROM cloud_users').get()).toEqual({server_admin:0});expect(f.cloud.database.db.prepare('SELECT count(*) count FROM cloud_bootstrap').get()).toEqual({count:0});
 }finally{await f.close();}
});

it('环境变量优先于 server-config.json 中的管理员配置',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-server-config-')),path=join(dir,'server-config.json');
 writeFileSync(path,JSON.stringify({host:'127.0.0.1',port:9090,admin:{email:'file@example.com',password:'file-password-123',name:'配置文件管理员'}}));
 try{
  const config=serverConfig({MYDOC_CONFIG_PATH:path,MYDOC_ADMIN_EMAIL:'env@example.com',MYDOC_ADMIN_PASSWORD:'env-password-123',MYDOC_ADMIN_NAME:'容器管理员'});
  expect(config.host).toBe('127.0.0.1');expect(config.port).toBe(9090);
  expect(config.admin).toEqual({email:'env@example.com',password:'env-password-123',name:'容器管理员'});
 }finally{rmSync(dir,{recursive:true,force:true});}
});

it('未设置管理员环境变量时使用配置文件回退值',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-server-config-')),path=join(dir,'server-config.json');
 writeFileSync(path,JSON.stringify({admin:{email:'file@example.com',password:'file-password-123',name:'配置文件管理员'}}));
 try{expect(serverConfig({MYDOC_CONFIG_PATH:path}).admin).toEqual({email:'file@example.com',password:'file-password-123',name:'配置文件管理员'});}
 finally{rmSync(dir,{recursive:true,force:true});}
});
