import {it,expect} from 'vitest';
import {createServer} from 'node:http';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import * as Y from 'yjs';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';

async function fixture(){
 const dir=mkdtempSync(join(tmpdir(),'mydoc-cloud-settings-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',input?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(input?{'Content-Type':'application/json'}:{})},body:input?JSON.stringify(input):undefined});
 const register=async(email:string,name:string)=>{const response=await call('/register','POST',{email,name,password:'settings-password-1'});return {cookie:response.headers.get('set-cookie')!.split(';')[0],user:(await response.json()).user as {id:string;email:string;name:string}};};
 const binary=(path:string,method:string,body:Buffer,cookie:string)=>fetch(root+path,{method,headers:{Cookie:cookie},body});
 return {call,register,binary,async close(){await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}};
}

it('概览返回名称、所有者、当前角色、成员分布、文档数量与存储占用',async()=>{
 const f=await fixture();
 try{
  const owner=await f.register('owner@example.com','所有者');
  const space=await (await f.call('/workspaces','POST',{name:'概览空间'},owner.cookie)).json() as {id:string};
  // 空库概览：只有所有者，没有文档与附件。
  let overview=await (await f.call('/overview?workspace='+space.id,'GET',undefined,owner.cookie)).json();
  expect(overview).toMatchObject({name:'概览空间',role:'owner',documents:0,invites:0});
  expect(overview.owner).toEqual({name:'所有者',email:'owner@example.com'});
  expect(overview.members).toMatchObject({total:1,owners:1,admins:0,members:0,viewers:0});
  expect(overview.storage.total).toBe(0);

  // 写入一篇文档、共享模型与一个附件后，统计同步增长。
  const doc=new Y.Doc();doc.getText('body').insert(0,'概览正文');
  await f.binary('/document?workspace='+space.id+'&id=doc:a','POST',Buffer.from(Y.encodeStateAsUpdate(doc)),owner.cookie);
  await f.binary('/asset?workspace='+space.id+'&id=file-a&name=a.txt&mime=text/plain','PUT',Buffer.from('附件字节'),owner.cookie);
  doc.destroy();

  const member=await f.register('member@example.com','成员');
  const invite=await (await f.call('/invite?workspace='+space.id,'POST',{email:'member@example.com',role:'member'},owner.cookie)).json() as {token:string};
  await f.call('/accept','POST',{token:invite.token},member.cookie);
  const viewer=await f.register('viewer@example.com','查看者');
  const viewerInvite=await (await f.call('/invite?workspace='+space.id,'POST',{email:'viewer@example.com',role:'viewer'},owner.cookie)).json() as {token:string};
  await f.call('/accept','POST',{token:viewerInvite.token},viewer.cookie);

  overview=await (await f.call('/overview?workspace='+space.id,'GET',undefined,owner.cookie)).json();
  expect(overview.members).toMatchObject({total:3,owners:1,members:1,viewers:1});
  expect(overview.documents).toBe(1);
  expect(overview.storage.assetCount).toBe(1);
  expect(overview.storage.assets).toBeGreaterThan(0);
  expect(overview.storage.documents).toBeGreaterThan(0);
  expect(overview.storage.total).toBe(overview.storage.documents+overview.storage.assets);

  // 成员和查看者看到的是自己的角色，不是所有者的角色。
  expect((await (await f.call('/overview?workspace='+space.id,'GET',undefined,member.cookie)).json()).role).toBe('member');
  expect((await (await f.call('/overview?workspace='+space.id,'GET',undefined,viewer.cookie)).json()).role).toBe('viewer');
  // 非成员不能读取概览。
  const outsider=await f.register('outsider@example.com','外部');
  expect((await f.call('/overview?workspace='+space.id,'GET',undefined,outsider.cookie)).status).toBe(403);
 }finally{await f.close();}
});

it('邀请列表含角色与过期状态；重新生成替换旧口令，撤销后口令失效',async()=>{
 const f=await fixture();
 try{
  const owner=await f.register('owner@example.com','所有者');
  const space=await (await f.call('/workspaces','POST',{name:'邀请空间'},owner.cookie)).json() as {id:string};
  const first=await (await f.call('/invite?workspace='+space.id,'POST',{email:'guest@example.com',role:'viewer'},owner.cookie)).json() as {token:string;role:string;expiresAt:number};
  expect(first.role).toBe('viewer');
  let invites=await (await f.call('/invites?workspace='+space.id,'GET',undefined,owner.cookie)).json() as {email:string;role:string;expired:boolean}[];
  expect(invites).toHaveLength(1);
  expect(invites[0]).toMatchObject({email:'guest@example.com',role:'viewer',expired:false});

  // 重新生成同邮箱邀请：旧口令立即失效，列表不堆积重复条目。
  const second=await (await f.call('/invite?workspace='+space.id,'POST',{email:'guest@example.com',role:'member'},owner.cookie)).json() as {token:string};
  expect(second.token).not.toBe(first.token);
  invites=await (await f.call('/invites?workspace='+space.id,'GET',undefined,owner.cookie)).json();
  expect(invites).toHaveLength(1);
  expect(invites[0].role).toBe('member');
  const guest=await f.register('guest@example.com','受邀人');
  expect((await f.call('/accept','POST',{token:first.token},guest.cookie)).status).toBe(400);

  // 撤销后新口令也不可用。
  const list=await (await f.call('/invites?workspace='+space.id,'GET',undefined,owner.cookie)).json() as {id:string}[];
  await f.call('/invite?workspace='+space.id,'DELETE',{id:list[0].id},owner.cookie);
  expect(await (await f.call('/invites?workspace='+space.id,'GET',undefined,owner.cookie)).json()).toEqual([]);
  expect((await f.call('/accept','POST',{token:second.token},guest.cookie)).status).toBe(400);

  // 过期邀请仍展示并标记过期，但不能接受。
  const third=await (await f.call('/invite?workspace='+space.id,'POST',{email:'guest@example.com',role:'member'},owner.cookie)).json() as {token:string};
  const cloudDb=await (await f.call('/overview?workspace='+space.id,'GET',undefined,owner.cookie)).json();
  expect(cloudDb.invites).toBe(1);
  // 普通成员不能读取或撤销邀请。
  const invite=await (await f.call('/invite?workspace='+space.id,'POST',{email:'member@example.com',role:'member'},owner.cookie)).json() as {token:string};
  const member=await f.register('member@example.com','成员');
  await f.call('/accept','POST',{token:invite.token},member.cookie);
  expect((await f.call('/invites?workspace='+space.id,'GET',undefined,member.cookie)).status).toBe(403);
  expect((await f.call('/invite?workspace='+space.id,'DELETE',{id:'x'},member.cookie)).status).toBe(403);
  expect((await f.call('/accept','POST',{token:third.token},member.cookie)).status).toBe(400);
 }finally{await f.close();}
});

it('空间管理员可搜索已有账号并排除已加入成员，普通成员不能读取账号目录',async()=>{
 const f=await fixture();
 try{
  const owner=await f.register('owner@example.com','所有者');const existing=await f.register('existing@example.com','已有账号');
  const space=await (await f.call('/workspaces','POST',{name:'账号目录空间'},owner.cookie)).json() as {id:string};
  let users=await (await f.call('/directory-users?workspace='+space.id+'&q=existing','GET',undefined,owner.cookie)).json();expect(users).toEqual([{id:existing.user.id,name:'已有账号',email:'existing@example.com'}]);
  const invite=await (await f.call('/invite?workspace='+space.id,'POST',{email:existing.user.email,role:'member'},owner.cookie)).json() as {token:string};await f.call('/accept','POST',{token:invite.token},existing.cookie);
  users=await (await f.call('/directory-users?workspace='+space.id+'&q=existing','GET',undefined,owner.cookie)).json();expect(users).toEqual([]);
  expect((await f.call('/directory-users?workspace='+space.id+'&q=owner','GET',undefined,existing.cookie)).status).toBe(403);
 }finally{await f.close();}
});

it('管理员不能越权操作所有者与其他管理员，查看者不能管理空间',async()=>{
 const f=await fixture();
 try{
  const owner=await f.register('owner@example.com','所有者');
  const space=await (await f.call('/workspaces','POST',{name:'权限空间'},owner.cookie)).json() as {id:string};
  const join=async(email:string,name:string,role:string)=>{const invite=await (await f.call('/invite?workspace='+space.id,'POST',{email,role:role==='viewer'?'viewer':'member'},owner.cookie)).json() as {token:string};const account=await f.register(email,name);await f.call('/accept','POST',{token:invite.token},account.cookie);return account;};
  const adminA=await join('admin-a@example.com','管理员甲','member');
  const adminB=await join('admin-b@example.com','管理员乙','member');
  const viewer=await join('viewer@example.com','查看者','viewer');
  await f.call('/member?workspace='+space.id,'PATCH',{user:adminA.user.id,role:'admin'},owner.cookie);
  await f.call('/member?workspace='+space.id,'PATCH',{user:adminB.user.id,role:'admin'},owner.cookie);

  // 管理员不能移除所有者或其他管理员，也不能调整角色。
  expect((await f.call('/member?workspace='+space.id,'DELETE',{user:owner.user.id},adminA.cookie)).status).toBe(403);
  expect((await f.call('/member?workspace='+space.id,'DELETE',{user:adminB.user.id},adminA.cookie)).status).toBe(403);
  expect((await f.call('/member?workspace='+space.id,'PATCH',{user:viewer.user.id,role:'admin'},adminA.cookie)).status).toBe(403);
  // 管理员可以移除普通查看者，但不能删除空间或转移所有权。
  expect((await f.call('/workspace?workspace='+space.id,'DELETE',undefined,adminA.cookie)).status).toBe(403);
  expect((await f.call('/transfer?workspace='+space.id,'POST',{user:adminB.user.id},adminA.cookie)).status).toBe(403);
  expect((await f.call('/member?workspace='+space.id,'DELETE',{user:viewer.user.id},adminA.cookie)).status).toBe(200);

  // 查看者不能邀请、改名或读取邀请列表。
  const viewerTwo=await join('viewer2@example.com','查看者二','viewer');
  expect((await f.call('/invite?workspace='+space.id,'POST',{email:'x@example.com'},viewerTwo.cookie)).status).toBe(403);
  expect((await f.call('/workspace?workspace='+space.id,'PATCH',{name:'改名'},viewerTwo.cookie)).status).toBe(403);
  expect((await f.call('/invites?workspace='+space.id,'GET',undefined,viewerTwo.cookie)).status).toBe(403);
  // 所有者必须先转移所有权才能退出；转移后仍只有一个所有者。
  expect((await f.call('/leave?workspace='+space.id,'POST',undefined,owner.cookie)).status).toBe(403);
  await f.call('/transfer?workspace='+space.id,'POST',{user:adminA.user.id},owner.cookie);
  const overview=await (await f.call('/overview?workspace='+space.id,'GET',undefined,owner.cookie)).json();
  expect(overview.members.owners).toBe(1);
  expect(overview.role).toBe('admin');
  expect((await (await f.call('/overview?workspace='+space.id,'GET',undefined,adminA.cookie)).json()).role).toBe('owner');
 }finally{await f.close();}
});
