import {it,expect} from 'vitest';
import {createServer,request} from 'node:http';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import * as Y from 'yjs';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';

it('文档授权结构迁移有版本记录，重启保留既有授权和文档',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-permission-migration-')),path=join(dir,'cloud.sqlite');let cloud=new CloudServer(path);
 try{cloud.database.db.prepare('INSERT INTO cloud_document_permissions VALUES (?,?,?,?)').run('workspace','doc:existing','member','viewer');const doc=new Y.Doc();doc.getText('body').insert(0,'升级保留正文');cloud.database.append('cloud:workspace','doc:existing',Y.encodeStateAsUpdate(doc));doc.destroy();cloud.close();cloud=new CloudServer(path);
  expect(cloud.database.db.prepare('SELECT version FROM cloud_migrations').all()).toEqual([{version:1},{version:2},{version:3}]);expect(cloud.database.db.prepare('SELECT role FROM cloud_document_permissions').get()).toEqual({role:'viewer'});expect(cloud.database.load('cloud:workspace','doc:existing')).not.toBeNull();
 }finally{cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('文档级 viewer/editor 权限覆盖空间默认角色并阻止实时写入',async()=>{
 let bodyStarted:(()=>void)|undefined;
 const dir=mkdtempSync(join(tmpdir(),'mydoc-document-permissions-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer((req,res)=>{void cloud.handler(req,res);if(req.headers['x-test-slow'])bodyStarted?.();});
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',input?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(input?{'Content-Type':'application/json'}:{})},body:input?JSON.stringify(input):undefined});
 const slow=async(path:string,method:string,cookie:string,bytes:Buffer,change:()=>Promise<void>)=>{
  let started!:()=>void;const ready=new Promise<void>(resolve=>{started=resolve;});bodyStarted=started;
  let finish!:()=>void;const response=new Promise<number>((resolve,reject)=>{const req=request(root+path,{method,headers:{Cookie:cookie,'X-Test-Slow':'1','Content-Length':bytes.length}},res=>{res.resume();res.on('end',()=>resolve(res.statusCode!));});req.on('error',reject);req.write(bytes.subarray(0,1));finish=()=>req.end(bytes.subarray(1));});
  await ready;await change();finish();return response;
 };
 try{
  const ownerResponse=await call('/register','POST',{email:'doc-owner@example.com',name:'Owner',password:'abcdefghijk'}),ownerCookie=ownerResponse.headers.get('set-cookie')!.split(';')[0],owner=(await ownerResponse.json()).user;
  const memberResponse=await call('/register','POST',{email:'doc-member@example.com',name:'Member',password:'abcdefghijk'}),memberCookie=memberResponse.headers.get('set-cookie')!.split(';')[0],member=(await memberResponse.json()).user;
  const workspace=await(await call('/workspaces','POST',{name:'文档权限空间'},ownerCookie)).json();
  const invite=await(await call('/invite?workspace='+workspace.id,'POST',{email:'doc-member@example.com'},ownerCookie)).json();
  expect((await call('/accept','POST',{token:invite.token},memberCookie)).status).toBe(200);
  const doc=new Y.Doc();doc.getText('body').insert(0,'初始');const path='/document?workspace='+workspace.id+'&id=doc:restricted';
  expect((await fetch(root+path,{method:'POST',headers:{Cookie:ownerCookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))})).status).toBe(200);
  expect((await fetch(root+path,{method:'POST',headers:{Cookie:memberCookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))})).status).toBe(200);
  expect(await slow(path,'POST',memberCookie,Buffer.from(Y.encodeStateAsUpdate(doc)),async()=>{
   expect((await call('/document-permissions?workspace='+workspace.id,'PUT',{doc:'doc:restricted',user:member.id,role:'viewer'},ownerCookie)).status).toBe(200);
  })).toBe(403);
  expect((await call('/document-permissions?workspace='+workspace.id,'DELETE',{doc:'doc:restricted',user:member.id},ownerCookie)).status).toBe(200);
  expect(await slow('/document-permissions?workspace='+workspace.id,'PUT',ownerCookie,Buffer.from(JSON.stringify({doc:'doc:restricted',user:member.id,role:'viewer'})),async()=>{
   cloud.database.db.prepare('UPDATE cloud_members SET role=? WHERE workspace=? AND user=?').run('member',workspace.id,owner.id);
  })).toBe(403);
  cloud.database.db.prepare('UPDATE cloud_members SET role=? WHERE workspace=? AND user=?').run('owner',workspace.id,owner.id);
  expect((await call('/document-permissions?workspace='+workspace.id,'PUT',{doc:'doc:missing',user:member.id,role:'viewer'},ownerCookie)).status).toBe(404);
  expect((await call('/document-permissions?workspace='+workspace.id,'PUT',{doc:'doc:restricted',user:owner.id,role:'viewer'},ownerCookie)).status).toBe(400);
  expect((await call('/document-permissions?workspace='+workspace.id+'&doc=doc:restricted','PUT',{doc:'doc:restricted',user:member.id,role:'viewer'},ownerCookie)).status).toBe(200);
  expect(await (await call('/document-permissions?workspace='+workspace.id+'&doc=doc:restricted', 'GET',undefined,ownerCookie)).json()).toHaveLength(1);
  expect((await call(path,'GET',undefined,memberCookie)).status).toBe(200);
  expect((await fetch(root+path,{method:'POST',headers:{Cookie:memberCookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))})).status).toBe(403);
  expect((await call(path,'DELETE',undefined,memberCookie)).status).toBe(403);
  const sync=await fetch(root+'/sync?workspace='+workspace.id+'&id=doc:restricted',{method:'POST',headers:{Cookie:memberCookie},body:Buffer.from(Y.encodeStateVector(new Y.Doc()))});
  expect((await sync.json()).role).toBe('viewer');
  expect((await (await call('/access?workspace='+workspace.id+'&id=doc:restricted','GET',undefined,memberCookie)).json()).role).toBe('viewer');
  expect((await call('/document-permissions?workspace='+workspace.id+'&doc=doc:restricted','PUT',{doc:'doc:restricted',user:member.id,role:'editor'},ownerCookie)).status).toBe(200);
  expect((await fetch(root+path,{method:'POST',headers:{Cookie:memberCookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))})).status).toBe(200);
  expect((await call('/document-permissions?workspace='+workspace.id+'&doc=doc:restricted','DELETE',{doc:'doc:restricted',user:member.id},ownerCookie)).status).toBe(200);
  expect((await fetch(root+path,{method:'POST',headers:{Cookie:memberCookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))})).status).toBe(200);
  expect((await call('/document-permissions?workspace='+workspace.id+'&doc=doc:restricted','PUT',{doc:'doc:restricted',user:member.id,role:'denied'},ownerCookie)).status).toBe(200);
  const visibility=await (await call('/document-visibility?workspace='+workspace.id+'&id=doc:restricted&id=doc:missing', 'GET', undefined, memberCookie)).json();expect(visibility).toEqual([{id:'doc:restricted',allowed:false},{id:'doc:missing',allowed:false}]);
  expect((await call(path,'GET',undefined,memberCookie)).status).toBe(404);
  expect((await fetch(root+path,{method:'POST',headers:{Cookie:memberCookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))})).status).toBe(404);
  const deniedSync=await fetch(root+'/sync?workspace='+workspace.id+'&id=doc:restricted',{method:'POST',headers:{Cookie:memberCookie},body:Buffer.from(Y.encodeStateVector(new Y.Doc()))});expect(deniedSync.status).toBe(404);
  expect((await call('/access?workspace='+workspace.id+'&id=doc:restricted','GET',undefined,memberCookie)).status).toBe(404);
  const deniedEvents=await fetch(root+'/events?workspace='+workspace.id+'&id=doc:restricted',{headers:{Cookie:memberCookie}});expect(deniedEvents.status).toBe(404);deniedEvents.body?.cancel();
  const assetPath='/asset?workspace='+workspace.id+'&id=restricted-asset&name=a.txt&mime=text/plain';expect((await fetch(root+assetPath,{method:'PUT',headers:{Cookie:ownerCookie},body:'asset'})).status).toBe(200);expect((await call(assetPath+'&doc=doc:restricted','GET',undefined,memberCookie)).status).toBe(404);expect((await call(assetPath+'&doc=doc:restricted','GET',undefined,ownerCookie)).status).toBe(200);
  expect((await call('/document-permissions?workspace='+workspace.id+'&doc=doc:restricted','DELETE',{doc:'doc:restricted',user:member.id},ownerCookie)).status).toBe(200);
  expect((await fetch(root+path,{method:'POST',headers:{Cookie:memberCookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))})).status).toBe(200);
  expect((await call('/member?workspace='+workspace.id,'PATCH',{user:member.id,role:'viewer'},ownerCookie)).status).toBe(200);
  expect((await call('/document-permissions?workspace='+workspace.id,'PUT',{doc:'doc:restricted',user:member.id,role:'editor'},ownerCookie)).status).toBe(400);
  expect((await fetch(root+path,{method:'POST',headers:{Cookie:memberCookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))})).status).toBe(403);
  expect((await call('/document-permissions?workspace='+workspace.id,'PUT',{doc:'doc:restricted',user:member.id,role:'viewer'},ownerCookie)).status).toBe(200);
  expect((await call(path,'DELETE',undefined,ownerCookie)).status).toBe(200);
  expect(cloud.database.db.prepare('SELECT count(*) count FROM cloud_document_permissions').get()).toEqual({count:0});
  doc.destroy();
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});
