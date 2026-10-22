import {it,expect} from 'vitest';
import {createServer} from 'node:http';
import {randomUUID} from 'node:crypto';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import * as Y from 'yjs';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {officeCommentRevision,validateOfficeComments,type OfficeCommentThread} from '../packages/common/nbstore/src/office-comments';
import {officeKindFromFilename} from '../packages/common/nbstore/src/office';
import {parseOfficeBackup} from '../packages/frontend/apps/web/server/office-backup';
async function fixture(){
 const dir=mkdtempSync(join(tmpdir(),'mydoc-office-comments-')),cloud=new CloudServer(join(dir,'db.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const register=async(name:string)=>{const response=await fetch(root+'/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:name+'@example.com',name,password:'test-password-123'})});return {user:(await response.json()).user.id as string,cookie:response.headers.get('set-cookie')!.split(';')[0]};};const owner=await register('owner'),other=await register('other');
 const call=(path:string,method='GET',data?:unknown,cookie=owner.cookie)=>fetch(root+path,{method,headers:{Cookie:cookie,'Content-Type':'application/json'},body:data instanceof Uint8Array?data as BodyInit:data===undefined?undefined:JSON.stringify(data)}),space=await(await call('/workspaces','POST',{name:'评论工作区'})).json();cloud.database.db.prepare("INSERT INTO cloud_members VALUES (?,?,'member')").run(space.id,other.user);
 const read=(id:string)=>{const doc=new Y.Doc();Y.applyUpdate(doc,cloud.database.load('cloud:'+space.id,id)!);const comments=doc.getMap<OfficeCommentThread>('officeComments');return {doc,comments};};
 return {cloud,owner,other,space,call,read,close:async()=>{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}};
}
it('各办公类型共享评论、回复、解决与备份恢复，回执幂等且保留作者',async()=>{
 const f=await fixture();try{for(const format of ['docx','xlsx','pptx','pdf','txt']){
  const resource=await(await f.call('/office-resources?workspace='+f.space.id,'POST',{...officeKindFromFilename('文件.'+format),name:'文件.'+format})).json(),q='?'+new URLSearchParams({workspace:f.space.id,id:resource.id}),threadId=randomUUID(),create={type:'create',threadId,operationId:randomUUID(),base:'',body:'共同评审😀 <script>'};
  expect((await f.call('/office-comment-operation'+q,'POST',create)).status).toBe(200);expect((await(await f.call('/office-comment-operation'+q,'POST',create)).json()).repeated).toBe(true);expect((await f.call('/office-comment-operation'+q,'POST',{...create,body:'重用编号'})).status).toBe(409);
  const first=f.read(resource.id),base=officeCommentRevision(first.comments.get(threadId)!);expect(first.comments.get(threadId)!.messages[0].author).toBe(f.owner.user);first.doc.destroy();
  const reply={type:'reply',threadId,operationId:randomUUID(),base,body:'成员回复'};expect((await f.call('/office-comment-operation'+q,'POST',reply,f.other.cookie)).status).toBe(200);expect((await f.call('/office-comment-operation'+q,'POST',{...reply,operationId:randomUUID()})).status).toBe(409);
  const second=f.read(resource.id),latest=officeCommentRevision(second.comments.get(threadId)!);expect(second.comments.get(threadId)!.messages).toHaveLength(2);second.doc.destroy();expect((await f.call('/office-comment-operation'+q,'POST',{type:'resolve',threadId,operationId:randomUUID(),base:latest})).status).toBe(200);
  const backup=await f.call('/office-backup'+q);expect(backup.status).toBe(200);const bytes=new Uint8Array(await backup.arrayBuffer()),parsed=parseOfficeBackup(bytes),snapshot=new Y.Doc();Y.applyUpdate(snapshot,parsed.state);expect(snapshot.getMap<OfficeCommentThread>('officeComments').get(threadId)!.resolved).toBe(true);snapshot.destroy();
  const restored=await f.call('/office-restore?workspace='+f.space.id+'&operationId='+randomUUID(),'POST',bytes);expect(restored.status).toBe(201);const recovered=f.read((await restored.json()).id);expect(recovered.comments.get(threadId)!.messages[1].authorName).toBe('other');expect(recovered.doc.getMap('officeCommentOperations').size).toBe(0);recovered.doc.destroy();
 }}finally{await f.close();}
},20000);
it('评论鉴权、线程竞争、原始 Yjs 伪造与非法内容均被拒绝，不写入半成品',async()=>{
 const f=await fixture();try{
  const resource=await(await f.call('/office-resources?workspace='+f.space.id,'POST',{...officeKindFromFilename('内容.txt'),name:'内容.txt'})).json(),q='?'+new URLSearchParams({workspace:f.space.id,id:resource.id}),threadId=randomUUID(),input={type:'create',threadId,operationId:randomUUID(),base:'',body:'真实评论'};expect((await f.call('/office-comment-operation'+q,'POST',input)).status).toBe(200);
  const current=f.read(resource.id),thread=current.comments.get(threadId)!,base=officeCommentRevision(thread);current.doc.destroy();const erase={type:'delete',threadId,operationId:randomUUID(),base};expect((await f.call('/office-comment-operation'+q,'POST',erase,f.other.cookie)).status).toBe(403);
  f.cloud.database.db.prepare("INSERT INTO cloud_document_permissions VALUES (?,?,?,'viewer')").run(f.space.id,resource.id,f.other.user);expect((await f.call('/office-comment-operation'+q,'POST',{type:'reply',threadId,base,operationId:randomUUID(),body:'无权回复'},f.other.cookie)).status).toBe(403);
  const forged=f.read(resource.id),vector=Y.encodeStateVector(forged.doc);forged.comments.set(threadId,{...thread,messages:[{...thread.messages[0],author:f.other.user,body:'伪造作者'}]});expect((await f.call('/document'+q,'POST',Y.encodeStateAsUpdate(forged.doc,vector))).status).toBe(400);forged.doc.destroy();
  const before=f.cloud.database.load('cloud:'+f.space.id,resource.id);expect((await f.call('/office-comment-operation'+q,'POST',{...input,operationId:randomUUID(),body:'x'.repeat(2001)})).status).toBe(400);expect(f.cloud.database.load('cloud:'+f.space.id,resource.id)).toEqual(before);
  expect((await f.call('/office-comment-operation'+q,'POST',erase)).status).toBe(200);expect((await(await f.call('/office-comment-operation'+q,'POST',erase)).json()).repeated).toBe(true);
  expect((await f.call('/office-resource'+q,'DELETE')).status).toBe(200);expect((await f.call('/office-comment-operation'+q,'POST',{...input,operationId:randomUUID()})).status).toBe(404);
 }finally{await f.close();}
},10000);
it('评论结构和配额校验不接受嵌套 CRDT 或伪造作者信息',()=>{
 const doc=new Y.Doc();doc.getMap('officeComments').set('fake',new Y.Map());expect(()=>validateOfficeComments(doc)).toThrow('评论内容');doc.destroy();
});

it('评论线程容量耗尽返回可操作错误，整次更新和操作回执回滚',async()=>{
 const f=await fixture();try{
  const resource=await(await f.call('/office-resources?workspace='+f.space.id,'POST',{...officeKindFromFilename('配额.txt'),name:'配额.txt'})).json(),q='?'+new URLSearchParams({workspace:f.space.id,id:resource.id}),threadId=randomUUID(),model=f.read(resource.id),now=Date.now();
  const thread:OfficeCommentThread={id:threadId,resolved:false,updated:now,messages:Array.from({length:100},(_,i)=>({id:'message-'+i,author:f.owner.user,authorName:'owner',body:'评论',created:now}))};model.comments.set(threadId,thread);f.cloud.database.append('cloud:'+f.space.id,resource.id,Y.encodeStateAsUpdate(model.doc));model.doc.destroy();const before=f.cloud.database.load('cloud:'+f.space.id,resource.id);
  const result=await f.call('/office-comment-operation'+q,'POST',{type:'reply',threadId,operationId:randomUUID(),base:officeCommentRevision(thread),body:'第101条'});expect(result.status).toBe(413);expect((await result.json()).error).toContain('评论');expect(f.cloud.database.load('cloud:'+f.space.id,resource.id)).toEqual(before);
 }finally{await f.close();}
},10000);

it('评论文字锚点跟随远端插入、回复保留定位，删除后成为孤立讨论，备份保持锚点',async()=>{
 const {captureOfficeTextAnchor,officeAnchorTarget}=await import('../packages/common/nbstore/src/office-comment-anchor');const f=await fixture();try{
  const resource=await(await f.call('/office-resources?workspace='+f.space.id,'POST',{...officeKindFromFilename('锚点.txt'),name:'锚点.txt'})).json(),q='?'+new URLSearchParams({workspace:f.space.id,id:resource.id}),initial=f.read(resource.id);initial.doc.getText('text').insert(0,'原始正文');f.cloud.database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(initial.doc)),'cloud:'+f.space.id,resource.id);const anchor=captureOfficeTextAnchor(initial.doc.getText('text'),0,initial.doc.getText('text'),2,'原始');initial.doc.destroy();const input={type:'create',threadId:randomUUID(),operationId:randomUUID(),base:'',body:'定位意见',anchor};expect((await f.call('/office-comment-operation'+q,'POST',input)).status).toBe(200);
  const before=f.read(resource.id),thread=before.comments.get(input.threadId)!;expect(officeAnchorTarget(before.doc,thread.anchor!)?.kind).toBe('text');expect((await f.call('/office-comment-operation'+q,'POST',{type:'reply',threadId:thread.id,operationId:randomUUID(),base:officeCommentRevision(thread),body:'回复意见'})).status).toBe(200);before.doc.destroy();const changed=f.read(resource.id);changed.doc.getText('text').insert(0,'远端');const target=officeAnchorTarget(changed.doc,changed.comments.get(thread.id)!.anchor!);expect(target).toMatchObject({kind:'text',start:2,end:4});changed.doc.getText('text').delete(2,2);expect(officeAnchorTarget(changed.doc,anchor)).toBeUndefined();changed.doc.destroy();
  const backup=await f.call('/office-backup'+q),parsed=parseOfficeBackup(new Uint8Array(await backup.arrayBuffer())),restored=new Y.Doc();try{Y.applyUpdate(restored,parsed.state);expect(restored.getMap<OfficeCommentThread>('officeComments').get(thread.id)?.anchor).toEqual(anchor);}finally{restored.destroy();}
  expect((await f.call('/office-comment-operation'+q,'POST',{...input,threadId:randomUUID(),operationId:randomUUID(),anchor:{kind:'cell',sheet:'missing',row:'1-1',column:'missing',label:'A1'}})).status).toBe(409);
 }finally{await f.close();}
});
