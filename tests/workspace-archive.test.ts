import {createOfficeState,createBlankPdf} from '../packages/frontend/apps/web/server/office-create';
import {it,expect,vi} from 'vitest';
import {createServer} from 'node:http';
import {randomUUID} from 'node:crypto';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import * as Y from 'yjs';
import {BlockStore} from '../packages/editor/framework/src/index';
import {markdownParser,writeDocument,readDocument} from '../packages/editor/blocks/src/document-format';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {archiveBinary} from '../packages/frontend/apps/web/server/workspace-archive';
import {prepareWorkspaceArchive} from '../packages/frontend/apps/web/server/workspace-archive-model';
import type {CloudWorkspaceArchive} from '../packages/common/nbstore/src/workspace-archive';
import {OfficeParserPool} from '../packages/frontend/apps/web/server/office-parser-process';

const parse=(bytes:Uint8Array)=>JSON.parse(Buffer.from(bytes).toString()) as CloudWorkspaceArchive;
const encode=(value:unknown)=>Buffer.from(JSON.stringify(value));
function normal(workspace:string){const doc=new Y.Doc(),store=new BlockStore(doc);store.initialize();store.ensureNote();writeDocument(store,markdownParser.parse('[内部页面](#workspace='+workspace+'&doc=page)'));const state=Y.encodeStateAsUpdate(doc);store.destroy();doc.destroy();return state;}
function archive():CloudWorkspaceArchive{
 const workspace='source',meta=new Y.Doc();meta.getMap('documents').set('page',{id:'page',title:'正文',createdAt:1,folderId:null});meta.getMap('trash').set('recycled',{meta:{id:'recycled',title:'回收正文',createdAt:1,folderId:null},deletedAt:2,collectionIds:[]});meta.getMap('uiPreferences').set('private','not copied');
 const text=new Y.Doc();text.getText('text').insert(0,'协作文本 🐈');const state=Y.encodeStateAsUpdate(text);text.destroy();
 const result:CloudWorkspaceArchive={type:'mydoc-team-workspace',version:1,exportedAt:1,workspace:{id:workspace,name:'完整备份'},documents:[{id:'workspace:source',state:archiveBinary(Y.encodeStateAsUpdate(meta))},{id:'doc:page',state:archiveBinary(normal(workspace))},{id:'doc:recycled',state:archiveBinary(normal(workspace))},{id:'doc:txt',state:archiveBinary(state)}],office:[{id:'doc:txt',kind:'text',name:'原件.txt',mime:'text/plain',sourceFormat:'txt',originalAssetId:'original',modelVersion:1,created:1,updated:1,deleted:0}],assets:[{id:'original',name:'原件.txt',mime:'text/plain',doc:'doc:txt',state:archiveBinary(Buffer.from('\ufeff原始文本\r\n'))}],deleted:['doc:purged']};meta.destroy();return result;
}
async function fixture(){
 const dir=mkdtempSync(join(tmpdir(),'mydoc-workspace-archive-')),cloud=new CloudServer(join(dir,'db.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const registered=await fetch(root+'/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:'archive@example.com',name:'Owner',password:'test-password-123'})}),cookie=registered.headers.get('set-cookie')!.split(';')[0],user=(await registered.json()).user.id as string;
 const call=(path:string,method='GET',value?:unknown,credential=cookie)=>fetch(root+path,{method,headers:{Cookie:credential,'Content-Type':'application/json'},body:value===undefined?undefined:JSON.stringify(value)});
 return {cloud,call,user,cookie,close:async()=>{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}};
}
it('完整模型校验保留正文、回收、原件并重映射内部引用，隔离个人设置',()=>{
 const source=archive(),prepared=parse(prepareWorkspaceArchive(encode(source),'target'));expect(prepared.workspace.id).toBe('target');expect(prepared.assets).toEqual(source.assets);expect(prepared.deleted).toEqual(source.deleted);
 const meta=new Y.Doc();Y.applyUpdate(meta,Buffer.from(prepared.documents.find(row=>row.id==='workspace:target')!.state.data,'base64'));expect(meta.getMap('trash').has('recycled')).toBe(true);expect(meta.getMap('uiPreferences').size).toBe(0);meta.destroy();
 const doc=new Y.Doc();Y.applyUpdate(doc,Buffer.from(prepared.documents.find(row=>row.id==='doc:page')!.state.data,'base64'));const store=new BlockStore(doc);let reference:unknown;readDocument(store).descendants(node=>{if(node.type.name==='doc_reference')reference=node.attrs.workspaceId;});expect(reference).toBe('target');store.destroy();doc.destroy();expect(source.workspace.id).toBe('source');
});
it('拒绝损坏校验值、缺失正文或原件、循环目录、未知根、错误类型和越界备份',()=>{
 const mutate=(change:(a:CloudWorkspaceArchive)=>void)=>{const value=archive();change(value);return()=>prepareWorkspaceArchive(encode(value),'target');};
 expect(mutate(a=>{a.assets[0].state.sha256='0'.repeat(64);})).toThrow('校验');expect(mutate(a=>{a.documents=a.documents.filter(row=>row.id!=='doc:page');})).toThrow('缺失');expect(mutate(a=>{a.assets=[];})).toThrow('原件');expect(mutate(a=>{a.office[0].sourceFormat='xlsx';})).toThrow();
 for(const root of ['folders','unknown','collections'])expect(mutate(a=>{const doc=new Y.Doc();Y.applyUpdate(doc,Buffer.from(a.documents[0].state.data,'base64'));doc.getMap(root).set('bad',root==='folders'?{id:'bad',name:'bad',parentId:'bad',createdAt:1}:{});a.documents[0].state=archiveBinary(Y.encodeStateAsUpdate(doc));doc.destroy();})).toThrow();
 expect(()=>prepareWorkspaceArchive(new Uint8Array(50*1024*1024+1))).toThrow('50 MB');
});
it('已删除办公资源连同原件和墓碑保留，孤立永久删除正文被拒绝',()=>{
 const a=archive();a.office[0].deleted=1;a.deleted.push('doc:txt');expect(parse(prepareWorkspaceArchive(encode(a))).office[0].deleted).toBe(1);a.deleted.push('doc:page');expect(()=>prepareWorkspaceArchive(encode(a))).toThrow();
});
it('受限进程支持团队备份任务并维持输入上限和失败恢复',async()=>{
 const pool=new OfficeParserPool();try{const result=parse(await pool.prepareWorkspaceArchive(encode(archive()),'new-space'));expect(result.workspace.id).toBe('new-space');await expect(pool.prepareWorkspaceArchive(Buffer.from('{}'))).rejects.toMatchObject({status:400});expect(pool.activeCount).toBe(0);}finally{pool.close();}
});
it('完整恢复隔离命名空间与成员，原件字节一致；重试、并发、冲突与删除后重放有确定结果',async()=>{
 const f=await fixture();try{
  const a=archive(),op=randomUUID(),path='/workspace-restore?operationId='+op;
  const [first,second]=await Promise.all([f.call(path,'POST',a),f.call(path,'POST',a)]);expect(first.status).toBe(201);expect(second.status).toBe(201);const one=await first.json(),two=await second.json();expect(one.id).toBe(two.id);expect(one.id).not.toBe('source');expect(one.role).toBe('owner');
  expect(f.cloud.database.db.prepare('SELECT user,role FROM cloud_members WHERE workspace=?').all(one.id)).toEqual([{user:f.user,role:'owner'}]);expect(f.cloud.database.db.prepare('SELECT count(*) n FROM cloud_document_permissions').get()).toEqual({n:0});
  const resource=f.cloud.database.db.prepare('SELECT generation,created_by FROM cloud_office_resources WHERE workspace=?').get(one.id) as {generation:string;created_by:string};expect(resource.generation).toHaveLength(36);expect(resource.created_by).toBe(f.user);
  const original=await f.call('/office-original?workspace='+one.id+'&id=doc:txt');expect(original.status).toBe(200);expect(Buffer.from(await original.arrayBuffer())).toEqual(Buffer.from(a.assets[0].state.data,'base64'));
  const backup=await f.call('/workspace-backup?workspace='+one.id);expect(backup.status).toBe(200);expect(backup.headers.get('content-disposition')).toContain('.mydoc-team.json');const exported=await backup.json() as CloudWorkspaceArchive;expect(exported.office).toHaveLength(1);expect(exported.assets[0].state).toEqual(a.assets[0].state);expect(exported.documents.map(row=>row.id)).toContain('doc:recycled');
  expect((await f.call(path,'POST',{...a,exportedAt:2})).status).toBe(409);
  expect((await f.call('/workspace?workspace='+one.id,'DELETE')).status).toBe(200);expect((await f.call(path,'POST',a)).status).toBe(410);expect(f.cloud.database.db.prepare('SELECT count(*) n FROM cloud_workspace_restores').get()).toEqual({n:1});
 }finally{await f.close();}
},20000);
it('恢复事务任一步失败无半成品；校验失败不写入；普通成员不能完整导出',async()=>{
 const f=await fixture();try{
  const counts=()=>['cloud_workspaces','cloud_members','cloud_workspace_restores','cloud_office_resources','documents','assets','asset_files','cloud_asset_owners'].map(table=>f.cloud.database.db.prepare('SELECT count(*) n FROM '+table).get()),before=counts();
  f.cloud.database.db.exec("CREATE TRIGGER fail_archive BEFORE INSERT ON cloud_workspace_restores BEGIN SELECT RAISE(ABORT,'restore failed'); END");expect((await f.call('/workspace-restore?operationId='+randomUUID(),'POST',archive())).status).toBe(400);expect(counts()).toEqual(before);f.cloud.database.db.exec('DROP TRIGGER fail_archive');
  const bad=archive();bad.assets=[];expect((await f.call('/workspace-restore?operationId='+randomUUID(),'POST',bad)).status).toBe(400);expect(counts()).toEqual(before);
  const restored=await(await f.call('/workspace-restore?operationId='+randomUUID(),'POST',archive())).json();f.cloud.database.db.prepare("UPDATE cloud_members SET role='member' WHERE workspace=?").run(restored.id);expect((await f.call('/workspace-backup?workspace='+restored.id)).status).toBe(403);expect((await f.call('/workspace-backup?workspace='+restored.id,'GET',undefined,'')).status).toBe(401);
 }finally{await f.close();}
},20000);
it('异步备份使用固定快照，撤权或会话注销后不返回内容',async()=>{
 const f=await fixture();try{
  const restored=await(await f.call('/workspace-restore?operationId='+randomUUID(),'POST',archive())).json();const pool=(f.cloud as unknown as {officeParser:OfficeParserPool}).officeParser;
  vi.spyOn(pool,'prepareWorkspaceArchive').mockImplementationOnce(async bytes=>{const doc=new Y.Doc();Y.applyUpdate(doc,f.cloud.database.load('cloud:'+restored.id,'doc:txt')!);doc.getText('text').insert(0,'后续改动');f.cloud.database.append('cloud:'+restored.id,'doc:txt',Y.encodeStateAsUpdate(doc));doc.destroy();return prepareWorkspaceArchive(bytes);});
  const fixed=await(await f.call('/workspace-backup?workspace='+restored.id)).json() as CloudWorkspaceArchive,copy=new Y.Doc();Y.applyUpdate(copy,Buffer.from(fixed.documents.find(row=>row.id==='doc:txt')!.state.data,'base64'));expect(copy.getText('text').toString()).toBe('协作文本 🐈');copy.destroy();
  const original=pool.prepareWorkspaceArchive.bind(pool);let captured:Uint8Array|undefined;
  vi.spyOn(pool,'prepareWorkspaceArchive').mockImplementationOnce(async bytes=>{captured=bytes;f.cloud.database.db.prepare("UPDATE cloud_members SET role='member' WHERE workspace=?").run(restored.id);return prepareWorkspaceArchive(bytes);});
  expect((await f.call('/workspace-backup?workspace='+restored.id)).status).toBe(403);expect(captured).toBeTruthy();
  f.cloud.database.db.prepare("UPDATE cloud_members SET role='owner' WHERE workspace=?").run(restored.id);
  vi.spyOn(pool,'prepareWorkspaceArchive').mockImplementationOnce(async bytes=>{f.cloud.database.db.prepare('DELETE FROM cloud_sessions').run();return original(bytes);});expect((await f.call('/workspace-backup?workspace='+restored.id)).status).toBe(401);
 }finally{await f.close();}
},20000);


it('全部办公类型和普通附件自包含往返，保留类型、备注、文本设置和删除原件',()=>{
 const a=archive();for(const [kind,sourceFormat,mime]of [['word','docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document'],['spreadsheet','xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],['spreadsheet','csv','text/csv'],['spreadsheet','tsv','text/tab-separated-values'],['presentation','pptx','application/vnd.openxmlformats-officedocument.presentationml.presentation'],['pdf','pdf','application/pdf']] as const){
  const id='doc:'+sourceFormat;a.documents.push({id,state:archiveBinary(createOfficeState(kind))});const originalAssetId=kind==='pdf'?'pdf-original':null;if(originalAssetId)a.assets.push({id:originalAssetId,name:'原件.pdf',mime,doc:id,state:archiveBinary(createBlankPdf())});a.office.push({id,kind,sourceFormat,mime,name:'文件.'+sourceFormat,originalAssetId,modelVersion:1,created:1,updated:2,deleted:0});
 }
 a.assets.push({id:'unbound',name:'二进制附件.bin',mime:'application/octet-stream',doc:null,state:archiveBinary(new Uint8Array([0,255,128]))});const result=parse(prepareWorkspaceArchive(encode(a),'target'));expect(result.office).toEqual(a.office);expect(result.office).toHaveLength(7);expect(result.assets).toEqual(a.assets);expect(result.documents).toHaveLength(a.documents.length);
 for(const row of result.documents){expect(row.state.sha256).toHaveLength(64);const doc=new Y.Doc();Y.applyUpdate(doc,Buffer.from(row.state.data,'base64'));doc.destroy();}
});

it('统一文件目录备份保留办公文件夹、收藏、标签和集合，拒绝损坏组织字段',()=>{
 const value=archive(),meta=new Y.Doc();Y.applyUpdate(meta,Buffer.from(value.documents[0].state.data,'base64'));meta.getMap('folders').set('project',{id:'project',name:'项目',parentId:null,createdAt:1});meta.getMap('tags').set('tag',{id:'tag',name:'标签',createdAt:1});meta.getMap('officeOrganization').set('doc:txt',{folderId:'project',favorite:true,tagIds:['tag']});meta.getMap('collections').set('all',{id:'all',name:'统一集合',createdAt:1,mode:'manual',documentIds:['page','doc:txt'],rules:{tagIds:[],tagMatch:'all',keyword:'',favorite:'any'}});value.documents[0].state=archiveBinary(Y.encodeStateAsUpdate(meta));
 const copy=parse(prepareWorkspaceArchive(encode(value),'target')),restored=new Y.Doc();Y.applyUpdate(restored,Buffer.from(copy.documents.find(row=>row.id==='workspace:target')!.state.data,'base64'));expect(restored.getMap('officeOrganization').get('doc:txt')).toEqual({folderId:'project',favorite:true,tagIds:['tag']});restored.destroy();
 for(const invalid of [{folderId:'missing'},{favorite:'yes'},{tagIds:['missing']},{title:'不允许覆盖资源名称'}]){meta.getMap('officeOrganization').set('doc:txt',invalid);value.documents[0].state=archiveBinary(Y.encodeStateAsUpdate(meta));expect(()=>prepareWorkspaceArchive(encode(value))).toThrow('组织信息无效');}meta.destroy();
});
