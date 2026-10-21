import {PersistentDocument,type DocumentStorage} from '../packages/common/nbstore/src/index';
import {CLOUD_SCHEMA_VERSION} from '../packages/frontend/apps/web/server/schema';
import {it,expect} from 'vitest';
import {createServer} from 'node:http';
import {mkdtempSync,rmSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import Database from 'better-sqlite3';
import * as Y from 'yjs';
import {SQLiteProjectDatabase} from '../packages/common/nbstore/server/database';
import {DatabaseBackups} from '../packages/common/nbstore/server/backups';
import {storageHandler} from '../packages/common/nbstore/server/http';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {upgradeCloudDatabase,LOCAL_OFFICE_BACKUP} from '../packages/frontend/apps/web/server/schema';
import {createOfficeState,createBlankPdf} from '../packages/frontend/apps/web/server/office-create';
import {exportDocxWord} from '../packages/frontend/apps/web/server/office-word-export';
import {exportPptxPresentation} from '../packages/frontend/apps/web/server/office-presentation-export';
import {exportXlsxText} from '../packages/frontend/apps/web/server/office-spreadsheet-export';
import {officeKindFromFilename} from '../packages/common/nbstore/src/office';
import {cellRevision} from '../packages/common/nbstore/src/spreadsheet-operation';
async function start(path:string){
 const db=new SQLiteProjectDatabase(path,p=>new Database(p),upgradeCloudDatabase,{backupExtension:LOCAL_OFFICE_BACKUP}),engine=new CloudServer(path,{database:db});
 const server=createServer((req,res)=>{const scope=String(req.headers['x-test-account']??'main');void storageHandler(db,false,undefined,scope==='main'?undefined:scope,(request,response,owner,check)=>engine.localOffice(request,response,owner,check))(req,res);});await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(route:string,method='GET',body?:unknown,scope='main',generation?:string)=>fetch(root+route,{method,headers:{'x-test-account':scope,...(scope.startsWith('user:')?{'x-mydoc-account':scope.slice(5)}:{}),...(generation?{'x-mydoc-generation':generation}:{})},body:body instanceof Uint8Array?body as BodyInit:body===undefined?undefined:JSON.stringify(body)});
 const register=(scope='main')=>{const doc=new Y.Doc();doc.getMap('workspaces').set('local',{id:'local',name:'个人',createdAt:1});doc.getMap('workspaces').set('other',{id:'other',name:'另一个个人工作区',createdAt:1});db.append(scope,'registry:local',Y.encodeStateAsUpdate(doc));doc.destroy();};
 return {db,engine,call,register,close:async()=>{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));engine.close();db.close();}};
}
const query=(id?:string,workspace='local')=>'?'+new URLSearchParams({workspace,...(id?{id}:{})});
it('个人办公全部类型创建/导入/编辑/原生导出，关闭服务和 SQLite 后原件与修改保持',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-local-office-')),path=join(dir,'db.sqlite');let f=await start(path);try{
  f.register();const created=new Map<string,string>(),originals=new Map<string,Buffer>();
  const word=new Y.Doc();word.getArray('paragraphs').push([new Y.Text('导入正文')]);const wordBytes=exportDocxWord(Y.encodeStateAsUpdate(word));word.destroy();
  const samples:Record<string,Uint8Array>={docx:wordBytes,xlsx:exportXlsxText(createOfficeState('spreadsheet')),pptx:exportPptxPresentation(createOfficeState('presentation')),pdf:createBlankPdf(),csv:Buffer.from('名称,编号\n样本,001\n'),tsv:Buffer.from('名称\t编号\n样本\t001\n'),txt:Buffer.from('\ufeff本地原件\r\n')};
  for(const extension of Object.keys(samples)){
   const format=officeKindFromFilename('新建.'+extension)!,response=await f.call('/office/office-resources'+query(),'POST',{...format,name:'新建.'+extension});expect(response.status).toBe(201);created.set(extension,(await response.json()).id);
   const imported=await f.call('/office/'+(extension==='txt'?'office-text-import':'office-import')+query()+'&name='+encodeURIComponent('导入.'+extension),'POST',samples[extension]);expect(imported.status).toBe(201);const resource=await imported.json();originals.set(resource.id,Buffer.from(samples[extension]));
   expect(Buffer.from(await(await f.call('/office/office-original'+query(resource.id))).arrayBuffer())).toEqual(Buffer.from(samples[extension]));
  }
  for(const extension of ['txt','docx','pptx']){const id=created.get(extension)!,doc=new Y.Doc();Y.applyUpdate(doc,new Uint8Array(await(await f.call('/office/document'+query(id))).arrayBuffer()));const vector=Y.encodeStateVector(doc);if(extension==='txt')doc.getText('text').insert(0,'重启后本地文本😀');if(extension==='docx')doc.getArray<Y.Text>('paragraphs').get(0).insert(0,'重启后Word正文');if(extension==='pptx')doc.getArray<Y.Map<unknown>>('objects').get(0).set('text','重启后演示正文');expect((await f.call('/office/document'+query(id),'POST',Y.encodeStateAsUpdate(doc,vector))).status).toBe(200);doc.destroy();}
  const sheetId=created.get('xlsx')!,sheet=new Y.Doc();Y.applyUpdate(sheet,new Uint8Array(await(await f.call('/office/document'+query(sheetId))).arrayBuffer()));const base=cellRevision(sheet.getArray('sheet'),0,0);sheet.destroy();expect((await f.call('/office/office-sheet-operation'+query(sheetId),'POST',{type:'setCell',row:0,column:0,value:'00123',cellType:'text',base,operationId:crypto.randomUUID()})).status).toBe(200);
  expect((await f.call('/office/office-pdf-operation'+query(created.get('pdf')!),'POST',{type:'create',annotationId:crypto.randomUUID(),base:'',operationId:crypto.randomUUID(),annotation:{type:'comment',page:1,x:50,y:50,content:'本地PDF批注'}})).status).toBe(200);
  expect(f.db.db.prepare('SELECT count(*) n FROM cloud_users').get()).toEqual({n:0});expect(f.db.db.prepare('SELECT count(*) n FROM cloud_workspaces').get()).toEqual({n:0});
  await f.close();f=await start(path);expect((await(await f.call('/office/office-resources'+query())).json())).toHaveLength(14);
  for(const [id,bytes]of originals)expect(Buffer.from(await(await f.call('/office/office-original'+query(id))).arrayBuffer())).toEqual(bytes);
  for(const [extension,format,scope]of [['docx','docx','paragraphs'],['xlsx','xlsx','text-grid'],['pptx','pptx','slides'],['csv','csv','raw-text'],['tsv','tsv','raw-text']] as const){const response=await f.call('/office/office-export'+query(created.get(extension))+'&format='+format+'&scope='+scope);expect(response.status).toBe(200);const bytes=new Uint8Array(await response.arrayBuffer());if(['docx','xlsx','pptx'].includes(extension))expect([...bytes.subarray(0,2)]).toEqual([80,75]);}
  const pdf=await f.call('/office/office-pdf-export'+query(created.get('pdf')));expect(pdf.status).toBe(200);expect(Buffer.from(await pdf.arrayBuffer()).subarray(0,5).toString()).toBe('%PDF-');
  const text=new Y.Doc();Y.applyUpdate(text,new Uint8Array(await(await f.call('/office/document'+query(created.get('txt')))).arrayBuffer()));expect(text.getText('text').toString()).toBe('重启后本地文本😀');text.destroy();
  const backup=new DatabaseBackups(f.db,p=>new Database(p));const saved=await backup.create();expect((await backup.preview(saved.id)).id).toBe(saved.id);const extra=await(await f.call('/office/office-resources'+query(),'POST',{...officeKindFromFilename('之后.txt'),name:'之后.txt'})).json();await backup.restore(saved.id);expect((await(await f.call('/office/office-resources'+query())).json()).some((r:{id:string})=>r.id===extra.id)).toBe(false);for(const [id,bytes]of originals)expect(Buffer.from(await(await f.call('/office/office-original'+query(id))).arrayBuffer())).toEqual(bytes);
 }finally{await f.close();rmSync(dir,{recursive:true,force:true});}
},30000);
it('个人办公按存储账户/工作区隔离；不能穿透本地边界，旧代际和恢复期间请求拒绝',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-local-boundary-')),f=await start(join(dir,'db.sqlite'));try{f.register('user:alice');f.register('user:bob');const response=await f.call('/office/office-resources'+query(),'POST',{...officeKindFromFilename('个人.txt'),name:'个人.txt'},'user:alice');expect(response.status).toBe(201);const resource=await response.json();expect(await(await f.call('/office/office-resources'+query(),'GET',undefined,'user:bob')).json()).toEqual([]);expect((await f.call('/office/document'+query(resource.id),'GET',undefined,'user:bob')).status).toBe(404);expect((await f.call('/office/document'+query(resource.id,'other'),'GET',undefined,'user:alice')).status).toBe(404);expect((await f.call('/office/admin/users'+query(),'GET',undefined,'user:alice')).status).toBe(404);expect((await f.call('/office/office-resources'+query(undefined,'missing'),'GET',undefined,'user:alice')).status).toBe(404);
  const generation=f.db.generation;f.db.generation=crypto.randomUUID();expect((await f.call('/office/office-resources'+query(),'GET',undefined,'user:alice',generation)).status).toBe(500);f.db.restoring=true;expect((await f.call('/office/office-resources'+query(),'GET',undefined,'user:alice')).status).toBe(500);f.db.restoring=false;
 }finally{await f.close();rmSync(dir,{recursive:true,force:true});}
});

it('旧版个人备份可恢复，清除其后办公资源；个人工作区删除仅清理所属办公命名空间',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-local-legacy-')),f=await start(join(dir,'db.sqlite'));try{
  f.register();const legacy=new SQLiteProjectDatabase(join(dir,'legacy.sqlite'),p=>new Database(p));legacy.append('main','registry:local',f.db.load('main','registry:local')!);const oldManager=new DatabaseBackups(legacy,p=>new Database(p));const old=await oldManager.create();oldManager.stop();legacy.close();
  const create=async(workspace:string)=>(await(await f.call('/office/office-resources'+query(undefined,workspace),'POST',{...officeKindFromFilename('数据.txt'),name:'数据.txt'})).json()).id;
  await create('local');const manager=new DatabaseBackups(f.db,p=>new Database(p));await manager.restore(old.id);expect(await(await f.call('/office/office-resources'+query())).json()).toEqual([]);expect(f.db.db.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:CLOUD_SCHEMA_VERSION});const source=await create('local'),other=await create('other');
  expect((await f.call('/office/office-purge'+query(),'DELETE')).status).toBe(200);expect((await f.call('/office/document'+query(source))).status).toBe(404);expect((await f.call('/office/document'+query(other,'other'))).status).toBe(200);expect(f.db.db.prepare('SELECT count(*) n FROM cloud_office_resources').get()).toEqual({n:1});
 }finally{await f.close();rmSync(dir,{recursive:true,force:true});}
});

it('本地后台读取成功不能清除尚未保存的写入错误，显式重试后才恢复已保存状态',async()=>{
 let fail=true,notify:(error:unknown)=>void=()=>{};const storage:DocumentStorage={load:async()=>null,append:async()=>{if(fail)throw new Error('写入失败');},remove:async()=>{},close(){},connect(_id,_doc,onError){notify=onError;return()=>{};}};
 const doc=new Y.Doc(),persistence=new PersistentDocument('doc:local-test',doc,storage);await persistence.open();doc.getText('text').insert(0,'必须保留的输入');await expect(persistence.flush()).rejects.toThrow('写入失败');notify(undefined);expect(persistence.state).toBe('error');await expect(persistence.flush()).rejects.toThrow('写入失败');fail=false;await persistence.retry();expect(persistence.state).toBe('saved');await persistence.destroy();doc.destroy();
});

it('个人办公整包跨工作区恢复包含原件、删除状态和命名版本；损坏包不产生半成品',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-local-portable-')),f=await start(join(dir,'db.sqlite'));try{
  f.register();const imported=await f.call('/office/office-text-import'+query()+'&name=source.txt','POST',Buffer.from('\ufeff原件\r\n'));expect(imported.status).toBe(201);const resource=await imported.json();
  expect((await f.call('/office/office-versions'+query(resource.id),'POST',{name:'迁移前',operationId:crypto.randomUUID()})).status).toBe(201);
  const archived=await f.call('/office/office-workspace-backup'+query());expect(archived.status).toBe(200);const archive=await archived.json();expect(archive.office).toHaveLength(1);expect(archive.versions).toHaveLength(1);expect(archive.assets).toHaveLength(1);
  const bad=structuredClone(archive);bad.assets[0].state.sha256='0'.repeat(64);expect((await f.call('/office/office-workspace-restore'+query(undefined,'other'),'POST',bad)).status).toBe(400);expect(await(await f.call('/office/office-resources'+query(undefined,'other'))).json()).toEqual([]);
  const restored=await f.call('/office/office-workspace-restore'+query(undefined,'other'),'POST',archive);expect(restored.status).toBe(200);const list=await(await f.call('/office/office-resources'+query(undefined,'other'))).json();expect(list).toHaveLength(1);expect(list[0].id).toBe(resource.id);expect(list[0].generation).not.toBe(resource.generation);
  expect(Buffer.from(await(await f.call('/office/office-original'+query(resource.id,'other'))).arrayBuffer())).toEqual(Buffer.from('\ufeff原件\r\n'));
  expect(await(await f.call('/office/office-versions'+query(resource.id,'other'))).json()).toHaveLength(1);
  expect((await f.call('/office/office-workspace-restore'+query(undefined,'other'),'POST',archive)).status).toBe(409);
  expect(await(await f.call('/office/office-resources'+query())).json()).toHaveLength(1);
  // A failing write rolls back the state, original ownership, resources and versions together.
  await f.call('/office/office-purge'+query(undefined,'other'),'DELETE');f.db.db.exec("CREATE TEMP TRIGGER reject_portable BEFORE INSERT ON cloud_office_versions BEGIN SELECT RAISE(ABORT,'test rollback'); END");
  expect((await f.call('/office/office-workspace-restore'+query(undefined,'other'),'POST',archive)).status).toBe(400);expect(await(await f.call('/office/office-resources'+query(undefined,'other'))).json()).toEqual([]);
  expect((f.db.db.prepare('SELECT count(*) n FROM cloud_asset_owners').get() as {n:number}).n).toBe(1);expect((f.db.db.prepare('SELECT count(*) n FROM cloud_office_versions').get() as {n:number}).n).toBe(1);
 }finally{await f.close();rmSync(dir,{recursive:true,force:true});}
},30000);

it('个人 v2 整包恢复将目录/笔记/办公内容/版本/原件一次提交，响应重试不重复，失败不留下半成品',async()=>{
 const {archiveBinary}=await import('../packages/frontend/apps/web/server/workspace-archive');const {BlockStore}=await import('../packages/editor/framework/src/index');const {writeMarkdown}=await import('../packages/editor/blocks/src/document-format');const dir=mkdtempSync(join(tmpdir(),'mydoc-personal-atomic-')),f=await start(join(dir,'db.sqlite'));try{
  f.register();const resource=await(await f.call('/office/office-text-import'+query()+'&name=atomic.txt','POST',Buffer.from('原件内容'))).json();await f.call('/office/office-versions'+query(resource.id),'POST',{name:'原稿',operationId:crypto.randomUUID()});const office=await(await f.call('/office/office-workspace-backup'+query())).json(),note=new Y.Doc(),store=new BlockStore(note);store.initialize();store.ensureNote();writeMarkdown(store,'普通文档与办公文件一起恢复');const noteUpdate=Buffer.from(Y.encodeStateAsUpdate(note)).toString('base64');store.destroy();note.destroy();
  const backup={type:'mydoc-workspace',version:2,workspaceId:'local',name:'完整个人空间',exportedAt:Date.now(),office,officeOrganization:[{id:resource.id,folderId:'folder-one',favorite:true,tagIds:['tag-one']}],documents:[{meta:{id:'source-note',title:'普通笔记',createdAt:1,folderId:'folder-one'},update:noteUpdate}],folders:[{id:'folder-one',name:'资料',parentId:null,createdAt:1}],tags:[{id:'tag-one',name:'标签',color:'blue',createdAt:1}],collections:[],assets:[]};
  const target=crypto.randomUUID(),operation=crypto.randomUUID(),path='/office/office-workspace-full-restore?'+new URLSearchParams({workspace:target,operationId:operation});const first=await f.call(path,'POST',backup);expect(first.status).toBe(201);const result=await first.json();expect(result.id).toBe(target);expect((await(await f.call(path,'POST',backup)).json()).repeated).toBe(true);
  const registry=new Y.Doc();Y.applyUpdate(registry,f.db.load('main','registry:local')!);expect(registry.getMap('workspaces').get(target)).toMatchObject({name:'完整个人空间（恢复）'});registry.destroy();const meta=new Y.Doc();Y.applyUpdate(meta,f.db.load('main','workspace:'+target)!);const restored=[...meta.getMap<{id:string;title:string}>('documents').values()][0];expect(restored.id).not.toBe('source-note');expect(f.db.load('main','doc:'+restored.id)).toBeTruthy();expect(meta.getMap('officeOrganization').get(resource.id)).toMatchObject({favorite:true,folderId:'folder-one'});meta.destroy();expect(await(await f.call('/office/office-resources'+query(undefined,target))).json()).toHaveLength(1);expect(Buffer.from(await(await f.call('/office/office-original'+query(resource.id,target))).arrayBuffer()).toString()).toBe('原件内容');expect(f.db.db.prepare('SELECT count(*) n FROM cloud_workspaces').get()).toEqual({n:0});
  const failed=crypto.randomUUID(),beforeRegistry=Buffer.from(f.db.load('main','registry:local')!),count=(table:string)=>(f.db.db.prepare('SELECT count(*) n FROM '+table).get() as {n:number}).n,tables=['documents','assets','cloud_office_resources','cloud_office_versions','cloud_asset_owners'],before=tables.map(count);f.db.db.exec("CREATE TEMP TRIGGER reject_atomic_registry BEFORE UPDATE ON documents WHEN NEW.id='registry:local' BEGIN SELECT RAISE(ABORT,'registry failure'); END");expect((await f.call('/office/office-workspace-full-restore?'+new URLSearchParams({workspace:failed,operationId:crypto.randomUUID()}),'POST',backup)).status).toBe(400);expect(tables.map(count)).toEqual(before);expect(Buffer.from(f.db.load('main','registry:local')!)).toEqual(beforeRegistry);expect(f.db.load('main','workspace:'+failed)).toBeNull();
 }finally{await f.close();rmSync(dir,{recursive:true,force:true});}
},30000);

it('个人物理备份不携带导出产物；协作版本 11 的旧个人备份（无导出表）仍可恢复，重启后补齐版本',async()=>{
 const {readFileSync,writeFileSync}=await import('node:fs');const {createHash}=await import('node:crypto');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-local-exports-')),path=join(dir,'db.sqlite');let f=await start(path);try{
  f.register();const resource=await(await f.call('/office/office-resources'+query(),'POST',{...officeKindFromFilename('方案.docx'),name:'方案.docx'})).json();
  expect((await f.call('/office/office-export'+query(resource.id)+'&format=docx&scope=paragraphs')).status).toBe(200);expect(await(await f.call('/office/office-exports'+query(resource.id))).json()).toHaveLength(1);
  const manager=new DatabaseBackups(f.db,p=>new Database(p)),saved=await manager.create(),file=join(manager.status().settings.directory,saved.id+'.sqlite');
  const snapshot=new Database(file);try{expect(snapshot.prepare('SELECT count(*) n FROM cloud_office_exports').get()).toEqual({n:0});}finally{snapshot.close();}
  // Turn the fresh snapshot into a version-11 personal backup: drop tables added later and their version records.
  const legacy=new Database(file);legacy.exec('DROP TABLE cloud_office_exports;DROP TABLE cloud_notifications;DROP TABLE cloud_office_activity;DELETE FROM cloud_migrations WHERE version>=12');legacy.close();const bytes=readFileSync(file),manifest=join(manager.status().settings.directory,saved.id+'.json');writeFileSync(manifest,JSON.stringify({...JSON.parse(readFileSync(manifest,'utf8')),bytes:bytes.length,sha256:createHash('sha256').update(bytes).digest('hex')}));
  await manager.restore(saved.id);manager.stop();expect(await(await f.call('/office/office-resources'+query())).json()).toHaveLength(1);expect(f.db.db.prepare('SELECT count(*) n FROM cloud_office_exports').get()).toEqual({n:0});
  await f.close();f=await start(path);expect(f.db.db.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:CLOUD_SCHEMA_VERSION});expect(f.db.db.prepare('SELECT count(*) n FROM cloud_notifications').get()).toEqual({n:0});expect((await f.call('/office/office-export'+query(resource.id)+'&format=docx&scope=paragraphs')).status).toBe(200);expect(await(await f.call('/office/office-exports'+query(resource.id))).json()).toHaveLength(1);
 }finally{await f.close();rmSync(dir,{recursive:true,force:true});}
},30000);

it('个人物理备份不携带通知与编辑记录；协作版本 12 的旧个人备份仍可恢复并在重启后升级',async()=>{
 const {readFileSync,writeFileSync}=await import('node:fs');const {createHash}=await import('node:crypto');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-local-collab-')),path=join(dir,'db.sqlite');let f=await start(path);try{
  f.register();const resource=await(await f.call('/office/office-resources'+query(),'POST',{...officeKindFromFilename('计划.docx'),name:'计划.docx'})).json();
  f.db.db.prepare("INSERT INTO cloud_notifications(id,user,workspace,resource,kind,actor,actor_name,thread,excerpt,created) VALUES ('n','u','w',?,'mention','a','甲','t','摘要',1)").run(resource.id);f.db.db.prepare("INSERT INTO cloud_office_activity(workspace,resource,user,user_name,started,last) VALUES ('w',?,'u','甲',1,2)").run(resource.id);
  const manager=new DatabaseBackups(f.db,p=>new Database(p)),saved=await manager.create(),file=join(manager.status().settings.directory,saved.id+'.sqlite');
  const snapshot=new Database(file);try{for(const table of ['cloud_notifications','cloud_office_activity'])expect(snapshot.prepare(`SELECT count(*) n FROM ${table}`).get()).toEqual({n:0});}finally{snapshot.close();}
  expect(f.db.db.prepare('SELECT count(*) n FROM cloud_notifications').get()).toEqual({n:1});
  const legacy=new Database(file);legacy.exec('DROP TABLE cloud_notifications;DROP TABLE cloud_office_activity;DELETE FROM cloud_migrations WHERE version>=13');legacy.close();const bytes=readFileSync(file),manifest=join(manager.status().settings.directory,saved.id+'.json');writeFileSync(manifest,JSON.stringify({...JSON.parse(readFileSync(manifest,'utf8')),bytes:bytes.length,sha256:createHash('sha256').update(bytes).digest('hex')}));
  await manager.restore(saved.id);manager.stop();expect(await(await f.call('/office/office-resources'+query())).json()).toHaveLength(1);
  await f.close();f=await start(path);expect(f.db.db.prepare('SELECT max(version) v FROM cloud_migrations').get()).toEqual({v:CLOUD_SCHEMA_VERSION});expect(f.db.db.prepare('SELECT count(*) n FROM cloud_office_activity').get()).toEqual({n:0});
 }finally{await f.close();rmSync(dir,{recursive:true,force:true});}
},30000);
