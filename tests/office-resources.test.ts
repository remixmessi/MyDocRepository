import {it,expect,vi} from 'vitest';
import {createServer} from 'node:http';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import * as Y from 'yjs';
import {zipSync,strToU8} from 'fflate';
import {FORMULA_CALCULATION_VERSION} from '../packages/common/nbstore/src/spreadsheet-formula';

it('办公资源目录支持 TXT、CSV/TSV 和 PPTX 元数据，并沿用空间权限',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-office-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const registered=await call('/register','POST',{email:'office@example.com',name:'Office',password:'abcdefghijk'}),cookie=registered.headers.get('set-cookie')!.split(';')[0];
  const space=await (await call('/workspaces','POST',{name:'办公协作'},cookie)).json() as {id:string};const query='?workspace='+space.id;
  expect(cloud.database.db.prepare('SELECT max(version) version FROM cloud_migrations').get()).toEqual({version:13});
  const created=await (await call('/office-resources'+query,'POST',{kind:'text',name:'说明.txt',mime:'text/plain',sourceFormat:'txt'},cookie)).json() as {id:string;kind:string;name:string};expect(created.kind).toBe('text');
  await call('/office-resources'+query,'POST',{kind:'spreadsheet',name:'数据.tsv',mime:'text/tab-separated-values',sourceFormat:'tsv'},cookie);
  await call('/office-resources'+query,'POST',{kind:'presentation',name:'汇报.pptx',mime:'application/vnd.openxmlformats-officedocument.presentationml.presentation',sourceFormat:'pptx'},cookie);
  const listed=await (await call('/office-resources'+query,'GET',undefined,cookie)).json() as {kind:string}[];expect(listed.map(row=>row.kind).sort()).toEqual(['presentation','spreadsheet','text']);
  expect((await call('/office-resource'+query+'&id='+encodeURIComponent(created.id),'PATCH',{name:'说明-改名.txt'},cookie)).status).toBe(200);
  await expect((await call('/office-resource'+query+'&id='+encodeURIComponent(created.id),'GET',undefined,cookie)).json()).resolves.toMatchObject({name:'说明-改名.txt'});
  expect((await call('/office-resource'+query+'&id='+encodeURIComponent(created.id),'DELETE',undefined,cookie)).status).toBe(200);
  expect((await call('/office-resource'+query+'&id='+encodeURIComponent(created.id),'GET',undefined,cookie)).status).toBe(404);
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('TXT 原件导入建立不可变原件与协同状态',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-office-import-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const registered=await call('/register','POST',{email:'import@example.com',name:'Importer',password:'abcdefghijk'}),cookie=registered.headers.get('set-cookie')!.split(';')[0];
  const space=await (await call('/workspaces','POST',{name:'原件导入'},cookie)).json() as {id:string};const query='?workspace='+space.id;
  const bytes=new TextEncoder().encode('\ufeff第一行\r\n第二行\n第三行');
  const imported=await (await fetch(root+'/office-text-import'+query+'&name='+encodeURIComponent('原件.txt')+'&lineEnding=crlf',{method:'POST',headers:{Cookie:cookie,'Content-Type':'application/octet-stream'},body:bytes})).json() as {id:string;originalAssetId:string};
  expect(imported.id).toMatch(/^doc:/);expect(imported.originalAssetId).toBeTruthy();
  const original=await fetch(root+'/office-original'+query+'&id='+encodeURIComponent(imported.id),{headers:{Cookie:cookie}});expect(original.status).toBe(200);expect(new Uint8Array(await original.arrayBuffer())).toEqual(bytes);
  const owner=cloud.database.db.prepare('SELECT doc FROM cloud_asset_owners WHERE workspace=? AND id=?').get(space.id,imported.originalAssetId) as {doc:string};expect(owner.doc).toBe(imported.id);
  const stored=cloud.database.load('cloud:'+space.id,imported.id)!;const doc=new Y.Doc();Y.applyUpdate(doc,stored);expect(doc.getText('text').toString()).toBe('第一行\n第二行\n第三行');expect(doc.getMap('textSettings').get('lineEnding')).toBe('crlf');doc.destroy();
  const replacement=new Uint8Array([1,2,3]);expect((await fetch(root+'/asset?workspace='+space.id+'&id='+encodeURIComponent(imported.originalAssetId),{method:'PUT',headers:{Cookie:cookie,'Content-Type':'application/octet-stream'},body:replacement})).status).toBe(409);
  expect((await fetch(root+'/office-text-import'+query+'&name=x.txt',{method:'POST',headers:{Cookie:cookie,'Content-Type':'application/octet-stream'},body:new Uint8Array([0xff,0xff])})).status).toBe(400);
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('统一办公导入接受 CSV、XLSX、DOCX、PDF、PPTX 并保留原件字节',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-office-binary-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const registered=await call('/register','POST',{email:'binary@example.com',name:'Binary',password:'abcdefghijk'}),cookie=registered.headers.get('set-cookie')!.split(';')[0];const space=await(await call('/workspaces','POST',{name:'格式导入'},cookie)).json() as {id:string};
  const xlsx=zipSync({'xl/worksheets/sheet1.xml':strToU8('<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>标题</t></is></c></row></sheetData></worksheet>')});
  const packageBytes=zipSync({'word/document.xml':strToU8('<w:document/>')});const pptPackage=zipSync({'ppt/slides/slide1.xml':strToU8('<p:sld/>')});
  const samples:[string,Uint8Array][]=[['数据.csv',new TextEncoder().encode('a,b\n1,2\n')],['表格.xlsx',xlsx],['文档.docx',packageBytes],['页面.pdf',new TextEncoder().encode('%PDF-1.7')],['演示.pptx',pptPackage]];
  for(const [name,bytes] of samples){const response=await fetch(root+'/office-import?workspace='+space.id+'&name='+encodeURIComponent(name),{method:'POST',headers:{Cookie:cookie,'Content-Type':'application/octet-stream'},body:bytes});expect(response.status).toBe(201);const resource=await response.json() as {id:string;kind:string};expect(resource.id).toMatch(/^doc:/);const original=await fetch(root+'/office-original?workspace='+space.id+'&id='+encodeURIComponent(resource.id),{headers:{Cookie:cookie}});expect(new Uint8Array(await original.arrayBuffer())).toEqual(bytes);}
  const snapshot=()=>JSON.stringify(['cloud_office_resources','cloud_asset_owners','assets','asset_files','documents'].map(table=>cloud.database.db.prepare('SELECT count(*) AS count FROM '+table).get()));
  const before=snapshot(),put=vi.spyOn(cloud.database.attachments,'put');
  for(const [name,bytes]of [
   ['entity.docx',zipSync({'word/document.xml':strToU8('<!DOCTYPE root [<!ENTITY x SYSTEM "file:///etc/passwd">]><root>&x;</root>')})],
   ['broken.pptx',zipSync({'ppt/slides/slide1.xml':strToU8('<p:sld><p:sp></p:sld>')})],
   ['too-many-rows.csv',strToU8('x\n'.repeat(10001))],
   ['shared.xlsx',zipSync({'xl/worksheets/sheet1.xml':strToU8('<worksheet><sheetData><row r="1"><c r="A1"><f t="shared" si="0">A2</f><v>99</v></c></row></sheetData></worksheet>')})],
   ['bad-index.xlsx',zipSync({'xl/worksheets/sheet1.xml':strToU8('<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>999</v></c></row></sheetData></worksheet>')})],
  ] as [string,Uint8Array][]){const rejected=await fetch(root+'/office-import?workspace='+space.id+'&name='+encodeURIComponent(name),{method:'POST',headers:{Cookie:cookie,'Content-Type':'application/octet-stream'},body:bytes});expect(rejected.status).toBe(400);expect(snapshot()).toBe(before);}
  expect(put).not.toHaveBeenCalled();put.mockRestore();
  const invalid=await fetch(root+'/office-import?workspace='+space.id+'&name=bad.docx',{method:'POST',headers:{Cookie:cookie,'Content-Type':'application/octet-stream'},body:new Uint8Array([1,2,3])});expect(invalid.status).toBe(400);
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
},20000);
it('TXT 同步拒绝非法模型，资源目录遵守撤权，删除后禁止提交',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-office-sync-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const register=async(name:string)=>{const r=await call('/register','POST',{email:name+'@example.com',name,password:'abcdefghijk'});return {cookie:r.headers.get('set-cookie')!.split(';')[0],id:(await r.json()).user.id};};
  const owner=await register('office-owner'),member=await register('office-member');const space=await(await call('/workspaces','POST',{name:'办公权限'},owner.cookie)).json();const query='?workspace='+space.id;
  const invite=await(await call('/invite'+query,'POST',{email:'office-member@example.com'},owner.cookie)).json();await call('/accept','POST',{token:invite.token},member.cookie);
  const resource=await(await call('/office-resources'+query,'POST',{kind:'text',name:'private.txt',mime:'text/plain',sourceFormat:'txt'},owner.cookie)).json();const path='/document'+query+'&id='+encodeURIComponent(resource.id);
  const write=(doc:Y.Doc,cookie=member.cookie)=>fetch(root+path,{method:'POST',headers:{Cookie:cookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))});
  const doc=new Y.Doc();doc.getText('text').insert(0,'原样\t中文😀\n');expect((await write(doc)).status).toBe(200);const confirmed=cloud.database.load('cloud:'+space.id,resource.id)!;
  const bad=new Y.Doc();bad.getMap('surface').set('bad','shape');expect((await write(bad)).status).toBe(400);expect(cloud.database.load('cloud:'+space.id,resource.id)).toEqual(confirmed);bad.destroy();
  const a=new Y.Doc(),b=new Y.Doc();a.getText('text');b.getText('text');Y.applyUpdate(a,confirmed);Y.applyUpdate(b,confirmed);a.getText('text').insert(0,'甲');b.getText('text').insert(b.getText('text').length,'乙');expect((await write(a,owner.cookie)).status).toBe(200);expect((await write(b)).status).toBe(200);
  const sync=await fetch(root+'/sync'+query+'&id='+encodeURIComponent(resource.id),{method:'POST',headers:{Cookie:member.cookie},body:Buffer.from(Y.encodeStateVector(b))});const delta=await sync.json();Y.applyUpdate(b,Buffer.from(delta.update,'base64'));expect(b.getText('text').toString()).toBe('甲原样\t中文😀\n乙');
  const permission='/document-permissions'+query;
  expect((await call(permission,'PUT',{doc:resource.id,user:member.id,role:'viewer'},owner.cookie)).status).toBe(200);expect((await write(b)).status).toBe(403);
  expect((await call(permission,'PUT',{doc:resource.id,user:member.id,role:'denied'},owner.cookie)).status).toBe(200);expect(await(await call('/office-resources'+query,'GET',undefined,member.cookie)).json()).toEqual([]);expect((await call(path,'GET',undefined,member.cookie)).status).toBe(404);
  const presentation=await(await call('/office-resources'+query,'POST',{kind:'presentation',name:'slides.pptx',mime:'application/octet-stream',sourceFormat:'pptx'},owner.cookie)).json();expect((await call('/document'+query+'&id='+encodeURIComponent(presentation.id),'GET',undefined,owner.cookie)).status).toBe(200);
  expect((await call('/office-resource'+query+'&id='+encodeURIComponent(resource.id),'DELETE',undefined,owner.cookie)).status).toBe(200);expect((await write(doc,owner.cookie)).status).toBe(410);
  for(const value of [doc,a,b])value.destroy();
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('新建办公文件为每种类型生成可编辑的合法初始模型和空白 PDF',async()=>{
 const {createOfficeState,createBlankPdf}=await import('../packages/frontend/apps/web/server/office-create');
 const {validateWordUpdate}=await import('../packages/frontend/apps/web/server/office-word');
 const {validateSpreadsheetUpdate}=await import('../packages/frontend/apps/web/server/office-spreadsheet');
 const {validatePresentationUpdate}=await import('../packages/frontend/apps/web/server/office-presentation');
 const {validatePdfUpdate}=await import('../packages/frontend/apps/web/server/office-pdf');
 const {validateOfficeTextUpdate}=await import('../packages/frontend/apps/web/server/office-text');
 for(const [kind,validate]of [['word',validateWordUpdate],['spreadsheet',validateSpreadsheetUpdate],['presentation',validatePresentationUpdate],['pdf',validatePdfUpdate],['text',validateOfficeTextUpdate]] as const){
  const update=createOfficeState(kind);expect(()=>validate(null,update)).not.toThrow();
  const doc=new Y.Doc();Y.applyUpdate(doc,update);
  if(kind==='word')expect(doc.getArray('paragraphs').length).toBe(1);
  if(kind==='spreadsheet'){expect(doc.getArray('sheet').length).toBe(20);expect((doc.getArray('sheet').get(0) as Y.Array<string>).length).toBe(8);}
  if(kind==='presentation')expect(doc.getArray('objects').length).toBe(1);
  doc.destroy();
 }
 const pdf=createBlankPdf().toString();const xref=Number(/startxref\n(\d+)/.exec(pdf)![1]);expect(pdf.slice(xref,xref+4)).toBe('xref');expect(pdf).toContain('/Count 1');
});

it('单元格语义提交检测竞争、支持幂等重试并拒绝通用更新绕过',async()=>{
 const {cellRevision,structureRevision}=await import('../packages/common/nbstore/src/spreadsheet-operation');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-sheet-ops-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 const doc=new Y.Doc();
 try{
  const registered=await call('/register','POST',{email:'cells@example.com',name:'Cells',password:'abcdefghijk'}),cookie=registered.headers.get('set-cookie')!.split(';')[0];const space=await(await call('/workspaces','POST',{name:'单元格竞争'},cookie)).json();const resource=await(await call('/office-resources?workspace='+space.id,'POST',{kind:'spreadsheet',name:'cells.csv',mime:'text/csv',sourceFormat:'csv'},cookie)).json();
  const query='?workspace='+space.id+'&id='+encodeURIComponent(resource.id),path='/office-sheet-operation'+query;
  const sheet=doc.getArray<Y.Array<string>>('sheet');Y.applyUpdate(doc,cloud.database.load('cloud:'+space.id,resource.id)!);
  const base=cellRevision(sheet,0,0),otherBase=cellRevision(sheet,0,1);
  const first={operationId:crypto.randomUUID(),type:'setCell',row:0,column:0,value:'甲',base};
  expect((await call(path,'POST',first,cookie)).status).toBe(200);
  const confirmed=cloud.database.load('cloud:'+space.id,resource.id)!;
  expect((await call(path,'POST',{...first,operationId:crypto.randomUUID(),value:'乙'},cookie)).status).toBe(409);expect(cloud.database.load('cloud:'+space.id,resource.id)).toEqual(confirmed);
  const retry=await call(path,'POST',first,cookie);expect(retry.status).toBe(200);expect((await retry.json()).repeated).toBe(true);
  expect((await call(path,'POST',{...first,value:'篡改'},cookie)).status).toBe(409);
  expect((await call(path,'POST',{...first,operationId:crypto.randomUUID(),column:1,value:'另一格',base:otherBase},cookie)).status).toBe(200);
  Y.applyUpdate(doc,cloud.database.load('cloud:'+space.id,resource.id)!);expect(sheet.get(0).toArray().slice(0,2)).toEqual(['甲','另一格']);
  const oldRevision=cellRevision(sheet,0,0);
  expect((await call(path,'POST',{...first,operationId:crypto.randomUUID(),value:'',base:oldRevision},cookie)).status).toBe(200);
  expect((await call(path,'POST',{...first,operationId:crypto.randomUUID(),value:'旧草稿'},cookie)).status).toBe(409);
  const structuralBase=structureRevision(doc);expect((await call(path,'POST',{operationId:crypto.randomUUID(),type:'appendRow',base:structuralBase},cookie)).status).toBe(409);
  Y.applyUpdate(doc,cloud.database.load('cloud:'+space.id,resource.id)!);
  expect((await call(path,'POST',{operationId:crypto.randomUUID(),type:'appendRow',base:structureRevision(doc)},cookie)).status).toBe(200);
  expect((await fetch(root+'/document'+query,{method:'POST',headers:{Cookie:cookie},body:Buffer.from(Y.encodeStateAsUpdate(doc))})).status).toBe(409);
  expect((await call(path,'POST',{...first,operationId:crypto.randomUUID(),row:99999},cookie)).status).toBe(400);
 }finally{doc.destroy();server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('解析等待期间服务可响应，提交前撤权会拒绝导入且不写入资源',async()=>{
 const {OfficeParserPool}=await import('../packages/frontend/apps/web/server/office-parser-process');
 let ready:()=>void=()=>{},release:()=>void=()=>{};const started=new Promise<void>(resolve=>ready=resolve),gate=new Promise<void>(resolve=>release=resolve);
 const parse=vi.spyOn(OfficeParserPool.prototype,'parseDelimited').mockImplementation(async()=>{ready();await gate;return new Uint8Array([0,0]);});
 const dir=mkdtempSync(join(tmpdir(),'mydoc-import-recheck-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const register=async(name:string)=>{const response=await call('/register','POST',{email:name+'@example.com',name,password:'abcdefghijk'});return {cookie:response.headers.get('set-cookie')!.split(';')[0],id:(await response.json()).user.id};};
  const owner=await register('parser-owner'),member=await register('parser-member');const space=await(await call('/workspaces','POST',{name:'解析权限'},owner.cookie)).json();
  const invite=await(await call('/invite?workspace='+space.id,'POST',{email:'parser-member@example.com'},owner.cookie)).json();await call('/accept','POST',{token:invite.token},member.cookie);
  const uploading=fetch(root+'/office-import?workspace='+space.id+'&name=cells.csv',{method:'POST',headers:{Cookie:member.cookie},body:'a,b\n1,2'});
  await started;expect((await call('/overview?workspace='+space.id,'GET',undefined,owner.cookie)).status).toBe(200);
  expect((await call('/member?workspace='+space.id,'PATCH',{user:member.id,role:'viewer'},owner.cookie)).status).toBe(200);release();expect((await uploading).status).toBe(403);
  for(const table of ['cloud_office_resources','cloud_asset_owners','assets','documents'])expect(cloud.database.db.prepare('SELECT count(*) AS count FROM '+table).get()).toEqual({count:0});
 }finally{release();parse.mockRestore();server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('DOCX 导出固定快照，保留原件，校验范围、查看权限与异步期间撤权',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-word-export-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const register=async(name:string)=>{const r=await call('/register','POST',{email:name+'@example.com',name,password:'abcdefghijk'});return {cookie:r.headers.get('set-cookie')!.split(';')[0],id:(await r.json()).user.id};};
  const owner=await register('export-owner'),viewer=await register('export-viewer');
  const space=await(await call('/workspaces','POST',{name:'Word 导出'},owner.cookie)).json(),query='?workspace='+space.id;
  const invite=await(await call('/invite'+query,'POST',{email:'export-viewer@example.com'},owner.cookie)).json();await call('/accept','POST',{token:invite.token},viewer.cookie);
  const original=zipSync({'word/document.xml':strToU8('<w:document><w:body><w:p><w:r><w:t>原始正文</w:t></w:r></w:p></w:body></w:document>')});
  const resource=await(await fetch(root+'/office-import'+query+'&name=source.docx',{method:'POST',headers:{Cookie:owner.cookie},body:original})).json();
  const path='/office-export'+query+'&id='+encodeURIComponent(resource.id)+'&format=docx&scope=paragraphs';
  expect((await call(path.replace('&scope=paragraphs',''),'GET',undefined,owner.cookie)).status).toBe(400);
  const before=Buffer.from(cloud.database.load('cloud:'+space.id,resource.id)!);
  await call('/document-permissions'+query,'PUT',{doc:resource.id,user:viewer.id,role:'viewer'},owner.cookie);
  const response=await call(path,'GET',undefined,viewer.cookie);expect(response.status).toBe(200);expect(response.headers.get('content-type')).toContain('wordprocessingml');expect(response.headers.get('x-mydoc-snapshot')).toMatch(/^[a-f0-9]{64}$/);expect(response.headers.get('cache-control')).toBe('no-store');
  const {encodeDocxWord}=await import('../packages/frontend/apps/web/server/office-word');
  const parsed=encodeDocxWord(new Uint8Array(await response.arrayBuffer()));expect(parsed.doc.getArray<Y.Text>('paragraphs').get(0).toString()).toBe('原始正文');parsed.doc.destroy();
  const {OfficeParserPool}=await import('../packages/frontend/apps/web/server/office-parser-process');
  const {exportDocxWord}=await import('../packages/frontend/apps/web/server/office-word-export');
  let captured:Uint8Array|undefined;
  const spy=vi.spyOn(OfficeParserPool.prototype,'exportDocx').mockImplementation(async bytes=>{
   captured=bytes.slice();const changed=new Y.Doc();Y.applyUpdate(changed,before);changed.getArray<Y.Text>('paragraphs').get(0).insert(0,'后续修改');cloud.database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(changed)),'cloud:'+space.id,resource.id);changed.destroy();return exportDocxWord(bytes);
  });
  try{const fixed=await call(path,'GET',undefined,owner.cookie);expect(fixed.status).toBe(200);const old=encodeDocxWord(new Uint8Array(await fixed.arrayBuffer()));expect(old.doc.getArray<Y.Text>('paragraphs').get(0).toString()).toBe('原始正文');old.doc.destroy();expect(Buffer.from(captured!)).toEqual(before);
   spy.mockImplementation(async bytes=>{cloud.database.db.prepare("UPDATE cloud_document_permissions SET role='denied' WHERE workspace=? AND doc=? AND user=?").run(space.id,resource.id,viewer.id);return exportDocxWord(bytes);});
   expect((await call(path,'GET',undefined,viewer.cookie)).status).toBe(404);
   spy.mockImplementation(async bytes=>{cloud.database.db.prepare('UPDATE cloud_office_resources SET deleted=1 WHERE workspace=? AND id=?').run(space.id,resource.id);return exportDocxWord(bytes);});
   expect((await call(path,'GET',undefined,owner.cookie)).status).toBe(404);
  }finally{spy.mockRestore();}
  expect(cloud.database.attachments.get('cloud:'+space.id,space.id,resource.originalAssetId)!.data).toEqual(Buffer.from(original));
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('办公 Presence 按文档授权过滤广播与重放，快照和撤权清理不泄露选区',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-office-presence-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});const controllers:AbortController[]=[];
 try{
  const registered=await call('/register','POST',{email:'presence-owner@example.com',name:'Owner',password:'abcdefghijk'}),owner=registered.headers.get('set-cookie')!.split(';')[0];const second=await call('/register','POST',{email:'presence-member@example.com',name:'Member',password:'abcdefghijk'}),member=second.headers.get('set-cookie')!.split(';')[0],memberId=(await second.json()).user.id;
  const space=await(await call('/workspaces','POST',{name:'Presence'},owner)).json(),query='?workspace='+space.id;const invite=await(await call('/invite'+query,'POST',{email:'presence-member@example.com'},owner)).json();await call('/accept','POST',{token:invite.token},member);
  const create=async(name:string)=>(await(await call('/office-resources'+query,'POST',{kind:'text',name,mime:'text/plain',sourceFormat:'txt'},owner)).json()).id as string;const privateId=await create('private.txt'),publicId=await create('public.txt'),doc=new Y.Doc(),position=Y.createRelativePositionFromTypeIndex(doc.getText('text'),0),payload=(client:number,clock=1)=>({client,clock,active:true,cursor:{anchor:position,head:position},user:{name:'Forged'}});
  const presence=(id:string)=>'/presence'+query+'&id='+id;
  expect((await call(presence(privateId),'POST',payload(777),owner)).status).toBe(200);expect((await(await call(presence(privateId),'GET',undefined,owner)).json())[0].state.user.name).toBe('Owner');
  // Typed focus for non-text editors is validated and relayed; malformed shapes are refused.
  const focused=[{kind:'cell',sheet:'first',row:'12-3',column:'column-first-0',label:'A1'},{kind:'slide',slide:'first',object:'5-9'},{kind:'pdf',page:3},{kind:'word',text:'12:4',anchor:position,head:position}];
  for(const [i,focus]of focused.entries()){expect((await call(presence(privateId),'POST',{...payload(777,10+i),focus},owner)).status).toBe(200);expect((await(await call(presence(privateId),'GET',undefined,owner)).json())[0].state.focus).toEqual(JSON.parse(JSON.stringify(focus)));}
  for(const focus of [{kind:'cell',sheet:'first',row:'x',column:'c',label:'A1'},{kind:'cell',sheet:'first',row:'1-1',column:'c',label:'<b>'},{kind:'pdf',page:0},{kind:'slide',slide:'first',object:'x'},{kind:'word',text:'1:1',anchor:{type:'bad'},head:position},{kind:'other'},{kind:'pdf',page:1,extra:true}])expect((await call(presence(privateId),'POST',{...payload(777,20),focus},owner)).status).toBe(400);
  await call('/document-permissions'+query,'PUT',{doc:privateId,user:memberId,role:'denied'},owner);expect((await call(presence(privateId),'GET',undefined,member)).status).toBe(404);expect((await call(presence(privateId),'POST',payload(779),member)).status).toBe(404);
  const stream=async(cookie:string,id:string)=>{const controller=new AbortController();controllers.push(controller);const response=await fetch(root+'/events'+query+'&id='+id,{headers:{Cookie:cookie},signal:controller.signal});const reader=response.body!.getReader(),initial=new TextDecoder().decode((await reader.read()).value);expect(initial).not.toContain('"client":777');return reader;};
  const memberStream=await stream(member,'*'),specificStream=await stream(owner,publicId);await call(presence(privateId),'POST',payload(777,2),owner);await call(presence(publicId),'POST',payload(778),owner);
  for(const reader of [memberStream,specificStream]){const data=new TextDecoder().decode((await reader.read()).value);expect(data).toContain('"client":778');expect(data).not.toContain('"client":777');}
  expect((await call(presence(publicId),'POST',payload(778),member)).status).toBe(409);expect((await call(presence(publicId),'POST',payload(779),member)).status).toBe(200);
  await call('/document-permissions'+query,'PUT',{doc:publicId,user:memberId,role:'viewer'},owner);expect((await(await call(presence(publicId),'GET',undefined,owner)).json()).some((p:{client:number})=>p.client===779)).toBe(false);expect((await call(presence(publicId),'POST',payload(779,4),member)).status).toBe(200);
  await call('/document-permissions'+query,'PUT',{doc:publicId,user:memberId,role:'denied'},owner);expect((await(await call(presence(publicId),'GET',undefined,owner)).json()).some((p:{client:number})=>p.client===779)).toBe(false);
  await call('/office-resource'+query+'&id='+privateId,'DELETE',undefined,owner);expect((await call(presence(privateId),'GET',undefined,owner)).status).toBe(404);expect((await call(presence(privateId),'POST',payload(777,3),owner)).status).toBe(404);
  await call('/document-permissions'+query,'PUT',{doc:publicId,user:memberId,role:'viewer'},owner);await call(presence(publicId),'POST',payload(778,5),owner);expect((await(await call(presence(publicId),'GET',undefined,member)).json()).some((p:{client:number;state:unknown})=>p.client===778&&p.state)).toBe(true);
  const future=Date.now()+31000,clock=vi.spyOn(Date,'now').mockReturnValue(future);try{expect((await(await call(presence(publicId),'GET',undefined,member)).json())).toEqual([]);}finally{clock.mockRestore();}
  await call('/logout','POST',undefined,owner);expect((await call(presence(publicId),'GET',undefined,owner)).status).toBe(401);expect((await(await call(presence(publicId),'GET',undefined,member)).json()).some((p:{client:number})=>p.client===778)).toBe(false);doc.destroy();
 }finally{for(const controller of controllers)controller.abort();server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('单文件备份保留原件和模型，跨空间恢复新身份，损坏、撤权及写入失败无半成品',async()=>{
 const {parseOfficeBackup,restoreOfficeBackup}=await import('../packages/frontend/apps/web/server/office-backup');const {createHash}=await import('node:crypto');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-office-backup-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const registration=await call('/register','POST',{email:'backup-owner@example.com',name:'Owner',password:'abcdefghijk'}),owner=registration.headers.get('set-cookie')!.split(';')[0];const registered=await call('/register','POST',{email:'backup-member@example.com',name:'Member',password:'abcdefghijk'}),member=registered.headers.get('set-cookie')!.split(';')[0],memberId=(await registered.json()).user.id;
  const source=await(await call('/workspaces','POST',{name:'来源'},owner)).json(),target=await(await call('/workspaces','POST',{name:'目标'},owner)).json(),query='?workspace='+source.id,targetQuery='?workspace='+target.id;const invite=await(await call('/invite'+query,'POST',{email:'backup-member@example.com'},owner)).json();await call('/accept','POST',{token:invite.token},member);
  const original=Buffer.from([0xff,0xfe,0x2d,0x4e,0x87,0x65]);const imported=await(await fetch(root+'/office-text-import'+query+'&name=说明.txt',{method:'POST',headers:{Cookie:owner},body:original})).json();const doc=new Y.Doc();Y.applyUpdate(doc,new Uint8Array(await(await call('/document'+query+'&id='+imported.id,'GET',undefined,owner)).arrayBuffer()));doc.getText('text').insert(0,'协作修改😀');await fetch(root+'/document'+query+'&id='+imported.id,{method:'POST',headers:{Cookie:owner},body:Buffer.from(Y.encodeStateAsUpdate(doc))});
  await call('/document-permissions'+query,'PUT',{doc:imported.id,user:memberId,role:'viewer'},owner);const exported=await call('/office-backup'+query+'&id='+imported.id,'GET',undefined,member);expect(exported.status).toBe(200);expect(exported.headers.get('content-disposition')).toContain('mydoc-office.json');const archive=await exported.json();expect(archive.resource).not.toHaveProperty('id');expect(archive).not.toHaveProperty('permissions');
  const restored=await(await call('/office-restore'+targetQuery,'POST',archive,owner)).json();expect(restored.id).not.toBe(imported.id);expect(restored.originalAssetId).not.toBe(imported.originalAssetId);expect(restored.generation).not.toBe(imported.generation);expect(restored.name).toBe('说明.txt');expect(cloud.database.db.prepare('SELECT doc FROM cloud_asset_owners WHERE workspace=? AND id=?').get(target.id,restored.originalAssetId)).toEqual({doc:restored.id});expect(Buffer.from(await(await call('/office-original'+targetQuery+'&id='+restored.id,'GET',undefined,owner)).arrayBuffer())).toEqual(original);const recovered=new Y.Doc();Y.applyUpdate(recovered,new Uint8Array(await(await call('/document'+targetQuery+'&id='+restored.id,'GET',undefined,owner)).arrayBuffer()));expect(recovered.getText('text').toString()).toBe(doc.getText('text').toString());expect(recovered.getMap('textSettings').toJSON()).toEqual(doc.getMap('textSettings').toJSON());recovered.destroy();doc.destroy();
  const operationId='11111111-1111-4111-8111-111111111111',retryPath='/office-restore'+targetQuery+'&operationId='+operationId;
  const concurrent=await Promise.all([call(retryPath,'POST',archive,owner),call(retryPath,'POST',archive,owner)]);expect(concurrent.map(r=>r.status)).toEqual([201,201]);const receipts=await Promise.all(concurrent.map(r=>r.json()));expect(receipts[0].id).toBe(receipts[1].id);
  expect((await call(retryPath,'POST',{...archive,exportedAt:archive.exportedAt+1},owner)).status).toBe(409);
  expect((await call('/office-restore'+targetQuery+'&operationId=bad','POST',archive,owner)).status).toBe(400);
  const memberRestorePath='/office-restore'+query+'&operationId='+operationId,memberRestored=await(await call(memberRestorePath,'POST',archive,member)).json();expect(memberRestored.id).not.toBe(receipts[0].id);
  await call('/document-permissions'+query,'PUT',{doc:memberRestored.id,user:memberId,role:'denied'},owner);expect((await call(memberRestorePath,'POST',archive,member)).status).toBe(404);
  await call('/office-resource'+targetQuery+'&id='+receipts[0].id,'DELETE',undefined,owner);expect((await call(retryPath,'POST',archive,owner)).status).toBe(410);
  const otherSpace=await(await call('/office-restore'+query+'&operationId='+operationId,'POST',archive,owner)).json();expect(otherSpace.id).not.toBe(memberRestored.id);expect(otherSpace.id).not.toBe(receipts[0].id);
  expect(cloud.database.db.prepare('SELECT count(*) n FROM cloud_office_restores').get()).toEqual({n:3});
  const counts=()=>['cloud_office_restores','cloud_office_resources','documents','assets','asset_files','cloud_asset_owners'].map(table=>(cloud.database.db.prepare('SELECT count(*) n FROM '+table).get() as {n:number}).n);const before=counts();
  for(const bad of [{...archive,version:3},{...archive,state:{...archive.state,sha256:'0'.repeat(64)}},{...archive,original:{...archive.original,data:'AAAA'}},{...archive,resource:{...archive.resource,permissions:['owner']}}]){expect((await call('/office-restore'+targetQuery,'POST',bad,owner)).status).toBe(400);expect(counts()).toEqual(before);}
  const bad=structuredClone(archive),invalid=new Y.Doc();invalid.getText('foreign').insert(0,'不允许');const bytes=Buffer.from(Y.encodeStateAsUpdate(invalid));bad.state={data:bytes.toString('base64'),sha256:createHash('sha256').update(bytes).digest('hex')};expect((await call('/office-restore'+targetQuery,'POST',bad,owner)).status).toBe(400);invalid.destroy();expect(counts()).toEqual(before);
  cloud.database.db.exec("CREATE TRIGGER fail_office_restore BEFORE INSERT ON cloud_office_resources BEGIN SELECT RAISE(ABORT,'injected restore failure'); END;");try{const parsed=parseOfficeBackup(Buffer.from(JSON.stringify(archive)));expect(()=>restoreOfficeBackup(cloud.database,target.id,'owner',parsed,parsed.state,{id:'rollback',sha256:'test'})).toThrow('injected');expect(counts()).toEqual(before);}finally{cloud.database.db.exec('DROP TRIGGER fail_office_restore');}
  cloud.database.db.exec("CREATE TRIGGER fail_restore_receipt BEFORE INSERT ON cloud_office_restores BEGIN SELECT RAISE(ABORT,'receipt failed'); END;");try{expect((await call('/office-restore'+targetQuery+'&operationId=22222222-2222-4222-8222-222222222222','POST',archive,owner)).status).toBe(400);expect(counts()).toEqual(before);}finally{cloud.database.db.exec('DROP TRIGGER fail_restore_receipt');}
  await call('/member'+query,'PATCH',{user:memberId,role:'viewer'},owner);expect((await call('/office-restore'+query,'POST',archive,member)).status).toBe(403);await call('/member'+query,'PATCH',{user:memberId,role:'member'},owner);await call('/document-permissions'+query,'PUT',{doc:imported.id,user:memberId,role:'denied'},owner);expect((await call('/office-backup'+query+'&id='+imported.id,'GET',undefined,member)).status).toBe(404);await call('/member'+query,'PATCH',{user:memberId,role:'viewer'},owner);expect((await call('/office-backup'+query+'&id='+imported.id,'GET',undefined,member)).status).toBe(404);
  await call('/office-resource'+query+'&id='+imported.id,'DELETE',undefined,owner);expect((await call('/office-backup'+query+'&id='+imported.id,'GET',undefined,owner)).status).toBe(404);expect((await call('/office-restore'+targetQuery,'POST',archive,owner)).status).toBe(201);
  for(const [kind,format,mime]of [['word','docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document'],['spreadsheet','xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],['pdf','pdf','application/pdf'],['presentation','pptx','application/vnd.openxmlformats-officedocument.presentationml.presentation']]){const resource=await(await call('/office-resources'+query,'POST',{kind,name:'备份.'+format,mime,sourceFormat:format},owner)).json();const response=await call('/office-backup'+query+'&id='+resource.id,'GET',undefined,owner);expect(response.status).toBe(200);const archive=await response.json(),restored=await call('/office-restore'+targetQuery,'POST',archive,owner);expect(restored.status).toBe(201);expect((await restored.json()).kind).toBe(kind);}
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
},20000);

it('XLSX 文本导出固定快照、查看者可下载，撤权或删除后拒绝响应',async()=>{
 const {encodeXlsxWorkbook}=await import('../packages/frontend/apps/web/server/office-spreadsheet');const {exportXlsxText}=await import('../packages/frontend/apps/web/server/office-spreadsheet-export');const {OfficeParserPool}=await import('../packages/frontend/apps/web/server/office-parser-process');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-xlsx-export-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const register=async(name:string)=>{const response=await call('/register','POST',{email:name+'@example.com',name,password:'abcdefghijk'});return {cookie:response.headers.get('set-cookie')!.split(';')[0],id:(await response.json()).user.id};};
  const owner=await register('xlsx-owner'),viewer=await register('xlsx-viewer'),space=await(await call('/workspaces','POST',{name:'XLSX 副本'},owner.cookie)).json(),query='?workspace='+space.id;
  const invite=await(await call('/invite'+query,'POST',{email:'xlsx-viewer@example.com',role:'viewer'},owner.cookie)).json();await call('/accept','POST',{token:invite.token},viewer.cookie);
  const original=Buffer.from('001,=1+1,12345678901234567890\r\n');const resource=await(await fetch(root+'/office-import'+query+'&name=数据.csv',{method:'POST',headers:{Cookie:owner.cookie},body:original})).json();const path='/office-export'+query+'&id='+resource.id+'&format=xlsx&scope=text-grid';
  expect((await call(path.replace('&scope=text-grid',''),'GET',undefined,owner.cookie)).status).toBe(400);expect((await call(path.replace('format=xlsx&scope=text-grid','format=docx&scope=paragraphs'),'GET',undefined,owner.cookie)).status).toBe(404);
  const response=await call(path,'GET',undefined,viewer.cookie);expect(response.status).toBe(200);expect(response.headers.get('content-type')).toContain('spreadsheetml.sheet');expect(response.headers.get('content-disposition')).toContain('.xlsx');expect(response.headers.get('x-mydoc-snapshot')).toMatch(/^[a-f0-9]{64}$/);expect(response.headers.get('cache-control')).toBe('no-store');
  const parsed=encodeXlsxWorkbook(new Uint8Array(await response.arrayBuffer()));expect(parsed.doc.getArray<Y.Array<string>>('sheet').get(0).toArray()).toEqual(['001','=1+1','12345678901234567890']);parsed.doc.destroy();
  const {parseDelimited}=await import('../packages/common/nbstore/src/office');
  for(const format of ['csv','tsv'])for(const mode of ['escaped','raw']){const exported=await call(path.replace('format=xlsx&scope=text-grid','format='+format+'&scope='+mode+'-text'),'GET',undefined,viewer.cookie);expect(exported.status).toBe(200);expect(exported.headers.get('x-mydoc-snapshot')).toMatch(/^[a-f0-9]{64}$/);expect(exported.headers.get('content-type')).toContain(format==='csv'?'text/csv':'text/tab-separated-values');const bytes=new Uint8Array(await exported.arrayBuffer());expect([...bytes.slice(0,3)]).toEqual([239,187,191]);expect(parseDelimited(new TextDecoder().decode(bytes),format==='csv'?',':'\t')[0]).toEqual(['001',mode==='escaped'?"'=1+1":'=1+1','12345678901234567890']);}
  expect((await call(path.replace('format=xlsx&scope=text-grid','format=csv&scope=unknown'),'GET',undefined,viewer.cookie)).status).toBe(400);
  const cachePath=path.replace('scope=text-grid','scope=formula-cache-grid');expect((await call(cachePath,'GET',undefined,viewer.cookie)).status).toBe(200);
  const {exportXlsxFormula}=await import('../packages/frontend/apps/web/server/office-spreadsheet-export');const cacheSpy=vi.spyOn(OfficeParserPool.prototype,'exportXlsxFormula').mockImplementation(async(bytes,_signal,cache)=>{cloud.database.db.prepare('INSERT INTO cloud_document_permissions VALUES (?,?,?,?)').run(space.id,resource.id,viewer.id,'denied');return exportXlsxFormula(bytes,cache);});
  try{expect((await call(cachePath,'GET',undefined,viewer.cookie)).status).toBe(404);cloud.database.db.prepare('DELETE FROM cloud_document_permissions WHERE workspace=? AND doc=? AND user=?').run(space.id,resource.id,viewer.id);cacheSpy.mockImplementation(async(bytes,_signal,cache)=>{cloud.database.db.prepare('UPDATE cloud_office_resources SET deleted=1 WHERE workspace=? AND id=?').run(space.id,resource.id);return exportXlsxFormula(bytes,cache);});expect((await call(cachePath,'GET',undefined,owner.cookie)).status).toBe(404);cloud.database.db.prepare('UPDATE cloud_office_resources SET deleted=0 WHERE workspace=? AND id=?').run(space.id,resource.id);}finally{cacheSpy.mockRestore();}
  const before=Buffer.from(cloud.database.load('cloud:'+space.id,resource.id)!);const spy=vi.spyOn(OfficeParserPool.prototype,'exportXlsxText').mockImplementation(async bytes=>{const changed=new Y.Doc();Y.applyUpdate(changed,before);changed.getArray<Y.Array<string>>('sheet').get(0).delete(0,1);changed.getArray<Y.Array<string>>('sheet').get(0).insert(0,['更新']);cloud.database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(changed)),'cloud:'+space.id,resource.id);changed.destroy();return exportXlsxText(bytes);});
  try{
   const fixed=await call(path,'GET',undefined,owner.cookie),snapshot=encodeXlsxWorkbook(new Uint8Array(await fixed.arrayBuffer()));expect(snapshot.doc.getArray<Y.Array<string>>('sheet').get(0).get(0)).toBe('001');snapshot.doc.destroy();
   spy.mockImplementation(async bytes=>{cloud.database.db.prepare('INSERT INTO cloud_document_permissions VALUES (?,?,?,?)').run(space.id,resource.id,viewer.id,'denied');return exportXlsxText(bytes);});expect((await call(path,'GET',undefined,viewer.cookie)).status).toBe(404);
   spy.mockImplementation(async bytes=>{cloud.database.db.prepare('UPDATE cloud_office_resources SET deleted=1 WHERE workspace=? AND id=?').run(space.id,resource.id);return exportXlsxText(bytes);});expect((await call(path,'GET',undefined,owner.cookie)).status).toBe(404);
  }finally{spy.mockRestore();}
  expect(cloud.database.attachments.get('cloud:'+space.id,space.id,resource.originalAssetId)!.data).toEqual(original);
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('CSV/TSV 导入选项保留原件、表头及设置，非法输入和查看者不创建资源',async()=>{
 const {encodeOfficeText}=await import('../packages/common/nbstore/src/office');const dir=mkdtempSync(join(tmpdir(),'mydoc-delimited-import-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const register=async(name:string)=>{const response=await call('/register','POST',{email:name+'@example.com',name,password:'abcdefghijk'});return {cookie:response.headers.get('set-cookie')!.split(';')[0],id:(await response.json()).user.id};};const owner=await register('csv-import-owner'),viewer=await register('csv-import-viewer'),space=await(await call('/workspaces','POST',{name:'CSV 导入设置'},owner.cookie)).json(),query='?workspace='+space.id;
  const invite=await(await call('/invite'+query,'POST',{email:'csv-import-viewer@example.com',role:'viewer'},owner.cookie)).json();await call('/accept','POST',{token:invite.token},viewer.cookie);
  const original=encodeOfficeText('表头\t编号\t表达式\r\n"中文"\t00123\t=1+1\r\n','utf-16be',false),path='/office-import'+query+'&name=数据.csv&encoding=utf-16be&delimiter=tab&quoteMode=none';const upload=(route:string,bytes:Uint8Array,cookie=owner.cookie)=>fetch(root+route,{method:'POST',headers:{Cookie:cookie},body:Buffer.from(bytes)});
  expect((await upload(path,original,viewer.cookie)).status).toBe(403);const response=await upload(path,original);expect(response.status).toBe(201);const resource=await response.json(),doc=new Y.Doc();try{Y.applyUpdate(doc,cloud.database.load('cloud:'+space.id,resource.id)!);expect(doc.getArray<Y.Array<string>>('sheet').toArray().map(row=>row.toArray())).toEqual([['表头','编号','表达式'],['"中文"','00123','=1+1']]);expect(doc.getMap('sheetImport').toJSON()).toEqual({encoding:'utf-16be',bom:false,delimiter:'\t',quoteMode:'none'});}finally{doc.destroy();}
  expect(Buffer.from(await(await call('/office-original'+query+'&id='+resource.id,'GET',undefined,owner.cookie)).arrayBuffer())).toEqual(Buffer.from(original));
  const counts=()=>['cloud_office_resources','cloud_asset_owners','assets','documents'].map(table=>(cloud.database.db.prepare('SELECT count(*) n FROM '+table).get() as {n:number}).n),before=counts();
  for(const [suffix,bytes]of [['&encoding=latin1',original],['&delimiter=comma&quoteMode=none',original],['&encoding=utf-8',new Uint8Array([255])],['&encoding=utf-16be',encodeOfficeText('x','utf-8',true)],['&encoding=utf-8',new TextEncoder().encode('"bad"x,1')]]){expect((await upload('/office-import'+query+'&name=bad.csv'+suffix,bytes as Uint8Array)).status).toBe(400);expect(counts()).toEqual(before);}
  const backup=await(await call('/office-backup'+query+'&id='+resource.id,'GET',undefined,owner.cookie)).json();const restored=await(await call('/office-restore'+query,'POST',backup,owner.cookie)).json(),copy=new Y.Doc();try{Y.applyUpdate(copy,cloud.database.load('cloud:'+space.id,restored.id)!);expect(copy.getMap('sheetImport').get('encoding')).toBe('utf-16be');}finally{copy.destroy();}
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('CSV/TSV 计算结果导出携带快照与计算版本，计算错误可改为原文且不写入状态',async()=>{
 const {OfficeParserPool}=await import('../packages/frontend/apps/web/server/office-parser-process'),{exportDelimitedText}=await import('../packages/frontend/apps/web/server/office-delimited-export');const dir=mkdtempSync(join(tmpdir(),'mydoc-calculated-export-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,method='GET',value?:unknown,cookie?:string)=>fetch(root+path,{method,headers:{...(cookie?{Cookie:cookie}:{}),...(value?{'Content-Type':'application/json'}:{})},body:value?JSON.stringify(value):undefined});
 try{
  const response=await call('/register','POST',{email:'calculate-owner@example.com',name:'Owner',password:'abcdefghijk'}),cookie=response.headers.get('set-cookie')!.split(';')[0],space=await(await call('/workspaces','POST',{name:'计算结果'},cookie)).json(),query='?workspace='+space.id,resource=await(await call('/office-resources'+query,'POST',{kind:'spreadsheet',name:'计算.xlsx',mime:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',sourceFormat:'xlsx'},cookie)).json();
  const doc=new Y.Doc(),row=new Y.Array<string>();row.push(['=1+1','001','=B1']);doc.getArray('sheet').push([row]);const {setSpreadsheetCellType}=await import('../packages/common/nbstore/src/spreadsheet-cell-type');setSpreadsheetCellType(doc,0,1,'text');const before=Buffer.from(Y.encodeStateAsUpdate(doc));cloud.database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(before,'cloud:'+space.id,resource.id);
  const path='/office-export'+query+'&id='+resource.id+'&format=csv&scope=raw-values',result=await call(path,'GET',undefined,cookie);expect(result.status).toBe(200);expect(result.headers.get('x-mydoc-calculation-version')).toBe(FORMULA_CALCULATION_VERSION);expect(result.headers.get('x-mydoc-snapshot')).toMatch(/^[a-f0-9]{64}$/);expect(result.headers.get('content-disposition')).toContain('.csv');expect(new TextDecoder().decode(await result.arrayBuffer())).toBe('2,001,001\r\n');expect(cloud.database.load('cloud:'+space.id,resource.id)).toEqual(before);
  const cachedResponse=await call(path.replace('format=csv&scope=raw-values','format=xlsx&scope=formula-cache-grid'),'GET',undefined,cookie);expect(cachedResponse.status).toBe(200);expect(cachedResponse.headers.get('x-mydoc-calculation-version')).toBe(FORMULA_CALCULATION_VERSION);expect(cachedResponse.headers.get('x-mydoc-snapshot')).toBe(result.headers.get('x-mydoc-snapshot'));const cachedFiles=(await import('../packages/frontend/apps/web/server/office-package')).readOfficePackage(new Uint8Array(await cachedResponse.arrayBuffer()));expect(new TextDecoder().decode(cachedFiles['xl/worksheets/sheet1.xml'])).toContain('<f>1+1</f><v>2</v>');
  const formulaResponse=await call(path.replace('format=csv&scope=raw-values','format=xlsx&scope=formula-grid'),'GET',undefined,cookie);expect(formulaResponse.status).toBe(200);expect(formulaResponse.headers.get('content-type')).toContain('spreadsheetml.sheet');const {readOfficePackage,readOfficeXml}=await import('../packages/frontend/apps/web/server/office-package');const formulaXml=readOfficeXml(readOfficePackage(new Uint8Array(await formulaResponse.arrayBuffer()))['xl/worksheets/sheet1.xml']);expect(formulaXml).toContain('<f>1+1</f></c>');expect(formulaXml).toContain('<f>B1</f></c>');expect(cloud.database.load('cloud:'+space.id,resource.id)).toEqual(before);
  const spy=vi.spyOn(OfficeParserPool.prototype,'exportDelimited').mockImplementation(async(bytes,format,mode,_signal,content)=>{row.delete(0,1);row.insert(0,['=99']);cloud.database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(doc)),'cloud:'+space.id,resource.id);return exportDelimitedText(bytes,format,mode,content);});try{const fixed=await call(path,'GET',undefined,cookie);expect(new TextDecoder().decode(await fixed.arrayBuffer())).toBe('2,001,001\r\n');}finally{spy.mockRestore();}
  row.delete(0,1);row.insert(0,['=1/0']);const broken=Buffer.from(Y.encodeStateAsUpdate(doc));cloud.database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(broken,'cloud:'+space.id,resource.id);doc.destroy();const failure=await call(path,'GET',undefined,cookie);expect(failure.status).toBe(400);expect((await failure.json()).error).toContain('A1 #DIV/0!');expect(cloud.database.load('cloud:'+space.id,resource.id)).toEqual(broken);const raw=await call(path.replace('raw-values','raw-text'),'GET',undefined,cookie);expect(raw.status).toBe(200);expect((await call(path.replace('format=csv&scope=raw-values','format=xlsx&scope=formula-grid'),'GET',undefined,cookie)).status).toBe(400);expect(raw.headers.has('x-mydoc-calculation-version')).toBe(false);expect(new TextDecoder().decode(await raw.arrayBuffer())).toBe('=1/0,001,=B1\r\n');
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
},20000);
