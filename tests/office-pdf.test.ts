import {it,expect,vi} from 'vitest';
import * as Y from 'yjs';
import {PDFDocument,PDFName,PDFArray,PDFDict,PDFHexString,degrees} from 'pdf-lib';
import {createHash,randomUUID} from 'node:crypto';
import {createServer} from 'node:http';
import {mkdtempSync,rmSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {exportAnnotatedPdf,packPdfExport} from '../packages/frontend/apps/web/server/office-pdf-export';
import {pdfAnnotationId,pdfAnnotationRevision,validPdfAnnotation} from '../packages/common/nbstore/src/office-pdf';
import {validatePdfUpdate} from '../packages/frontend/apps/web/server/office-pdf';
import {OfficeParserPool} from '../packages/frontend/apps/web/server/office-parser-process';
async function original(){const pdf=await PDFDocument.create();const page=pdf.addPage([400,500]);page.drawText('Immutable original');page.setCropBox(10,20,380,460);page.setRotation(degrees(90));return pdf.save();}
function state(){const doc=new Y.Doc();for(const data of [{page:1,type:'comment',x:30,y:50,content:'中文附注 🐈',id:'comment-1'},{page:1,type:'highlight',x:40,y:80,width:120,height:20,content:'高亮',id:'highlight-1'},{page:1,type:'ink',x:40,y:110,width:60,height:30,points:[40,110,70,140,100,120],content:'手绘',id:'ink-1'}]){const a=new Y.Map();for(const [key,value]of Object.entries(data))a.set(key,value);doc.getArray('annotations').push([a]);}const bytes=Y.encodeStateAsUpdate(doc);doc.destroy();return bytes;}
it('PDF 原生批注副本保留页内容、裁剪/旋转和 Unicode，不修改输入',async()=>{
 const source=await original(),update=state(),hash=createHash('sha256').update(source).digest('hex'),bytes=await exportAnnotatedPdf(packPdfExport(update,source)),pdf=await PDFDocument.load(bytes);expect(pdf.getPages()).toHaveLength(1);const page=pdf.getPage(0);expect(page.getRotation().angle).toBe(90);expect(page.getCropBox()).toEqual({x:10,y:20,width:380,height:460});expect(page.node.Contents()).toBeTruthy();const array=page.node.lookup(PDFName.of('Annots'),PDFArray);expect(array.size()).toBe(3);const first=array.lookup(0,PDFDict);expect(first.lookup(PDFName.of('Contents'),PDFHexString).decodeText()).toBe('中文附注 🐈');expect(array.lookup(1,PDFDict).lookup(PDFName.of('QuadPoints'),PDFArray).asArray().map(v=>v.toString())).toEqual(['40','100','160','100','40','80','160','80']);expect(array.lookup(2,PDFDict).lookup(PDFName.of('InkList'),PDFArray).size()).toBe(1);expect(createHash('sha256').update(source).digest('hex')).toBe(hash);
});
it('非法页码、裁剪外坐标、损坏 PDF 和缺少路径拒绝，导出在受限进程内执行',async()=>{
 const source=await original(),doc=new Y.Doc();Y.applyUpdate(doc,state());const a=doc.getArray<Y.Map<unknown>>('annotations').get(0);a.set('page',2);await expect(exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source))).rejects.toThrow('页码');a.set('page',1);a.set('x',0);await expect(exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source))).rejects.toThrow('裁剪');doc.destroy();expect(validPdfAnnotation({page:1,type:'ink',x:0,y:0,content:''},true)).toBe(false);
 const pool=new OfficeParserPool();try{const bytes=await pool.exportAnnotatedPdf(packPdfExport(state(),source));expect((await PDFDocument.load(bytes)).getPageCount()).toBe(1);await expect(pool.exportAnnotatedPdf(packPdfExport(state(),new Uint8Array([1,2,3])))).rejects.toMatchObject({status:400});}finally{pool.close();}
});
it('批注唯一 ID、非法根和不完整依赖校验，历史无 ID 附注仍可读取',()=>{
 const doc=new Y.Doc(),a=new Y.Map();for(const [key,value]of Object.entries({page:1,type:'comment',x:0,y:0,content:'旧附注'}))a.set(key,value);doc.getArray('annotations').push([a]);expect(pdfAnnotationId(a)).toMatch(/^legacy-/);expect(()=>validatePdfUpdate(null,Y.encodeStateAsUpdate(doc))).not.toThrow();doc.getMap('bad').set('x',1);expect(()=>validatePdfUpdate(null,Y.encodeStateAsUpdate(doc))).toThrow('根');doc.destroy();
});
it('PDF 语义修改竞争、幂等重放、原件保护、只读与原生导出',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-pdf-')),cloud=new CloudServer(join(dir,'db.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const register=await fetch(root+'/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:'pdf@example.com',name:'PDF',password:'test-password-123'})}),cookie=register.headers.get('set-cookie')!.split(';')[0];const call=(path:string,method='GET',data?:unknown)=>fetch(root+path,{method,headers:{Cookie:cookie,'Content-Type':'application/json'},body:data===undefined?undefined:JSON.stringify(data)});
 try{
  const space=await(await call('/workspaces','POST',{name:'PDF'})).json(),source=await original(),imported=await(await fetch(root+'/office-import?workspace='+space.id+'&name=test.pdf',{method:'POST',headers:{Cookie:cookie},body:source.slice().buffer})).json(),q='?workspace='+space.id+'&id='+imported.id;
  const annotation={page:1,type:'comment',x:30,y:50,content:'第一条'},operation={type:'create',operationId:randomUUID(),annotationId:randomUUID(),base:'',annotation};const first=await call('/office-pdf-operation'+q,'POST',operation);expect(first.status).toBe(200);expect((await(await call('/office-pdf-operation'+q,'POST',operation)).json()).repeated).toBe(true);
  const doc=new Y.Doc();Y.applyUpdate(doc,cloud.database.load('cloud:'+space.id,imported.id)!);const base=pdfAnnotationRevision(doc.getArray<Y.Map<unknown>>('annotations').get(0));doc.destroy();
  const update={...operation,type:'update',operationId:randomUUID(),base,annotation:{...annotation,content:'已修改'}};expect((await call('/office-pdf-operation'+q,'POST',update)).status).toBe(200);expect((await call('/office-pdf-operation'+q,'POST',{...update,operationId:randomUUID(),annotation:{...annotation,content:'旧草稿'}})).status).toBe(409);expect((await call('/office-pdf-operation'+q,'POST',{...update,annotation:{...annotation,content:'改动编号'}})).status).toBe(409);
  const raw=await fetch(root+'/document'+q,{method:'POST',headers:{Cookie:cookie},body:state().slice().buffer});expect(raw.status).toBe(409);const exported=await call('/office-pdf-export'+q);expect(exported.status).toBe(200);expect(exported.headers.get('x-mydoc-snapshot')).toHaveLength(64);const pdf=await PDFDocument.load(await exported.arrayBuffer());expect(pdf.getPage(0).node.lookup(PDFName.of('Annots'),PDFArray).lookup(0,PDFDict).lookup(PDFName.of('Contents'),PDFHexString).decodeText()).toBe('已修改');expect(Buffer.from(await(await call('/office-original'+q)).arrayBuffer())).toEqual(Buffer.from(source));
  const before=cloud.database.load('cloud:'+space.id,imported.id);cloud.database.db.exec("CREATE TRIGGER fail_pdf BEFORE UPDATE ON documents BEGIN SELECT RAISE(ABORT,'fail'); END");expect((await call('/office-pdf-operation'+q,'POST',{...operation,annotationId:randomUUID(),operationId:randomUUID()})).status).toBe(500);expect(cloud.database.load('cloud:'+space.id,imported.id)).toEqual(before);cloud.database.db.exec('DROP TRIGGER fail_pdf');
  cloud.database.db.prepare("UPDATE cloud_members SET role='viewer' WHERE workspace=?").run(space.id);expect((await call('/office-pdf-operation'+q,'POST',{...operation,operationId:randomUUID()})).status).toBe(403);expect((await call('/office-pdf-export'+q)).status).toBe(200);
  const pool=(cloud as unknown as {officeParser:OfficeParserPool}).officeParser;vi.spyOn(pool,'exportAnnotatedPdf').mockImplementationOnce(async packet=>{cloud.database.db.prepare('DELETE FROM cloud_members WHERE workspace=?').run(space.id);return exportAnnotatedPdf(packet);});expect((await call('/office-pdf-export'+q)).status).toBe(403);
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
},20000);


it('数字签名阻止重写，已有标准批注保留，路径越界和伪造作者拒绝',async()=>{
 const original=await PDFDocument.create(),page=original.addPage([400,500]);const old=original.context.obj({Type:'Annot',Subtype:'Text',Rect:[30,30,50,50],Contents:PDFHexString.fromText('原有附注')});page.node.addAnnot(original.context.register(old));
 const preserved=await PDFDocument.load(await exportAnnotatedPdf(packPdfExport(state(),await original.save())));expect(preserved.getPage(0).node.lookup(PDFName.of('Annots'),PDFArray).size()).toBe(4);
 original.context.register(original.context.obj({Type:'Sig',Filter:'Adobe.PPKLite'}));await expect(exportAnnotatedPdf(packPdfExport(state(),await original.save()))).rejects.toThrow('签名');expect(validPdfAnnotation({page:1,type:'comment',x:1,y:1,content:'x',author:'forged'},true)).toBe(false);
 const doc=new Y.Doc();Y.applyUpdate(doc,state());doc.getArray<Y.Map<unknown>>('annotations').get(2).set('points',[40,110,9999,120]);const clean=await PDFDocument.create();clean.addPage([400,500]);await expect(exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),await clean.save()))).rejects.toThrow('路径');doc.destroy();
});

it('PDF 标准文本/勾选表单在受限进程读取，协作值原生往返及扁平化，拒绝缺字字体和 XFA',async()=>{
 const {inspectPdfForms,pdfFieldId}=await import('../packages/frontend/apps/web/server/office-pdf-forms');const pdf=await PDFDocument.create(),page=pdf.addPage([400,500]),form=pdf.getForm();const name=form.createTextField('姓名');name.setText('Original');name.setMaxLength(20);name.addToPage(page,{x:40,y:350,width:200,height:24});const check=form.createCheckBox('同意');check.addToPage(page,{x:40,y:300,width:18,height:18});const source=await pdf.save(),specs=JSON.parse(new TextDecoder().decode(await inspectPdfForms(source)));expect(specs.map((field:{name:string})=>field.name)).toEqual(['姓名','同意']);const doc=new Y.Doc();doc.getMap('pdfFields').set(pdfFieldId('姓名'),{type:'text',value:'Alice',revision:'operation-123456789'});doc.getMap('pdfFields').set(pdfFieldId('同意'),{type:'checkbox',value:true,revision:'operation-223456789'});const a=new Y.Map();for(const [key,value]of Object.entries({page:1,type:'highlight',x:40,y:250,width:100,height:15,content:'mark'}))a.set(key,value);doc.getArray('annotations').push([a]);const pool=new OfficeParserPool();try{
  expect(JSON.parse(Buffer.from(await pool.pdfFormSpecs(source)).toString())).toEqual(specs);const exported=await pool.exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source)),reopened=await PDFDocument.load(exported);expect(reopened.getForm().getTextField('姓名').getText()).toBe('Alice');expect(reopened.getForm().getCheckBox('同意').isChecked()).toBe(true);
  const flattened=await pool.exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source),undefined,true),flat=await PDFDocument.load(flattened);expect(flat.getForm().getFields()).toHaveLength(0);expect(flat.getPage(0).node.Annots()?.size()??0).toBe(0);expect((await PDFDocument.load(source)).getForm().getTextField('姓名').getText()).toBe('Original');
  // The pool ships the bundled Unicode font, so Chinese values get a real appearance; emoji are refused by name.
  doc.getMap('pdfFields').set(pdfFieldId('姓名'),{type:'text',value:'中文',revision:'operation-323456789'});const chinese=await PDFDocument.load(await pool.exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source)));expect(chinese.getForm().getTextField('姓名').getText()).toBe('中文');
  doc.getMap('pdfFields').set(pdfFieldId('姓名'),{type:'text',value:'中文😀',revision:'operation-423456789'});await expect(pool.exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source))).rejects.toThrow('无法显示的字符');const xfa=await PDFDocument.load(source);xfa.getForm().acroForm.dict.set(PDFName.of('XFA'),xfa.context.obj([]));await expect(inspectPdfForms(await xfa.save({updateFieldAppearances:false}))).rejects.toThrow('XFA');
 }finally{doc.destroy();pool.close();}
},15000);
it('PDF 表单字段级竞争、响应丢失幂等、类型/只读校验和备份恢复',async()=>{
 const {pdfFormRevision}=await import('../packages/common/nbstore/src/pdf-forms');const {submitPdfForm,inspectPdfForms}=await import('../packages/frontend/apps/web/server/office-pdf-forms');const {parseOfficeBackup}=await import('../packages/frontend/apps/web/server/office-backup');const directory=mkdtempSync(join(tmpdir(),'mydoc-pdf-forms-')),cloud=new CloudServer(join(directory,'db.sqlite')),server=createServer(cloud.handler);await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 try{const register=await fetch(root+'/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:'forms@example.com',name:'Form user',password:'test-password-123'})}),cookie=register.headers.get('set-cookie')!.split(';')[0],call=(path:string,data?:unknown)=>fetch(root+path,{method:data===undefined?'GET':'POST',headers:{Cookie:cookie,'Content-Type':'application/json'},body:data===undefined?undefined:JSON.stringify(data)}),space=await(await call('/workspaces',{name:'表单'})).json();const original=await PDFDocument.create(),page=original.addPage([400,500]),text=original.getForm().createTextField('subject');text.addToPage(page,{x:40,y:350,width:200,height:24});const locked=original.getForm().createCheckBox('locked');locked.enableReadOnly();locked.addToPage(page,{x:40,y:300,width:18,height:18});const bytes=await original.save();const resource=await(await fetch(root+'/office-import?workspace='+space.id+'&name=form.pdf',{method:'POST',headers:{Cookie:cookie},body:bytes})).json(),query='?'+new URLSearchParams({workspace:space.id,id:resource.id}),specs=await(await call('/office-pdf-forms'+query)).json(),spec=specs.find((s:{name:string})=>s.name==='subject'),input={fieldId:spec.id,operationId:randomUUID(),base:pdfFormRevision(spec),value:'First'};
 expect((await call('/office-pdf-form-operation'+query,input)).status).toBe(200);expect((await(await call('/office-pdf-form-operation'+query,input)).json()).repeated).toBe(true);expect((await call('/office-pdf-form-operation'+query,{...input,operationId:randomUUID(),value:'Second'})).status).toBe(409);expect((await call('/office-pdf-form-operation'+query,{...input,operationId:randomUUID(),value:true})).status).toBe(400);const lockedSpec=specs.find((s:{name:string})=>s.name==='locked');expect((await call('/office-pdf-form-operation'+query,{fieldId:lockedSpec.id,operationId:randomUUID(),base:pdfFormRevision(lockedSpec),value:true})).status).toBe(403);
 const backup=await call('/office-backup'+query);expect(backup.status).toBe(200);const parsed=parseOfficeBackup(new Uint8Array(await backup.arrayBuffer())),model=new Y.Doc();Y.applyUpdate(model,parsed.state);expect(model.getMap<{value:string}>('pdfFields').get(spec.id)?.value).toBe('First');model.destroy();const exported=await call('/office-pdf-export'+query);expect(exported.status).toBe(200);expect((await PDFDocument.load(await exported.arrayBuffer())).getForm().getTextField('subject').getText()).toBe('First');
 }finally{server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(directory,{recursive:true,force:true});}
},15000);

it('原件原生批注：受限进程识别附注/高亮/手绘并一次性转为可协作编辑，导出以协作内容替换，其他批注按外观扁平化且保留链接',async()=>{
 const {readNativeAnnotations}=await import('../packages/frontend/apps/web/server/office-pdf-native');const {importNativePdfAnnotations}=await import('../packages/frontend/apps/web/server/office-pdf-operation');
 const pdf=await PDFDocument.create(),page=pdf.addPage([400,500]);page.drawText('Native annotations');
 const appearance=pdf.context.register(pdf.context.flateStream('0 0 1 rg 0 0 50 20 re f',{Type:'XObject',Subtype:'Form',BBox:[0,0,50,20]}));
 const add=(dict:Record<string,unknown>)=>page.node.addAnnot(pdf.context.register(pdf.context.obj(dict as never)));
 add({Type:'Annot',Subtype:'Text',Rect:[30,400,50,420],Contents:PDFHexString.fromText('原件附注 😀'),T:PDFHexString.fromText('原作者')});
 add({Type:'Annot',Subtype:'Highlight',Rect:[40,300,160,320],QuadPoints:[40,320,160,320,40,300,160,300],Contents:PDFHexString.fromText('原高亮')});
 add({Type:'Annot',Subtype:'Ink',Rect:[40,200,100,230],InkList:[[40,200,70,230,100,210]]});
 add({Type:'Annot',Subtype:'Ink',Rect:[200,200,260,230],InkList:[[200,200,210,210],[220,220,230,230]]});
 add({Type:'Annot',Subtype:'Square',Rect:[250,100,300,120],AP:{N:appearance}});
 add({Type:'Annot',Subtype:'Link',Rect:[10,10,60,30],A:{S:'URI',URI:PDFHexString.fromText('https://example.com/')}});
 const source=await pdf.save(),natives=readNativeAnnotations(await PDFDocument.load(source));
 expect(natives.map(native=>[native.subtype,!!native.annotation])).toEqual([['Text',true],['Highlight',true],['Ink',true],['Ink',false],['Square',false]]);expect(natives[0].annotation).toMatchObject({type:'comment',content:'原件附注 😀',x:30,y:400});expect(natives[0].author).toBe('原作者');expect(natives[3].reason).toContain('多笔画');
 const pool=new OfficeParserPool();try{expect(JSON.parse(Buffer.from(await pool.pdfNativeAnnotations(source)).toString())).toHaveLength(5);}finally{pool.close();}
 const dir=mkdtempSync(join(tmpdir(),'mydoc-pdf-native-')),cloud=new CloudServer(join(dir,'cloud.sqlite'));try{
  const doc=new Y.Doc();doc.getArray('annotations');doc.getMap('pdfOperations').set('$schema','0'.repeat(64));cloud.database.append('cloud:space','pdf',Buffer.from(Y.encodeStateAsUpdate(doc)));
  const operationId=randomUUID(),first=importNativePdfAnnotations(cloud.database,'space','pdf','actor',operationId,natives);Y.applyUpdate(doc,first.update);expect(first.imported).toBe(3);
  expect(importNativePdfAnnotations(cloud.database,'space','pdf','actor',operationId,natives).repeated).toBe(true);expect(()=>importNativePdfAnnotations(cloud.database,'space','pdf','actor',randomUUID(),natives)).toThrow('已导入');
  const annotations=doc.getArray<Y.Map<unknown>>('annotations');expect(annotations.length).toBe(3);expect(annotations.get(0).get('author')).toBe('原作者');expect(()=>validatePdfUpdate(null,Y.encodeStateAsUpdate(doc))).not.toThrow();
  // Edit one imported annotation and delete another, as the semantic operations would.
  annotations.get(0).set('content','协作修改后的附注');annotations.delete(1,1);
  const exported=await PDFDocument.load(await exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source))),list=exported.getPage(0).node.lookup(PDFName.of('Annots'),PDFArray),subtypes=list.asArray().map((_,i)=>(list.lookup(i,PDFDict).get(PDFName.of('Subtype')) as PDFName).decodeText());
  expect(subtypes.sort()).toEqual(['Ink','Ink','Link','Square','Text']);const texts=list.asArray().map((_,i)=>list.lookup(i,PDFDict)).filter(dict=>dict.get(PDFName.of('Subtype'))===PDFName.of('Text'));expect(texts).toHaveLength(1);expect((texts[0].lookup(PDFName.of('Contents')) as PDFHexString).decodeText()).toBe('协作修改后的附注');
  annotations.get(0).set('content','Edited note');const flat=await PDFDocument.load(await exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source),true)),flatList=flat.getPage(0).node.lookup(PDFName.of('Annots'),PDFArray);
  expect(flatList.size()).toBe(1);expect((flatList.lookup(0,PDFDict).get(PDFName.of('Subtype')) as PDFName).decodeText()).toBe('Link');
  const resources=flat.getPage(0).node.Resources()!.lookup(PDFName.of('XObject'),PDFDict);expect(resources.keys().some(key=>key.decodeText().startsWith('MyDocAnnot'))).toBe(true);
  for(const bad of [(m:Y.Map<unknown>)=>m.set('rp-1','x'),(m:Y.Map<unknown>)=>m.set('$imported',2),(m:Y.Map<unknown>)=>m.set('r1-0-R','bad id!')]){const copy=new Y.Doc();try{Y.applyUpdate(copy,Y.encodeStateAsUpdate(doc));bad(copy.getMap('pdfSource'));expect(()=>validatePdfUpdate(null,Y.encodeStateAsUpdate(copy))).toThrow('导入记录');}finally{copy.destroy();}}
  doc.destroy();
 }finally{cloud.close();rmSync(dir,{recursive:true,force:true});}
 const missing=await PDFDocument.create(),p=missing.addPage([200,200]);p.node.addAnnot(missing.context.register(missing.context.obj({Type:'Annot',Subtype:'Stamp',Rect:[10,10,50,50]})));
 const empty=new Y.Doc();empty.getArray('annotations');try{await expect(exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(empty),await missing.save()),true)).rejects.toThrow('缺少外观');}finally{empty.destroy();}
},20000);

it('批注讨论导出为 PDF 原生附注：挂在所讨论的批注旁、回复以 IRT 成线程、已解决写入审阅状态，扁平化副本不含讨论',async()=>{
 const source=await original(),doc=new Y.Doc();Y.applyUpdate(doc,state());const annotation=doc.getArray<Y.Map<unknown>>('annotations').get(1),comments=doc.getMap('officeComments');
 const message=(id:string,name:string,body:string,created:number)=>({id,author:'u'+id,authorName:name,body,created});
 comments.set('a',{id:'a',resolved:true,updated:1700000003000,anchor:{kind:'pdf',annotation:pdfAnnotationId(annotation),page:1},messages:[message('m1','王小明','数据来源？',1700000001000),message('m2','Lisa Chen','见附录 A',1700000002000)]});
 comments.set('b',{id:'b',resolved:false,updated:1700000004000,messages:[message('m3','王小明','整份意见',1700000004000)]});
 try{
  const read=async(bytes:Uint8Array)=>{const pdf=await PDFDocument.load(bytes),annots=pdf.getPage(0).node.Annots()!;return Array.from({length:annots.size()},(_,i)=>annots.lookup(i,PDFDict)).filter(dict=>String(dict.get(PDFName.of('NM'))?.toString()??'').length&&(dict.lookup(PDFName.of('NM')) as PDFHexString).decodeText().startsWith('mydoc-discussion'));};
  const notes=await read(await exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source)));expect(notes).toHaveLength(4);
  const text=(dict:PDFDict,key:string)=>(dict.lookup(PDFName.of(key)) as PDFHexString|undefined)?.decodeText();
  expect(notes.map(note=>[text(note,'T'),text(note,'Contents')])).toEqual([['王小明','数据来源？'],['Lisa Chen','见附录 A'],['王小明','已解决'],['王小明','整份意见']]);
  const [root,reply,state0,floating]=notes;expect(reply.get(PDFName.of('IRT'))).toBeDefined();expect(root.get(PDFName.of('IRT'))).toBeUndefined();expect(floating.get(PDFName.of('IRT'))).toBeUndefined();expect(state0.get(PDFName.of('IRT'))?.toString()).toBe(reply.get(PDFName.of('IRT'))?.toString());
  expect(text(state0,'State')).toBe('Accepted');expect(state0.get(PDFName.of('F'))?.toString()).toBe('2');expect(root.get(PDFName.of('M'))?.toString()).toContain('D:20231114221321Z');
  // Next to the highlight it discusses; the unanchored note sits at the top-left of the crop box.
  const rect=(dict:PDFDict)=>(dict.lookup(PDFName.of('Rect')) as PDFArray).asArray().map(value=>Number(value.toString()));expect(rect(root).slice(0,2)).toEqual([40,80]);expect(rect(floating).slice(0,2)).toEqual([10,460]);
  // Flattened copies keep only page content, so discussions (interactive notes) are not written.
  doc.getArray<Y.Map<unknown>>('annotations').get(0).set('content','ascii note');const flat=await PDFDocument.load(await exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source),true));expect(flat.getPage(0).node.Annots()).toBeUndefined();
 }finally{doc.destroy();}
});
