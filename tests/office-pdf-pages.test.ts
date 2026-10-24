import {it,expect} from 'vitest';
import * as Y from 'yjs';
import {PDFDocument,PDFName,PDFArray} from 'pdf-lib';
import {mkdtempSync,rmSync,readFileSync} from 'node:fs';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {exportAnnotatedPdf,packPdfExport} from '../packages/frontend/apps/web/server/office-pdf-export';
import {submitPdfOperation} from '../packages/frontend/apps/web/server/office-pdf-operation';
import {validatePdfUpdate,emptyPdfUpdate} from '../packages/frontend/apps/web/server/office-pdf';
import {validPdfPages,validPdfWatermark,validPdfAnnotation,parsePdfPageRange,pdfIdentityPages,pdfPagesRevision,pdfWatermarkRevision,pdfInkPaths,type PdfPageEntry} from '../packages/common/nbstore/src/office-pdf';
const font=readFileSync(new URL('../packages/frontend/apps/web/server/fonts/NotoSansSC-Regular.ttf',import.meta.url));
async function threePages(){const pdf=await PDFDocument.create();for(const [i,size]of [[400,500],[500,400],[300,300]].entries()){const page=pdf.addPage(size as [number,number]);page.drawText('Page '+(i+1),{x:20,y:20});}const form=pdf.getForm(),field=form.createTextField('name');field.setText('A');field.addToPage(pdf.getPage(1),{x:20,y:60,width:120,height:20});return pdf.save();}
const pageTexts=async(bytes:Uint8Array)=>{const pdf=await PDFDocument.load(bytes);return pdf.getPages().map(page=>({width:Math.round(page.getWidth()),height:Math.round(page.getHeight()),rotate:page.getRotation().angle}));};

it('页面模型校验：编号唯一、旋转为 90 的倍数、空白页尺寸、水印字段；页码范围解析；签名多笔画',()=>{
 expect(validPdfPages(pdfIdentityPages(3))).toBe(true);expect(validPdfPages([{id:'a',blank:{width:612,height:792},rotate:90}])).toBe(true);
 for(const bad of [[],[{id:'a',source:1},{id:'a',source:2}],[{id:'a',source:0}],[{id:'a',source:1,rotate:45}],[{id:'a',blank:{width:10,height:792}}],[{id:'a',source:1,extra:1}],[{id:'bad id',source:1}]])expect(validPdfPages(bad)).toBe(false);
 expect(validPdfWatermark({text:'机密',color:'#FF0000',opacity:.3,size:60,rotation:45})).toBe(true);for(const bad of [{text:'',color:'#FF0000',opacity:.3,size:60,rotation:45},{text:'机密',color:'red',opacity:.3,size:60,rotation:45},{text:'机密',color:'#FF0000',opacity:2,size:60,rotation:45},{text:'机密',color:'#FF0000',opacity:.3,size:60,rotation:30}])expect(validPdfWatermark(bad)).toBe(false);
 expect(parsePdfPageRange('1-3, 5，2',6)).toEqual([1,2,3,5]);expect(parsePdfPageRange('7',6)).toBeUndefined();expect(parsePdfPageRange('a',6)).toBeUndefined();
 const signature={page:1,type:'ink' as const,x:0,y:0,width:10,height:10,content:'签名',points:[0,0,1,1,5,5,6,6,7,7],strokes:[2,3]};expect(validPdfAnnotation(signature,true)).toBe(true);expect(pdfInkPaths(signature)).toEqual([[0,0,1,1],[5,5,6,6,7,7]]);expect(validPdfAnnotation({...signature,strokes:[2,2]},true)).toBe(false);expect(validPdfAnnotation({...signature,type:'comment'},true)).toBe(false);
});

it('页面与水印语义操作：整体比较替换（并发调整返回 409）、幂等重试、通用更新与非法值拒绝',()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-pdf-pages-')),cloud=new CloudServer(join(dir,'cloud.sqlite')),{doc,update}=emptyPdfUpdate();try{
  cloud.database.append('cloud:space','pdf',Buffer.from(update));const call=(input:Parameters<typeof submitPdfOperation>[4],user='actor')=>{const result=submitPdfOperation(cloud.database,'space','pdf',user,input);Y.applyUpdate(doc,result.update);return result;};
  const first:PdfPageEntry[]=[{id:'p3',source:3},{id:'p1',source:1,rotate:90}],op={operationId:crypto.randomUUID(),type:'pages' as const,base:'',pages:first};call(op);expect(pdfPagesRevision(doc)).toBe(JSON.stringify(first));expect(call(op).repeated).toBe(true);
  expect(()=>call({operationId:crypto.randomUUID(),type:'pages',base:'',pages:pdfIdentityPages(3)},'other')).toThrow('其他成员调整');
  call({operationId:crypto.randomUUID(),type:'pages',base:JSON.stringify(first),pages:[...first,{id:'b1',blank:{width:400,height:500}}]},'other');
  expect(()=>call({operationId:crypto.randomUUID(),type:'pages',base:pdfPagesRevision(doc),pages:[{id:'x',source:1},{id:'x',source:2}]})).toThrow('页面操作无效');
  const mark={text:'内部资料',color:'#C00000',opacity:.25,size:48,rotation:45};call({operationId:crypto.randomUUID(),type:'watermark',base:'',watermark:mark});expect(pdfWatermarkRevision(doc)).toBe(JSON.stringify(mark));
  expect(()=>call({operationId:crypto.randomUUID(),type:'watermark',base:'',watermark:null})).toThrow('其他成员修改');call({operationId:crypto.randomUUID(),type:'watermark',base:JSON.stringify(mark),watermark:null});expect(pdfWatermarkRevision(doc)).toBe('');
  const forged=new Y.Doc();Y.applyUpdate(forged,Y.encodeStateAsUpdate(doc));forged.getMap('pdfPages').set('arrangement',[{id:'a',source:-1}]);expect(()=>validatePdfUpdate(null,Y.encodeStateAsUpdate(forged))).toThrow('页面设置无效');forged.destroy();
 }finally{doc.destroy();cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('导出按页面排列：重排、删除、旋转与空白页，删去含表单的页时扁平化表单，批注随页、讨论改挂首页，水印写入每页，提取指定页',async()=>{
 const source=await threePages(),doc=new Y.Doc();try{
  doc.getMap('pdfOperations').set('$schema','0'.repeat(64));const add=(page:number,content:string,extra:object={})=>{const value=new Y.Map<unknown>();for(const [key,data]of Object.entries({id:'a'+page,page,type:'comment',x:30,y:30,content,...extra}))value.set(key,data);doc.getArray<Y.Map<unknown>>('annotations').push([value]);};
  add(1,'第一页附注');add(2,'被删页附注');add(3,'签名',{type:'ink',width:40,height:20,points:[30,30,40,40,50,30,60,40,70,30],strokes:[2,3]});
  doc.getMap('officeComments').set('t1',{id:'t1',resolved:false,updated:1,anchor:{kind:'pdf',annotation:'a2',page:2},messages:[{id:'m1',author:'u',authorName:'甲',body:'讨论被删页',created:1}]});
  doc.getMap('pdfPages').set('arrangement',[{id:'p3',source:3,rotate:90},{id:'b',blank:{width:200,height:300}},{id:'p1',source:1}]);
  doc.getMap('pdfPages').set('watermark',{text:'机密',color:'#FF0000',opacity:.3,size:40,rotation:45});
  const state=Y.encodeStateAsUpdate(doc);validatePdfUpdate(null,state);const bytes=await exportAnnotatedPdf(packPdfExport(state,source),false,font);if(process.env.MYDOC_PDF_CHECK)(await import('node:fs')).writeFileSync(process.env.MYDOC_PDF_CHECK,bytes);
  expect(await pageTexts(bytes)).toEqual([{width:300,height:300,rotate:90},{width:200,height:300,rotate:0},{width:400,height:500,rotate:0}]);
  const pdf=await PDFDocument.load(bytes),annots=pdf.getPages().map(page=>(page.node.Annots()?.asArray()??[]).map(ref=>{const dict=pdf.context.lookup(ref) as import('pdf-lib').PDFDict;return [String(dict.get(PDFName.of('Subtype'))),dict.lookup(PDFName.of('InkList')) instanceof PDFArray?(dict.lookup(PDFName.of('InkList')) as PDFArray).size():0];}));
  // The discussion on a deleted page moves to the new first page (original page 3), beside the two-stroke signature.
  expect(annots[0]).toEqual([['/Ink',2],['/Text',0]]);expect(annots[1]).toEqual([]);expect(annots[2].map(a=>a[0])).toEqual(['/Text']);
  // The form's page was deleted, so the form is flattened rather than left with a dangling widget.
  expect(pdf.getForm().getFields()).toHaveLength(0);
  const contents=pdf.getPages().map(page=>page.node.Contents());expect(contents.every(Boolean)).toBe(true);
  // 提取 positions 3 and 1 (in that order of the arrangement).
  const extract=await exportAnnotatedPdf(packPdfExport(state,source),false,font,[1,3]);expect(await pageTexts(extract)).toEqual([{width:300,height:300,rotate:90},{width:400,height:500,rotate:0}]);
  await expect(exportAnnotatedPdf(packPdfExport(state,source),false,font,[9])).rejects.toThrow('超出范围');
  // Without an arrangement the export is unchanged in page count, and the watermark still applies.
  doc.getMap('pdfPages').delete('arrangement');const plain=await exportAnnotatedPdf(packPdfExport(Y.encodeStateAsUpdate(doc),source),true,font);expect((await pageTexts(plain)).length).toBe(3);
 }finally{doc.destroy();}
});
