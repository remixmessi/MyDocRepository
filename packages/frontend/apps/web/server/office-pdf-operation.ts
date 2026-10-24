import * as Y from 'yjs';
import {createHash} from 'node:crypto';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {pdfAnnotationId,pdfAnnotationRevision,validPdfAnnotation,validPdfPages,validPdfWatermark,pdfPagesRevision,pdfWatermarkRevision,type PdfOperation} from '../../../../common/nbstore/src/office-pdf';
import {validatePdfUpdate} from './office-pdf';
export class PdfOperationError extends Error{constructor(message:string,readonly status=400){super(message);}}
export function submitPdfOperation(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,input:PdfOperation){
 if(input?.type==='pages'||input?.type==='watermark')return submitPdfSetting(database,workspace,id,user,input);
 if(!input||!['create','update','delete'].includes(input.type)||typeof input.operationId!=='string'||! /^[a-zA-Z0-9-]{16,80}$/.test(input.operationId)||typeof input.annotationId!=='string'||! /^[a-zA-Z0-9-]{1,100}$/.test(input.annotationId)||typeof input.base!=='string'||input.base.length>131072||input.type!=='delete'&&!validPdfAnnotation(input.annotation,true)||input.type==='delete'&&input.annotation!==undefined)throw new PdfOperationError('PDF 批注操作无效');
 return database.db.transaction(()=>{
  const current=database.load('cloud:'+workspace,id);if(!current)throw new PdfOperationError('PDF 不存在',404);const doc=new Y.Doc();try{
   Y.applyUpdate(doc,current);const annotations=doc.getArray<Y.Map<unknown>>('annotations'),receipts=doc.getMap<string>('pdfOperations'),key=user+':'+input.operationId,signature=createHash('sha256').update(JSON.stringify(input)).digest('hex');
   if(receipts.has(key)){if(receipts.get(key)!==signature)throw new PdfOperationError('批注操作编号已用于其他内容',409);return {update:current,repeated:true};}if(receipts.size>=10000)throw new PdfOperationError('批注操作回执达到上限，请备份为新文件后继续');
   const index=annotations.toArray().findIndex(a=>pdfAnnotationId(a)===input.annotationId),target=index>=0?annotations.get(index):undefined;
   if(input.type==='create'){if(target||input.base!=='')throw new PdfOperationError('批注已存在',409);}else if(!target||pdfAnnotationRevision(target)!==input.base)throw new PdfOperationError('批注已被修改或删除，草稿保留；请载入最新内容后重试',409);
   const vector=Y.encodeStateVector(doc);
   if(input.type==='delete')annotations.delete(index,1);else{
    const next=target??new Y.Map<unknown>();for(const key of target?[...next.keys()]:[])if(!['id','author'].includes(key))next.delete(key);for(const [key,value]of Object.entries(input.annotation!))next.set(key,value);next.set('id',input.annotationId);if(!target){next.set('author',user);annotations.push([next]);}
   }
   receipts.set(key,signature);const update=Y.encodeStateAsUpdate(doc,vector);validatePdfUpdate(null,Y.encodeStateAsUpdate(doc));
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(doc)),'cloud:'+workspace,id);database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);return {update,repeated:false};
  }finally{doc.destroy();}
 }).immediate();
}

/** One-time import of the original's native Text/Highlight/Ink annotations; the server supplies the parsed list. */
export function importNativePdfAnnotations(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,operationId:string,natives:import('./office-pdf-native').PdfNativeAnnotation[]){
 if(typeof operationId!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(operationId))throw new PdfOperationError('PDF 批注操作无效');
 return database.db.transaction(()=>{
  const current=database.load('cloud:'+workspace,id);if(!current)throw new PdfOperationError('PDF 不存在',404);const doc=new Y.Doc();try{
   Y.applyUpdate(doc,current);const receipts=doc.getMap<string>('pdfOperations'),key=user+':'+operationId,signature=createHash('sha256').update('import-native').digest('hex'),source=doc.getMap<string|number>('pdfSource');
   if(receipts.has(key)){if(receipts.get(key)!==signature)throw new PdfOperationError('批注操作编号已用于其他内容',409);return {update:current,repeated:true,imported:0};}
   if(source.has('$imported'))throw new PdfOperationError('原件批注已导入过，请刷新查看',409);if(receipts.size>=10000)throw new PdfOperationError('批注操作回执达到上限，请备份为新文件后继续');
   const annotations=doc.getArray<Y.Map<unknown>>('annotations'),vector=Y.encodeStateVector(doc),importable=natives.filter(native=>native.annotation);
   source.set('$imported',1);for(const native of importable){const value=new Y.Map<unknown>(),annotationId='native-'+createHash('sha256').update(id+':'+native.key).digest('hex').slice(0,32);for(const [field,data]of Object.entries(native.annotation!))value.set(field,data);value.set('id',annotationId);value.set('author',native.author??user);annotations.push([value]);source.set(native.key,annotationId);}
   receipts.set(key,signature);const state=Y.encodeStateAsUpdate(doc);try{validatePdfUpdate(null,state);}catch(error){throw new PdfOperationError(error instanceof Error?error.message:'原件批注无法导入');}
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(state),'cloud:'+workspace,id);database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);return {update:Y.encodeStateAsUpdate(doc,vector),repeated:false,imported:importable.length};
  }finally{doc.destroy();}
 }).immediate();
}

/**
 * Page arrangement and watermark: the whole new value with the stored JSON as base (compare-and-set), so two members
 * reordering or deleting pages at once cannot interleave into an arrangement neither of them saw.
 */
function submitPdfSetting(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,input:Extract<PdfOperation,{type:'pages'|'watermark'}>){
 const pages=input.type==='pages';
 if(typeof input.operationId!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(input.operationId)||typeof input.base!=='string'||input.base.length>1024*1024||Object.keys(input).some(key=>!['operationId','type','base',pages?'pages':'watermark','officeGeneration'].includes(key))||(pages?!validPdfPages(input.pages):input.watermark!==null&&!validPdfWatermark(input.watermark)))throw new PdfOperationError(pages?'PDF 页面操作无效':'PDF 水印设置无效');
 return database.db.transaction(()=>{
  const current=database.load('cloud:'+workspace,id);if(!current)throw new PdfOperationError('PDF 不存在',404);const doc=new Y.Doc();try{
   Y.applyUpdate(doc,current);const receipts=doc.getMap<string>('pdfOperations'),key=user+':'+input.operationId,signature=createHash('sha256').update(JSON.stringify({...input,officeGeneration:undefined})).digest('hex');
   if(receipts.has(key)){if(receipts.get(key)!==signature)throw new PdfOperationError('操作编号已用于其他内容',409);return {update:current,repeated:true};}if(receipts.size>=10000)throw new PdfOperationError('操作回执达到上限，请备份为新文件后继续');
   if((pages?pdfPagesRevision(doc):pdfWatermarkRevision(doc))!==input.base)throw new PdfOperationError(pages?'页面已被其他成员调整，请核对后重试':'水印已被其他成员修改，请核对后重试',409);
   const vector=Y.encodeStateVector(doc),map=doc.getMap('pdfPages');
   if(pages)map.set('arrangement',input.pages);else if(input.watermark)map.set('watermark',input.watermark);else map.delete('watermark');
   receipts.set(key,signature);const update=Y.encodeStateAsUpdate(doc,vector);try{validatePdfUpdate(null,Y.encodeStateAsUpdate(doc));}catch(error){throw new PdfOperationError(error instanceof Error?error.message:'PDF 页面设置无效');}
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(Y.encodeStateAsUpdate(doc)),'cloud:'+workspace,id);database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);return {update,repeated:false};
  }finally{doc.destroy();}
 }).immediate();
}
