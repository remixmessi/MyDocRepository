import {PDFDocument,PDFTextField,PDFCheckBox,PDFName,PDFDict} from 'pdf-lib';
import {PdfFonts} from './office-pdf-font';
import {createHash} from 'node:crypto';
import * as Y from 'yjs';
import type {SQLiteProjectDatabase} from '../../../../common/nbstore/server/database';
import {pdfFormRevision,validatePdfFields,type PdfFormSpec,type PdfFormOperation,type PdfFormValue} from '../../../../common/nbstore/src/pdf-forms';
import {validatePdfUpdate} from './office-pdf';
export const pdfFieldId=(name:string)=>createHash('sha256').update(name).digest('hex');
export async function inspectPdfForms(bytes:Uint8Array){
 const pdf=await PDFDocument.load(bytes,{updateMetadata:false,throwOnInvalidObject:true});if(pdf.isEncrypted)throw new Error('加密 PDF 不支持表单编辑');if(pdf.catalog.lookupMaybe(PDFName.of('AcroForm'),PDFDict)?.has(PDFName.of('XFA')))throw new Error('XFA 表单暂不支持编辑');const form=pdf.getForm();if(form.hasXFA())throw new Error('XFA 表单暂不支持编辑');const fields=form.getFields();if(fields.length>1000)throw new Error('PDF 表单超过 1000 个字段');
 const result:PdfFormSpec[]=[],names=new Set<string>();let size=0;
 for(const field of fields){const name=field.getName();if(field.acroField.dict.has(PDFName.of('AA')))throw new Error('含自动计算或动作的 PDF 表单暂不支持协作编辑');if(!name||name.length>1000||names.has(name))throw new Error('PDF 表单字段名称无效或重复');names.add(name);if(!(field instanceof PDFTextField||field instanceof PDFCheckBox))continue;
  const entry:PdfFormSpec={id:pdfFieldId(name),name,type:field instanceof PDFTextField?'text':'checkbox',value:field instanceof PDFTextField?field.getText()??'':field.isChecked(),readOnly:field.isReadOnly(),maxLength:field instanceof PDFTextField?Math.min(10000,field.getMaxLength()??10000):0,multiline:field instanceof PDFTextField&&field.isMultiline()};if(typeof entry.value==='string'&&entry.value.length>10000)throw new Error('PDF 表单原始值超过 10000 字符');result.push(entry);size+=JSON.stringify(entry).length;if(size>2*1024*1024)throw new Error('PDF 表单定义超过 2 MB');
 }
 return new TextEncoder().encode(JSON.stringify(result));
}
export class PdfFormError extends Error{constructor(message:string,readonly status=400){super(message);}}
export function submitPdfForm(database:SQLiteProjectDatabase,workspace:string,id:string,user:string,input:PdfFormOperation,specs:PdfFormSpec[]){
 if(!input||Object.keys(input).some(key=>!['operationId','fieldId','base','value'].includes(key))||! /^[a-zA-Z0-9-]{16,80}$/.test(input.operationId??'')||! /^[a-f0-9]{64}$/.test(input.fieldId??'')||typeof input.base!=='string'||input.base.length>20000)throw new PdfFormError('PDF 表单操作无效');const spec=specs.find(field=>field.id===input.fieldId);if(!spec||spec.readOnly)throw new PdfFormError('字段不存在、不支持或为只读',403);if(spec.type==='text'?typeof input.value!=='string'||input.value.length>spec.maxLength:typeof input.value!=='boolean')throw new PdfFormError('PDF 表单值类型或长度无效');
 return database.db.transaction(()=>{
  const current=database.load('cloud:'+workspace,id);if(!current)throw new PdfFormError('PDF 不存在',404);const doc=new Y.Doc();try{
   Y.applyUpdate(doc,current);const fields=doc.getMap<PdfFormValue>('pdfFields'),receipts=doc.getMap<string>('pdfOperations'),key=user+':'+input.operationId,signature=createHash('sha256').update(JSON.stringify([input.fieldId,input.base,input.value])).digest('hex');
   if(receipts.has(key)){if(receipts.get(key)!==signature)throw new PdfFormError('表单操作编号已用于其他内容',409);return {update:current,repeated:true};}if(receipts.size>=10000)throw new PdfFormError('PDF 操作记录已满，请备份恢复为新文件',409);
   if(pdfFormRevision(spec,fields.get(spec.id))!==input.base)throw new PdfFormError('字段已被修改，草稿保留；请核对最新值后重新提交',409);
   const vector=Y.encodeStateVector(doc);fields.set(spec.id,{type:spec.type,value:input.value,revision:user+':'+input.operationId});receipts.set(key,signature);const state=Y.encodeStateAsUpdate(doc);validatePdfUpdate(null,state);
   database.db.prepare('UPDATE documents SET data=? WHERE scope=? AND id=?').run(Buffer.from(state),'cloud:'+workspace,id);database.db.prepare('UPDATE cloud_office_resources SET updated=? WHERE workspace=? AND id=?').run(Date.now(),workspace,id);return {update:Y.encodeStateAsUpdate(doc,vector),repeated:false};
  }finally{doc.destroy();}
 }).immediate();
}
/** Values and appearances are separate: each text appearance uses Helvetica when it can, otherwise the Unicode font. */
export async function applyPdfFormValues(pdf:PDFDocument,doc:Y.Doc,fonts=new PdfFonts(pdf,undefined)){
 validatePdfFields(doc);if(!doc.share.has('pdfFields'))return;
 if(pdf.catalog.lookupMaybe(PDFName.of('AcroForm'),PDFDict)?.has(PDFName.of('XFA')))throw new Error('XFA 表单不支持导出修改');const form=pdf.getForm();if(form.hasXFA())throw new Error('XFA 表单不支持导出修改');const fields=form.getFields();
 for(const [id,raw]of doc.getMap<PdfFormValue>('pdfFields')){const field=fields.find(field=>pdfFieldId(field.getName())===id);if(!field||field.isReadOnly())throw new Error('PDF 原件中的表单字段缺失或只读');
  if(raw.type==='text'&&field instanceof PDFTextField){if(field.getMaxLength()!==undefined&&String(raw.value).length>field.getMaxLength()!)throw new Error('表单内容超过原件长度限制');field.setText(String(raw.value));field.updateAppearances(await fonts.for(String(raw.value),`表单字段“${field.getName()}”`));}
  else if(raw.type==='checkbox'&&field instanceof PDFCheckBox){if(raw.value)field.check();else field.uncheck();field.updateAppearances();}else throw new Error('PDF 表单字段类型不匹配');
 }
}
