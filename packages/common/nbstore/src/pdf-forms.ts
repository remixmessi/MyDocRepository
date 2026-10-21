import type * as Y from 'yjs';
export interface PdfFormSpec {id:string;name:string;type:'text'|'checkbox';value:string|boolean;readOnly:boolean;maxLength:number;multiline:boolean;}
export interface PdfFormValue {type:'text'|'checkbox';value:string|boolean;revision:string;}
export interface PdfFormOperation {operationId:string;fieldId:string;base:string;value:string|boolean;officeGeneration?:string;}
export function pdfFormRevision(spec:PdfFormSpec,value?:PdfFormValue){return value?.revision??JSON.stringify([spec.type,spec.value]);}
export function validatePdfFields(doc:Y.Doc){
 if(!doc.share.has('pdfFields'))return;const fields=doc.getMap('pdfFields');let size=0;if(fields._start||fields.size>1000)throw new Error('PDF 表单字段数量或结构无效');
 for(const [id,value]of fields){if(!/^[a-f0-9]{64}$/.test(id)||!value||typeof value!=='object'||Array.isArray(value)||Object.keys(value).sort().join(',')!=='revision,type,value')throw new Error('PDF 表单字段无效');const field=value as PdfFormValue;if(!['text','checkbox'].includes(field.type)||typeof field.revision!=='string'||! /^[a-zA-Z0-9:-]{1,200}$/.test(field.revision)||(field.type==='text'?typeof field.value!=='string'||field.value.length>10000:typeof field.value!=='boolean'))throw new Error('PDF 表单值无效');size+=JSON.stringify(value).length;if(size>2*1024*1024)throw new Error('PDF 表单值超过 2 MB');}
}
